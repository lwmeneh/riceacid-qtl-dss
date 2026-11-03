"""RiceAcid-QTL DSS: evidence-traceable, cascading research browser.
The uploaded workbook is currently an aluminium-tolerance research database.
"""
from pathlib import Path
import io
import re
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go

st.set_page_config(page_title="RiceAcid-QTL DSS", page_icon="🌾", layout="wide")
ROOT = Path(__file__).resolve().parent
DEFAULT = ROOT / "data" / "Al_MetaQTL_Database_Input.xlsx"
ELEMENTS = {"Al": "Aluminium toxicity", "Cd": "Cadmium toxicity", "Ca": "Calcium nutrition", "Mg": "Magnesium nutrition", "Fe": "Iron stress/nutrition", "Mn": "Manganese stress/nutrition", "Zn": "Zinc nutrition/toxicity"}

@st.cache_data(show_spinner="Reading workbook...")
def load_book(raw):
    xls = pd.ExcelFile(io.BytesIO(raw))
    return {s: pd.read_excel(xls, sheet_name=s).dropna(how="all") for s in xls.sheet_names}

def table(name):
    return db.get(name, pd.DataFrame()).copy()

def rows(name, column=None, values=None):
    d = table(name)
    if column is None or values is None or column not in d.columns:
        return d if column is None else d.iloc[0:0]
    values = {str(v).strip() for v in values if pd.notna(v)}
    return d[d[column].astype("string").str.strip().isin(values).fillna(False)]

def ids(d, col):
    if col not in d: return set()
    return set(d[col].dropna().astype(str).str.strip()) - {"", "nan", "None"}

def select(label, values, key, default=None):
    opts = sorted({str(v) for v in values if pd.notna(v) and str(v).strip() and str(v) != 'nan'})
    if not opts:
        st.info(f"No available options for {label.lower()}.")
        return None
    return st.selectbox(label, opts, index=opts.index(default) if default in opts else 0, key=key)

def display(d, title, key, note=None):
    st.subheader(f"{title} ({len(d):,})")
    if note: st.caption(note)
    if d.empty:
        st.info("No linked records in this workbook. This does not establish biological absence.")
        return
    st.dataframe(d, hide_index=True, use_container_width=True)
    st.download_button("Download this table (CSV)", d.to_csv(index=False).encode("utf-8-sig"), file_name=f"{key}.csv", mime="text/csv", key=f"dl_{key}")

def source_panel(frames, key):
    src_ids = set()
    for d in frames:
        src_ids |= ids(d, 'source_id')
    s = rows('Sources', 'source_id', src_ids)
    display(s, 'Original source records', key+'_sources', 'Source URLs and file references are reproduced from the uploaded workbook; verify access and provenance.')
    if not s.empty:
        for _, r in s.iterrows():
            url = str(r.get('file_url', '')).strip()
            if url.startswith(('https://','http://')):
                st.markdown(f"Source **{r.get('source_id','')}**: [Open recorded source]({url})")
            doi = str(r.get('study_doi','')).strip()
            if doi and doi.lower() != 'nan':
                st.markdown(f"DOI: [ {doi} ](https://doi.org/{doi})")

def show_interval(m):
    start = pd.to_numeric(pd.Series([m.get('left_anchor_bp')]), errors='coerce').iloc[0]
    end = pd.to_numeric(pd.Series([m.get('right_anchor_bp')]), errors='coerce').iloc[0]
    if pd.isna(start) or pd.isna(end) or end < start:
        st.warning('Valid physical interval unavailable for this MQTL.')
        return
    fig = go.Figure()
    fig.add_trace(go.Scatter(x=[start/1e6,end/1e6], y=[1,1], mode='lines+markers', line=dict(width=12,color='#17816b'), marker=dict(size=10), name='MQTL interval'))
    fig.update_layout(height=180, title=f"MQTL physical interval · chromosome {m.get('chromosome')} · {m.get('genome_assembly', 'assembly not specified')}", xaxis_title='Physical position (Mb)', yaxis=dict(visible=False,range=[0.5,1.5]), margin=dict(l=15,r=15,t=60,b=30))
    st.plotly_chart(fig, use_container_width=True)

st.sidebar.title('🌾 RiceAcid-QTL DSS')
up = st.sidebar.file_uploader('Upload research workbook (.xlsx)', type=['xlsx'])
if up is not None: raw = up.getvalue()
elif DEFAULT.exists(): raw = DEFAULT.read_bytes()
else:
    st.title('Rice Acid-Soil QTL & Meta-QTL Decision Support System')
    st.info('Upload Al_MetaQTL_Database_Input.xlsx using the left sidebar to begin.')
    st.stop()
try:
    db = load_book(raw)
except Exception as exc:
    st.error(f'Workbook could not be opened: {exc}')
    st.stop()

st.title('Rice Acid-Soil QTL & Meta-QTL Decision Support System')
st.caption('Evidence-traceable research explorer • original data retained • not a de novo meta-analysis or experimental validation engine')
element = st.selectbox('1 · Select soil element / stress', list(ELEMENTS), format_func=lambda e: f'{e} — {ELEMENTS[e]}')
if element != 'Al':
    st.warning(f'{ELEMENTS[element]} is reserved for future datasets. The uploaded workbook contains aluminium-tolerance evidence only; no cross-element assignment has been made.')
    st.stop()

mq = table('MQTLs')
if mq.empty or 'mqtl_id' not in mq:
    st.error('The workbook has no MQTLs sheet with mqtl_id.')
    st.stop()
chromosomes = sorted(mq['chromosome'].dropna().unique().tolist()) if 'chromosome' in mq else []
chrom = st.selectbox('2 · Filter by rice chromosome', ['All'] + chromosomes, format_func=lambda v: 'All chromosomes' if v=='All' else f'Chromosome {v}')
if chrom != 'All': mq = mq[mq.chromosome.eq(chrom)]
selected = select('3 · Select Meta-QTL (MQTL)', ids(mq,'mqtl_id'), 'mqtl_select')
if not selected: st.stop()
m = mq.loc[mq.mqtl_id.astype(str).eq(selected)].iloc[0]

members = rows('MQTL_QTLs', 'mqtl_id', {selected})
qtl_ids = ids(members, 'qtl_id')
projected_ids = ids(members, 'projected_qtl_id')
original = rows('QTLs', 'qtl_id', qtl_ids)
projected = rows('Projected_QTLs', 'projected_qtl_id', projected_ids)
links = rows('MQTL_Genes', 'mqtl_id', {selected})
linked_genes = rows('Genes', 'gene_id', ids(links,'gene_id'))

c1,c2,c3,c4=st.columns(4)
c1.metric('Selected MQTL', selected)
c2.metric('Linked original QTLs',len(original))
c3.metric('Linked projected QTLs',len(projected))
c4.metric('Linked genes',len(linked_genes))
show_interval(m)

sections = st.tabs(['MQTL & QTLs','Genes & traits','Expression','Protein & enrichment','qPCR & primers','Markers','Source & provenance','Database inventory'])
with sections[0]:
    display(pd.DataFrame([m]), 'Selected meta-QTL', 'mqtl')
    display(members, 'MQTL–QTL relationship records', 'mqtl_members')
    display(original, 'Original QTL records', 'original_qtls', 'Only explicitly linked qtl_id records are shown.')
    display(projected, 'Projected QTL records', 'projected_qtls', 'Only explicitly linked projected_qtl_id records are shown.')
with sections[1]:
    display(links, 'MQTL–gene associations', 'mqtl_gene_links')
    display(linked_genes, 'Candidate genes in selected MQTL', 'mqtl_genes')
    selected_gene = select('4 · Select candidate gene', ids(linked_genes,'gene_id'), 'gene_select')
    if selected_gene:
        st.session_state['last_gene'] = selected_gene
        display(rows('Genes','gene_id',{selected_gene}), 'Selected gene annotation', 'selected_gene')
    trait_codes = ids(original,'trait_code')
    display(rows('Traits','trait_code',trait_codes), 'Traits associated with linked original QTLs', 'mqtl_traits', 'Trait associations are derived through linked QTLs; they are not automatically gene-specific.')

# A gene selection shared by all evidence tabs, independent of which tab is open.
with st.sidebar:
    st.divider()
    gene = select('Candidate gene for evidence drill-down', ids(linked_genes,'gene_id'), 'global_gene')
    st.caption('Evidence tabs follow this gene selection. Other records remain visible at MQTL level.')

with sections[2]:
    if gene:
        expr = rows('Expression','gene_id',{gene})
        display(expr, 'Gene expression records', 'expression')
        display(rows('Expression_Datasets','expression_dataset_id',ids(expr,'expression_dataset_id')), 'Linked expression datasets', 'expression_datasets')
        st.caption('Expression metrics and scales are presented as supplied; do not compare across studies without harmonization.')
with sections[3]:
    if gene:
        protein = rows('Protein_Features','gene_id',{gene})
        promoter = rows('Promoter_Sites','gene_id',{gene})
        eg = rows('Enrichment_Genes','gene_id',{gene})
        enrichment = rows('Enrichment','enrichment_id',ids(eg,'enrichment_id'))
        display(protein,'Protein features','protein')
        display(promoter,'Promoter regulatory motifs','promoter')
        display(enrichment,'Enrichment terms linked through gene membership','enrichment')
        st.caption('Enrichment is gene-set-level evidence, not proof of an individual gene’s function.')
with sections[4]:
    if gene:
        genes_row = rows('Genes','gene_id',{gene})
        gene_names = {gene}
        for col in ('symbol','symbol_aliases'):
            if col in genes_row:
                for v in genes_row[col].dropna().astype(str):
                    gene_names.update(n.strip() for n in re.split(r'[,;/]',v) if n.strip())
        qsum = rows('qPCR_Summary','gene_id',{gene})
        # The workbook has qPCR targets that may not be mapped to gene IDs.
        if not qsum.empty:
            gene_names |= ids(qsum,'target_name')
        display(qsum,'qPCR summaries explicitly linked by gene ID','qpcr_summary')
        target_qpcr = rows('qPCR_Summary','target_name',gene_names)
        if not target_qpcr.empty:
            display(target_qpcr,'Potential target-name matches (verify identity)','qpcr_target_matches','Target-name matching is suggestive, not a verified gene mapping.')
        target_names = ids(qsum,'target_name') | ids(target_qpcr,'target_name')
        display(rows('qPCR_Values','target_name',target_names),'qPCR replicate values for matched targets','qpcr_values')
        display(rows('Primers','target_name',target_names),'Primer sequences for matched targets','primers')
        st.caption('Some qPCR records lack gene_id. Unmatched records are accessible in the inventory and should not be inferred to validate this gene.')
with sections[5]:
    marker_names = set()
    for d in (pd.DataFrame([m]),members):
        for col in ('left_marker','right_marker','left_marker_as_supplied','right_marker_as_supplied'):
            marker_names |= ids(d,col)
    marker_rows = rows('Markers','marker_name',marker_names)
    display(marker_rows,'Exact marker-name matches','markers','Marker names can recur across maps; study and chromosome must be checked before assigning a physical position.')
    st.caption('Marker flanks reported in the MQTL and membership tables:')
    display(members[[c for c in ['qtl_id','projected_qtl_id','left_marker_as_supplied','right_marker_as_supplied','source_id'] if c in members]],'MQTL membership marker flanks','marker_flanks')
with sections[6]:
    frames = [pd.DataFrame([m]),members,original,projected,links,linked_genes]
    if gene:
        expr=rows('Expression','gene_id',{gene})
        frames.extend([expr,rows('Expression_Datasets','expression_dataset_id',ids(expr,'expression_dataset_id')),rows('Protein_Features','gene_id',{gene}),rows('Promoter_Sites','gene_id',{gene}),rows('Enrichment_Genes','gene_id',{gene}),rows('qPCR_Summary','gene_id',{gene})])
    source_panel(frames,'trace')
    display(rows('Studies','study_id',ids(original,'study_id')),'Contributing studies','linked_studies')
    st.caption('The source chain may be incomplete where the original workbook lacks identifiers or mappings.')
with sections[7]:
    inventory=pd.DataFrame([{'sheet':name,'records':len(d),'columns':len(d.columns),'missing_cells':int(d.isna().sum().sum())} for name,d in db.items()])
    display(inventory,'Workbook inventory and missingness','inventory')
    sheet=select('Inspect any original workbook sheet',list(db),'inventory_sheet')
    if sheet: display(table(sheet),f'Unfiltered sheet: {sheet}','all_'+sheet)
    st.warning('This application explores source-reported associations; it does not perform statistical meta-QTL inference, confirm candidate-gene function, or harmonize genome assemblies.')
