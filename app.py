"""RiceAcid-QTL DSS | Research explorer for acid-soil rice genetics.
Run: streamlit run app.py
Place the workbook in data/Al_MetaQTL_Database_Input.xlsx.
"""
from pathlib import Path
import io
import re
import html
import pandas as pd
import numpy as np
import streamlit as st
import plotly.graph_objects as go
import plotly.express as px

st.set_page_config(page_title="RiceAcid | QTL Discovery", page_icon="🌾", layout="wide", initial_sidebar_state="expanded")
ROOT = Path(__file__).resolve().parent
DEFAULT = ROOT / "data" / "Al_MetaQTL_Database_Input.xlsx"
ELEMENTS = {"Al":"Aluminium toxicity", "Cd":"Cadmium toxicity", "Ca":"Calcium nutrition", "Mg":"Magnesium nutrition", "Fe":"Iron stress / nutrition", "Mn":"Manganese stress / nutrition", "Zn":"Zinc nutrition / toxicity"}

st.markdown("""<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&family=Manrope:wght@400;500;600;700;800&display=swap');
html,body,[class*="css"], [data-testid="stApp"]{font-family:'DM Sans',sans-serif}
[data-testid="stAppViewContainer"]{background:#f5f8f6;color:#17392e}
[data-testid="stSidebar"]{background:linear-gradient(175deg,#103e33 0%,#092b25 100%)}
[data-testid="stSidebar"] *{color:#e9f6ef!important}
[data-testid="stSidebar"] [data-baseweb="select"] *{color:#17392e!important}
[data-testid="stSidebar"] input{color:#17392e!important}
[data-testid="stSidebar"] [data-testid="stFileUploader"]{border:1px dashed #8cbaa9;border-radius:12px;padding:8px}
.block-container{max-width:1480px;padding-top:1.6rem;padding-bottom:4rem}
h1,h2,h3{font-family:'Manrope',sans-serif;letter-spacing:-.025em;color:#153d30}
h1{font-weight:800;font-size:2.15rem!important}
h2{font-size:1.35rem!important}h3{font-size:1.08rem!important}
.hero{background:linear-gradient(115deg,#104b3a,#14745a);color:#fff;border-radius:20px;padding:28px 32px;margin-bottom:20px;box-shadow:0 10px 30px #0a49331a}
.hero h1{color:white!important;margin:0;font-size:2.1rem!important}.hero p{color:#d2e9df;margin:10px 0 0;font-size:1rem}
.hero .eyebrow{font-size:.74rem;letter-spacing:.15em;text-transform:uppercase;color:#a6d9c2;font-weight:700;margin-bottom:9px}
.card{background:white;border:1px solid #e1ebe4;border-radius:15px;padding:19px 22px;box-shadow:0 3px 12px #113b2810;margin:6px 0 16px}
.card .k{color:#6b8577;font-size:.76rem;text-transform:uppercase;letter-spacing:.065em;font-weight:700}
.card .v{color:#134d39;font-size:1.72rem;font-weight:800;margin-top:6px}
.card .sub{color:#789185;font-size:.79rem;margin-top:4px}
.pill{display:inline-block;border-radius:99px;background:#e3f3e9;color:#116749;font-size:.77rem;padding:6px 12px;font-weight:700;margin:3px 4px 3px 0}
.soft-note{background:#eef7f0;border-left:4px solid #2a9b70;border-radius:8px;padding:12px 16px;color:#24523f;margin:10px 0 20px}
.step{font-size:.78rem;color:#698475;font-weight:700;text-transform:uppercase;letter-spacing:.07em;margin-bottom:5px}
[data-testid="stMetric"]{background:white;border:1px solid #e1ebe4;padding:14px 18px;border-radius:14px}
[data-testid="stMetricValue"]{color:#165e46}
[data-testid="stTabs"] button{font-weight:700}
[data-testid="stDataFrame"]{border:1px solid #e2ebe4;border-radius:10px;overflow:hidden}
.stButton>button[kind="primary"]{background:#146f53;border-color:#146f53;border-radius:10px}
@media(max-width:768px){.hero{padding:20px}.hero h1{font-size:1.45rem!important}.block-container{padding:1rem}}
</style>""",unsafe_allow_html=True)

@st.cache_data(show_spinner="Reading scientific workbook…",max_entries=3)
def read_workbook(content:bytes):
    x = pd.ExcelFile(io.BytesIO(content))
    return {name:pd.read_excel(x,sheet_name=name).dropna(how="all") for name in x.sheet_names}

def sheet(name):return DB.get(name,pd.DataFrame()).copy()
def tokens(frame,col):
    if frame.empty or col not in frame.columns:return set()
    return {str(x).strip() for x in frame[col].dropna() if str(x).strip() not in ('','nan','None','<NA>')}
def subset(name,col,values):
    frame=sheet(name)
    if frame.empty or col not in frame.columns or not values:return frame.iloc[:0]
    vals={str(x).strip() for x in values}
    return frame[frame[col].astype('string').str.strip().isin(vals).fillna(False)]
def pretty(x):
    if x is None:
        return '—'
    try:
        if bool(pd.isna(x)):
            return '—'
    except (TypeError, ValueError):
        pass
    return str(x)
def stat(label,value,sub=''):
    st.markdown(f'<div class="card"><div class="k">{html.escape(str(label))}</div><div class="v">{html.escape(str(value))}</div><div class="sub">{html.escape(str(sub))}</div></div>',unsafe_allow_html=True)
def grid(frame,label,key,desc=None):
    st.markdown(f'### {label} <span style="font-size:.78rem;color:#789185;font-weight:500">({len(frame):,})</span>',unsafe_allow_html=True)
    if desc:st.caption(desc)
    if frame.empty:
        st.info('No explicitly linked records are available for this selection. This does not imply biological absence.')
        return
    st.dataframe(frame,hide_index=True,use_container_width=True,height=min(420,max(120,42+len(frame)*34)))
    st.download_button('↓ Export CSV',frame.to_csv(index=False).encode('utf-8-sig'),file_name=f'{key}.csv',mime='text/csv',key='export_'+key)
def choice(label,options,key,empty='No linked records'):
    opts=sorted({str(v) for v in options if v is not None and str(v).strip() not in ('','nan','None')},key=str.casefold)
    if not opts:st.caption(empty);return None
    return st.selectbox(label,opts,key=key)
def source_table(frames):
    source_ids=set()
    for f in frames:source_ids |= tokens(f,'source_id') | tokens(f,'candidate_source_id') | tokens(f,'annotation_source_id') | tokens(f,'membership_source_id')
    sources=subset('Sources','source_id',source_ids)
    grid(sources,'Original source records','sources')
    if not sources.empty:
        for _,r in sources.iterrows():
            title=str(r.get('source_id','Source'))
            url=str(r.get('file_url','')).strip()
            doi=str(r.get('study_doi','')).strip()
            if url.startswith(('https://','http://')):st.markdown(f'**{title}:** [Open recorded source]({url})')
            if doi and doi.lower()!='nan':
                doi_url=doi if doi.startswith('http') else 'https://doi.org/'+doi.removeprefix('doi:')
                st.markdown(f'**{title}:** [DOI / publication]({doi_url})')
    return sources

def interval_figure(record,genes):
    start=pd.to_numeric(pd.Series([record.get('left_anchor_bp')]),errors='coerce').iloc[0]
    end=pd.to_numeric(pd.Series([record.get('right_anchor_bp')]),errors='coerce').iloc[0]
    if pd.isna(start) or pd.isna(end) or end<=start:
        st.info('This MQTL does not have a valid physical interval in the workbook.')
        return
    fig=go.Figure()
    fig.add_trace(go.Scatter(x=[start/1e6,end/1e6],y=[2,2],mode='lines+markers',line=dict(color='#12825f',width=16),marker=dict(color='#0c6548',size=11),name='MQTL interval',hovertemplate='Position: %{x:.3f} Mb<extra>MQTL</extra>'))
    g=genes.copy()
    if not g.empty and {'start_bp','end_bp'}.issubset(g.columns):
        g['s']=pd.to_numeric(g['start_bp'],errors='coerce');g['e']=pd.to_numeric(g['end_bp'],errors='coerce')
        g=g.dropna(subset=['s','e']);g=g[(g.e>=start)&(g.s<=end)].head(250)
        if not g.empty:
            g['mid']=(g.s+g.e)/2/1e6
            fig.add_trace(go.Scatter(x=g.mid,y=[1]*len(g),mode='markers',marker=dict(color='#e3a44d',size=8,opacity=.75),text=g['gene_id'] if 'gene_id' in g else ['Gene']*len(g),hovertemplate='Gene %{text}<br>%{x:.3f} Mb<extra></extra>',name='Genes (up to 250)'))
    fig.update_layout(height=240,margin=dict(l=15,r=15,t=30,b=25),paper_bgcolor='rgba(0,0,0,0)',plot_bgcolor='rgba(0,0,0,0)',xaxis=dict(title='Physical position (Mb)',showgrid=True,gridcolor='#e7eee9'),yaxis=dict(tickvals=[1,2],ticktext=['Genes','MQTL'],range=[.5,2.5],showgrid=False),legend=dict(orientation='h',y=1.22),font=dict(color='#315a49'))
    st.plotly_chart(fig,use_container_width=True)

st.sidebar.markdown('## 🌾 RiceAcid-QTL')
st.sidebar.caption('GENOMICS · META-QTL · EVIDENCE')
st.sidebar.divider()
with st.sidebar.expander('Data connection',expanded=False):
    st.caption('Local database is loaded automatically when present.')
    uploaded=st.file_uploader('Optional replacement Excel workbook',type=['xlsx'])
    st.code(str(DEFAULT),language=None)
if uploaded is not None:raw=uploaded.getvalue();data_mode='Uploaded workbook'
elif DEFAULT.is_file():raw=DEFAULT.read_bytes();data_mode='Local Excel workbook'
else:
    st.markdown('<div class="hero"><div class="eyebrow">Research decision-support system</div><h1>RiceAcid · QTL Explorer</h1><p>Discover loci, genes, molecular evidence and original sources.</p></div>',unsafe_allow_html=True)
    st.warning('No workbook found. Place Al_MetaQTL_Database_Input.xlsx in the repository data folder, or upload it using the sidebar.')
    st.code(str(DEFAULT),language=None)
    st.info('For Streamlit Cloud, the local Windows data folder is not accessible. Use the sidebar uploader or configure private server-side storage.')
    st.stop()
try:DB=read_workbook(raw)
except Exception as exc:st.error(f'Unable to read Excel workbook: {exc}');st.stop()

st.sidebar.success(f'Connected · {data_mode}')
st.sidebar.caption(f'{len(DB)} workbook sheets available')
page=st.sidebar.radio('Explore', ['Overview','Discovery workflow','Database catalogue'],index=1)
st.sidebar.divider()
st.sidebar.caption('Research use only · Links indicate source-reported associations, not experimental validation.')

st.markdown('<div class="hero"><div class="eyebrow">Rice acid-soil research intelligence</div><h1>RiceAcid · QTL Explorer</h1><p>From soil stress to meta-QTL, candidate gene, molecular evidence and traceable scientific source.</p></div>',unsafe_allow_html=True)

mq_all=sheet('MQTLs');qtl_all=sheet('QTLs');genes_all=sheet('Genes')
if page=='Overview':
    st.markdown('## Research database at a glance')
    a,b,c,d=st.columns(4)
    with a:stat('Meta-QTLs',f'{len(mq_all):,}','Reported MQTL records')
    with b:stat('Original QTLs',f'{len(qtl_all):,}','Source-reported loci')
    with c:stat('Candidate genes',f'{len(genes_all):,}','Gene inventory')
    with d:stat('Published studies',f'{len(sheet("Studies")):,}','Source studies')
    left,right=st.columns([1.5,1])
    with left:
        st.markdown('### Meta-QTL distribution across chromosomes')
        if not mq_all.empty and 'chromosome' in mq_all:
            cnt=mq_all['chromosome'].astype(str).value_counts().sort_index().reset_index()
            cnt.columns=['Chromosome','MQTLs']
            fig=px.bar(cnt,x='Chromosome',y='MQTLs',color_discrete_sequence=['#188568'])
            fig.update_layout(plot_bgcolor='rgba(0,0,0,0)',paper_bgcolor='rgba(0,0,0,0)',height=350)
            st.plotly_chart(fig,use_container_width=True)
    with right:
        st.markdown('### Evidence coverage')
        for label,name in [('Projected QTLs','Projected_QTLs'),('Expression records','Expression'),('Protein features','Protein_Features'),('Enrichment terms','Enrichment'),('qPCR summaries','qPCR_Summary'),('Molecular markers','Markers')]:
            st.markdown(f'**{label}** · {len(sheet(name)):,}')
    st.markdown('<div class="soft-note">Select <b>Discovery workflow</b> from the sidebar to explore an element → MQTL → QTL → gene → evidence → source.</div>',unsafe_allow_html=True)
    st.stop()

if page=='Database catalogue':
    st.markdown('## Workbook catalogue & quality control')
    inv=pd.DataFrame([{'Sheet':n,'Records':len(d),'Fields':len(d.columns),'Missing cells':int(d.isna().sum().sum())} for n,d in DB.items()])
    grid(inv,'Imported sheets','inventory')
    name=choice('Inspect source sheet',DB.keys(),'sheet_inventory')
    if name:grid(sheet(name),f'Original data · {name}','sheet_'+name)
    st.stop()

st.markdown('## Guided scientific discovery')
st.markdown('<div class="soft-note">Each selection narrows the downstream evidence. The original records and identifiers remain unchanged; missing evidence is clearly marked.</div>',unsafe_allow_html=True)

left,right=st.columns([1,1])
with left:
    st.markdown('<div class="step">Step 01 · Soil element / stress</div>',unsafe_allow_html=True)
    element=st.selectbox('Select chemical element',list(ELEMENTS),format_func=lambda k:f'{k} · {ELEMENTS[k]}')
with right:
    st.markdown('<div class="step">Data availability</div>',unsafe_allow_html=True)
    st.markdown(f'<div class="card"><div class="k">Current evidence domain</div><div class="v" style="font-size:1.2rem">{"Aluminium tolerance" if element=="Al" else "Dataset not yet connected"}</div><div class="sub">{"Curated workbook available" if element=="Al" else "Future element-specific import required"}</div></div>',unsafe_allow_html=True)
if element!='Al':
    st.info(f'{ELEMENTS[element]} has no imported QTL dataset yet. Aluminium records will not be reused for this element.')
    st.stop()
if mq_all.empty or 'mqtl_id' not in mq_all:
    st.error('The workbook requires an MQTLs sheet containing mqtl_id.');st.stop()

c1,c2=st.columns([1,2])
with c1:
    st.markdown('<div class="step">Step 02 · Chromosome</div>',unsafe_allow_html=True)
    chroms=sorted(tokens(mq_all,'chromosome'),key=lambda x:(int(x) if x.isdigit() else 999,x))
    chrom=st.selectbox('Filter rice chromosome',['All chromosomes']+chroms,format_func=lambda x:f'Chromosome {x}' if x!='All chromosomes' else x)
with c2:
    st.markdown('<div class="step">Step 03 · Meta-QTL</div>',unsafe_allow_html=True)
    filtered=mq_all if chrom=='All chromosomes' else subset('MQTLs','chromosome',{chrom})
    mqtl=choice('Select MQTL ID',tokens(filtered,'mqtl_id'),'mqtl_pick')
if not mqtl:st.stop()
record=subset('MQTLs','mqtl_id',{mqtl}).iloc[0]
member=subset('MQTL_QTLs','mqtl_id',{mqtl})
original=subset('QTLs','qtl_id',tokens(member,'qtl_id'))
projected=subset('Projected_QTLs','projected_qtl_id',tokens(member,'projected_qtl_id'))
links=subset('MQTL_Genes','mqtl_id',{mqtl})
linked_genes=subset('Genes','gene_id',tokens(links,'gene_id'))

st.markdown('### Selected MQTL · Evidence snapshot')
a,b,c,d=st.columns(4)
with a:stat('MQTL',mqtl,'Selected region')
with b:stat('Original QTLs',len(original),'Explicit relationships')
with c:stat('Projected QTLs',len(projected),'Explicit relationships')
with d:stat('Linked genes',len(linked_genes),'Source-reported associations')
with st.container(border=True):
    st.markdown('#### Chromosome interval & gene positions')
    interval_figure(record,linked_genes)
    cols=st.columns(4)
    for col,(label,key) in zip(cols,[('Chromosome','chromosome'),('Peak (cM)','peak_cm'),('CI width (cM)','ci95_width_cm'),('Mean PVE (%)','mean_pve_pct')]):
        with col:st.metric(label,pretty(record.get(key)))
    st.caption(f'Genome assembly: {pretty(record.get("genome_assembly"))}. cM and bp are distinct coordinate systems; the physical plot uses supplied bp anchors only.')

st.markdown('### Select locus and gene for evidence drill-down')
qcol,gcol=st.columns(2)
with qcol:
    qt=st.radio('QTL category',['Original QTL','Projected QTL'],horizontal=True)
    active_qtl=original if qt=='Original QTL' else projected
    qfield='qtl_id' if qt=='Original QTL' else 'projected_qtl_id'
    selected_qtl=choice('Select linked QTL (optional)',tokens(active_qtl,qfield),'qtl_pick')
with gcol:
    selected_gene=choice('Select candidate gene',tokens(linked_genes,'gene_id'),'gene_pick')

# Gene evidence is always restricted by the selected gene. QTL selection is shown independently;
# there is no asserted QTL-to-gene link beyond shared MQTL membership.
selected_qtl_frame=subset('QTLs','qtl_id',{selected_qtl}) if qt=='Original QTL' else subset('Projected_QTLs','projected_qtl_id',{selected_qtl})
gene_frame=subset('Genes','gene_id',{selected_gene})
expr=subset('Expression','gene_id',{selected_gene})
expr_ds=subset('Expression_Datasets','expression_dataset_id',tokens(expr,'expression_dataset_id'))
protein=subset('Protein_Features','gene_id',{selected_gene})
promoter=subset('Promoter_Sites','gene_id',{selected_gene})
enrich_members=subset('Enrichment_Genes','gene_id',{selected_gene})
enrich=subset('Enrichment','enrichment_id',tokens(enrich_members,'enrichment_id'))
qpcr=subset('qPCR_Summary','gene_id',{selected_gene})

st.markdown('### Molecular evidence and provenance')
tabs=st.tabs(['MQTL & QTLs','Genes & traits','Expression','Protein & promoter','Enrichment','qPCR & primers','Markers','Original sources'])
with tabs[0]:
    grid(pd.DataFrame([record]),'Selected meta-QTL','mqtl_record')
    grid(member,'Constituent QTL relationships','mqtl_qtl_members')
    grid(original,'Original QTLs','original_qtls')
    grid(projected,'Projected QTLs','projected_qtls')
    grid(selected_qtl_frame,'Selected QTL record','selected_qtl')
with tabs[1]:
    grid(gene_frame,'Selected candidate gene','gene_record')
    grid(links,'All MQTL–gene relationships','mqtl_gene_links')
    trait_ids=tokens(original,'trait_code')
    if qt=='Original QTL' and selected_qtl:trait_ids=tokens(selected_qtl_frame,'trait_code')
    grid(subset('Traits','trait_code',trait_ids),'Traits supported by linked original QTLs','traits', 'Traits are linked through QTL records, not automatically attributed to the selected gene.')
with tabs[2]:
    grid(expr,'Gene expression observations','expression','Expression values and scales are preserved as reported; cross-study comparisons require harmonization.')
    grid(expr_ds,'Associated expression datasets','expression_datasets')
    if not expr.empty and 'expression_value_as_supplied' in expr:
        v=pd.to_numeric(expr['expression_value_as_supplied'],errors='coerce')
        plot=expr.loc[v.notna()].copy()
        if not plot.empty:
            plot['value']=v[v.notna()]
            plot['dataset']=plot.get('expression_dataset_id',pd.Series(['Record']*len(plot),index=plot.index)).astype(str)
            fig=px.bar(plot,x='dataset',y='value',color_discrete_sequence=['#19876a'],title='Reported expression values (not normalized across datasets)')
            fig.update_layout(plot_bgcolor='rgba(0,0,0,0)',paper_bgcolor='rgba(0,0,0,0)')
            st.plotly_chart(fig,use_container_width=True)
with tabs[3]:
    grid(protein,'Protein features and annotations','protein')
    grid(promoter,'Promoter regulatory motifs','promoter')
with tabs[4]:
    grid(enrich_members,'Gene–enrichment memberships','enrichment_members')
    grid(enrich,'Linked enrichment terms','enrichment_terms','Enrichment reflects gene-set membership; it is not direct proof of gene function.')
with tabs[5]:
    grid(qpcr,'qPCR records explicitly linked by gene ID','qpcr_gene')
    possible={selected_gene} if selected_gene else set()
    if not gene_frame.empty:
        for field in ['symbol','symbol_aliases']:
            for val in tokens(gene_frame,field):possible.update(x.strip() for x in re.split(r'[,;/]',val) if x.strip())
    target=subset('qPCR_Summary','target_name',possible)
    if not target.empty:grid(target,'Potential target-name matches (requires confirmation)','qpcr_possible','A matching target name does not independently establish gene identity.')
    target_names=tokens(qpcr,'target_name')|tokens(target,'target_name')
    grid(subset('qPCR_Values','target_name',target_names),'qPCR replicate values','qpcr_values')
    grid(subset('Primers','target_name',target_names),'Primer sequences','primers')
with tabs[6]:
    grid(subset('MQTL_Markers','mqtl_id',{mqtl}),'MQTL-associated markers','mqtl_markers')
    grid(subset('QTL_Markers','qtl_id',tokens(original,'qtl_id')),'Original QTL-associated markers','qtl_markers')
    names=set()
    for f in (pd.DataFrame([record]),member):
        for col in ['left_marker','right_marker','left_marker_as_supplied','right_marker_as_supplied']:names|=tokens(f,col)
    grid(subset('Markers','marker_name',names),'Exact marker-name matches','marker_lookup','Names alone may not uniquely identify markers across genetic maps; confirm chromosome and study.')
with tabs[7]:
    studies=subset('Studies','study_id',tokens(original,'study_id'))
    grid(studies,'Contributing published studies','studies')
    frames=[pd.DataFrame([record]),member,original,projected,links,gene_frame,expr,expr_ds,protein,promoter,enrich,enrich_members,qpcr,studies]
    source_table(frames)
    st.markdown('<div class="soft-note"><b>Evidence integrity:</b> All linked records are based on explicit identifiers in the workbook. Source references can be incomplete; check the original publication before drawing biological conclusions.</div>',unsafe_allow_html=True)

st.divider()
st.caption('RiceAcid-QTL DSS · Local-first research interface · Source data remain unmodified · Not a validated de novo meta-QTL or functional validation engine.')
