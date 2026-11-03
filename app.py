"""RiceAcid-QTL DSS — research explorer; no de novo meta-analysis."""
from pathlib import Path
import io
import pandas as pd
import streamlit as st
import plotly.express as px

st.set_page_config(page_title='RiceAcid-QTL DSS', page_icon='🌾', layout='wide')
ROOT=Path(__file__).resolve().parent
DEFAULT=ROOT/'data'/'Al_MetaQTL_Database_Input.xlsx'
@st.cache_data(show_spinner='Loading research database...')
def load_book(blob: bytes):
    x=pd.ExcelFile(io.BytesIO(blob))
    return {name:pd.read_excel(x,sheet_name=name).dropna(how='all') for name in x.sheet_names}

def show_table(df, key, cols=None):
    if cols: df=df[[c for c in cols if c in df.columns]]
    st.dataframe(df,hide_index=True,use_container_width=True)
    st.download_button('Download CSV',df.to_csv(index=False).encode('utf-8-sig'),file_name=f'{key}.csv',mime='text/csv',key='dl_'+key)

def numeric(s): return pd.to_numeric(s,errors='coerce')

st.sidebar.title('🌾 RiceAcid-QTL DSS')
source=st.sidebar.file_uploader('Optional: replace workbook',type=['xlsx'])
if source is not None:
    blob = source.getvalue()
elif DEFAULT.is_file():
    blob = DEFAULT.read_bytes()
else:
    st.title('Rice Acid-Soil QTL Research DSS')
    st.info('Upload Al_MetaQTL_Database_Input.xlsx using the sidebar to start exploring your research data.')
    st.caption('Your workbook is processed in the current Streamlit session. Do not upload confidential data to an untrusted deployment.')
    st.stop()
try: tables=load_book(blob)
except Exception as e:
    st.error(f'Could not read workbook: {e}');st.stop()
page=st.sidebar.radio('Navigate',['Dashboard','Meta-QTL Explorer','Original QTL Explorer','Candidate Genes','Functional Evidence','Literature & Provenance','Data Quality & Export'])
st.sidebar.caption('Research prototype • descriptive exploration only. Gene ranking is heuristic, not experimental validation.')

if page=='Dashboard':
    st.title('Rice Acid-Soil QTL Research DSS')
    st.caption('Evidence-based exploration of aluminium-tolerance QTLs and meta-QTLs')
    metrics=[('Meta-QTLs','MQTLs'),('Original QTLs','QTLs'),('Genes','Genes'),('Studies','Studies')]
    for col,(label,key) in zip(st.columns(4),metrics):col.metric(label,len(tables.get(key,[])))
    mq=tables['MQTLs'].copy()
    fig=px.histogram(mq,x='chromosome',color='retained' if 'retained' in mq else None,barmode='group',title='Published meta-QTL distribution by chromosome')
    st.plotly_chart(fig,use_container_width=True)
    st.info('The workbook contains source-reported intervals and candidate evidence. No new meta-QTL statistics are computed by this app.')
    st.subheader('Available research datasets')
    show_table(pd.DataFrame({'Sheet':list(tables),'Records':[len(v) for v in tables.values()]}),'dataset_inventory')

elif page=='Meta-QTL Explorer':
    st.title('Meta-QTL Explorer')
    df=tables['MQTLs'].copy()
    chrom=st.multiselect('Chromosomes',sorted(df.chromosome.dropna().unique().tolist()),default=sorted(df.chromosome.dropna().unique().tolist()))
    df=df[df.chromosome.isin(chrom)]
    if 'retained' in df:
        status=st.selectbox('Publication retention',['All','Retained only','Non-retained only'])
        if status!='All':df=df[numeric(df.retained).eq(1 if status=='Retained only' else 0)]
    show_table(df,'meta_qtls',['mqtl_id','chromosome','retained','cluster_qtl_n','peak_cm','ci95_width_cm','mean_pve_pct','left_anchor_bp','right_anchor_bp','genome_assembly','source_id'])
    if len(df):
        chart=df.copy();chart['start_mb']=numeric(chart.left_anchor_bp)/1e6;chart['end_mb']=numeric(chart.right_anchor_bp)/1e6
        chart=chart.dropna(subset=['start_mb','end_mb']);chart=chart[chart.end_mb>=chart.start_mb]
        if len(chart):
            fig=px.scatter(chart,x='start_mb',y='chromosome',hover_name='mqtl_id',size='cluster_qtl_n',title='MQTL left anchors (Mb); bubble size = constituent QTL count')
            st.plotly_chart(fig,use_container_width=True)
        chosen=st.selectbox('Inspect meta-QTL',df.mqtl_id.astype(str).tolist())
        st.subheader('Constituent QTLs')
        rel=tables.get('MQTL_QTLs',pd.DataFrame());show_table(rel[rel.mqtl_id.astype(str)==chosen],'mqtl_members')
        st.subheader('Genes associated with this MQTL')
        links=tables.get('MQTL_Genes',pd.DataFrame());genes=tables['Genes']
        if {'mqtl_id','gene_id'}.issubset(links):
            ids=links.loc[links.mqtl_id.astype(str)==chosen,'gene_id'].dropna().astype(str).unique()
            show_table(genes[genes.gene_id.astype(str).isin(ids)],'mqtl_gene_list',['gene_id','chromosome','start_bp','end_bp','symbol','description','expression_candidate','recurrent_candidate','candidate_record_status'])

elif page=='Original QTL Explorer':
    st.title('Original QTL Explorer')
    q=tables['QTLs'].copy()
    study=st.selectbox('Study',['All']+sorted(q.study_id.dropna().astype(str).unique().tolist()))
    if study!='All':q=q[q.study_id.astype(str)==study]
    trait=st.selectbox('Trait',['All']+sorted(q.trait_code.dropna().astype(str).unique().tolist()))
    if trait!='All':q=q[q.trait_code.astype(str)==trait]
    show_table(q,'original_qtls',['qtl_id','study_id','qtl_name_as_supplied','trait_code','chromosome','lod_as_supplied','pve_as_supplied','pve_input_scale','original_peak_cm','original_ci_start_cm','original_ci_stop_cm','source_id','source_record'])
    st.warning('PVE input scales vary in the source data. Original CI start/stop may be reversed. These values are displayed as supplied, not normalized or repaired.')

elif page=='Candidate Genes':
    st.title('Candidate Gene Prioritization')
    g=tables['Genes'].copy()
    chrom=st.selectbox('Chromosome',['All']+sorted(g.chromosome.dropna().unique().tolist()))
    if chrom!='All':g=g[g.chromosome==chrom]
    query=st.text_input('Search gene ID, symbol, description')
    if query:
        mask=g[[c for c in ['gene_id','symbol','description'] if c in g]].fillna('').astype(str).apply(lambda c:c.str.contains(query,case=False,regex=False)).any(axis=1)
        g=g[mask]
    st.caption('Evidence-priority score: +2 if expression candidate; +2 if recurrent candidate; +1 if interval inventory. This is an illustrative sorting heuristic, not a validated biological score.')
    for col in ['expression_candidate','recurrent_candidate','in_interval_inventory']:
        if col not in g:g[col]=0
    g['evidence_priority']=2*numeric(g.expression_candidate).fillna(0)+2*numeric(g.recurrent_candidate).fillna(0)+numeric(g.in_interval_inventory).fillna(0)
    g=g.sort_values(['evidence_priority','gene_id'],ascending=[False,True])
    show_table(g,'ranked_candidate_genes',['gene_id','chromosome','start_bp','end_bp','symbol','description','expression_candidate','recurrent_candidate','in_interval_inventory','evidence_priority','candidate_record_status'])
    if len(g):
        chosen=st.selectbox('Gene details',g.gene_id.astype(str).tolist())
        for sheet in ['Expression','Protein_Features','Promoter_Sites','qPCR_Summary','MQTL_Genes']:
            tab=tables.get(sheet,pd.DataFrame())
            if 'gene_id' in tab:
                matched=tab[tab.gene_id.astype(str)==chosen]
                if len(matched):
                    with st.expander(f'{sheet} ({len(matched)})'):show_table(matched,sheet+'_gene')

elif page=='Functional Evidence':
    st.title('Functional Evidence')
    st.warning('Expression differences, protein predictions and promoter motifs are supporting evidence; they do not by themselves establish gene function or experimental validation.')
    options=['Expression','qPCR_Summary','qPCR_Values','Protein_Features','Promoter_Sites','Enrichment','Primers']
    name=st.selectbox('Evidence dataset',options)
    d=tables.get(name,pd.DataFrame())
    term=st.text_input('Filter all columns (literal text)')
    if term:
        d=d[d.fillna('').astype(str).apply(lambda c:c.str.contains(term,case=False,regex=False)).any(axis=1)]
    show_table(d,'evidence_'+name)

elif page=='Literature & Provenance':
    st.title('Literature and Source Provenance')
    name=st.selectbox('Source table',['Studies','Sources','Data_Notes','Fields'])
    show_table(tables[name],'sources_'+name)
    st.caption('Source identifiers and original record locators are retained; independently verify literature before treating derived records as authoritative.')

else:
    st.title('Data Quality and Export')
    st.subheader('Missing values and duplicate identifiers')
    rows=[]
    for name,d in tables.items():
        candidates=[c for c in d.columns if c.lower() in ['mqtl_id','qtl_id','gene_id','study_id','source_id']]
        key=candidates[0] if candidates else None
        rows.append({'table':name,'records':len(d),'columns':len(d.columns),'missing_cells':int(d.isna().sum().sum()),'duplicate_primary_ids':int(d[key].duplicated().sum()) if key else None})
    show_table(pd.DataFrame(rows),'quality_summary')
    st.subheader('Known scientific cautions')
    st.markdown('- Missing reference assembly identifiers must be resolved before genome-coordinate integration.\n- PVE scales are not uniform across original QTLs.\n- Original genetic-map confidence interval endpoints may need review.\n- Literature-derived candidate status is not equivalent to functional validation.\n- Do not merge meta-QTL intervals from different genome assemblies without verified liftover.')
    st.subheader('Download any complete table')
    name=st.selectbox('Table',list(tables))
    show_table(tables[name],'full_'+name)
