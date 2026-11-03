"""RiceAcid-QTL DSS | Research explorer for acid-soil rice genetics.
Run: streamlit run app.py
Place the workbook in data/Al_MetaQTL_Database_Input.xlsx.
"""
from pathlib import Path
import io
import re
import html
import base64
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

/* v6 · Editorial research dashboard design */
[data-testid="stAppViewContainer"]{background:radial-gradient(circle at 88% 6%,#e7f5ef 0%,transparent 32%),#f4f7fb!important}
[data-testid="stSidebar"]{border-right:1px solid #ffffff26!important}
[data-testid="stSidebar"] .stRadio label{padding:9px 12px!important;border-radius:12px!important;transition:background .2s ease}
[data-testid="stSidebar"] .stRadio label:hover{background:#ffffff1c!important}
.block-container{max-width:1550px;padding-top:1.1rem!important}
.hero{position:relative;overflow:hidden;min-height:198px;display:flex;flex-direction:column;justify-content:center;border:1px solid #ffffff35!important;box-shadow:0 20px 44px #123e3526!important}
.hero:after{content:'◉  ◇  ⌬  ✦';position:absolute;right:35px;top:27px;font-size:38px;letter-spacing:15px;color:#ffffff32;transform:rotate(-12deg);pointer-events:none}
.hero h1{font-size:clamp(1.8rem,3vw,2.9rem)!important;line-height:1.18!important;max-width:750px}
.hero p{max-width:720px;font-size:1.03rem!important;line-height:1.55}
.card{position:relative;overflow:hidden;transition:transform .2s,box-shadow .2s;border-radius:19px!important}
.card:hover{transform:translateY(-3px);box-shadow:0 15px 30px #193c3d1b!important}
.card:after{content:'✦';position:absolute;right:18px;top:12px;color:#bedbd0;font-size:25px}
.card .v{font-size:1.95rem!important}
[data-testid="stTabs"] [role="tablist"]{gap:5px;flex-wrap:wrap;background:#e9eff4;border-radius:15px;padding:6px}
[data-testid="stTabs"] button{background:white;border-radius:10px;padding:9px 13px;min-height:42px}
[data-testid="stTabs"] button[aria-selected="true"]{background:#dff3e8!important;box-shadow:0 2px 7px #1a5b3d1a}
[data-testid="stSelectbox"] [data-baseweb="select"]>div{border-radius:11px!important;min-height:45px}
[data-testid="stVerticalBlockBorderWrapper"]>div{border-radius:17px!important}
[data-testid="stDataFrame"]{box-shadow:0 4px 15px #183c3510}
[data-testid="stDownloadButton"] button{border-radius:10px!important}
[data-testid="stSidebar"] [data-testid="stImage"] img{max-width:104px!important}
.visual-kicker{font-size:.72rem;font-weight:800;letter-spacing:.14em;text-transform:uppercase;color:#27765b;margin:15px 0 9px}
.visual-title{font-size:1.42rem;font-weight:800;letter-spacing:-.035em;color:#183b36;margin:10px 0 16px}
.visual-grid{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:13px;margin:10px 0 22px}
.visual-tile{background:#fff;border:1px solid #e3e9ee;border-radius:16px;padding:17px;box-shadow:0 8px 22px #1e45510d}
.visual-tile .ico{font-size:27px;margin-bottom:7px}
.visual-tile .num{font-size:27px;font-weight:800;letter-spacing:-.04em;color:#184b3d}
.visual-tile .lab{font-size:12px;color:#637781;font-weight:700}
.visual-flow{display:flex;flex-wrap:wrap;gap:7px;margin:8px 0 19px;align-items:center}
.visual-node{background:#fff;border:1px solid #d9e6e0;border-radius:13px;padding:11px 13px;font-size:12px;font-weight:700;color:#244a40;box-shadow:0 3px 9px #173f3310}
.visual-arrow{color:#169f79;font-weight:900}
.image-feature{background:#0d5141;border-radius:18px;overflow:hidden;display:flex;min-height:135px;margin:12px 0 18px;box-shadow:0 10px 25px #1a56441c}
.image-feature .photo{width:29%;background-position:center;background-size:cover}
.image-feature .copy{padding:20px 23px;color:#fff;flex:1}
.image-feature .copy strong{display:block;font-size:19px;line-height:1.3;margin:5px 0}
.image-feature .copy span{color:#d1efe4;font-size:13px;line-height:1.5}
.sidebar-landscape{height:120px;background-size:cover;background-position:center;border-radius:13px;border:1px solid #ffffff4d;margin:15px 0 8px;box-shadow:0 10px 25px #0002}
@media(max-width:900px){.visual-grid{grid-template-columns:repeat(2,minmax(0,1fr))}.image-feature .photo{width:35%}}
@media(max-width:550px){.visual-grid{grid-template-columns:repeat(2,minmax(0,1fr));gap:7px}.visual-tile{padding:12px}.image-feature .photo{display:none}.hero{min-height:170px}.hero:after{display:none}}
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

def asset_uri(filename):
    path = ROOT / "assets" / filename
    if not path.is_file():
        return ""
    mime = "image/png"
    return "data:" + mime + ";base64," + base64.b64encode(path.read_bytes()).decode("ascii")

RICE_IMAGE = asset_uri("rice_landscape.png")
DNA_IMAGE = asset_uri("rice_dna_banner.png")

st.sidebar.image(str(ROOT / "assets" / "institute_logo.png"), width=112)
st.sidebar.markdown('## RiceAcid-QTL DSS')
st.sidebar.markdown(f'<div class="sidebar-landscape" style="background-image:url({RICE_IMAGE})"></div>', unsafe_allow_html=True)
st.sidebar.caption('RICE GENOMICS  /  ACID-SOIL TOLERANCE')
st.sidebar.markdown('<div style="font-size:.8rem;opacity:.85;padding:2px 0 10px">Research discovery and molecular evidence platform</div>', unsafe_allow_html=True)
PALETTES = {
    'Emerald & Gold': ('#103f36','#0e705c','#f0b957','#f4f8f5','#e6f1eb'),
    'Navy & Teal': ('#142a4a','#126d88','#f3ba68','#f3f7fb','#e4edf7'),
    'Plum & Coral': ('#3d284d','#795080','#f4a68c','#faf6fa','#f1e8f4'),
    'Terracotta & Sage': ('#553a32','#9a5746','#c7aa68','#faf7f2','#f0e9dd'),
}
with st.sidebar.expander('🎨 Appearance', expanded=False):
    palette_name=st.selectbox('Color palette',list(PALETTES),key='palette_choice')
    st.caption('Changes apply immediately; all scientific data remain unchanged.')
ink, accent, highlight, canvas, soft=PALETTES[palette_name]
st.markdown(f"""<style>
[data-testid="stAppViewContainer"] {{background:{canvas}!important}}
[data-testid="stSidebar"] {{background:linear-gradient(165deg,{ink}, {accent})!important}}
.hero {{background:linear-gradient(115deg,{ink}, {accent})!important;border-top:5px solid {highlight}!important}}
.hero .eyebrow {{color:{highlight}!important}}
.card {{border-color:{soft}!important;box-shadow:0 5px 22px #172f3912!important}}
.card .v, [data-testid="stMetricValue"] {{color:{accent}!important}}
.soft-note {{background:{soft}!important;border-left-color:{highlight}!important}}
.pill {{background:{soft}!important;color:{ink}!important}}
.stButton>button[kind="primary"] {{background:{accent}!important;border-color:{accent}!important}}
[data-testid="stTabs"] button[aria-selected="true"] {{color:{accent}!important;border-bottom-color:{highlight}!important}}
[data-testid="stSidebar"] [data-testid="stImage"] img {{background:#ffffff;border-radius:16px;padding:10px;box-shadow:0 7px 25px #0002}}
[data-testid="stSidebar"] [data-testid="stWidgetLabel"] p {{font-weight:600}}
</style>""",unsafe_allow_html=True)
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
st.sidebar.caption('Research use only · Source-reported associations are not necessarily experimentally validated.')

st.markdown('<div class="hero"><div class="eyebrow">🌾 RICE GENOMICS  /  MOLECULAR BREEDING  /  ACID-SOIL RESILIENCE</div><h1>RiceAcid <span style="color:#f8d57e">QTL Explorer</span></h1><p>Discover genomic regions, connect candidate genes to molecular evidence, and trace every finding back to its original source.</p><div style="margin-top:15px"><span class="pill">🧬 Meta-QTL</span><span class="pill">🔬 Functional evidence</span><span class="pill">📚 Literature provenance</span></div></div>',unsafe_allow_html=True)
st.markdown(f'<div class="image-feature"><div class="photo" style="background-image:url({DNA_IMAGE})"></div><div class="copy"><span>INTEGRATED SCIENTIFIC RESOURCE</span><strong>From soil stress to gene-level evidence</strong><span>Aluminium tolerance is available in the current workbook. Other elements can be added as independent, curated evidence collections.</span></div></div>',unsafe_allow_html=True)
st.caption('Institutional identity · Research portal | Element → MQTL → QTL → Gene → Evidence → Literature')

mq_all=sheet('MQTLs');qtl_all=sheet('QTLs');genes_all=sheet('Genes')
if page=='Overview':
    st.markdown('<div class="visual-kicker">📊 Scientific knowledgebase</div><div class="visual-title">Research database at a glance</div>',unsafe_allow_html=True)
    st.markdown(f'''<div class="visual-grid">
      <div class="visual-tile"><div class="ico">🧬</div><div class="num">{len(mq_all):,}</div><div class="lab">Meta-QTLs</div></div>
      <div class="visual-tile"><div class="ico">📍</div><div class="num">{len(qtl_all):,}</div><div class="lab">Original QTLs</div></div>
      <div class="visual-tile"><div class="ico">🌱</div><div class="num">{len(genes_all):,}</div><div class="lab">Candidate genes</div></div>
      <div class="visual-tile"><div class="ico">📚</div><div class="num">{len(sheet('Studies')):,}</div><div class="lab">Published studies</div></div>
    </div>''',unsafe_allow_html=True)
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

st.markdown('<div class="visual-kicker">🧭 Guided exploration</div><div class="visual-title">Discover evidence, step by step</div>',unsafe_allow_html=True)
st.markdown('<div class="visual-flow">'+''.join('<span class="visual-node">'+a+'</span>' + ('<span class="visual-arrow">→</span>' if i<7 else '') for i,a in enumerate(['🧪 Element','🧬 MQTL','📍 QTLs','🌱 Candidate gene','🍃 Traits','🔬 Molecular evidence','🧫 Primers & markers','📚 Sources']))+'</div>',unsafe_allow_html=True)
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

st.markdown('### 🧬 Selected MQTL · Evidence snapshot')
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

st.markdown('### 🔎 Select locus and candidate gene')
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


# -----------------------------------------------------------------------------
# Evidence navigator: linked context + independent catalogue, with provenance.
# -----------------------------------------------------------------------------
def safe_values(frame, column):
    return sorted(tokens(frame, column), key=str.casefold)

def choose_evidence(label, frame, field, key):
    opts = safe_values(frame, field)
    if not opts:
        st.info(f"No available {label.lower()} identifiers in this scope.")
        return None
    return st.selectbox(label, opts, key=key)

def evidence_scope(title, linked, full, key):
    st.markdown(f"#### {title}")
    if not linked.empty:
        st.success(f"{len(linked):,} directly linked record(s) in the current selection.")
    else:
        st.warning("No direct records for the current selection. This does not imply biological absence.")
    options = ["Linked to selected gene"] if not linked.empty else []
    options += ["Browse complete evidence catalogue"]
    scope = st.radio("Evidence scope", options, key="scope_" + key, horizontal=True)
    data = linked if scope.startswith("Linked") else full
    if scope.startswith("Browse"):
        st.caption("Independent catalogue: these records are NOT automatically evidence for the selected MQTL or gene.")
    return data

def evidence_sources(*frames, key):
    ids = set()
    for frame in frames:
        for field in ["source_id", "candidate_source_id", "annotation_source_id", "membership_source_id", "numeric_source_id", "caption_source_id", "curated_source_id", "tm_source_id", "localization_source_id", "sequence_source_id"]:
            ids |= tokens(frame, field)
    records = subset("Sources", "source_id", ids)
    if records.empty:
        st.caption("No explicit source identifier was resolved for these records.")
        return
    with st.expander(f"Original sources and provenance ({len(records)})", expanded=False):
        grid(records, "Source records", "source_" + key)
        for _, source_row in records.iterrows():
            source_id = pretty(source_row.get("source_id"))
            url = str(source_row.get("file_url", "")).strip()
            doi = str(source_row.get("study_doi", "")).strip()
            if url.startswith(("https://", "http://")):
                st.markdown(f"- **{html.escape(source_id)}:** [Source file]({url})")
            if doi and doi.lower() not in ("nan", "none"):
                doi_url = doi if doi.startswith("http") else "https://doi.org/" + doi.removeprefix("doi:")
                if doi_url.startswith("https://"):
                    st.markdown(f"- **{html.escape(source_id)}:** [Publication / DOI]({doi_url})")

def row_select(frame, field, selected, key):
    if selected is None:
        return frame.iloc[:0]
    return frame[frame[field].astype("string").str.strip().eq(str(selected)).fillna(False)] if field in frame else frame.iloc[:0]

st.markdown('### 🔬 Molecular evidence navigator')
st.caption('Follow linked evidence where available, or browse each complete evidence table independently. Every record retains its original source identifiers.')
tabs = st.tabs(['🧬 MQTL & QTLs','🌱 Genes & traits','📈 Expression','🧪 Protein & promoter','🕸️ Enrichment','🔬 qPCR & primers','📍 Markers','📚 Original sources'])

with tabs[0]:
    grid(pd.DataFrame([record]), 'Selected meta-QTL', 'mqtl_record')
    grid(member, 'Constituent QTL relationships', 'mqtl_qtl_members')
    grid(original, 'Original QTLs', 'original_qtls')
    grid(projected, 'Projected QTLs', 'projected_qtls')
    grid(selected_qtl_frame, 'Selected QTL record', 'selected_qtl')
    evidence_sources(pd.DataFrame([record]), member, original, projected, key='qtl')

with tabs[1]:
    grid(gene_frame, 'Selected candidate gene', 'gene_record')
    grid(links, 'MQTL–gene relationships', 'mqtl_gene_links')
    trait_ids = tokens(selected_qtl_frame if qt == 'Original QTL' and selected_qtl else original, 'trait_code')
    trait_data = subset('Traits', 'trait_code', trait_ids)
    grid(trait_data, 'Traits from linked original QTLs', 'traits', 'A trait linked to a QTL is not automatically experimentally demonstrated for every gene in its interval.')
    evidence_sources(gene_frame, links, trait_data, key='traits')

with tabs[2]:
    data = evidence_scope('Gene expression', expr, sheet('Expression'), 'expression')
    gene = choose_evidence('Expression gene ID', data, 'gene_id', 'expr_gene')
    gene_data = row_select(data, 'gene_id', gene, 'expr_rows')
    ds_id = choose_evidence('Expression dataset', gene_data, 'expression_dataset_id', 'expr_dataset')
    rows = row_select(gene_data, 'expression_dataset_id', ds_id, 'expr_ds_rows')
    grid(rows, 'Selected expression observations', 'expression_detail')
    datasets = subset('Expression_Datasets', 'expression_dataset_id', {ds_id} if ds_id else set())
    grid(datasets, 'Dataset and experimental context', 'expression_dataset_detail')
    if not rows.empty and 'expression_value_as_supplied' in rows:
        numbers = pd.to_numeric(rows['expression_value_as_supplied'], errors='coerce').dropna()
        if not numbers.empty:
            st.metric('Reported expression value (first numeric record)', f'{numbers.iloc[0]:g}')
            st.caption('Value and scale are as supplied; no cross-study normalization is implied.')
    evidence_sources(rows, datasets, key='expression')

with tabs[3]:
    subtabs = st.tabs(['Protein features', 'Promoter motifs'])
    with subtabs[0]:
        data = evidence_scope('Protein features', protein, sheet('Protein_Features'), 'protein')
        gene = choose_evidence('Protein gene ID', data, 'gene_id', 'protein_gene')
        rows = row_select(data, 'gene_id', gene, 'protein_rows')
        transcript = choose_evidence('Transcript ID', rows, 'transcript_id', 'protein_transcript')
        detail = row_select(rows, 'transcript_id', transcript, 'protein_detail') if transcript else rows
        grid(detail, 'Protein annotation, domains, localization and sequence', 'protein_details')
        evidence_sources(detail, key='protein')
    with subtabs[1]:
        data = evidence_scope('Promoter sites', promoter, sheet('Promoter_Sites'), 'promoter')
        gene = choose_evidence('Promoter gene ID', data, 'gene_id', 'promoter_gene')
        rows = row_select(data, 'gene_id', gene, 'promoter_rows')
        motif = choose_evidence('Motif name', rows, 'motif_name', 'promoter_motif')
        detail = row_select(rows, 'motif_name', motif, 'promoter_detail') if motif else rows
        grid(detail, 'Promoter motifs and reported positions', 'promoter_details')
        evidence_sources(detail, key='promoter')

with tabs[4]:
    all_members = sheet('Enrichment_Genes')
    members = evidence_scope('Gene–enrichment memberships', enrich_members, all_members, 'enrichment')
    gene = choose_evidence('Gene ID for enrichment', members, 'gene_id', 'enrich_gene')
    gene_members = row_select(members, 'gene_id', gene, 'enrich_gene_rows')
    eid = choose_evidence('Enrichment term ID', gene_members, 'enrichment_id', 'enrich_term')
    detail_members = row_select(gene_members, 'enrichment_id', eid, 'enrich_member_detail')
    terms = subset('Enrichment', 'enrichment_id', {eid} if eid else set())
    grid(detail_members, 'Gene membership evidence', 'enrichment_membership')
    grid(terms, 'Enrichment statistics and pathway annotation', 'enrichment_detail', 'Enrichment indicates membership in an analyzed gene set, not functional validation.')
    evidence_sources(detail_members, terms, key='enrichment')

with tabs[5]:
    st.caption('qPCR summaries, replicate values and primer records are linked by target_name. Gene identity is confirmed only where the source explicitly supplies gene_id.')
    all_qpcr = sheet('qPCR_Summary')
    aliases = {selected_gene} if selected_gene else set()
    if not gene_frame.empty:
        for field in ['symbol', 'symbol_aliases']:
            for val in tokens(gene_frame, field):
                aliases.update(t.strip() for t in re.split(r'[,;/]', val) if t.strip())
    linked_qpcr = subset('qPCR_Summary', 'gene_id', {selected_gene} if selected_gene else set())
    # Alias-only matches are not treated as confirmed gene links.
    data = evidence_scope('qPCR results', linked_qpcr, all_qpcr, 'qpcr')
    target = choose_evidence('qPCR target', data, 'target_name', 'qpcr_target')
    summary = row_select(data, 'target_name', target, 'qpcr_summary')
    values = subset('qPCR_Values', 'target_name', {target} if target else set())
    primers = subset('Primers', 'target_name', {target} if target else set())
    grid(summary, 'qPCR treatment, tissue, genotype and summary values', 'qpcr_summary_detail')
    grid(values, 'qPCR replicate values', 'qpcr_replicates')
    grid(primers, 'Forward and reverse primer sequences', 'primer_details')
    if target and target in aliases and linked_qpcr.empty:
        st.warning('The target name resembles a gene symbol/alias, but the workbook does not confirm the gene-ID link.')
    evidence_sources(summary, values, primers, key='qpcr')

with tabs[6]:
    m_markers = subset('MQTL_Markers', 'mqtl_id', {mqtl})
    q_markers = subset('QTL_Markers', 'qtl_id', tokens(original, 'qtl_id'))
    st.markdown('#### MQTL and QTL marker relationships')
    grid(m_markers, 'MQTL-associated markers', 'mqtl_markers')
    grid(q_markers, 'Original QTL-associated markers', 'qtl_markers')
    all_marker = sheet('Markers')
    names = tokens(m_markers, 'marker_name') | tokens(q_markers, 'marker_name')
    linked_marker = subset('Markers', 'marker_name', names)
    data = evidence_scope('Genetic map markers', linked_marker, all_marker, 'markers')
    marker = choose_evidence('Marker name', data, 'marker_name', 'marker_name_choice')
    detail = row_select(data, 'marker_name', marker, 'marker_detail')
    grid(detail, 'Marker position and study context', 'marker_detail')
    evidence_sources(m_markers, q_markers, detail, key='markers')

with tabs[7]:
    studies = subset('Studies', 'study_id', tokens(original, 'study_id'))
    grid(studies, 'Contributing published studies', 'studies')
    evidence_sources(pd.DataFrame([record]), member, original, projected, links, gene_frame, expr, expr_ds, protein, promoter, enrich, enrich_members, qpcr, studies, key='all')
    st.info('When exploring an independent evidence catalogue, open that section’s source expander to see its own original references.')

st.divider()
st.caption('RiceAcid-QTL DSS · Source records are unchanged · Independent catalogue records are not automatically associated with the selected MQTL.')
