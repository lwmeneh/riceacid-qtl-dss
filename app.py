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

/* v10: deliberately restrained, non-screenshot portal styling */
[data-testid="stAppViewContainer"]{background:#f2f6f8!important}
[data-testid="stSidebar"]{background:linear-gradient(180deg,#072f3a,#075768)!important;border-right:3px solid #d8ae53!important}
[data-testid="stSidebar"] [data-testid="stImage"] img{width:112px!important;max-width:112px!important;height:112px!important;object-fit:contain!important;border-radius:18px!important;background:#fff!important;padding:5px!important}
.block-container{max-width:1680px!important;padding-top:1.2rem!important}
[data-testid="stTabs"] [role="tablist"]{background:#e6edf2;border-radius:13px;padding:6px;gap:4px;flex-wrap:wrap}
[data-testid="stTabs"] button{border-radius:9px;background:white;padding:8px 12px;font-weight:700}
[data-testid="stTabs"] button[aria-selected="true"]{background:#0b6659!important;color:#fff!important}
[data-testid="stTabs"] button[aria-selected="true"] p{color:#fff!important}
[data-testid="stSelectbox"] [data-baseweb="select"]>div{border-radius:10px!important}
.portal-hero{display:grid;grid-template-columns:132px minmax(0,1fr) 230px;gap:24px;align-items:center;background:linear-gradient(112deg,#062f3b 0%,#09515b 65%,#0c7164 100%);border-radius:22px;min-height:214px;padding:24px 28px;box-shadow:0 14px 32px #063c452a;border-bottom:5px solid #f0bb55;margin-bottom:18px;overflow:hidden}
.portal-logo{width:126px;height:126px;background:#fff;border-radius:18px;display:flex;align-items:center;justify-content:center;padding:10px;box-shadow:0 10px 28px #001d2945}
.portal-logo img{width:100%;height:100%;object-fit:contain}
.portal-eyebrow{font-size:11px;color:#f5d17d;letter-spacing:.13em;font-weight:800;text-transform:uppercase}
.portal-title{font-size:clamp(29px,3.1vw,48px);line-height:1.12;font-weight:850;color:white;letter-spacing:-.045em;margin:9px 0}
.portal-title em{color:#f6d17a;font-style:normal}
.portal-subtitle{font-size:15px;color:#d9f2ec;line-height:1.5}
.portal-chips{display:flex;flex-wrap:wrap;gap:8px;margin-top:14px}
.portal-chips span{font-size:11px;font-weight:700;color:#fff;background:#ffffff1c;border:1px solid #ffffff44;border-radius:99px;padding:5px 11px}
.portal-art{display:flex;align-items:center;justify-content:center;background:radial-gradient(circle,#ffffff1d,transparent 72%);height:150px}
.portal-art img{max-width:100%;max-height:120px;object-fit:contain;filter:drop-shadow(0 8px 14px #001c2550)}
.research-metrics{display:grid;grid-template-columns:repeat(7,minmax(0,1fr));gap:10px;margin:12px 0 18px}
.research-metric{border-radius:15px;padding:13px 11px;color:#fff;min-height:106px;box-shadow:0 6px 16px #113a4922}
.research-icon{font-size:24px}.research-name{font-size:11px;font-weight:700;margin-top:6px}.research-count{font-size:27px;font-weight:850}
.metric-green{background:linear-gradient(135deg,#148f5b,#086341)}.metric-purple{background:linear-gradient(135deg,#9558be,#60339a)}
.metric-blue{background:linear-gradient(135deg,#3b8fdb,#1254a3)}.metric-orange{background:linear-gradient(135deg,#ef9c35,#bf640c)}
.metric-pink{background:linear-gradient(135deg,#ce629c,#973b79)}.metric-cyan{background:linear-gradient(135deg,#299bc7,#0d6998)}
.metric-teal{background:linear-gradient(135deg,#219c9a,#08716f)}
.workflow-header{background:linear-gradient(90deg,#0a7563,#084857);border-radius:14px 14px 0 0;padding:13px 17px;color:#fff;font-weight:800;margin-top:8px}
.reference-flow{display:flex;gap:6px;align-items:center;justify-content:space-between;background:#fff;border:1px solid #dbe8e6;padding:17px 12px;border-radius:0 0 14px 14px;margin-bottom:18px;box-shadow:0 7px 16px #143c4510}
.reference-step{display:flex;flex:1;min-width:0;flex-direction:column;align-items:center;text-align:center;font-size:11px;color:#173b49;gap:7px}
.step-circle{width:52px;height:52px;border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:24px}
.step-color-0{background:#ddf5e8}.step-color-1{background:#e9e5fc}.step-color-2{background:#e1f0fc}.step-color-3{background:#f4e7f9}
.flow-chevron{color:#148d74;font-weight:900}
.visual-grid{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:12px;margin:12px 0 20px}
.visual-tile{background:white;border:1px solid #dce8e7;border-radius:15px;padding:17px}
.visual-tile .ico{font-size:26px}.visual-tile .num{font-size:27px;font-weight:800;color:#0b6659}.visual-tile .lab{font-size:12px;color:#60757b}
.visual-kicker{font-size:12px;letter-spacing:.12em;text-transform:uppercase;font-weight:800;color:#167b67}
.visual-title{font-size:23px;font-weight:800;color:#153c46}
.visual-flow{display:flex;flex-wrap:wrap;gap:8px;margin:12px 0 20px;align-items:center}
.visual-node{padding:10px 13px;background:#fff;border:1px solid #dce8e7;border-radius:11px;font-weight:700;font-size:12px}
.visual-arrow{color:#138b72;font-weight:900}
.image-feature{display:flex;min-height:125px;border-radius:15px;overflow:hidden;background:#0b594b;margin:12px 0 18px}
.image-feature .photo{width:28%;background-size:cover;background-position:center}
.image-feature .copy{padding:20px;color:white}.image-feature .copy strong{display:block;font-size:19px;margin:6px 0}.image-feature .copy span{font-size:13px;color:#d6f1e7}
.sidebar-landscape{height:120px;background-position:center;background-size:cover;border-radius:12px;margin:14px 0}
@media(max-width:1180px){.research-metrics{grid-template-columns:repeat(4,1fr)}.portal-hero{grid-template-columns:105px 1fr}.portal-art{display:none}.portal-logo{width:100px;height:100px}}
@media(max-width:700px){.research-metrics{grid-template-columns:repeat(2,1fr)}.portal-hero{grid-template-columns:1fr;padding:18px}.portal-logo{width:82px;height:82px}.portal-title{font-size:31px}.reference-flow{overflow-x:auto;justify-content:flex-start}.reference-step{min-width:75px}.visual-grid{grid-template-columns:repeat(2,1fr)}}

/* v11 Biotechnology Studio: deliberately no screenshot imagery */
[data-testid="stAppViewContainer"]{background:linear-gradient(160deg,#edf4f6 0%,#f9fbfc 55%,#f2f6f9 100%)!important}
[data-testid="stSidebar"]{background:linear-gradient(180deg,#082b3b 0%,#093d4d 50%,#08483e 100%)!important;border-right:1px solid #174d56!important}
[data-testid="stSidebar"] [data-testid="stImage"] img{height:auto!important;width:100%!important;max-width:100%!important;object-fit:cover!important;background:transparent!important;padding:0!important}
.sidebar-brand{display:flex;gap:11px;align-items:center;padding:7px 0 14px;color:white}
.sidebar-brand .brand-mark{display:grid;place-items:center;width:39px;height:39px;border-radius:12px;background:linear-gradient(135deg,#c3eaa0,#38ad91);color:#063a38;font-size:27px}
.sidebar-brand strong{display:block;font-family:Manrope,sans-serif;font-size:21px;letter-spacing:-.06em}
.sidebar-brand small{display:block;color:#b9d7d5;font-size:9px;letter-spacing:.17em;font-weight:800}
[data-testid="stSidebar"] [role="radiogroup"] label{border-radius:11px;padding:8px 11px;margin:2px 0;background:#ffffff0b;border:1px solid #ffffff13}
[data-testid="stSidebar"] [role="radiogroup"] label:has(input:checked){background:#2eab8a44;border-color:#81d2b4}
[data-testid="stSidebar"] [data-testid="stSidebarUserContent"]{padding-top:1.4rem}
.block-container{max-width:1760px!important;padding-top:1.2rem!important}
.portal-hero{position:relative;isolation:isolate;grid-template-columns:104px minmax(0,1fr) 330px;min-height:228px;padding:27px 32px;gap:25px;border-radius:22px;background:linear-gradient(108deg,#062f43 0%,#064a56 63%,#0d6559 100%);border-bottom:0;box-shadow:0 17px 45px #073c4a25;overflow:hidden}
.portal-hero:after{content:'';position:absolute;right:-80px;top:-100px;width:450px;height:380px;border:1px solid #a2ffda30;border-radius:50%;box-shadow:0 0 0 48px #a2ffda08,0 0 0 105px #a2ffda08;z-index:-1}
.portal-logo{width:104px;height:104px;padding:11px;border-radius:20px;background:#fff;box-shadow:0 7px 24px #002d3a55;align-self:start}
.portal-title{font-size:clamp(32px,3.1vw,50px);letter-spacing:-.055em;margin:8px 0 6px}
.portal-eyebrow{font-size:10px;letter-spacing:.17em;color:#f5d48d}
.portal-subtitle{font-size:15px;color:#e5f7f3}
.portal-chips span{background:#ffffff17;border-color:#ffffff31;color:#e6faf2}
.portal-art{height:170px;border-radius:18px;overflow:hidden;background:#083f4a}
.portal-art img{width:100%;height:100%;max-height:100%;object-fit:cover;filter:none}
.research-metrics{gap:11px;margin:17px 0 22px}
.research-metric{min-height:124px;padding:16px 15px;border-radius:17px;position:relative;overflow:hidden;box-shadow:0 7px 17px #163a4814}
.research-metric:after{content:'';position:absolute;right:-23px;top:-27px;width:110px;height:110px;border:1px solid #ffffff30;border-radius:50%;box-shadow:0 0 0 20px #ffffff09}
.research-icon{font-size:22px}.research-name{font-size:11px;margin-top:10px;color:#fff}.research-count{font-size:30px;line-height:1.15;margin-top:5px}
.workflow-header{font-size:15px;padding:15px 20px;background:linear-gradient(95deg,#0c6b63,#0b4c66)}
.reference-flow{padding:20px 16px 16px;border-radius:0 0 18px 18px;box-shadow:0 12px 24px #17424a0c}
.step-circle{width:57px;height:57px;font-size:23px;box-shadow:inset 0 -3px 0 #0000000a}
.reference-step{font-size:12px;font-weight:650}
[data-testid="stVerticalBlockBorderWrapper"]>div{border-color:#dce9ed!important;border-radius:16px!important}
[data-testid="stSelectbox"] [data-baseweb="select"]>div{min-height:45px;border:1px solid #d8e5e9;border-radius:12px!important;background:#fff}
[data-testid="stTabs"] [role="tablist"]{background:#e8f0f3;gap:5px;padding:7px;border-radius:14px;flex-wrap:wrap}
[data-testid="stTabs"] button{background:#fff!important;border:1px solid #d9e6e9;border-radius:10px!important;font-size:12px!important;padding:9px 13px!important}
[data-testid="stTabs"] button[aria-selected="true"]{background:#075c65!important;border-color:#075c65!important;color:white!important}
[data-testid="stTabs"] button[aria-selected="true"] p{color:white!important}
[data-testid="stDataFrame"]{border-radius:12px;box-shadow:0 4px 12px #123c480b}
[data-testid="stMetric"]{background:#fff;border:1px solid #dce8ed;border-radius:13px}
.stButton>button,.stDownloadButton>button{border-radius:11px!important;font-weight:700!important}
.lab-heading{display:flex;align-items:center;gap:11px;margin:16px 0 10px;font-family:Manrope,sans-serif;color:#103b4c;font-weight:850;font-size:23px}
.lab-heading span{background:#d9f4eb;color:#097c67;padding:8px 11px;border-radius:12px;font-size:19px}
.lab-context{background:#fff;border:1px solid #d9e7ec;border-radius:15px;padding:15px 18px;margin:12px 0 16px;box-shadow:0 5px 17px #0d455010}
.lab-context .label{font-size:10px;letter-spacing:.13em;text-transform:uppercase;color:#617f8c;font-weight:800}
.lab-context .value{font-size:16px;font-weight:800;color:#0d4955;margin-top:4px}
.evidence-band{display:grid;grid-template-columns:repeat(5,1fr);gap:9px;margin:12px 0 22px}
.evidence-cell{background:white;border:1px solid #dce9ed;border-radius:13px;padding:12px 13px;box-shadow:0 3px 10px #143c4b0b}
.evidence-cell b{display:block;color:#0a5b65;font-size:21px}.evidence-cell small{font-size:11px;color:#607c89}
.biotech-note{background:#e7f4f5;border:1px solid #cbe5e5;border-left:4px solid #0e9586;border-radius:12px;padding:12px 16px;color:#1b5361;margin:12px 0}
@media(max-width:1180px){.portal-hero{grid-template-columns:90px 1fr 230px}.portal-logo{width:90px;height:90px}.research-metrics{grid-template-columns:repeat(4,1fr)}}
@media(max-width:800px){.portal-hero{grid-template-columns:70px 1fr;gap:14px;padding:18px}.portal-logo{width:70px;height:70px}.portal-art{grid-column:1/-1;height:100px}.research-metrics{grid-template-columns:repeat(2,1fr)}.evidence-band{grid-template-columns:repeat(2,1fr)}.portal-title{font-size:30px}.reference-flow{overflow-x:auto;justify-content:flex-start}.reference-step{min-width:80px}}

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
    mime = "image/jpeg" if filename.lower().endswith((".jpg", ".jpeg")) else "image/png"
    return "data:" + mime + ";base64," + base64.b64encode(path.read_bytes()).decode("ascii")

RICE_IMAGE = asset_uri("rice_landscape.png")
LOGO_IMAGE = asset_uri("institute_logo_padded.png")
DNA_IMAGE = asset_uri("biotech_banner_art.png")



st.sidebar.markdown('<div class="sidebar-brand"><span class="brand-mark">◈</span><div><strong>RiceAcid</strong><small>GENOMICS WORKSPACE</small></div></div>',unsafe_allow_html=True)

st.sidebar.caption('MOLECULAR BREEDING · ACID-SOIL RICE')
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
/* Fixed reference-design canvas */
/* Fixed teal sidebar */
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
page=st.sidebar.radio('🧭 Navigation', ['Overview','Discovery workflow','Gene intelligence','Evidence analytics','Database catalogue'],index=1)
st.sidebar.divider()
st.sidebar.caption('Research use only · Source-reported associations are not necessarily experimentally validated.')

# Purpose-built HTML banner: separate brand assets, not a screenshot of the interface.
st.markdown(f"""<div class="portal-hero">
<div class="portal-logo"><img src="{LOGO_IMAGE}" alt="ICAR logo"></div>
<div><div class="portal-eyebrow">ICAR Research Complex for NEH Region · Umiam, Meghalaya</div>
<div class="portal-title">Rice<em>Acid-QTL</em> DSS</div>
<div class="portal-subtitle">A translational genomics workbench for QTL discovery, candidate genes and molecular validation</div>
<div class="portal-chips"><span>◈ QTL atlas</span><span>⌬ Gene evidence</span><span>◉ Multi-omics</span><span>↗ Literature provenance</span></div></div>
<div class="portal-art"><img src="{DNA_IMAGE}" alt="Rice and DNA illustration"></div>
</div>""", unsafe_allow_html=True)
mq_all=sheet('MQTLs');qtl_all=sheet('QTLs');genes_all=sheet('Genes')
# Counts are calculated from the uploaded workbook, never from design mockups.
metric_items = [
    ('🧬', 'Meta-QTLs', len(mq_all), 'metric-green'),
    ('📍', 'Original QTLs', len(qtl_all), 'metric-purple'),
    ('🔗', 'Projected QTLs', len(sheet('Projected_QTLs')), 'metric-blue'),
    ('🌾', 'Candidate genes', len(genes_all), 'metric-orange'),
    ('📊', 'Expression records', len(sheet('Expression')), 'metric-pink'),
    ('🧪', 'qPCR targets', len(sheet('qPCR_Summary')), 'metric-cyan'),
    ('🧷', 'Molecular markers', len(sheet('Markers')), 'metric-teal'),
]
metric_html=''.join(f'<div class="research-metric {color}"><div class="research-icon">{ico}</div><div class="research-name">{html.escape(label)}</div><div class="research-count">{n:,}</div></div>' for ico,label,n,color in metric_items)
st.markdown('<div class="research-metrics">'+metric_html+'</div>',unsafe_allow_html=True)
if page=='Overview':
    st.markdown('<div class="visual-kicker">📊 Scientific knowledgebase</div><div class="visual-title">Research database at a glance</div>',unsafe_allow_html=True)
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


# Advanced exploration uses explicit workbook links only. These analyses do not
# infer novel MQTLs or claim functional validation.
def gene_evidence_index():
    base=sheet('Genes')
    if base.empty or 'gene_id' not in base:return pd.DataFrame()
    out=base[['gene_id']].drop_duplicates().copy()
    for name, frame, field in [
        ('MQTL regions',sheet('MQTL_Genes'),'gene_id'),
        ('Expression records',sheet('Expression'),'gene_id'),
        ('Protein records',sheet('Protein_Features'),'gene_id'),
        ('Promoter sites',sheet('Promoter_Sites'),'gene_id'),
        ('Enrichment memberships',sheet('Enrichment_Genes'),'gene_id'),
        ('qPCR summaries',sheet('qPCR_Summary'),'gene_id')]:
        if field in frame:
            counts=frame.groupby(field,dropna=True).size().rename(name)
            out=out.join(counts,on='gene_id')
        else:out[name]=0
        out[name]=out[name].fillna(0).astype(int)
    evidence_cols=['Expression records','Protein records','Promoter sites','Enrichment memberships','qPCR summaries']
    out['Evidence categories present']=(out[evidence_cols]>0).sum(axis=1)
    out['Evidence coverage (%)']=(out['Evidence categories present']/len(evidence_cols)*100).round(0).astype(int)
    return out

def gene_profile(gid):
    result={}
    for label,table in [('Gene','Genes'),('MQTL membership','MQTL_Genes'),('Expression','Expression'),('Protein','Protein_Features'),('Promoter','Promoter_Sites'),('Enrichment membership','Enrichment_Genes'),('qPCR','qPCR_Summary')]:
        result[label]=subset(table,'gene_id',{gid})
    result['Enrichment terms']=subset('Enrichment','enrichment_id',tokens(result['Enrichment membership'],'enrichment_id'))
    result['Expression datasets']=subset('Expression_Datasets','expression_dataset_id',tokens(result['Expression'],'expression_dataset_id'))
    return result

def gene_report(gid,parts):
    lines=[f'# RiceAcid-QTL DSS | Gene evidence dossier',f'Gene: {gid}', 'Source: locally loaded research workbook', 'Evidence coverage is descriptive, not a biological confidence score.', '']
    for name,frame in parts.items():
        lines.extend([f'## {name}',f'Records: {len(frame)}',''])
        if not frame.empty:
            lines.append(frame.head(100).to_csv(index=False))
    return '\n'.join(lines)

if page=='Gene intelligence':
    st.markdown('## 🧬 Candidate-gene intelligence')
    st.caption('Compare reported evidence across candidate genes. Counts are descriptive, not rankings of biological causality.')
    idx=gene_evidence_index()
    if idx.empty:st.info('No gene inventory available.');st.stop()
    search=st.text_input('Search gene ID or symbol',placeholder='e.g. LOC_Os...')
    if search:
        ids=sheet('Genes')
        matches=ids[[c for c in ['gene_id','symbol','description'] if c in ids]].astype(str).apply(lambda c:c.str.contains(search,case=False,regex=False,na=False)).any(axis=1)
        idx=idx[idx.gene_id.isin(ids.loc[matches,'gene_id'])]
    c1,c2,c3=st.columns(3)
    with c1:st.metric('Genes in current view',f'{len(idx):,}')
    with c2:st.metric('With expression evidence',f'{(idx["Expression records"]>0).sum():,}')
    with c3:st.metric('With qPCR records',f'{(idx["qPCR summaries"]>0).sum():,}')
    st.dataframe(idx.sort_values(['Evidence categories present','MQTL regions'],ascending=False),hide_index=True,use_container_width=True,height=340)
    st.download_button('Export evidence coverage table',idx.to_csv(index=False).encode('utf-8-sig'),'gene_evidence_coverage.csv','text/csv')
    gid=choice('Inspect candidate gene',tokens(idx,'gene_id'),'intelligence_gene')
    if gid:
        parts=gene_profile(gid)
        gene=parts['Gene']
        if not gene.empty:
            st.markdown('### Gene annotation')
            st.dataframe(gene,hide_index=True,use_container_width=True)
        c1,c2=st.columns([1,1])
        with c1:
            cov=idx[idx.gene_id==gid]
            if not cov.empty:
                fields=['Expression records','Protein records','Promoter sites','Enrichment memberships','qPCR summaries']
                fig=go.Figure(go.Bar(x=[int(cov.iloc[0][f]>0) for f in fields],y=fields,orientation='h',marker_color=['#20a77c','#2c80cc','#df9b36','#9968c9','#d9689b']))
                fig.update_layout(title='Evidence availability by category (0/1)',height=310,xaxis=dict(range=[0,1.1],tickvals=[0,1]),margin=dict(l=5,r=5,t=50,b=15),paper_bgcolor='rgba(0,0,0,0)',plot_bgcolor='rgba(0,0,0,0)')
                st.plotly_chart(fig,use_container_width=True)
        with c2:
            expr=parts['Expression']
            if not expr.empty:
                plot=expr.copy();plot['value']=pd.to_numeric(plot.get('expression_value_as_supplied'),errors='coerce')
                plot=plot.dropna(subset=['value'])
                if not plot.empty:
                    fig=px.bar(plot.head(35),x='expression_dataset_id',y='value',color='metric_as_supplied' if 'metric_as_supplied' in plot else None,title='Reported expression values (original scales)')
                    fig.update_layout(height=310,paper_bgcolor='rgba(0,0,0,0)',plot_bgcolor='rgba(0,0,0,0)')
                    st.plotly_chart(fig,use_container_width=True)
                    st.caption('Different studies may use different scales; values must not be compared as normalized measurements.')
        subtabs=st.tabs(list(parts))
        for tab,(label,frame) in zip(subtabs,parts.items()):
            with tab:grid(frame,label,'int_'+label.replace(' ','_'))
        st.download_button('⬇ Download gene evidence dossier',gene_report(gid,parts).encode('utf8'),f'{gid}_dossier.md','text/markdown')
    st.stop()

if page=='Evidence analytics':
    st.markdown('## 📊 Evidence analytics & provenance')
    st.caption('Data availability, missingness and literature-linked records from the imported workbook.')
    coverage=gene_evidence_index()
    a,b=st.columns([1,1])
    with a:
        if not coverage.empty:
            categories=['Expression records','Protein records','Promoter sites','Enrichment memberships','qPCR summaries']
            bars=pd.DataFrame({'Category':categories,'Genes with records':[(coverage[k]>0).sum() for k in categories]})
            fig=px.bar(bars,x='Genes with records',y='Category',orientation='h',color='Category',title='Gene-level evidence coverage')
            fig.update_layout(showlegend=False,height=370,paper_bgcolor='rgba(0,0,0,0)',plot_bgcolor='rgba(0,0,0,0)')
            st.plotly_chart(fig,use_container_width=True)
    with b:
        m=sheet('MQTLs')
        if not m.empty and 'chromosome' in m:
            dist=m.groupby('chromosome').size().reset_index(name='MQTL count')
            fig=px.bar(dist,x='chromosome',y='MQTL count',title='Meta-QTL distribution',color_discrete_sequence=['#118c70'])
            fig.update_layout(height=370,paper_bgcolor='rgba(0,0,0,0)',plot_bgcolor='rgba(0,0,0,0)')
            st.plotly_chart(fig,use_container_width=True)
    inventory=pd.DataFrame([{'Table':name,'Records':len(frame),'Columns':len(frame.columns),'Missing values':int(frame.isna().sum().sum())} for name,frame in DB.items()])
    grid(inventory,'Database completeness overview','advanced_inventory')
    grid(sheet('Data_Notes'),'Curation notes and data limitations','advanced_notes')
    st.stop()

if page=='Database catalogue':
    st.markdown('## Workbook catalogue & quality control')
    inv=pd.DataFrame([{'Sheet':n,'Records':len(d),'Fields':len(d.columns),'Missing cells':int(d.isna().sum().sum())} for n,d in DB.items()])
    grid(inv,'Imported sheets','inventory')
    name=choice('Inspect source sheet',DB.keys(),'sheet_inventory')
    if name:grid(sheet(name),f'Original data · {name}','sheet_'+name)
    st.stop()

st.markdown('<div class="workflow-header">🧬 &nbsp; Discovery workflow: From element to molecular evidence</div>',unsafe_allow_html=True)
workflow=[('🧪','Element'),('🎯','Meta-QTL'),('📍','QTLs'),('🧬','Candidate gene'),('🌱','Traits'),('🔬','Molecular evidence'),('🧫','Primers / markers'),('📚','Sources')]
st.markdown('<div class="reference-flow">'+''.join(f'<div class="reference-step"><span class="step-circle step-color-{i%4}">{ico}</span><b>{label}</b></div>'+('<span class="flow-chevron">➜</span>' if i<len(workflow)-1 else '') for i,(ico,label) in enumerate(workflow))+'</div>',unsafe_allow_html=True)
st.caption('Selections are linked only when the workbook contains an explicit relationship. Unlinked evidence can be browsed separately.')
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

# Biotechnologist workbench: show the exact scope before exploring downstream tables.
st.markdown('<div class="lab-heading"><span>⌬</span> Molecular evidence workbench</div>',unsafe_allow_html=True)
context_cols=st.columns(3)
for col, label, val in zip(context_cols,['ACTIVE META-QTL','SELECTED QTL','CANDIDATE GENE'],[mqtl,selected_qtl or 'Not selected',selected_gene or 'Not selected']):
    with col:
        st.markdown(f'<div class="lab-context"><div class="label">{html.escape(label)}</div><div class="value">{html.escape(str(val))}</div></div>',unsafe_allow_html=True)
counts=[('Expression',len(expr)),('Protein',len(protein)),('Promoter',len(promoter)),('Enrichment',len(enrich_members)),('qPCR',len(qpcr))]
st.markdown('<div class="evidence-band">'+''.join(f'<div class="evidence-cell"><b>{n:,}</b><small>{html.escape(label)} linked records</small></div>' for label,n in counts)+'</div>',unsafe_allow_html=True)
st.markdown('<div class="biotech-note">Evidence counts indicate records directly linked by gene ID. Use each tab’s <b>complete catalogue</b> to inspect independent studies. Missing evidence is not evidence of biological absence.</div>',unsafe_allow_html=True)
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

# End of advanced DSS v9.
