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
def pretty(x):return '—' if pd.isna(x) else str(x)
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
    opts=sorted({str(v) for v in options if pd.notna(v) and str(v).strip() not in ('','nan')},key=str.casefold)
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

def interval_figure(record, genes):
    fig = go.Figure()
    return fig
