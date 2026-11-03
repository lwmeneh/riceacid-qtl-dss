# RiceAcid-QTL DSS v9

An interactive scientific research explorer with an independent, clean banner, live Excel-derived KPIs, MQTL/QTL/gene drill-down, evidence tables, gene intelligence, provenance and descriptive evidence analytics.

## Local installation

1. Put `Al_MetaQTL_Database_Input.xlsx` inside `data/`.
2. Install: `python -m pip install -r requirements.txt`
3. Start: `python -m streamlit run app.py`
4. Open http://localhost:8501

Your `data/` directory is excluded from Git by `.gitignore`. GitHub Desktop handles code changes only. A Streamlit Cloud deployment cannot read a workbook stored only on your Windows computer.

## Scientific interpretation

Evidence coverage counts presence of records; it is not a statistical confidence score. qPCR records require an explicit gene ID to be treated as linked. Expression values from different datasets are not normalized or compared. This app explores existing meta-QTLs; it does not recompute meta-QTL statistics.
