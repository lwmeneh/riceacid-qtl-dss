# RiceAcid-QTL DSS — Biotechnology Studio v11

A polished Streamlit research portal with QTL discovery, gene evidence, multi-omics catalogues and provenance.

## Local setup

Place `Al_MetaQTL_Database_Input.xlsx` in `data/` (kept out of public GitHub), then run:

```bash
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

The app does not infer biological validation from evidence counts. Non-aluminium element datasets must be imported separately.

## GitHub safety

Use `.gitignore` to exclude `data/`. GitHub Desktop tracks code but does not make local data available to Streamlit Cloud.
