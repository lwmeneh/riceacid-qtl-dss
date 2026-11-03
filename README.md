# RiceAcid-QTL DSS

A responsive, source-traceable rice acid-soil QTL/meta-QTL research browser. Current workbook: aluminium tolerance. Other element selectors are reserved for future element-specific data.

## Run

```bash
pip install -r requirements.txt
streamlit run app.py
```

Upload `Al_MetaQTL_Database_Input.xlsx` in the sidebar. Alternatively place it in `data/Al_MetaQTL_Database_Input.xlsx` for local/private use. **Do not publish unpublished research workbooks in public GitHub repositories.**

## Deploy to Streamlit Cloud

Push `app.py`, `requirements.txt`, and `README.md` to your GitHub repository. Select `main` and `app.py` in Streamlit Community Cloud. Upload the workbook after the app starts.

## Scientific linkage rules

- MQTL -> MQTL_QTLs -> QTLs and Projected_QTLs via explicit IDs.
- MQTL -> MQTL_Genes -> Genes via explicit IDs.
- Gene -> Expression -> Expression_Datasets; Gene -> Protein_Features and Promoter_Sites.
- Gene -> Enrichment_Genes -> Enrichment (gene-set-level evidence).
- Gene -> qPCR_Summary via gene_id, with clearly labeled unverified target-name matches -> qPCR_Values and Primers.
- MQTL/QTL marker names -> Markers; chromosome and study context require verification.
- Source IDs -> Sources, and study IDs -> Studies.

No fabricated gene links, meta-QTL statistics, functional validation, or data for other elements. Missing links are reported as missing.
