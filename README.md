# RiceAcid-QTL DSS · Reference Design v7

Streamlit research explorer matching the supplied dashboard design. KPI values are computed from the Excel workbook, not the illustrative reference image.

## Install locally (Windows)

1. Copy `app.py`, `requirements.txt`, and the complete `assets/` directory into your GitHub Desktop repository.
2. Keep `data/Al_MetaQTL_Database_Input.xlsx` on your computer. Add `data/` to `.gitignore` to prevent publishing private data.
3. Run `python -m pip install -r requirements.txt` and `python -m streamlit run app.py`.
4. Visit `http://localhost:8501`.

The hosted Streamlit Cloud app cannot read your Windows-local workbook. Use the optional upload there or configure a secure hosted database.

The header is a decorative visual based on the supplied design reference. It is not scientific evidence. The live data views below it use workbook records. No additional genes, loci, measurements, or links are fabricated.
