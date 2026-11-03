# RiceAcid-QTL DSS — Visual Edition v6

An improved Streamlit application with a branded sidebar, original rice imagery, colorful research metric tiles, molecular symbols, guided workflow, and evidence tabs. Scientific records are read from the original Excel workbook without altering their values.

## Local setup
1. Copy `app.py`, `requirements.txt`, and the **entire `assets` directory** into your GitHub Desktop repository.
2. Place `Al_MetaQTL_Database_Input.xlsx` in `data/`.
3. Run `python -m pip install -r requirements.txt` then `python -m streamlit run app.py`.

The `data/` folder is gitignored by default. For Streamlit Cloud, configure separate hosted data storage or upload a workbook through the sidebar.

Images are decorative illustrations, not scientific evidence. The source data and gene relationships remain unchanged.
