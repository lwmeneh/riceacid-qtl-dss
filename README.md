# RiceAcid-QTL DSS 🌾

A research-focused, mobile-responsive Streamlit decision-support dashboard for published rice aluminium-tolerance QTLs, meta-QTLs, candidate genes and supporting molecular evidence.

## Run locally

```bash
python -m venv .venv
# Windows: .venv\Scripts\activate
# macOS/Linux: source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

The default dataset is `data/Al_MetaQTL_Database_Input.xlsx`. You can upload a replacement `.xlsx` file from the sidebar with the same schema.

## Deploy to GitHub and Streamlit Community Cloud

1. Create a new GitHub repository named `riceacid-qtl-dss`.
2. Upload `app.py`, `requirements.txt`, `.streamlit/config.toml`, and the `data/` folder.
3. On https://share.streamlit.io, connect your GitHub account, choose the repository, set the entrypoint to `app.py`, and deploy.
4. Open the resulting HTTPS URL on a phone. It is a responsive web app, not a native Android/iOS app.

**Data privacy:** The bundled workbook will be visible to anyone with access to a public repository. If your dataset is unpublished, proprietary, or confidential, do not publish the workbook publicly. Use a private repository and an appropriately secured hosting environment instead. The sample app has no login or authorization layer.

## Supported views

- Dashboard: dataset counts and chromosome distribution
- Meta-QTL Explorer: published intervals, constituent QTLs, linked genes
- Original QTL Explorer: source statistics, original positions and scales
- Candidate Genes: annotation search and transparent illustrative evidence sorting
- Functional Evidence: qPCR, expression, promoter motifs, protein predictions, enrichment and primers
- Literature & Provenance: original studies, source references, data notes
- Data Quality & Export: missing data, duplicate ID checks, downloadable CSV tables

## Scientific limitations

This is a **read-only evidence explorer**, not a validated meta-QTL inference engine, genomic prediction model, or breeding recommendation algorithm. It does not calculate new consensus QTLs, infer causal genes, perform genome assembly liftover, or establish functional validation. The heuristic gene evidence score is illustrative only and must not be interpreted as a statistically validated rank. Existing source annotations and units are preserved. In particular, PVE fraction vs percentage scales and reversed original CI endpoints must be checked before downstream analyses. Reference genome assembly identifiers may be absent.

## Future development

1. PostgreSQL with normalized foreign keys, user accounts, audit trail and curated import approvals.
2. Genome assembly harmonization, marker liftover and JBrowse 2.
3. Benchmark meta-QTL analysis in R with reproducible model diagnostics.
4. Evidence-weighted candidate prioritization with documented uncertainty.
5. Separate Cd, Ca, Mg, Mn and Fe study modules and cross-trait integration.
6. Native mobile app only if offline data collection or device-specific features become necessary.

## Attribution

The workbook retains source citations and identifiers. Cite the underlying publications, including Jaiswal et al. (2024), *Environmental and Experimental Botany* 228, 106030, DOI: 10.1016/j.envexpbot.2024.106030, where appropriate. Confirm permissions before public redistribution of compiled supplementary material.
