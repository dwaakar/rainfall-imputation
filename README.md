
Rainfall Imputation — Figures & DOCX Summary

Generate nine manuscript‑ready visualizations and a DOCX summary for the rainfall imputation pipeline.

Features
- Missingness matrices
- Heatmaps
- Treemap
- Waffle charts
- Density strips
- Flow diagram
- Tile calendar
- Adjacency correlation matrix
- Elevation grid map

Usage
python Day10.py --outdir outputs_all --year_for_matrix 2020 --window Full

Inputs
- cv_results_metrics_all.csv
- trend_preservation_all.csv
- Station_deets.csv
- daily rf1990-2023.csv

Outputs
- PNG figures in outputs_all/figs
- DOCX summary in outputs_all/docx


