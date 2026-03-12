Rainfall Imputation — Figures & DOCX Summary
Generate nine manuscript‑ready, non‑line/non‑bar/non‑scatter visualizations and a DOCX summary to document and communicate results from a rainfall imputation pipeline. This script ingests cross‑validation metrics, trend‑preservation diagnostics, station metadata, and daily rainfall, then emits publication‑quality figures plus a step‑by‑step Word summary.

Highlights

9 specialty plots (tile heatmaps, treemap, waffle charts, density strips, flow diagram, tile calendar, adjacency matrix, raster‑style grid map).
Elevation‑band analysis with representative station selection.
Fully headless plotting (matplotlib Agg) for CI/servers.
One‑click DOCX narrative summary.



Table of Contents

#features
#visuals-produced
#input-files
#outputs
#installation
#quick-start
#command-line-options
#data-schemas
#how-it-works-pipeline
#project-structure
#troubleshooting
#faq
#references--inspiration
#license


Features

Robust I/O: Safe CSV loading with warnings for missing inputs.
Elevation awareness: Automatic elevation‑banding (0–800, 800–1500, 1500–2500, ≥2500 m) and per‑band representative station selection (by completeness).
Headless rendering: Uses matplotlib’s "Agg" backend to run on servers/CI without a display.
Flexible windows: Predefined analysis windows: Full (1990–2023), Last20, Last10, Last5.
Reproducible outputs: Deterministic figure file names and folders; DOCX summary captures pipeline decisions and metrics.


Visuals Produced

Missingness Pattern Matrix (tile map) — per representative station and year
Method × Missingness × GapType Heatmaps — tile heatmaps for metrics (NSE, RMSE)
Elevation‑Band Treemap — area sized by evaluated pairs, labeled with band winner
Trend‑Sign Agreement Waffle Charts — per method
Monthly Totals Density Strips (“distribution glyphs”) — per representative station
Process Flow Diagram — tile/arrow schematic of the CV pipeline
Monsoon Tile Calendar (Rain/No‑Rain) — for selected months and year
Adjacency Matrix of Station Correlations — |r| among stations (sorted)
Elevation Raster‑Style Grid Map — stations colored by elevation band; reps outlined


All figures are saved as .png at 300 DPI into outputs_all/figs/ (or your chosen --outdir).


Input Files
Place these files in the working directory before running the script:

cv_results_metrics_all.csv — Cross‑validation results and metrics
trend_preservation_all.csv — Trend diagnostics (e.g., Sen’s slope sign agreement)
Station_deets.csv — Station metadata (name, lat, lon, elevation)
daily rf1990-2023.csv — Daily rainfall time series (wide format)
(optional) station_completeness_by_window.csv — If missing, it will be computed
(optional) imputed_monthly_totals.csv — For future extensions


Outputs
Generated under your --outdir (default outputs_all/):
outputs_all/
├─ figs/
│  ├─ missingness_matrix_<Band>_<Station>_<Year>.png
│  ├─ heatmap_NSE_random.png
│  ├─ heatmap_NSE_block.png
│  ├─ heatmap_RMSE_random.png
│  ├─ heatmap_RMSE_block.png
│  ├─ treemap_elev_band.png
│  ├─ waffle_trend_<Method>.png
│  ├─ density_strip_<Band>_<Station>.png
│  ├─ flow_diagram.png
│  ├─ tile_calendar_<Band>_<Station>_<Year>.png
│  └─ adj_corr_matrix.png
├─ docx/
│  └─ code_step_by_step_summary.docx
└─ station_completeness_by_window.csv   # (computed if absent)


Installation

Python: 3.8+ recommended

Install dependencies (example with pip):
Shellpip install numpy pandas matplotlib scipy python-docxShow more lines

If you use conda:
Shellconda create -n rainfigs python=3.11 numpy pandas matplotlib scipy -yconda activate rainfigspip install python-docxShow more lines


Quick Start

Put all required CSV files in the repo root (or the directory where you’ll run the script).
Run:

Shellpython Day10.py --outdir outputs_all --year_for_matrix 2020 --window FullShow more lines

Explore figures under outputs_all/figs/ and the DOCX in outputs_all/docx/.


Command‑Line Options
--outdir OUTDIR           Output directory (default: outputs_all)
--year_for_matrix YEAR    Year for missingness matrix & tile calendar (default: 2020)
--window {Full,Last20,Last10,Last5}
                          Evaluation window for picking rep stations (default: Full)


Data Schemas
1) daily rf1990-2023.csv

Columns: Time, <Station1>, <Station2>, …
Time: daily timestamps (YYYY-MM-DD or parseable datetime)
Values: daily rainfall (mm); non‑negative; missing as blank/NaN

Example:
CSVTime,ST001,ST002,ST0031990-01-01,0.0,1.2,1990-01-02,,0.0,5.4...Show more lines
2) Station_deets.csv

Columns: Stations, Latitude, Longitude, Elevation (meters)

CSVStations,Latitude,Longitude,ElevationST001,27.71,85.32,1350ST002,28.12,84.01,720...Show more lines
3) cv_results_metrics_all.csv
Minimum useful columns (others are preserved if present):

Station, Method, GapType (random/block),
MissingFrac (0.1–0.9),
Metrics such as NSE, RMSE, MAE, R2, PBIAS, Precision, Recall,
ValidPairs (count of scored gaps)

4) trend_preservation_all.csv

Station, Method, SenSignAgree (0–1), plus other trend fields if available.

5) station_completeness_by_window.csv (optional)
If absent, it is computed from rainfall + metadata using predefined windows:

Full: 1990–2023, Last20: 2003–2023, Last10: 2013–2023, Last5: 2018–2023.


How It Works (Pipeline)


Load & QC

Reads all CSVs, converts Time to datetime.
Computes station completeness by window if missing.
Assigns elevation bands (0–800, 800–1500, 1500–2500, ≥2500 m).



Representative Stations

For each elevation band, selects the station with ≥90% completeness (in chosen window) and max completeness (tie‑break by elevation then name).



Figures

Missingness matrices for reps in the selected --year_for_matrix.
Heatmaps of metrics across Method × MissingFrac for each GapType.
Treemap sized by ValidPairs per band, labeled by winning method (highest median metric).
Waffle charts (trend sign agreement per method).
Density strips of monthly totals (observed).
Flow diagram of the CV pipeline.
Monsoon tile calendar (rain/no‑rain days, months 6–9).
Adjacency matrix of |Pearson r| among up to 40 stations in window.
Elevation grid map (raster‑style tiles; representative stations outlined).



DOCX Summary

Writes code_step_by_step_summary.docx summarizing inputs/outputs, QC, neighbor selection, masks, methods (LR/MLR/IDW/ElevIDW/NR/SA/OK/CokrigElev/XGB), metrics (RMSE/MAE/R²/NSE/PBIAS/Precision/Recall), and trend analysis (Kendall τ, Sen’s slope, sign agreement).




Project Structure
.
├─ Day10.py                    # main script
├─ daily rf1990-2023.csv       # input
├─ Station_deets.csv           # input
├─ cv_results_metrics_all.csv  # input
├─ trend_preservation_all.csv  # input
└─ outputs_all/                # generated by the script
   ├─ figs/
   └─ docx/


Troubleshooting

“Missing file” warnings: Ensure all required CSVs are in the working directory; names must match exactly.
No representatives found: The default completeness threshold is 90%. Either improve input data coverage or change the code/threshold to be more permissive.
Empty figures: Check that time windows and station IDs match across files. For correlation matrices, you need sufficient overlapping days (the script uses simple checks like ≥30 rows before computing r).
Headless environments: The script forces matplotlib.use("Agg"); no GUI required.
Fonts/labels: If labels overlap in dense plots, consider tweaking figure sizes in the code blocks where the figsize is defined.


FAQ
Q: Can I use a different year for the tile calendar and missingness matrices?
Yes — pass --year_for_matrix 2018 (for example).
Q: How are elevation bands defined?
0–800 m, 800–1500 m, 1500–2500 m, ≥2500 m; unknown if elevation missing.
Q: Which metric determines the “winning method” in the treemap?
By default median NSE across relevant records; you can change the metric argument in fig_elev_band_treemap(...).
Q: How are representative stations selected?
One per elevation band in the chosen window, ≥90% completeness, then by highest completeness, then elevation (desc), then station name (asc).

References & Inspiration

Galicia case: Missingness matrices; CV flow — Vidal‑Paz et al., 2023
Johor basin: Performance heat maps; PDFs — Sa’adi et al., 2023
Susurluk basin: Elevation stratification; homogeneity — Hırca & Türkkan, 2024

(These are cited as conceptual inspiration embedded in comments within the script.)
