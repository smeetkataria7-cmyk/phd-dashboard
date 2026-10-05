# PHD Dataset Dashboard

Self-contained HTML dashboard built from the **first 7 sheets** of an Excel workbook
(sugarcane trial data: variety master, field metadata, three satellite sources, weather, Brix/Pol ground truth).

Open `dashboard.html` in a browser (needs internet only for the Chart.js CDN).

## Rebuild for a new dataset

```bash
pip install openpyxl
python build_dashboard.py "NEW DATASET.xlsx"            # -> dashboard.html
python build_dashboard.py "NEW DATASET.xlsx" -n 7 -o out.html
```

`build_dashboard.py` reads the workbook and injects the data into `dashboard_template.html`
(keep both files in the same folder).

## What is in the dashboard

- **Overview** – KPIs and domain charts (field map, juice quality, vegetation index vs crop age, climate, variety master).
- **Relations** – three relations: Satellite A ↔ weather, Satellite B ↔ weather, Satellite C ↔ ground truth
  (correlation heatmaps, strongest relationships, scatter explorers, index-vs-rainfall timeline).
- **Prediction** – predict Brix (or Pol / CCS / Purity) and harvest time from satellite indices (+ optional rainfall/weather):
  - 5 selectable ML models: Ridge, k-NN, Random Forest, Gradient Boosting, Kernel Ridge (RBF), with a side-by-side comparison
  - "Satellite + weather" or "Satellite only (no weather)" input modes, optional spectral bands
  - plot-wise and variety-wide harvest-time tables and ripening curves
  - an input form: type current index / weather values to get approximate Brix and optimum harvest date
- **One tab per sheet** – chart builder, column profile, searchable data table with CSV export.
- Auto / Light / Dark theme toggle.

## Notes

- `dashboard.html` embeds the data from the workbook's first 7 sheets.
- Models run in the browser (pure JavaScript). Validation is leave-one-plot-out and leave-one-date-out;
  with ~84 lab samples on 12 dates, treat predictions as approximate.
