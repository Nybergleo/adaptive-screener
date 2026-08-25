# Adaptive Screener

Start in BQuant with:

```python
00_adaptive_screener.ipynb
```

This app runs a reusable adaptive Bloomberg BQL screen for:

- reported LTM EBIT / operating margin
- reported LTM EBITA margin
- quarterly EBIT-margin stability, calculated as `std(oper_margin(fpt=Q, fpo=range(...)))`
- blended 12-month forward P/E
- reported LTM EBIT in the selected currency

The default geography is Western Europe. The current universe source is Bloomberg active primary equities, with geography set to Western Europe or North America.

Generated files are written under:

```text
data_cache/margin_stability_screener_v02/
```

Preset files are saved as JSON here:

```text
data_cache/margin_stability_screener_v02/margin_stability_presets.json
```

In an interactive BQuant notebook project, saving, loading, and deleting presets updates that local JSON file. In a published BQAP app, runtime file writes may not persist after the app is closed, so shared/default presets should be included in the project before publishing.

Main v2 files:

```text
00_adaptive_screener.ipynb
adaptive_screener_app.py
```

## Local Excel Generator

To try the local Excel workbook generator, run:

```powershell
python adaptive_excel_generator.py
```

Use a Python environment with `openpyxl` installed. The non-zipped `.xls` export itself is plain XML, but the optional `.xlsx` export uses `openpyxl`.

Then open:

```text
http://127.0.0.1:8765
```

The default download is a non-zipped `.xls` XML workbook. A modern `.xlsx` download is also available as a secondary option.
