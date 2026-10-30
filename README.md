# DaqIdea Graph Panel

A local desktop viewer for DaqIdea Excel data dumps. It reads the active worksheet, uses its `Time` column as seconds, and exposes numeric columns as plot signals. Every imported signal is plotted automatically.

## Run

```powershell
python -m pip install -r Desktop/requirements.txt
python Desktop/daqidea_plotter.py
```

## Web Version

The browser-based app is in [`web/`](web/README.md). To run it locally:

```powershell
cd web
npm install
npm run dev
```

## Graph panel features

- Open one or more `.xlsx` or `.xlsm` dumps; workbooks load in the background and all numeric signals appear in the color-coded list and graph automatically. Signal colors are generated dynamically, with no eight-signal palette limit.
- The header ribbon is removed; `Open dump` is at the bottom of the Signals & Values panel. Save and CSV export are grouped in the file-actions menu.
- Use each signal's `Show` checkbox to hide or restore it, and its axis selector to assign it to the left or right Y axis. `Select all` and `Unselect all` control the whole list.
- A single compact, horizontally scrollable signal table combines color-coded channels, visibility, Y-axis assignment, live values, cursor A/B readings, and delta. Double-click a signal name to edit its display name.
- Signals are grouped beneath collapsible workbook-name headings, including when a single dump is loaded.
- Use the zoom icon menu for area, X-only, or Y-only drag zoom and fit commands. Scroll or use a trackpad gesture to zoom around the pointer; hold Shift to zoom X only or Ctrl to zoom Y only. The Matplotlib navigation toolbar uses dark-tinted Home, Back, Forward, Pan, Zoom, Subplots, and Save icons.
- `Separate diagrams` stacks each visible signal in the same graph window with an independent, autoscaled Y axis and shared time axis; toggle back to the overlay with `Overlay graph`.
- Use the sample-points icon to toggle a signal-colored dot at every raw sample.
- Place cursors A and B by selecting a cursor and clicking the graph; grab a cursor line and drag it to reposition it. The measurement table shows nearest-sample values and the B - A difference.
- The toolbar shows the cursor time difference, Δt = B - A, in the selected seconds or milliseconds unit.
- Enable `Live cursor` to turn the pointer into a crosshair that follows the graph and updates each visible signal's live value and time readout.
- With Live cursor active, select a signal row to show its nearest sample time and value in a popup on the graph.
- Display time in seconds or milliseconds; saved graph images include the loaded workbook name and an overlay legend. Save as PNG/PDF/SVG or export plotted data to CSV from the file-actions menu.

Large signals are downsampled for display to keep redraws responsive while retaining peaks; cursor readings and CSV exports use the original samples. Signal names are taken directly from workbook headers. Numeric values stored as text are converted automatically; non-numeric columns are skipped.