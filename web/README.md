# DaqIdea Graph Panel Web

Browser-based version of the desktop graph panel. Excel workbooks are parsed locally in the browser; they are not uploaded to a server.

## Run

```powershell
cd web
npm install
npm run dev
```

Open the local URL printed by Vite. Use **Open dump** to select one or more `.xlsx` or `.xlsm` DaqIdea workbooks.

## Included

- Grouped file/signal table with editable names, visibility, axis assignment, live values, cursor A/B values, and delta.
- Ctrl/Cmd-click or Shift-click to select multiple signals and toggle their visibility; each workbook heading has its own show/hide control.
- Drag `.xlsx`/`.xlsm` files onto the app to load them, and drag signal rows to reorder within a workbook.
- The info button lists keyboard shortcuts for opening, exporting, cursor placement, sample dots, graph modes, and axis fitting.
- Interactive Plotly overlay or stacked diagrams with per-signal Y ranges.
- Mouse/trackpad zoom, axis-specific zoom modes, fit commands, draggable A/B cursors, live cursor readout, and sample dots.
- PNG graph export with the source workbook name as the image title, plus visible-signal CSV export.
