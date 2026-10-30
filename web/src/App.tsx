import { useEffect, useMemo, useRef, useState } from "react";
import {
  Activity,
  Check,
  ChevronDown,
  CircleDot,
  Download,
  Eye,
  EyeOff,
  FileImage,
  FolderOpen,
  GripVertical,
  Info,
  Layers2,
  Save,
  ScanLine,
  X,
  ZoomIn,
} from "lucide-react";
import * as XLSX from "@e965/xlsx";
import Plotly from "plotly.js-dist-min";
import type { Config, Data, Layout, PlotlyHTMLElement } from "plotly.js";

type AxisSide = "left" | "right";
type CursorName = "A" | "B";

interface Signal {
  id: string;
  fileId: string;
  fileName: string;
  name: string;
  times: number[];
  values: number[];
  color: string;
  axis: AxisSide;
  visible: boolean;
}

interface Dataset {
  id: string;
  name: string;
  signals: Signal[];
}

interface PlotPoint {
  x?: number | string;
  y?: number;
  curveNumber?: number;
  data?: { uid?: string };
}

interface PlotEvent extends Record<string, unknown> {
  points?: PlotPoint[];
}

interface PlotEventTarget extends HTMLElement {
  on: (name: string, callback: (event: PlotEvent) => void) => void;
  removeListener: (name: string, callback: (event: PlotEvent) => void) => void;
}

type PlotYAxisRef = NonNullable<Layout["annotations"]>[number]["yref"];

const COLORS = [
  "#50d8b0",
  "#ff9d72",
  "#f5c75e",
  "#76b8f2",
  "#df82d4",
  "#c2db6a",
  "#ed7682",
  "#64d3cf",
];

const PLOT_BG = "#101715";
const GRID = "#26332f";
const TEXT = "#dbe7e1";

function signalColor(index: number): string {
  if (index < COLORS.length) return COLORS[index];
  return `hsl(${(index * 137.508) % 360} 70% 64%)`;
}

function plotTraceUid(signalId: string, suffix = ""): string {
  return `${signalId.replace(/[^a-zA-Z0-9_-]/g, "_")}${suffix}`;
}

function numeric(value: unknown): number | null {
  if (value == null || typeof value === "boolean" || value === "") return null;
  const number = typeof value === "number" ? value : Number(value);
  return Number.isFinite(number) ? number : null;
}

async function parseWorkbook(file: File, colorOffset: number): Promise<Dataset> {
  const workbook = XLSX.read(await file.arrayBuffer(), { type: "array", cellDates: false });
  const sheetName = workbook.SheetNames[0];
  if (!sheetName) throw new Error("Workbook has no worksheets.");
  const rows = XLSX.utils.sheet_to_json<unknown[]>(workbook.Sheets[sheetName], {
    header: 1,
    defval: null,
    raw: true,
  });
  if (rows.length < 2) throw new Error("Worksheet has no signal samples.");

  const headers = rows[0].map((value) => String(value ?? "").trim());
  const timeIndex = headers.findIndex((value) => ["time", "time [s]", "time(s)"].includes(value.toLowerCase()));
  const resolvedTimeIndex = timeIndex < 0 ? 0 : timeIndex;
  const columns = headers
    .map((name, index) => ({ name, index }))
    .filter(({ name, index }) => name && index !== resolvedTimeIndex);
  const times: number[] = [];
  const samples = new Map<number, number[]>();
  for (const column of columns) samples.set(column.index, []);

  for (const row of rows.slice(1)) {
    const time = numeric(row[resolvedTimeIndex]);
    if (time == null) continue;
    times.push(time);
    for (const column of columns) {
      samples.get(column.index)!.push(numeric(row[column.index]) ?? Number.NaN);
    }
  }
  if (times.length === 0) throw new Error("No numeric time samples were found.");

  const fileId = crypto.randomUUID();
  const signals = columns.flatMap(({ name, index }, signalIndex) => {
    const values = samples.get(index)!;
    if (!values.some(Number.isFinite)) return [];
    return [{
      id: `${fileId}:${index}`,
      fileId,
      fileName: file.name,
      name,
      times,
      values,
      color: signalColor(colorOffset + signalIndex),
      axis: "left" as const,
      visible: true,
    }];
  });
  return { id: fileId, name: file.name, signals };
}

function nearestIndex(signal: Signal, time: number): number {
  let low = 0;
  let high = signal.times.length;
  while (low < high) {
    const middle = (low + high) >>> 1;
    if (signal.times[middle] < time) low = middle + 1;
    else high = middle;
  }
  if (low <= 0) return 0;
  if (low >= signal.times.length) return signal.times.length - 1;
  return Math.abs(signal.times[low] - time) < Math.abs(signal.times[low - 1] - time) ? low : low - 1;
}

function sampleAt(signal: Signal, time: number): number | null {
  const value = signal.values[nearestIndex(signal, time)];
  return Number.isFinite(value) ? value : null;
}

function displaySamples(signal: Signal, pointLimit: number): { x: number[]; y: number[] } {
  if (signal.times.length <= pointLimit) return { x: signal.times, y: signal.values };
  const bucketSize = Math.max(1, Math.ceil(signal.times.length / Math.max(1, pointLimit / 2)));
  const x: number[] = [];
  const y: number[] = [];
  for (let start = 0; start < signal.times.length; start += bucketSize) {
    const end = Math.min(signal.times.length, start + bucketSize);
    let minIndex = -1;
    let maxIndex = -1;
    for (let index = start; index < end; index += 1) {
      if (!Number.isFinite(signal.values[index])) continue;
      if (minIndex < 0 || signal.values[index] < signal.values[minIndex]) minIndex = index;
      if (maxIndex < 0 || signal.values[index] > signal.values[maxIndex]) maxIndex = index;
    }
    if (minIndex < 0) {
      x.push(signal.times[start]);
      y.push(Number.NaN);
      continue;
    }
    for (const index of [...new Set([start, end - 1, minIndex, maxIndex])].sort((a, b) => a - b)) {
      if (Number.isFinite(signal.values[index])) {
        x.push(signal.times[index]);
        y.push(signal.values[index]);
      }
    }
  }
  return { x, y };
}

function fmt(value: number | null | undefined): string {
  return value == null || !Number.isFinite(value) ? "" : Number(value.toPrecision(7)).toString();
}

function csvCell(value: string | number): string {
  const text = String(value);
  return /[",\r\n]/.test(text) ? `"${text.replaceAll('"', '""')}"` : text;
}

export default function App() {
  const [datasets, setDatasets] = useState<Dataset[]>([]);
  const [selectedId, setSelectedId] = useState<string | null>(null);
  const [selectedIds, setSelectedIds] = useState<string[]>([]);
  const [timeUnit, setTimeUnit] = useState<"s" | "ms">("s");
  const [stacked, setStacked] = useState(false);
  const [samplesVisible, setSamplesVisible] = useState(false);
  const [cursorA, setCursorA] = useState<number | null>(null);
  const [cursorB, setCursorB] = useState<number | null>(null);
  const [cursorMode, setCursorMode] = useState<CursorName | null>(null);
  const [liveCursor, setLiveCursor] = useState(false);
  const [liveTime, setLiveTime] = useState<number | null>(null);
  const [zoomAxis, setZoomAxis] = useState<"both" | "x" | "y">("both");
  const [status, setStatus] = useState("Open a DaqIdea workbook to begin.");
  const [loading, setLoading] = useState(false);
  const [editingId, setEditingId] = useState<string | null>(null);
  const [editingName, setEditingName] = useState("");
  const [shortcutsOpen, setShortcutsOpen] = useState(false);
  const [fileDropActive, setFileDropActive] = useState(false);
  const [draggedSignalId, setDraggedSignalId] = useState<string | null>(null);
  const [dropTargetSignalId, setDropTargetSignalId] = useState<string | null>(null);
  const fileInput = useRef<HTMLInputElement>(null);
  const plotElement = useRef<HTMLDivElement>(null);
  const signals = useMemo(() => datasets.flatMap((dataset) => dataset.signals), [datasets]);
  const visibleSignals = useMemo(() => signals.filter((signal) => signal.visible), [signals]);
  const selectedSignal = signals.find((signal) => signal.id === selectedId) ?? null;
  const scale = timeUnit === "ms" ? 1000 : 1;
  const deltaTime = cursorA != null && cursorB != null ? (cursorB - cursorA) * scale : null;

  function updateSignal(id: string, update: Partial<Signal>) {
    setDatasets((current) => current.map((dataset) => ({
      ...dataset,
      signals: dataset.signals.map((signal) => signal.id === id ? { ...signal, ...update } : signal),
    })));
  }

  function selectSignal(id: string, event: { shiftKey: boolean; ctrlKey: boolean; metaKey: boolean }) {
    let nextSelection: string[];
    if (event.shiftKey && selectedId) {
      const start = signals.findIndex((signal) => signal.id === selectedId);
      const end = signals.findIndex((signal) => signal.id === id);
      nextSelection = start < 0 || end < 0
        ? [id]
        : signals.slice(Math.min(start, end), Math.max(start, end) + 1).map((signal) => signal.id);
    } else if (event.ctrlKey || event.metaKey) {
      nextSelection = selectedIds.includes(id)
        ? selectedIds.filter((selected) => selected !== id)
        : [...selectedIds, id];
      if (nextSelection.length === 0) nextSelection = [id];
    } else {
      nextSelection = [id];
    }
    setSelectedId(id);
    setSelectedIds(nextSelection);
  }

  function setSelectedVisible(visible: boolean) {
    if (selectedIds.length === 0) {
      setAllVisible(visible);
      return;
    }
    const selected = new Set(selectedIds);
    setDatasets((current) => current.map((dataset) => ({
      ...dataset,
      signals: dataset.signals.map((signal) => selected.has(signal.id) ? { ...signal, visible } : signal),
    })));
  }

  function reorderSignal(sourceId: string, targetId: string) {
    if (sourceId === targetId) return;
    setDatasets((current) => current.map((dataset) => {
      const from = dataset.signals.findIndex((signal) => signal.id === sourceId);
      const to = dataset.signals.findIndex((signal) => signal.id === targetId);
      if (from < 0 || to < 0) return dataset;
      const reordered = [...dataset.signals];
      const [moved] = reordered.splice(from, 1);
      reordered.splice(to, 0, moved);
      return { ...dataset, signals: reordered };
    }));
  }

  async function openFiles(files: FileList | File[]) {
    const selectedFiles = Array.from(files).filter((file) => /\.(xlsx|xlsm)$/i.test(file.name));
    if (selectedFiles.length === 0) return;
    setLoading(true);
    setStatus(`Loading ${selectedFiles.length} workbook${selectedFiles.length === 1 ? "" : "s"}...`);
    try {
      let colorOffset = signals.length;
      const parsed: Dataset[] = [];
      for (const file of selectedFiles) {
        parsed.push(await parseWorkbook(file, colorOffset));
        colorOffset += parsed.at(-1)!.signals.length;
        await new Promise<void>((resolve) => window.setTimeout(resolve, 0));
      }
      setDatasets((current) => [...current, ...parsed]);
      const firstSignalId = parsed.flatMap((dataset) => dataset.signals)[0]?.id ?? null;
      setSelectedId((current) => current ?? firstSignalId);
      setSelectedIds((current) => current.length ? current : firstSignalId ? [firstSignalId] : []);
      setStatus(`Loaded ${parsed.length} dump${parsed.length === 1 ? "" : "s"}.`);
    } catch (error) {
      setStatus(error instanceof Error ? error.message : "Could not read the selected workbook.");
    } finally {
      setLoading(false);
      if (fileInput.current) fileInput.current.value = "";
    }
  }

  function setAllVisible(visible: boolean) {
    setDatasets((current) => current.map((dataset) => ({
      ...dataset,
      signals: dataset.signals.map((signal) => ({ ...signal, visible })),
    })));
  }

  useEffect(() => {
    const onKeyDown = (event: KeyboardEvent) => {
      const target = event.target as HTMLElement | null;
      const editing = target?.matches("input, textarea, select, [contenteditable='true']");
      if (editing && event.key !== "Escape") return;
      const command = event.ctrlKey || event.metaKey;
      const key = event.key.toLowerCase();
      if (command && key === "o") {
        event.preventDefault();
        fileInput.current?.click();
      } else if (command && key === "s") {
        event.preventDefault();
        saveImage();
      } else if (command && key === "e") {
        event.preventDefault();
        exportCsv();
      } else if (command && key === "a") {
        event.preventDefault();
        setSelectedIds(signals.map((signal) => signal.id));
        setSelectedId(signals[0]?.id ?? null);
      } else if (!command && key === "i") {
        setShortcutsOpen(true);
      } else if (!command && key === "a") {
        setCursorMode("A");
      } else if (!command && key === "b") {
        setCursorMode("B");
      } else if (!command && key === "l") {
        setLiveCursor((current) => !current);
        setLiveTime(null);
      } else if (!command && key === "p") {
        setSamplesVisible((current) => !current);
      } else if (!command && key === "g") {
        setStacked((current) => !current);
      } else if (!command && key === "1") {
        fitAxes("all");
      } else if (!command && key === "2") {
        fitAxes("x");
      } else if (!command && key === "3") {
        fitAxes("y");
      } else if (key === "escape") {
        setCursorMode(null);
        setShortcutsOpen(false);
      }
    };
    window.addEventListener("keydown", onKeyDown);
    return () => window.removeEventListener("keydown", onKeyDown);
  }, [signals, setSelectedIds, setSelectedId]);

  function onFileDragOver(event: React.DragEvent<HTMLElement>) {
    if (!Array.from(event.dataTransfer.types).includes("Files")) return;
    event.preventDefault();
    setFileDropActive(true);
  }

  function onFileDrop(event: React.DragEvent<HTMLElement>) {
    if (!Array.from(event.dataTransfer.types).includes("Files")) return;
    event.preventDefault();
    setFileDropActive(false);
    void openFiles(event.dataTransfer.files);
  }

  function commitRename(id: string) {
    const name = editingName.trim();
    if (name) updateSignal(id, { name });
    setEditingId(null);
  }

  function exportCsv() {
    if (visibleSignals.length === 0) {
      setStatus("Show at least one signal before exporting.");
      return;
    }
    const headers = visibleSignals.flatMap((signal) => [`${signal.fileName} / ${signal.name} time (s)`, `${signal.fileName} / ${signal.name} value`]);
    const rows: (string | number)[][] = [headers];
    const count = Math.max(...visibleSignals.map((signal) => signal.times.length));
    for (let index = 0; index < count; index += 1) {
      rows.push(visibleSignals.flatMap((signal) => [
        signal.times[index] ?? "",
        Number.isFinite(signal.values[index]) ? signal.values[index] : "",
      ]));
    }
    const csv = rows.map((row) => row.map((value) => csvCell(value)).join(",")).join("\r\n");
    const url = URL.createObjectURL(new Blob(["\ufeff", csv], { type: "text/csv;charset=utf-8" }));
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = "daqidea-signals.csv";
    anchor.click();
    URL.revokeObjectURL(url);
    setStatus("CSV exported.");
  }

  function saveImage() {
    if (!plotElement.current) return;
    const plot = plotElement.current as unknown as PlotlyHTMLElement;
    const title = datasets.map((dataset) => dataset.name).join(" · ") || "DaqIdea Graph Panel";
    void (async () => {
      try {
        await Plotly.relayout(plot, {
          title: { text: title, x: 0.5, xanchor: "center", font: { color: TEXT, size: 14 } },
          showlegend: !stacked && visibleSignals.length > 0,
        });
        await Plotly.downloadImage(plot, {
          format: "png",
          filename: "daqidea-graph",
          width: 1600,
          height: Math.max(900, stacked ? visibleSignals.length * 260 : 900),
          scale: 1,
        });
      } finally {
        await Plotly.relayout(plot, { title: { text: "" }, showlegend: false });
      }
    })();
  }

  function fitAxes(axis: "all" | "x" | "y") {
    const plot = plotElement.current as unknown as PlotlyHTMLElement | null;
    if (!plot) return;
    const updates: Record<string, unknown> = {};
    if (axis !== "y") updates["xaxis.autorange"] = true;
    if (axis !== "x") {
      if (stacked) visibleSignals.forEach((_signal, index) => { updates[`yaxis${index ? index + 1 : ""}.autorange`] = true; });
      else {
        updates["yaxis.autorange"] = true;
        updates["yaxis2.autorange"] = true;
      }
    }
    void Plotly.relayout(plot, updates);
  }

  function placeCursorAt(event: React.MouseEvent<HTMLDivElement>) {
    if (!cursorMode || !plotElement.current) return;
    const graph = plotElement.current as unknown as {
      _fullLayout?: { xaxis?: { _offset: number; p2d: (pixel: number) => number } };
    };
    const axis = graph._fullLayout?.xaxis;
    if (!axis) return;
    const bounds = plotElement.current.getBoundingClientRect();
    const displayedTime = axis.p2d(event.clientX - bounds.left - axis._offset);
    if (!Number.isFinite(displayedTime)) return;
    const seconds = displayedTime / scale;
    if (cursorMode === "A") setCursorA(seconds);
    else setCursorB(seconds);
    setCursorMode(null);
  }

  const traces: Data[] = [];
  const layoutAxes: Record<string, unknown> = {};
  const pointsPerSignal = Math.max(180, Math.min(2400, Math.floor((window.innerWidth - 430) * 1.6 / Math.max(1, visibleSignals.length))));

  if (stacked) {
    visibleSignals.forEach((signal, index) => {
      const axisIndex = index + 1;
      const yref = axisIndex === 1 ? "y" : `y${axisIndex}`;
      const ykey = axisIndex === 1 ? "yaxis" : `yaxis${axisIndex}`;
      const domainBottom = 1 - axisIndex / visibleSignals.length;
      const domainTop = 1 - (axisIndex - 1) / visibleSignals.length;
      layoutAxes[ykey] = {
        domain: [domainBottom + 0.012, domainTop - 0.012],
        anchor: "x",
        title: { text: "" },
        gridcolor: GRID,
        zerolinecolor: GRID,
        tickfont: { color: "#91a69d", size: 10 },
        automargin: true,
        fixedrange: zoomAxis === "x",
      };
      const points = displaySamples(signal, pointsPerSignal);
      traces.push({
        type: "scattergl",
        mode: "lines",
        uid: plotTraceUid(signal.id),
        name: signal.name,
        x: points.x.map((time) => time * scale),
        y: points.y,
        yaxis: yref,
        line: { color: signal.color, width: 1.35 },
        hovertemplate: `${signal.name}<br>t=%{x:.6g} s<br>value=%{y:.7g}<extra></extra>`,
      });
      if (samplesVisible) traces.push({
        type: "scattergl",
        mode: "markers",
        uid: plotTraceUid(signal.id, "_samples"),
        name: `${signal.name} samples`,
        x: signal.times.map((time) => time * scale),
        y: signal.values,
        yaxis: yref,
        marker: { color: signal.color, size: 5 },
        hoverinfo: "skip",
        showlegend: false,
      });
    });
  } else {
    for (const signal of visibleSignals) {
      const points = displaySamples(signal, pointsPerSignal);
      const yref = signal.axis === "right" ? "y2" : "y";
      traces.push({
        type: "scattergl",
        mode: "lines",
        uid: plotTraceUid(signal.id),
        name: signal.name,
        x: points.x.map((time) => time * scale),
        y: points.y,
        yaxis: yref,
        line: { color: signal.color, width: 1.35 },
        hovertemplate: `${signal.name}<br>t=%{x:.6g} s<br>value=%{y:.7g}<extra></extra>`,
      });
      if (samplesVisible) traces.push({
        type: "scattergl",
        mode: "markers",
        uid: plotTraceUid(signal.id, "_samples"),
        name: `${signal.name} samples`,
        x: signal.times.map((time) => time * scale),
        y: signal.values,
        yaxis: yref,
        marker: { color: signal.color, size: 5 },
        hoverinfo: "skip",
        showlegend: false,
      });
    }
    layoutAxes.yaxis = { title: { text: "Left axis", font: { color: "#50d8b0", size: 11 } }, gridcolor: GRID, zerolinecolor: GRID, tickfont: { color: "#91a69d", size: 10 }, automargin: true, fixedrange: zoomAxis === "x" };
    layoutAxes.yaxis2 = { title: { text: "Right axis", font: { color: "#ff9d72", size: 11 } }, overlaying: "y", side: "right", gridcolor: "transparent", tickfont: { color: "#91a69d", size: 10 }, automargin: true, fixedrange: zoomAxis === "x" };
  }

  const shapes: NonNullable<Layout["shapes"]> = [];
  if (cursorA != null) shapes.push({ type: "line", xref: "x", yref: "paper", x0: cursorA * scale, x1: cursorA * scale, y0: 0, y1: 1, line: { color: "#f4eee1", dash: "dash", width: 1.2 } });
  if (cursorB != null) shapes.push({ type: "line", xref: "x", yref: "paper", x0: cursorB * scale, x1: cursorB * scale, y0: 0, y1: 1, line: { color: "#f3c75f", dash: "dash", width: 1.2 } });
  if (liveCursor && liveTime != null) shapes.push({ type: "line", xref: "x", yref: "paper", x0: liveTime * scale, x1: liveTime * scale, y0: 0, y1: 1, line: { color: "#85c9ff", dash: "dot", width: 1.3 } });

  const annotations: NonNullable<Layout["annotations"]> = [];
  if (liveCursor && liveTime != null && selectedSignal?.visible) {
    const index = nearestIndex(selectedSignal, liveTime);
    const value = selectedSignal.values[index];
    if (Number.isFinite(value)) {
      const selectedAxisIndex = visibleSignals.findIndex((signal) => signal.id === selectedSignal.id);
      const selectedYRef = (stacked
        ? selectedAxisIndex === 0 ? "y" : `y${selectedAxisIndex + 1}`
        : selectedSignal.axis === "right" ? "y2" : "y") as PlotYAxisRef;
      annotations.push({
        x: selectedSignal.times[index] * scale,
        y: value,
        xref: "x",
        yref: selectedYRef,
        text: `${selectedSignal.name}<br>t = ${fmt(selectedSignal.times[index] * scale)} ${timeUnit}<br>value = ${fmt(value)}`,
        showarrow: true,
        arrowhead: 2,
        ax: 22,
        ay: -28,
        bgcolor: "#18221f",
        bordercolor: selectedSignal.color,
        borderwidth: 1,
        font: { color: TEXT, size: 11 },
      });
    }
  }

  const layout: Partial<Layout> = {
    autosize: true,
    height: stacked ? Math.max(420, visibleSignals.length * 210) : undefined,
    paper_bgcolor: PLOT_BG,
    plot_bgcolor: PLOT_BG,
    font: { family: "IBM Plex Sans, sans-serif", color: TEXT, size: 11 },
    margin: { l: 62, r: stacked ? 22 : 66, t: stacked ? 16 : 24, b: 48 },
    xaxis: {
      title: { text: `Time [${timeUnit}]`, font: { color: "#c5d3cf", size: 11 } },
      anchor: (stacked && visibleSignals.length > 1 ? `y${visibleSignals.length}` : "y") as NonNullable<Layout["xaxis"]>["anchor"],
      gridcolor: GRID,
      zerolinecolor: GRID,
      tickfont: { color: "#91a69d", size: 10 },
      automargin: true,
      fixedrange: zoomAxis === "y",
    },
    hovermode: "closest",
    showlegend: false,
    dragmode: "zoom",
    shapes,
    annotations,
    uirevision: stacked ? "daqidea-stacked" : "daqidea-overlay",
    ...layoutAxes,
  };

  const config: Partial<Config> = {
    responsive: true,
    scrollZoom: true,
    displayModeBar: false,
    displaylogo: false,
    edits: { shapePosition: true },
  };

  const interactionRef = useRef({
    onClick: (_event: PlotEvent) => {},
    onRelayout: (_event: Record<string, unknown>) => {},
  });
  const plotReady = useRef(false);
  const chartRevision = JSON.stringify([
    stacked,
    samplesVisible,
    timeUnit,
    zoomAxis,
    signals.map((signal) => [signal.id, signal.name, signal.axis, signal.visible, signal.color]),
  ]);

  interactionRef.current.onClick = (event) => {
    const x = numeric(event.points?.[0]?.x);
    if (x == null || !cursorMode) return;
    if (cursorMode === "A") setCursorA(x / scale);
    else setCursorB(x / scale);
    setCursorMode(null);
  };
  interactionRef.current.onRelayout = (event) => {
    const aShapeIndex = 0;
    const bShapeIndex = cursorA != null ? 1 : 0;
    const nextA = numeric(event[`shapes[${aShapeIndex}].x0`]);
    const nextB = numeric(event[`shapes[${bShapeIndex}].x0`]);
    if (cursorA != null && nextA != null) setCursorA(nextA / scale);
    if (cursorB != null && nextB != null) setCursorB(nextB / scale);
  };

  useEffect(() => {
    const element = plotElement.current;
    if (!element) return;
    const target = element as unknown as PlotEventTarget;
    let active = true;
    const click = (event: PlotEvent) => interactionRef.current.onClick(event);
    const relayout = (event: Record<string, unknown>) => interactionRef.current.onRelayout(event);

    void Plotly.react(element as unknown as PlotlyHTMLElement, traces, layout as Layout, config as Config).then(() => {
      if (!active) return;
      plotReady.current = true;
      target.on("plotly_click", click);
      target.on("plotly_relayout", relayout);
      void Plotly.relayout(target as unknown as PlotlyHTMLElement, { shapes, annotations });
    });
    return () => {
      active = false;
      plotReady.current = false;
      target.removeListener("plotly_click", click);
      target.removeListener("plotly_relayout", relayout);
    };
  }, [chartRevision]);

  useEffect(() => {
    const element = plotElement.current;
    if (!element || !liveCursor) return;
    const updatePointer = (event: PointerEvent) => {
      const graph = element as unknown as {
        _fullLayout?: {
          xaxis?: { _offset: number; _length: number; p2d: (pixel: number) => number };
          _size?: { t: number; h: number };
        };
      };
      const layout = graph._fullLayout;
      const axis = layout?.xaxis;
      const size = layout?._size;
      if (!axis || !size) return;
      const bounds = element.getBoundingClientRect();
      const pixelX = event.clientX - bounds.left - axis._offset;
      const plotTop = bounds.top + size.t;
      const plotBottom = plotTop + size.h;
      if (
        pixelX < 0 || pixelX > axis._length ||
        event.clientY < plotTop || event.clientY > plotBottom
      ) {
        setLiveTime(null);
        return;
      }
      const displayedTime = axis.p2d(pixelX);
      setLiveTime(Number.isFinite(displayedTime) ? displayedTime / scale : null);
    };
    const leavePlot = () => setLiveTime(null);
    element.addEventListener("pointermove", updatePointer, true);
    element.addEventListener("pointerleave", leavePlot);
    return () => {
      element.removeEventListener("pointermove", updatePointer, true);
      element.removeEventListener("pointerleave", leavePlot);
    };
  }, [liveCursor, scale, stacked, visibleSignals.length]);

  useEffect(() => {
    if (!plotReady.current || !plotElement.current) return;
    void Plotly.relayout(plotElement.current as unknown as PlotlyHTMLElement, { shapes, annotations });
  }, [cursorA, cursorB, liveCursor, liveTime, selectedId, timeUnit, stacked, visibleSignals]);

  const bottomDelta = deltaTime == null ? "Δt: --" : `Δt: ${fmt(deltaTime)} ${timeUnit}`;
  const bottomTime = liveTime == null ? "t: --" : `t: ${fmt(liveTime * scale)} ${timeUnit}`;

  return (
    <main
      className={`app-shell ${fileDropActive ? "is-file-drop-active" : ""}`}
      onDragEnter={onFileDragOver}
      onDragOver={onFileDragOver}
      onDragLeave={(event) => {
        if (!event.currentTarget.contains(event.relatedTarget as Node | null)) setFileDropActive(false);
      }}
      onDrop={onFileDrop}
    >
      <aside className="sidebar">
        <div className="sidebar-heading">
          <span className="eyebrow">DAQIDEA</span>
          <h1>Signals &amp; values</h1>
        </div>

        <div className="signal-table-wrap">
          <table className="signal-table">
            <thead>
              <tr><th>Signal</th><th>Y</th><th>On</th><th>Now</th><th>A</th><th>B</th><th>Δ</th></tr>
            </thead>
            <tbody>
              {datasets.map((dataset) => (
                <DatasetRows
                  key={dataset.id}
                  dataset={dataset}
                  selectedIds={selectedIds}
                  liveTime={liveTime}
                  cursors={{ A: cursorA, B: cursorB }}
                  onSelect={selectSignal}
                  onUpdate={updateSignal}
                  onToggleDataset={(visible) => setDatasets((current) => current.map((item) => item.id === dataset.id ? { ...item, signals: item.signals.map((signal) => ({ ...signal, visible })) } : item))}
                  draggedSignalId={draggedSignalId}
                  dropTargetSignalId={dropTargetSignalId}
                  setDraggedSignalId={setDraggedSignalId}
                  setDropTargetSignalId={setDropTargetSignalId}
                  onReorder={reorderSignal}
                  onRenameStart={(signal) => { setEditingId(signal.id); setEditingName(signal.name); }}
                  editingId={editingId}
                  editingName={editingName}
                  setEditingName={setEditingName}
                  onRename={commitRename}
                  onRenameCancel={() => setEditingId(null)}
                />
              ))}
            </tbody>
          </table>
          {datasets.length === 0 && (
            <div className="empty-table">
              <Activity size={17} />
              <span>No signals loaded</span>
            </div>
          )}
        </div>

        <div className="sidebar-actions">
          <button className="icon-button" title={selectedIds.length ? "Show selected signals" : "Show all signals"} aria-label={selectedIds.length ? "Show selected signals" : "Show all signals"} onClick={() => setSelectedVisible(true)}><Check size={16} /></button>
          <button className="icon-button" title={selectedIds.length ? "Hide selected signals" : "Hide all signals"} aria-label={selectedIds.length ? "Hide selected signals" : "Hide all signals"} onClick={() => setSelectedVisible(false)}><EyeOff size={16} /></button>
          <span className="action-spacer" />
          <button className="open-files-button" onClick={() => fileInput.current?.click()}>
            {loading ? <span className="spinner" /> : <FolderOpen size={17} />}
            {loading ? "Loading…" : "Open dump"}
          </button>
          <input ref={fileInput} type="file" accept=".xlsx,.xlsm" multiple hidden onChange={(event) => void openFiles(event.currentTarget.files ?? [])} />
        </div>
      </aside>

      <section className="workspace">
        <div className="toolbar">
          <div className="toolbar-group">
            <label className="unit-control"><span>TIME</span><select value={timeUnit} onChange={(event) => setTimeUnit(event.target.value as "s" | "ms")}><option value="s">s</option><option value="ms">ms</option></select></label>
            <div className="toolbar-divider" />
            <details className="tool-menu">
              <summary className="icon-button" title="Zoom and fit options" aria-label="Zoom and fit options"><ZoomIn size={17} /></summary>
              <div className="menu-panel">
                <button onClick={() => setZoomAxis("both")}>Area zoom <span>Both axes</span></button>
                <button onClick={() => setZoomAxis("x")}>X-axis zoom <span>Time only</span></button>
                <button onClick={() => setZoomAxis("y")}>Y-axis zoom <span>Value only</span></button>
                <div className="menu-rule" />
                <button onClick={() => fitAxes("all")}>Fit all axes</button>
                <button onClick={() => fitAxes("x")}>Fit X axis</button>
                <button onClick={() => fitAxes("y")}>Fit Y axes</button>
              </div>
            </details>
            <details className="tool-menu">
              <summary className="icon-button" title="Cursor tools" aria-label="Cursor tools"><ScanLine size={17} /></summary>
              <div className="menu-panel">
                <button onClick={() => setCursorMode("A")}>Place cursor A</button>
                <button onClick={() => setCursorMode("B")}>Place cursor B</button>
                <button onClick={() => { setLiveCursor((value) => !value); setLiveTime(null); }}>Toggle live cursor</button>
                <button onClick={() => { setCursorA(null); setCursorB(null); setCursorMode(null); }}>Clear cursors</button>
              </div>
            </details>
            <button className={`icon-button ${samplesVisible ? "is-active" : ""}`} title="Toggle sample points" aria-label="Toggle sample points" onClick={() => setSamplesVisible((value) => !value)}><CircleDot size={17} /></button>
            <button className={`icon-button ${stacked ? "is-active" : ""}`} title="Toggle separate diagrams" aria-label="Toggle separate diagrams" onClick={() => setStacked((value) => !value)}><Layers2 size={17} /></button>
          </div>
          <div className="toolbar-group">
            <span className="signal-count">{visibleSignals.length} / {signals.length} signals</span>
            <button className="icon-button info-button" title="Keyboard shortcuts and help" aria-label="Keyboard shortcuts and help" onClick={() => setShortcutsOpen(true)}><Info size={17} /></button>
            <details className="tool-menu">
              <summary className="icon-button" title="Save and export" aria-label="Save and export"><Save size={17} /></summary>
              <div className="menu-panel menu-panel-right">
                <button onClick={saveImage}><FileImage size={15} /> Save graph image</button>
                <button onClick={exportCsv}><Download size={15} /> Export visible signals</button>
              </div>
            </details>
          </div>
        </div>

        <div className="plot-shell">
          {signals.length === 0 ? (
            <div className="plot-empty">
              <div className="empty-glyph"><Activity size={21} /></div>
              <strong>Ready for a data dump</strong>
              <span>Open one or more DaqIdea Excel files to plot their signals.</span>
            </div>
          ) : (
            <div
              ref={plotElement}
              className={`plotly-host ${liveCursor ? "is-live-cursor" : ""}`}
              onClick={placeCursorAt}
            />
          )}
        </div>
        <footer className="status-bar">
          <span className="status-message">{status}</span>
          <span>{bottomTime}</span>
          <span>{bottomDelta}</span>
          {cursorMode && <span className="status-mode">Click graph to place cursor {cursorMode}</span>}
        </footer>
      </section>
      {fileDropActive && (
        <div className="file-drop-overlay" aria-live="polite">
          <div className="file-drop-card"><FolderOpen size={24} /><strong>Drop DaqIdea workbooks</strong><span>Release to add .xlsx or .xlsm dumps</span></div>
        </div>
      )}
      {shortcutsOpen && (
        <div className="dialog-backdrop" onMouseDown={(event) => { if (event.target === event.currentTarget) setShortcutsOpen(false); }}>
          <section className="shortcuts-dialog" role="dialog" aria-modal="true" aria-labelledby="shortcuts-title">
            <header><div><span className="eyebrow">REFERENCE</span><h2 id="shortcuts-title">Keyboard shortcuts</h2></div><button className="icon-button" aria-label="Close shortcuts" onClick={() => setShortcutsOpen(false)}><X size={17} /></button></header>
            <div className="shortcut-list">
              <Shortcut keys="Ctrl / Cmd + O" action="Open DaqIdea workbooks" />
              <Shortcut keys="Ctrl / Cmd + S" action="Save graph image" />
              <Shortcut keys="Ctrl / Cmd + E" action="Export visible signals to CSV" />
              <Shortcut keys="Ctrl / Cmd + A" action="Select all signals" />
              <Shortcut keys="A" action="Place cursor A" />
              <Shortcut keys="B" action="Place cursor B" />
              <Shortcut keys="L" action="Toggle live cursor" />
              <Shortcut keys="P" action="Toggle sample dots" />
              <Shortcut keys="G" action="Toggle stacked diagrams" />
              <Shortcut keys="1 / 2 / 3" action="Fit all / X / Y axes" />
              <Shortcut keys="Esc" action="Cancel cursor placement or close this panel" />
            </div>
            <footer>Use Ctrl/Cmd-click or Shift-click to select multiple signals. Drag a signal row to reorder it within its file.</footer>
          </section>
        </div>
      )}
    </main>
  );
}

function Shortcut({ keys, action }: { keys: string; action: string }) {
  return <div className="shortcut-row"><kbd>{keys}</kbd><span>{action}</span></div>;
}

interface DatasetRowsProps {
  dataset: Dataset;
  selectedIds: string[];
  liveTime: number | null;
  cursors: { A: number | null; B: number | null };
  onSelect: (id: string, event: { shiftKey: boolean; ctrlKey: boolean; metaKey: boolean }) => void;
  onUpdate: (id: string, update: Partial<Signal>) => void;
  onToggleDataset: (visible: boolean) => void;
  draggedSignalId: string | null;
  dropTargetSignalId: string | null;
  setDraggedSignalId: (id: string | null) => void;
  setDropTargetSignalId: (id: string | null) => void;
  onReorder: (sourceId: string, targetId: string) => void;
  editingId: string | null;
  editingName: string;
  setEditingName: (name: string) => void;
  onRenameStart: (signal: Signal) => void;
  onRename: (id: string) => void;
  onRenameCancel: () => void;
}

function DatasetRows({
  dataset,
  selectedIds,
  liveTime,
  cursors,
  onSelect,
  onUpdate,
  onToggleDataset,
  draggedSignalId,
  dropTargetSignalId,
  setDraggedSignalId,
  setDropTargetSignalId,
  onReorder,
  editingId,
  editingName,
  setEditingName,
  onRenameStart,
  onRename,
  onRenameCancel,
}: DatasetRowsProps) {
  const [expanded, setExpanded] = useState(true);
  const allVisible = dataset.signals.length > 0 && dataset.signals.every((signal) => signal.visible);

  return <>
    <tr className="dataset-heading-row">
      <td colSpan={7}>
        <button className="dataset-heading" onClick={() => setExpanded((value) => !value)} aria-expanded={expanded}>
          <ChevronDown className={expanded ? "chevron" : "chevron is-collapsed"} size={13} />
          <span>{dataset.name}</span>
          <small>{dataset.signals.length}</small>
        </button>
        <button
          className="dataset-visibility"
          title={allVisible ? "Hide all signals in this dump" : "Show all signals in this dump"}
          aria-label={allVisible ? "Hide all signals in this dump" : "Show all signals in this dump"}
          onClick={(event) => { event.stopPropagation(); onToggleDataset(!allVisible); }}
        >{allVisible ? <Eye size={14} /> : <EyeOff size={14} />}</button>
      </td>
    </tr>
    {expanded && dataset.signals.map((signal) => {
      const liveValue = liveTime != null ? sampleAt(signal, liveTime) : null;
      const valueA = cursors.A != null ? sampleAt(signal, cursors.A) : null;
      const valueB = cursors.B != null ? sampleAt(signal, cursors.B) : null;
      const delta = valueA != null && valueB != null ? valueB - valueA : null;
      const editing = editingId === signal.id;
      return <tr
        key={signal.id}
        className={`signal-row ${selectedIds.includes(signal.id) ? "is-selected" : ""} ${signal.visible ? "" : "is-hidden"} ${draggedSignalId === signal.id ? "is-dragging" : ""} ${dropTargetSignalId === signal.id ? "is-drop-target" : ""}`}
        draggable
        onClick={(event) => onSelect(signal.id, event)}
        onDragStart={(event) => {
          event.dataTransfer.effectAllowed = "move";
          event.dataTransfer.setData("text/plain", signal.id);
          setDraggedSignalId(signal.id);
        }}
        onDragOver={(event) => {
          if (draggedSignalId && draggedSignalId !== signal.id) {
            event.preventDefault();
            event.dataTransfer.dropEffect = "move";
            setDropTargetSignalId(signal.id);
          }
        }}
        onDragLeave={() => { if (dropTargetSignalId === signal.id) setDropTargetSignalId(null); }}
        onDrop={(event) => {
          event.preventDefault();
          const sourceId = draggedSignalId ?? event.dataTransfer.getData("text/plain");
          if (sourceId) onReorder(sourceId, signal.id);
          setDraggedSignalId(null);
          setDropTargetSignalId(null);
        }}
        onDragEnd={() => { setDraggedSignalId(null); setDropTargetSignalId(null); }}
      >
        <td className="signal-name-cell">
          <GripVertical className="signal-drag-handle" size={13} aria-hidden="true" />
          <span className="signal-swatch" style={{ backgroundColor: signal.color }} />
          {editing ? (
            <input
              className="rename-input"
              value={editingName}
              autoFocus
              onChange={(event) => setEditingName(event.target.value)}
              onBlur={() => onRename(signal.id)}
              onClick={(event) => event.stopPropagation()}
              onKeyDown={(event) => {
                if (event.key === "Enter") onRename(signal.id);
                if (event.key === "Escape") onRenameCancel();
              }}
            />
          ) : (
            <span className="signal-name" title={signal.name} onDoubleClick={(event) => { event.stopPropagation(); onRenameStart(signal); }}>{signal.name}</span>
          )}
        </td>
        <td className="axis-cell" title={`Use ${signal.axis === "left" ? "right" : "left"} Y axis`} onClick={(event) => { event.stopPropagation(); onUpdate(signal.id, { axis: signal.axis === "left" ? "right" : "left" }); }}>
          {signal.axis === "left" ? "L" : "R"}
        </td>
        <td className="visibility-cell">
          <button className="cell-icon" title={signal.visible ? "Hide signal" : "Show signal"} aria-label={signal.visible ? "Hide signal" : "Show signal"} onClick={(event) => { event.stopPropagation(); onUpdate(signal.id, { visible: !signal.visible }); }}>
            {signal.visible ? <Eye size={14} /> : <EyeOff size={14} />}
          </button>
        </td>
        <td>{fmt(liveValue)}</td>
        <td>{fmt(valueA)}</td>
        <td>{fmt(valueB)}</td>
        <td>{fmt(delta)}</td>
      </tr>;
    })}
  </>;
}