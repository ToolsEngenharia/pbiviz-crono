import powerbi from "powerbi-visuals-api";
import { GanttTask, GanttViewModel, DEFAULT_SETTINGS } from "./interfaces";
import DataView = powerbi.DataView;
import IVisualHost = powerbi.extensibility.visual.IVisualHost;

export function parseDataView(dataView: DataView, host: IVisualHost): GanttViewModel | null {
  if (!dataView?.table?.rows || dataView.table.rows.length === 0) return null;

  const columns = dataView.table.columns;
  const rows    = dataView.table.rows;

  const idx = {
    taskName:     columns.findIndex(c => c.roles?.taskName),
    taskId:       columns.findIndex(c => c.roles?.taskId),
    wbs:          columns.findIndex(c => c.roles?.wbs),
    outlineLevel: columns.findIndex(c => c.roles?.outlineLevel),
    plannedStart: columns.findIndex(c => c.roles?.plannedStart),
    plannedEnd:   columns.findIndex(c => c.roles?.plannedEnd),
    baselineStart:columns.findIndex(c => c.roles?.baselineStart),
    baselineEnd:  columns.findIndex(c => c.roles?.baselineEnd),
    progress:     columns.findIndex(c => c.roles?.progress),
    dependencies: columns.findIndex(c => c.roles?.dependencies),
    isMilestone:  columns.findIndex(c => c.roles?.isMilestone),
  };

  const tasks: GanttTask[] = [];

  // ── DIAGNOSTIC LOG — remove after debugging ──
  console.log("[GANTT-PARSER] column indices:", idx);
  if (rows.length > 0) {
    const r0 = rows[0];
    console.log("[GANTT-PARSER] first row raw values:", {
      taskId: idx.taskId >= 0 ? r0[idx.taskId] : "N/A (idx=-1)",
      taskIdType: idx.taskId >= 0 ? typeof r0[idx.taskId] : "N/A",
      wbs: idx.wbs >= 0 ? r0[idx.wbs] : "N/A",
      deps: idx.dependencies >= 0 ? r0[idx.dependencies] : "N/A",
    });
  }

  rows.forEach((row, rowIndex) => {
    const name         = idx.taskName >= 0     ? String(row[idx.taskName]     ?? `Tarefa ${rowIndex + 1}`) : `Tarefa ${rowIndex + 1}`;
    const taskIdRaw    = idx.taskId >= 0       ? String(row[idx.taskId]       ?? "") : "";
    const taskId       = normalizeId(taskIdRaw);
    const wbs          = idx.wbs >= 0          ? String(row[idx.wbs]          ?? "") : "";
    const outlineLevel = idx.outlineLevel >= 0 ? Number(row[idx.outlineLevel] ?? 1)  : 1;
    const plannedStart = idx.plannedStart >= 0 ? parseDate(row[idx.plannedStart]) : null;
    const plannedEnd   = idx.plannedEnd >= 0   ? parseDate(row[idx.plannedEnd])   : null;

    if (!plannedStart || !plannedEnd) return;

    // Milestone: field can be boolean true, string "true"/"1", or number 1
    const msRaw    = idx.isMilestone >= 0 ? row[idx.isMilestone] : null;
    // Power BI boolean columns can arrive as: boolean true/false,
    // number 1/0, or string "True"/"False" (capital T/F from DAX).
    // Any non-empty string that is NOT explicitly "false"/"0" must NOT be treated as true.
    let isMilestone = false;
    if (msRaw === true  || msRaw === 1)  isMilestone = true;
    else if (msRaw === false || msRaw === 0) isMilestone = false;
    else if (typeof msRaw === "string") {
      const ms = msRaw.trim().toLowerCase();
      isMilestone = ms === "true" || ms === "1";
    }

    const baselineStart  = idx.baselineStart >= 0 ? parseDate(row[idx.baselineStart]) : null;
    const baselineEnd    = idx.baselineEnd >= 0   ? parseDate(row[idx.baselineEnd])   : null;
    const rawProgress    = idx.progress >= 0 ? Number(row[idx.progress] ?? 0) : 0;
    const progress       = clamp(rawProgress > 0 && rawProgress <= 1 ? rawProgress * 100 : rawProgress, 0, 100);
    const depsRaw        = idx.dependencies >= 0 ? String(row[idx.dependencies] ?? "") : "";
    const dependencies   = depsRaw ? parseDependencies(depsRaw) : [];

    const selectionId = host
      .createSelectionIdBuilder()
      .withTable(dataView.table, rowIndex)
      .createSelectionId();

    tasks.push({
      id: taskId || wbs || name,
      taskId,
      name, wbs, outlineLevel,
      isSummary: false,
      isMilestone,
      plannedStart, plannedEnd,
      baselineStart, baselineEnd,
      progress, dependencies, selectionId,
      isVisible: true,
    });
  });

  if (tasks.length === 0) return null;

  tasks.sort((a, b) => compareWbs(a.wbs, b.wbs));

  tasks.forEach((task, i) => {
    const next = tasks[i + 1];
    task.isSummary = !task.isMilestone && !!next && next.outlineLevel > task.outlineLevel;
  });

  const allDates = tasks.flatMap(t =>
    [t.plannedStart, t.plannedEnd, t.baselineStart, t.baselineEnd]
  ).filter(Boolean) as Date[];

  let minDate = new Date(Math.min(...allDates.map(d => d.getTime())));
  let maxDate = new Date(Math.max(...allDates.map(d => d.getTime())));
  // Padding using local-time arithmetic
  minDate = new Date(minDate.getFullYear(), minDate.getMonth(), minDate.getDate() - 3);
  maxDate = new Date(maxDate.getFullYear(), maxDate.getMonth(), maxDate.getDate() + 5);

  // Always include today in the domain so the "Hoje" line is always visible
  const today = new Date();
  const todayNorm = new Date(today.getFullYear(), today.getMonth(), today.getDate());
  if (todayNorm < minDate) minDate = new Date(todayNorm.getFullYear(), todayNorm.getMonth(), todayNorm.getDate() - 2);
  if (todayNorm > maxDate) maxDate = new Date(todayNorm.getFullYear(), todayNorm.getMonth(), todayNorm.getDate() + 2);

  return { tasks, settings: { ...DEFAULT_SETTINGS }, minDate, maxDate };
}

function parseDate(value: powerbi.PrimitiveValue): Date | null {
  if (value === null || value === undefined || value === "") return null;

  let d: Date;
  if (value instanceof Date) {
    d = value;
  } else if (typeof value === "number") {
    d = new Date(value);
  } else {
    // Try to parse string — "2026-08-03T00:00:00" or "2026-08-03"
    d = new Date(value as string);
  }

  if (isNaN(d.getTime())) return null;

  // Normalize to LOCAL midnight so d3 time scales (which use local time)
  // position bars correctly regardless of server/UTC timezone in the raw value.
  return new Date(d.getFullYear(), d.getMonth(), d.getDate(), 0, 0, 0, 0);
}

/**
 * Normalize an ID string: trim, remove trailing .0/.00, collapse whitespace.
 * Handles Power BI sending numbers as "929.0" or " 929 " etc.
 */
function normalizeId(s: string): string {
  return s.trim().replace(/\.0+$/, "");
}

function clamp(val: number, min: number, max: number): number {
  return Math.max(min, Math.min(max, val));
}

/**
 * Parse dependency string from MS Project / P6 / etc.
 * Handles: "1000;873", "938,941", "1000IT", "500TT+5d", "200FF-3", "100SS+10d"
 * Strips link-type suffixes (IT/TT/II/TI/SS/SF/FS/FF) and optional lag (+/-Nd).
 * Splits by ";" or ",".
 */
function parseDependencies(raw: string): string[] {
  // Split by semicolon or comma
  return raw.split(/[;,]/)
    .map(d => {
      // Remove whitespace
      let s = d.trim();
      // Strip optional lag: +5d, -3d, +10, -2, +5 dias, etc.
      s = s.replace(/[+\-]\s*\d+\s*[dD]?\w*$/, "");
      // Strip link-type suffix (case-insensitive): IT, TT, II, TI, SS, SF, FS, FF
      s = s.replace(/\s*(IT|TT|II|TI|SS|SF|FS|FF)\s*$/i, "");
      return normalizeId(s);
    })
    .filter(Boolean);
}

function compareWbs(a: string, b: string): number {
  if (!a && !b) return 0;
  if (!a) return 1;
  if (!b) return -1;
  const sa = a.split(".").map(s => parseInt(s, 10) || 0);
  const sb = b.split(".").map(s => parseInt(s, 10) || 0);
  const len = Math.max(sa.length, sb.length);
  for (let i = 0; i < len; i++) {
    const d = (sa[i] ?? -1) - (sb[i] ?? -1);
    if (d !== 0) return d;
  }
  return 0;
}