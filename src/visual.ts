import powerbi from "powerbi-visuals-api";
import * as d3 from "d3";
import { FormattingSettingsService } from "powerbi-visuals-utils-formattingmodel";
import { GanttFormattingSettings } from "./formattingSettings";
import { parseDataView } from "./dataParser";
import { GanttViewModel, GanttTask } from "./interfaces";

import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions      = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisual                  = powerbi.extensibility.visual.IVisual;
import IVisualHost              = powerbi.extensibility.visual.IVisualHost;
import ISelectionManager        = powerbi.extensibility.ISelectionManager;
import IVisualEventService      = powerbi.extensibility.IVisualEventService;

import "../style/visual.less";

// ── Zoom levels ───────────────────────────────────────────────────────────────
interface ZoomLevel {
  minPx:          number;
  label:          string;
  topInterval:    d3.CountableTimeInterval;
  topFmt:         (d: Date) => string;
  bottomInterval: d3.CountableTimeInterval;
  bottomFmt:      (d: Date, i: number, all: Date[]) => string;
}

const ZOOM_LEVELS: ZoomLevel[] = [
  {
    minPx: 0, label: "ANO / SEMESTRE",
    topInterval: d3.timeYear,
    topFmt: d => `FY${String(d.getFullYear()).slice(-2)}`,
    bottomInterval: d3.timeMonth.every(6)! as d3.CountableTimeInterval,
    bottomFmt: (d) => `S${d.getMonth() < 6 ? 1 : 2}`,
  },
  {
    minPx: 2, label: "ANO / MÊS",
    topInterval: d3.timeYear,
    topFmt: d => `FY${String(d.getFullYear()).slice(-2)}`,
    bottomInterval: d3.timeMonth,
    bottomFmt: (d) => d3.timeFormat("%b")(d),
  },
  {
    minPx: 12, label: "MÊS / DIA",
    topInterval: d3.timeMonth,
    topFmt: d => d3.timeFormat("%b %Y")(d),
    bottomInterval: d3.timeDay,
    bottomFmt: (d) => d3.timeFormat("%d")(d),
  },
];

function pickZoom(pxPerDay: number): ZoomLevel {
  let z = ZOOM_LEVELS[0];
  for (const level of ZOOM_LEVELS) {
    if (pxPerDay >= level.minPx) z = level;
  }
  return z;
}

// ── Generate ticks covering full domain (including start) ──────────────────
function fullTicks(
  scale: d3.ScaleTime<number, number>,
  interval: d3.CountableTimeInterval
): Date[] {
  const [d0, d1] = scale.domain() as Date[];
  // floor start to interval boundary
  const start = interval.floor(d0);
  const raw   = interval.range(start, d1);
  // ensure domain start is represented
  if (raw.length === 0 || raw[0].getTime() > d0.getTime()) {
    raw.unshift(start);
  }
  return raw;
}

// ─────────────────────────────────────────────────────────────────────────────
export class Visual implements IVisual {
  private host:             IVisualHost;
  private selectionManager: ISelectionManager;
  private events:           IVisualEventService;
  private fmtService:       FormattingSettingsService;
  private fmtSettings:      GanttFormattingSettings;

  // DOM structure
  // root
  //  └── gantt-root (flex row)
  //        ├── sidebar (flex col: sideHeader + sideBody)
  //        └── chart-area (flex col, flex:1)
  //              ├── chart-header-wrap (overflow:hidden, fixed height)
  //              │     └── headerSvg  (width = chartW, scrolled by JS)
  //              └── chart-body-wrap  (overflow:auto, flex:1)
  //                    └── bodySvg    (width = chartW, height = totalH)

  private root:           HTMLElement;
  private sidebar:        HTMLElement;
  private sideBody:       HTMLElement;
  private chartHeaderWrap:HTMLElement;
  private chartBodyWrap:  HTMLElement;
  private headerSvg:      d3.Selection<SVGSVGElement, unknown, null, undefined>;
  private bodySvg:        d3.Selection<SVGSVGElement, unknown, null, undefined>;
  private tooltip:        HTMLElement;
  private todayPill:      HTMLElement;
  private zoomBtns:       HTMLButtonElement[] = [];

  private viewModel: GanttViewModel | null = null;
  private collapsed  = new Set<string>();
  private statusFilter: string | null = null;  // null = show all
  private filterBtns: HTMLButtonElement[] = [];

  private pxPerDay   = 0.5;
  private firstLoad   = true;   // auto-fit only once
  // Fixed zoom steps — index 0 = fit-to-width (computed), 1 = ANO/MÊS, 2 = MÊS/DIA
  private readonly ZOOM_STEPS_FIXED = [5, 30];
  private zoomIdx    = 0;      // 0 = fit, 1+ = ZOOM_STEPS_FIXED[zoomIdx-1]

  /** Compute pxPerDay for fit-to-width mode (zoomIdx 0) */
  private computeFitPxPerDay(): number {
    if (!this.viewModel) return 0.5;
    const { minDate, maxDate } = this.viewModel;
    const daySpan  = Math.max(Math.ceil((maxDate.getTime() - minDate.getTime()) / 86400000), 1);
    const s        = this.fmtSettings;
    const wbsW     = (this.fmtSettings.layout.showWbs.value && this.showWbs)
      ? s.layout.wbsColumnWidth.value : 0;
    // Subtract scrollbar width (~12px) to avoid horizontal scroll when vertical scroll is present
    const scrollbarW = 12;
    const availW   = Math.max(this.vpWidth - wbsW - this.sidebarLabelW - scrollbarW, 60);
    return Math.max(0.1, availW / daySpan);
  }

  /** Resolve pxPerDay for the current zoomIdx */
  private resolvePxPerDay(): number {
    if (this.zoomIdx === 0) return this.computeFitPxPerDay();
    return this.ZOOM_STEPS_FIXED[this.zoomIdx - 1];
  }

  private vpWidth  = 800;
  private vpHeight = 400;

  // Sidebar state
  private showWbs      = true;
  private sidebarLabelW = 220;   // resizable (task name column width)
  private isDragging   = false;
  private dragStartX   = 0;
  private dragStartW   = 0;

  private readonly HEADER_H  = 52;
  private readonly TOP_ROW_H = 26;
  private readonly BOT_ROW_H = 26;
  private readonly BAR_R     = 4;
  private readonly INDENT_PX = 12;   // reduced indent per level

  constructor(options: VisualConstructorOptions) {
    this.host             = options.host;
    this.events           = options.host.eventService;
    this.selectionManager = options.host.createSelectionManager();
    this.fmtService       = new FormattingSettingsService();
    this.fmtSettings      = new GanttFormattingSettings();

    this.selectionManager.registerOnSelectCallback(() => this.syncHighlight());

    this.root = options.element;
    this.root.innerHTML = "";
    this.root.style.cssText = "position:relative;overflow:hidden;width:100%;height:100%;box-sizing:border-box;";

    // ── Outer flex row ───────────────────────────────────────────────────────
    const wrapper = document.createElement("div");
    wrapper.className = "gantt-root";
    this.root.appendChild(wrapper);

    // ── Sidebar ──────────────────────────────────────────────────────────────
    this.sidebar = document.createElement("div");
    this.sidebar.className = "gantt-sidebar";
    wrapper.appendChild(this.sidebar);

    const sideHeaderEl = document.createElement("div");
    sideHeaderEl.className = "sidebar-header";
    this.sidebar.appendChild(sideHeaderEl);

    this.sideBody = document.createElement("div");
    this.sideBody.className = "sidebar-body";
    this.sidebar.appendChild(this.sideBody);

    // ── Chart area (right side, flex col) ─────────────────────────────────────
    const chartArea = document.createElement("div");
    chartArea.style.cssText = "flex:1;display:flex;flex-direction:column;overflow:hidden;min-width:0;";
    wrapper.appendChild(chartArea);

    // Fixed header strip
    this.chartHeaderWrap = document.createElement("div");
    this.chartHeaderWrap.className = "chart-header-wrap";
    chartArea.appendChild(this.chartHeaderWrap);

    const hSvgEl = document.createElementNS("http://www.w3.org/2000/svg", "svg");
    hSvgEl.classList.add("gantt-header-svg");
    this.chartHeaderWrap.appendChild(hSvgEl);
    this.headerSvg = d3.select(hSvgEl);

    // Scrollable body
    this.chartBodyWrap = document.createElement("div");
    this.chartBodyWrap.className = "chart-body-wrap";
    chartArea.appendChild(this.chartBodyWrap);

    const bSvgEl = document.createElementNS("http://www.w3.org/2000/svg", "svg");
    bSvgEl.classList.add("gantt-body-svg");
    this.chartBodyWrap.appendChild(bSvgEl);
    this.bodySvg = d3.select(bSvgEl);

    // Sync horizontal scroll: body → header
    this.chartBodyWrap.addEventListener("scroll", () => {
      this.chartHeaderWrap.scrollLeft = this.chartBodyWrap.scrollLeft;
      this.sideBody.scrollTop         = this.chartBodyWrap.scrollTop;
    });
    this.sideBody.addEventListener("scroll", () => {
      this.chartBodyWrap.scrollTop = this.sideBody.scrollTop;
    });

    // Ctrl+wheel zoom
    this.chartBodyWrap.addEventListener("wheel", (ev: WheelEvent) => {
      if (!ev.ctrlKey) return;
      ev.preventDefault();
      this.adjustZoom(ev.deltaY < 0 ? 1 : -1);
    }, { passive: false });

    // Click background = clear selection
    this.bodySvg.on("click", () => {
      this.selectionManager.clear();
      this.syncHighlight();
    });

    // ── Status filter bar (collapsible) ─────────────────────────────────────
    const filterBar = document.createElement("div");
    filterBar.className = "filter-bar collapsed";
    this.root.appendChild(filterBar);

    const FILTER_OPTIONS: { key: string | null; label: string; cls: string; dot: string }[] = [
      { key: null,            label: "Todos",      cls: "all",        dot: "" },
      { key: "em_andamento",  label: "Andamento",  cls: "inprogress", dot: "#3b82f6" },
      { key: "atrasado",      label: "Atrasado",   cls: "late",       dot: "#ef4444" },
      { key: "adiantado",     label: "Adiantado",  cls: "early",      dot: "#10b981" },
      { key: "concluida",     label: "Concluída",  cls: "done",       dot: "#6b7280" },
      { key: "no_prazo",      label: "No prazo",   cls: "neutral",    dot: "#9ca3af" },
    ];

    const isFilterExpanded = () => !filterBar.classList.contains("collapsed");
    const collapseFilter   = () => filterBar.classList.add("collapsed");
    const expandFilter     = () => filterBar.classList.remove("collapsed");

    // Click outside → collapse
    document.addEventListener("mousedown", (ev: MouseEvent) => {
      if (isFilterExpanded() && !filterBar.contains(ev.target as Node)) {
        collapseFilter();
      }
    });

    FILTER_OPTIONS.forEach(opt => {
      const btn = document.createElement("button");
      btn.className = "filter-btn " + opt.cls + (opt.key === this.statusFilter ? " active" : "");
      if (opt.dot) {
        const dot = document.createElement("span");
        dot.className = "filter-dot";
        dot.style.background = opt.dot;
        btn.appendChild(dot);
      }
      const txt = document.createTextNode(opt.label);
      btn.appendChild(txt);

      // Chevron only on active button (visual hint that it's expandable)
      const chevron = document.createElement("span");
      chevron.className = "filter-chevron";
      chevron.innerHTML = "&#9660;";
      btn.appendChild(chevron);

      btn.addEventListener("click", (ev) => {
        ev.stopPropagation();
        if (btn.classList.contains("active")) {
          // Toggle expand/collapse
          if (isFilterExpanded()) collapseFilter();
          else expandFilter();
          return;
        }
        // Select this filter
        this.statusFilter = opt.key;
        this.filterBtns.forEach(b => b.classList.remove("active"));
        btn.classList.add("active");
        collapseFilter();
        if (this.viewModel) {
          this.computeVisibility();
          this.render(this.vpWidth, this.vpHeight);
        }
      });
      filterBar.appendChild(btn);
      this.filterBtns.push(btn);
    });

    // ── Zoom bar ─────────────────────────────────────────────────────────────
    const zoomBar = document.createElement("div");
    zoomBar.className = "zoom-bar";
    this.root.appendChild(zoomBar);

    // Toggle segmentado de zoom
    const zoomToggle = document.createElement("div");
    zoomToggle.className = "zoom-toggle";
    zoomBar.appendChild(zoomToggle);

    const ZOOM_TOGGLE_LABELS = ["Ano", "Mês", "Dia"];
    ZOOM_TOGGLE_LABELS.forEach((label, idx) => {
      const btn = document.createElement("button");
      btn.className = "zoom-toggle-btn" + (idx === this.zoomIdx ? " active" : "");
      btn.textContent = label;
      btn.title = ["Ajustar ao viewport (Ctrl+scroll)", "Zoom: ano / mês", "Zoom: mês / dia"][idx];
      btn.addEventListener("click", () => {
        if (this.zoomIdx === idx) return;
        this.zoomIdx  = idx;
        this.pxPerDay = this.resolvePxPerDay();
        this.updateZoomToggle();
        if (this.viewModel) {
          const scrollMax = this.chartBodyWrap.scrollWidth - this.chartBodyWrap.clientWidth;
          const ratio     = scrollMax > 0 ? this.chartBodyWrap.scrollLeft / scrollMax : 0;
          this.computeVisibility();
          this.render(this.vpWidth, this.vpHeight);
          const newScrollMax = this.chartBodyWrap.scrollWidth - this.chartBodyWrap.clientWidth;
          this.chartBodyWrap.scrollLeft   = Math.round(ratio * newScrollMax);
          this.chartHeaderWrap.scrollLeft = this.chartBodyWrap.scrollLeft;
        }
      });
      zoomToggle.appendChild(btn);
      this.zoomBtns.push(btn);
    });

    // Divider
    const divider = document.createElement("div");
    divider.className = "zoom-divider";
    zoomBar.appendChild(divider);

    const btnToday = document.createElement("button");
    btnToday.className = "zoom-today-btn";
    btnToday.innerHTML = "Hoje";
    btnToday.title = "Ir para hoje";
    btnToday.addEventListener("click", () => this.scrollToToday());
    zoomBar.appendChild(btnToday);

    // ── WBS toggle button (on sidebar header) ──────────────────────────────
    const wbsToggle = document.createElement("button");
    wbsToggle.className = "wbs-toggle-btn";
    wbsToggle.title = "Ocultar/exibir coluna WBS";
    wbsToggle.innerHTML = "WBS ✕";
    wbsToggle.addEventListener("click", () => {
      this.showWbs = !this.showWbs;
      wbsToggle.innerHTML = this.showWbs ? "WBS ✕" : "WBS ☰";
      wbsToggle.title = this.showWbs ? "Ocultar coluna WBS" : "Exibir coluna WBS";
      if (this.viewModel) this.render(this.vpWidth, this.vpHeight);
    });
    this.root.appendChild(wbsToggle);

    // ── Sidebar resize handle ────────────────────────────────────────────────
    const resizeHandle = document.createElement("div");
    resizeHandle.className = "sidebar-resize-handle";
    resizeHandle.title = "Arraste para redimensionar";
    this.root.appendChild(resizeHandle);

    resizeHandle.addEventListener("mousedown", (ev: MouseEvent) => {
      this.isDragging  = true;
      this.dragStartX  = ev.clientX;
      this.dragStartW  = this.sidebarLabelW;
      document.body.style.cursor = "col-resize";
      document.body.style.userSelect = "none";
      ev.preventDefault();
    });

    window.addEventListener("mousemove", (ev: MouseEvent) => {
      if (!this.isDragging) return;
      const delta = ev.clientX - this.dragStartX;
      const wbsW  = (this.fmtSettings.layout.showWbs.value && this.showWbs)
        ? this.fmtSettings.layout.wbsColumnWidth.value : 0;
      const minW  = 80;
      const maxW  = Math.max(this.vpWidth - wbsW - 200, 120);
      this.sidebarLabelW = Math.min(maxW, Math.max(minW, this.dragStartW + delta));
      if (this.viewModel) this.render(this.vpWidth, this.vpHeight);
    });

    window.addEventListener("mouseup", () => {
      if (!this.isDragging) return;
      this.isDragging = false;
      document.body.style.cursor = "";
      document.body.style.userSelect = "";
    });

    // ── Tooltip ───────────────────────────────────────────────────────────────
    this.tooltip = document.createElement("div");
    this.tooltip.className = "gantt-tooltip";
    this.root.appendChild(this.tooltip);

    // "Hoje" pill — HTML element, positioned absolutely relative to root
    // Updated on every render + scroll so it always aligns with bodySvg
    this.todayPill = document.createElement("div");
    this.todayPill.className = "today-pill";
    this.todayPill.textContent = "Hoje";
    this.root.appendChild(this.todayPill);
  }

  // ── Zoom ──────────────────────────────────────────────────────────────────
  // ── Scroll to today ───────────────────────────────────────────────────────
  private scrollToToday(): void {
    if (!this.viewModel) return;
    const { minDate, maxDate } = this.viewModel;
    const _n    = new Date();
    const today = new Date(_n.getFullYear(), _n.getMonth(), _n.getDate());
    if (today < minDate || today > maxDate) return;

    const daySpan = Math.max(Math.ceil((maxDate.getTime() - minDate.getTime()) / 86400000), 1);
    const chartW  = Math.round(daySpan * this.pxPerDay);
    const todayX  = Math.round(((today.getTime() - minDate.getTime()) / 86400000) * this.pxPerDay);
    const wrapW   = this.chartBodyWrap.clientWidth;
    const target  = Math.max(0, Math.min(todayX - wrapW / 2, chartW - wrapW));

    this.chartBodyWrap.scrollTo({ left: target, behavior: "smooth" });
  }


  // ── Scroll to a specific task (double-click) ────────────────────────────
  private scrollToTask(task: GanttTask): void {
    if (!this.viewModel) return;
    const { minDate, maxDate } = this.viewModel;
    const daySpan = Math.max(Math.ceil((maxDate.getTime() - minDate.getTime()) / 86400000), 1);
    const chartW  = Math.round(daySpan * this.pxPerDay);
    const wrapW   = this.chartBodyWrap.clientWidth;

    // Center horizontally on the task bar midpoint
    const taskMid = (task.plannedStart.getTime() + task.plannedEnd.getTime()) / 2;
    const taskX   = Math.round(((taskMid - minDate.getTime()) / 86400000) * this.pxPerDay);
    const targetX = Math.max(0, Math.min(taskX - wrapW / 2, chartW - wrapW));

    this.chartBodyWrap.scrollTo({ left: targetX, behavior: "smooth" });
  }

  private updateZoomToggle(): void {
    this.zoomBtns.forEach((btn, i) => btn.classList.toggle("active", i === this.zoomIdx));
  }

    private adjustZoom(direction: number): void {
    const maxIdx = this.ZOOM_STEPS_FIXED.length; // 0=fit, 1..N=fixed
    const newIdx = Math.min(maxIdx, Math.max(0,
      this.zoomIdx + (direction > 0 ? 1 : -1)));
    if (newIdx === this.zoomIdx) return;
    this.zoomIdx  = newIdx;
    this.pxPerDay = this.resolvePxPerDay();
    this.updateZoomToggle();

    if (this.viewModel) {
      // Preserve scroll position proportionally so chart doesn't jump
      const scrollMax = this.chartBodyWrap.scrollWidth - this.chartBodyWrap.clientWidth;
      const ratio     = scrollMax > 0 ? this.chartBodyWrap.scrollLeft / scrollMax : 0;

      this.computeVisibility();
      this.render(this.vpWidth, this.vpHeight);

      // Restore proportional scroll position after render changed content width
      const newScrollMax = this.chartBodyWrap.scrollWidth - this.chartBodyWrap.clientWidth;
      this.chartBodyWrap.scrollLeft  = Math.round(ratio * newScrollMax);
      this.chartHeaderWrap.scrollLeft = this.chartBodyWrap.scrollLeft;
    }
  }

  // ── PBI update ────────────────────────────────────────────────────────────
  public update(options: VisualUpdateOptions): void {
    this.events.renderingStarted(options);
    try {
      const dataView = options.dataViews?.[0];
      this.fmtSettings = this.fmtService.populateFormattingSettingsModel(GanttFormattingSettings, dataView);

      if (!dataView) { this.clear(); this.events.renderingFinished(options); return; }

      this.viewModel = parseDataView(dataView, this.host);
      if (!this.viewModel) { this.renderEmptyState(); this.events.renderingFinished(options); return; }

      this.vpWidth  = options.viewport.width;
      this.vpHeight = options.viewport.height;

      // Auto-fit only on first load — start at fit-to-width (zoomIdx 0)
      if (this.firstLoad) {
        this.firstLoad = false;
        this.zoomIdx   = 0;
      }

      // Recalculate pxPerDay when in fit-to-width mode (responds to viewport resize)
      this.pxPerDay = this.resolvePxPerDay();

      this.computeVisibility();
      this.render(this.vpWidth, this.vpHeight);
      this.events.renderingFinished(options);
    } catch (e) {
      this.events.renderingFailed(options, e as string);
    }
  }

  public getFormattingModel(): powerbi.visuals.FormattingModel {
    return this.fmtService.buildFormattingModel(this.fmtSettings);
  }

  // ── Task status ────────────────────────────────────────────────────────────
  private getTaskStatus(task: GanttTask): string {
    const _n = new Date();
    const today = new Date(_n.getFullYear(), _n.getMonth(), _n.getDate());
    if (task.isMilestone) {
      if (today > task.plannedEnd) return "atrasado";
      return "no_prazo";
    }
    if (task.progress >= 100 && today <= task.plannedEnd) return "adiantado";
    if (task.progress >= 100) return "concluida";
    if (today > task.plannedEnd) return "atrasado";
    if (task.progress > 0) return "em_andamento";
    return "no_prazo";
  }

  // ── Collapse + status filter ──────────────────────────────────────────────
  private computeVisibility(): void {
    if (!this.viewModel) return;
    const tasks = this.viewModel.tasks;
    tasks.forEach(t => { t.isVisible = true; });

    // 1) Collapse
    for (let i = 0; i < tasks.length; i++) {
      if (!this.collapsed.has(tasks[i].id)) continue;
      const pl = tasks[i].outlineLevel;
      for (let j = i + 1; j < tasks.length; j++) {
        if (tasks[j].outlineLevel <= pl) break;
        tasks[j].isVisible = false;
      }
    }

    // 2) Status filter
    if (this.statusFilter) {
      const filter = this.statusFilter;
      // First pass: mark all non-matching tasks (including summaries)
      tasks.forEach(t => {
        if (!t.isVisible) return;
        if (this.getTaskStatus(t) !== filter) t.isVisible = false;
      });
      // Second pass (bottom-up): restore summaries that have at least one visible descendant
      for (let i = tasks.length - 1; i >= 0; i--) {
        const t = tasks[i];
        if (!t.isSummary || t.isVisible) continue;
        const pl = t.outlineLevel;
        for (let j = i + 1; j < tasks.length; j++) {
          if (tasks[j].outlineLevel <= pl) break;
          if (tasks[j].isVisible) { t.isVisible = true; break; }
        }
      }
    }
  }

  // ── Position the "Hoje" pill (HTML element, always pixel-perfect) ─────────
  private updateTodayPillPosition(): void {
    if (!this.viewModel) { this.todayPill.style.display = "none"; return; }
    const s      = this.fmtSettings;
    const wbsW   = (this.fmtSettings.layout.showWbs.value && this.showWbs)
      ? s.layout.wbsColumnWidth.value : 0;
    const labelW = this.sidebarLabelW;
    const sidebarW = wbsW + labelW;

    const { minDate, maxDate } = this.viewModel;
    const _n    = new Date();
    const today = new Date(_n.getFullYear(), _n.getMonth(), _n.getDate());

    const daySpan  = Math.max(Math.ceil((maxDate.getTime() - minDate.getTime()) / 86400000), 1);
    const contentW = Math.round(daySpan * this.pxPerDay);
    const wrapW    = Math.max(this.vpWidth - sidebarW, 60);
    const chartW   = Math.max(contentW, wrapW);

    const xScale   = d3.scaleTime().domain([minDate, maxDate]).range([0, contentW]);
    const tx       = xScale(today);

    // Position = sidebar + tx - scrollLeft
    const scrollLeft = this.chartBodyWrap.scrollLeft;
    const left       = sidebarW + tx - scrollLeft;

    // Hide if scrolled out of visible area
    const pillW = 36;
    if (left + pillW / 2 < sidebarW || left - pillW / 2 > this.vpWidth) {
      this.todayPill.style.display = "none";
      return;
    }

    this.todayPill.style.display  = "block";
    this.todayPill.style.left     = `${Math.round(left - pillW / 2)}px`;
    this.todayPill.style.top      = `4px`;

    // Update the dashed line position too — rendered as bodySvg line
    // (already drawn correctly by xScale; only pill needs repositioning)
    // Suppress headerSvg pill — it's now handled by this HTML element
    void chartW; // suppress unused warning
  }

  // ── Render ────────────────────────────────────────────────────────────────
  private render(vpWidth: number, vpHeight: number): void {
    if (!this.viewModel) return;

    // Recalculate pxPerDay in fit-to-width mode (zoomIdx 0) so sidebar
    // resize, WBS toggle, and viewport changes all keep content fitted
    if (this.zoomIdx === 0) {
      this.pxPerDay = this.computeFitPxPerDay();
    }

    const { tasks, minDate, maxDate } = this.viewModel;
    const s            = this.fmtSettings;
    const rowH         = s.layout.rowHeight.value;
    const showDeps          = s.layout.showDependencies.value;
    const showToday         = s.layout.showToday.value;
    const showBaseline      = s.layout.showBaseline.value;
    const showStatusLabels  = s.layout.showStatusLabels.value;
    const showStatusBar     = s.layout.showStatusBar.value;
    const showWbsSetting    = s.layout.showWbs.value;
    // WBS: panel toggle AND in-chart button must both be on
    const wbsActive    = showWbsSetting && this.showWbs;
    const wbsW         = wbsActive ? s.layout.wbsColumnWidth.value : 0;
    const labelW       = this.sidebarLabelW;
    const cPlanned     = s.colors.plannedBarColor.value.value;
    const cSummary     = s.colors.summaryBarColor.value.value;
    const cBaseline    = s.colors.baselineBarColor.value.value;
    const cProgress    = s.colors.progressColor.value.value;
    const cToday       = s.colors.todayLineColor.value.value;
    const cMilestone   = s.colors.milestoneColor.value.value;
    const cMsBaseline  = s.colors.milestoneBaselineColor.value.value;
    const cDepLine     = s.colors.dependencyLineColor.value.value;

    // ── DIAGNOSTIC — remove after debugging ──
    console.log("[GANTT-COLOR] dependencyLineColor:", cDepLine, "| raw:", s.colors.dependencyLineColor);

    const visibleTasks = tasks.filter(t => t.isVisible);
    const daySpan      = Math.max(Math.ceil((maxDate.getTime() - minDate.getTime()) / 86400000), 1);
    const sidebarW     = wbsW + labelW;
    const wrapW        = Math.max(vpWidth - sidebarW, 60);
    // contentW: exact pixels based on pxPerDay — xScale ALWAYS uses this, never stretched
    const contentW     = Math.round(daySpan * this.pxPerDay);
    // svgW: in fit mode (zoomIdx 0) use contentW to avoid horizontal scroll;
    // in other modes, at least wrapW so background fills the visible area
    const chartW       = this.zoomIdx === 0 ? contentW : Math.max(contentW, wrapW);
    const totalH       = visibleTasks.length * rowH;
    const bodyH        = vpHeight - this.HEADER_H;

    const zl = pickZoom(this.pxPerDay);
    this.updateZoomToggle();

    // xScale maps to contentW — bars are always proportional to pxPerDay
    const xScale   = d3.scaleTime().domain([minDate, maxDate]).range([0, contentW]);
    const topTicks = fullTicks(xScale, zl.topInterval);
    const botTicks = fullTicks(xScale, zl.bottomInterval);

    // ── Sidebar ──────────────────────────────────────────────────────────────
    this.sidebar.style.cssText =
      `width:${sidebarW}px;min-width:${sidebarW}px;max-width:${sidebarW}px;height:${vpHeight}px;`;

    // Position resize handle at right edge of sidebar
    const resizeHandleEl = this.root.querySelector(".sidebar-resize-handle") as HTMLElement;
    if (resizeHandleEl) {
      resizeHandleEl.style.cssText =
        `position:absolute;left:${sidebarW - 3}px;top:0;width:6px;height:${vpHeight}px;` +
        `cursor:col-resize;z-index:20;`;
    }

    // WBS toggle button — hide if panel toggle is off
    const wbsToggleEl = this.root.querySelector(".wbs-toggle-btn") as HTMLElement;
    if (wbsToggleEl) {
      if (!showWbsSetting) {
        wbsToggleEl.style.display = "none";
      } else {
        wbsToggleEl.style.display = "";
        wbsToggleEl.style.cssText =
          `position:absolute;top:${(this.HEADER_H - 22) / 2}px;left:${sidebarW - 62}px;` +
          `z-index:25;`;
      }
    }

    // Status filter bar — hide if disabled from panel
    const filterBarEl = this.root.querySelector(".filter-bar") as HTMLElement;
    if (filterBarEl) filterBarEl.style.display = showStatusBar ? "" : "none";

    const sideHeaderEl = this.sidebar.querySelector(".sidebar-header") as HTMLElement;
    sideHeaderEl.innerHTML = "";
    sideHeaderEl.style.height = `${this.HEADER_H}px`;
    sideHeaderEl.style.minHeight = `${this.HEADER_H}px`;

    if (wbsActive) {
      const wbsHead = document.createElement("div");
      wbsHead.className = "col-wbs";
      wbsHead.style.cssText = `width:${wbsW}px;min-width:${wbsW}px;height:${this.HEADER_H}px;`;
      wbsHead.textContent = "WBS";
      sideHeaderEl.appendChild(wbsHead);
    }

    const nameHead = document.createElement("div");
    nameHead.className = "col-name";
    nameHead.textContent = "Tarefa";
    sideHeaderEl.appendChild(nameHead);

    // Sidebar rows
    this.sideBody.innerHTML = "";
    this.sideBody.style.height = `${bodyH}px`;

    tasks.forEach(task => {
      const row = document.createElement("div");
      row.className = "sidebar-row";
      if (task.isSummary)  row.classList.add("is-summary");
      if (!task.isVisible) row.classList.add("is-hidden");
      row.style.height = `${rowH}px`;
      row.dataset.id   = task.id;

      if (wbsActive) {
        const wbsCell = document.createElement("div");
        wbsCell.className = "col-wbs";
        wbsCell.style.cssText = `width:${wbsW}px;min-width:${wbsW}px;height:${rowH}px;`;
        wbsCell.textContent = task.wbs || "";
        row.appendChild(wbsCell);
      }

      const nameCell = document.createElement("div");
      nameCell.className = "col-name";
      nameCell.style.height = `${rowH}px`;

      const indentW = (task.outlineLevel - 1) * this.INDENT_PX;
      if (indentW > 0) {
        const sp = document.createElement("div");
        sp.className = "task-indent";
        sp.style.width = `${indentW}px`;
        nameCell.appendChild(sp);
      }

      if (task.isSummary) {
        const btn = document.createElement("div");
        btn.className = "collapse-btn" + (this.collapsed.has(task.id) ? " collapsed" : "");
        btn.innerHTML = "&#9660;";
        btn.addEventListener("click", ev => {
          ev.stopPropagation();
          if (this.collapsed.has(task.id)) this.collapsed.delete(task.id);
          else this.collapsed.add(task.id);
          this.computeVisibility();
          this.render(vpWidth, vpHeight);
        });
        nameCell.appendChild(btn);
      } else if (task.isMilestone) {
        const diamond = document.createElement("div");
        diamond.className = "task-diamond";
        nameCell.appendChild(diamond);
      } else {
        const b = document.createElement("div");
        b.className = "task-bullet";
        nameCell.appendChild(b);
      }

      const ns = document.createElement("span");
      ns.className = "task-name-text" + (task.isSummary ? " summary" : "");
      ns.textContent = task.name;
      ns.title = task.name;
      nameCell.appendChild(ns);

      // Status badge — only for non-summary, non-milestone tasks
      if (showStatusLabels && !task.isSummary && !task.isMilestone) {
        const STATUS_DEF: Record<string, { cls: string; label: string }> = {
          adiantado:    { cls: "early",      label: "Adiantado" },
          concluida:    { cls: "done",       label: "Concluída" },
          atrasado:     { cls: "late",       label: "Atrasado" },
          em_andamento: { cls: "inprogress", label: "Andamento" },
          no_prazo:     { cls: "ontime",     label: "No prazo" },
        };
        const st = this.getTaskStatus(task);
        const def = STATUS_DEF[st];
        if (def) {
          const badge = document.createElement("span");
          badge.className = `task-status-badge ${def.cls}`;
          badge.textContent = def.label;
          nameCell.appendChild(badge);
        }
      }

      row.appendChild(nameCell);

      row.addEventListener("click", ev => {
        if (!task.isVisible) return;
        ev.stopPropagation();
        this.selectionManager.select(task.selectionId, ev.ctrlKey || ev.metaKey)
          .then(() => this.syncHighlight());
      });
      row.addEventListener("contextmenu", ev => {
        ev.preventDefault();
        this.selectionManager.showContextMenu(task.selectionId, { x: ev.clientX, y: ev.clientY });
      });
      row.addEventListener("dblclick", ev => {
        ev.stopPropagation();
        this.scrollToTask(task);
      });
      this.sideBody.appendChild(row);
    });

    // ── Chart layout dims ────────────────────────────────────────────────────
    // Header wrap: fixed height, overflow hidden (JS scrolls scrollLeft)
    this.chartHeaderWrap.style.cssText =
      `width:${wrapW}px;height:${this.HEADER_H}px;overflow:hidden;flex-shrink:0;`;

    // Body wrap: remaining height, overflow-x hidden in fit mode (no horizontal scroll needed)
    const overflowX = this.zoomIdx === 0 ? "hidden" : "auto";
    this.chartBodyWrap.style.cssText =
      `width:${wrapW}px;height:${bodyH}px;overflow-x:${overflowX};overflow-y:auto;flex:1;`;

    // ── Header SVG ───────────────────────────────────────────────────────────
    this.headerSvg.attr("width", chartW).attr("height", this.HEADER_H);
    this.headerSvg.selectAll("*").remove();

    // Background
    this.headerSvg.append("rect")
      .attr("x", 0).attr("y", 0).attr("width", chartW).attr("height", this.HEADER_H)
      .attr("fill", "#f3f4f6");

    // Separator lines
    this.headerSvg.append("line")
      .attr("x1", 0).attr("x2", chartW)
      .attr("y1", this.TOP_ROW_H).attr("y2", this.TOP_ROW_H)
      .attr("stroke", "#d1d5db").attr("stroke-width", 1);

    this.headerSvg.append("line")
      .attr("x1", 0).attr("x2", chartW)
      .attr("y1", this.HEADER_H).attr("y2", this.HEADER_H)
      .attr("stroke", "#c9cdd4").attr("stroke-width", 1.5);

    // Top row: render each cell between consecutive ticks
    topTicks.forEach((d, i) => {
      const x1    = xScale(d);
      const x2    = i + 1 < topTicks.length
        ? xScale(topTicks[i + 1])
        : Math.min(xScale(zl.topInterval.offset(d, 1)), chartW);
      const cellW = x2 - x1;

      // Alternating background
      this.headerSvg.append("rect")
        .attr("x", x1).attr("y", 0)
        .attr("width", cellW).attr("height", this.TOP_ROW_H)
        .attr("fill", i % 2 === 0 ? "#f3f4f6" : "#e9edf2");

      // Divider between cells
      if (i > 0) {
        this.headerSvg.append("line")
          .attr("x1", x1).attr("x2", x1)
          .attr("y1", 0).attr("y2", this.TOP_ROW_H)
          .attr("stroke", "#c9cdd4").attr("stroke-width", 1);
      }

      // Label — center within cell, always show full FY label
      const label = zl.topFmt(d);
      const minWidthForLabel = 20;
      if (cellW >= minWidthForLabel) {
        // Use SVG clipPath per cell so text doesn't bleed into next cell
        const clipId = `clip-top-${i}`;
        this.headerSvg.append("clipPath").attr("id", clipId)
          .append("rect").attr("x", x1 + 2).attr("y", 0)
          .attr("width", Math.max(cellW - 4, 0)).attr("height", this.TOP_ROW_H);

        this.headerSvg.append("text")
          .attr("x", x1 + cellW / 2)
          .attr("y", this.TOP_ROW_H / 2 + 4)
          .attr("text-anchor", "middle")
          .attr("clip-path", `url(#clip-top-${i})`)
          .attr("fill", "#1e293b")
          .attr("font-size", "11px").attr("font-weight", "700")
          .attr("font-family", "Segoe UI, sans-serif")
          .text(label);
      }
    });

    // Bottom row
    const isDayView = this.pxPerDay >= 12;
    const today     = new Date();

    botTicks.forEach((d, i) => {
      const x1    = xScale(d);
      const x2    = i + 1 < botTicks.length
        ? xScale(botTicks[i + 1])
        : Math.min(xScale(zl.bottomInterval.offset(d, 1)), chartW);
      const cellW = Math.max(x2 - x1, 0);
      const isWE  = isDayView && (d.getDay() === 0 || d.getDay() === 6);
      // Suppress label when today falls in this cell (arrow/pill already marks it)
      const _t = new Date();
      const todayLocal  = new Date(_t.getFullYear(), _t.getMonth(), _t.getDate());
      const todayInCell = showToday
        && todayLocal >= minDate && todayLocal <= maxDate
        && todayLocal >= d
        && (i + 1 >= botTicks.length || todayLocal < botTicks[i + 1]);

      this.headerSvg.append("rect")
        .attr("x", x1).attr("y", this.TOP_ROW_H)
        .attr("width", cellW).attr("height", this.BOT_ROW_H)
        .attr("fill", isWE ? "#dde3ec" : (i % 2 === 0 ? "#f3f4f6" : "#ebedf0"));

      if (i > 0) {
        this.headerSvg.append("line")
          .attr("x1", x1).attr("x2", x1)
          .attr("y1", this.TOP_ROW_H).attr("y2", this.HEADER_H)
          .attr("stroke", "#d1d5db").attr("stroke-width", 1);
      }

      const label = zl.bottomFmt(d, i, botTicks);
      const minW  = label.length * 5.5;
      if (cellW >= minW && !todayInCell) {
        this.headerSvg.append("text")
          .attr("x", x1 + cellW / 2)
          .attr("y", this.TOP_ROW_H + this.BOT_ROW_H / 2 + 4)
          .attr("text-anchor", "middle")
          .attr("fill", isWE ? "#94a3b8" : "#4b5563")
          .attr("font-size", "10px").attr("font-weight", "600")
          .attr("font-family", "Segoe UI, sans-serif")
          .text(label);
      }
    });

    // ── Body SVG ─────────────────────────────────────────────────────────────
    this.bodySvg.attr("width", chartW).attr("height", Math.max(totalH, bodyH));
    this.bodySvg.selectAll("*").remove();

    const fullH = Math.max(totalH, bodyH);

    // Defs
    const defs = this.bodySvg.append("defs");
    defs.append("clipPath").attr("id", "bars-clip")
      .append("rect").attr("x", 0).attr("y", 0).attr("width", chartW).attr("height", fullH);
    defs.append("marker").attr("id", "dep-arrow")
      .attr("viewBox", "0 0 8 8").attr("refX", 7).attr("refY", 4)
      .attr("markerWidth", 6).attr("markerHeight", 6).attr("orient", "auto")
      .append("path").attr("d", "M0,1 L7,4 L0,7 z").attr("fill", cDepLine);

    // Row backgrounds
    visibleTasks.forEach((task, i) => {
      this.bodySvg.append("rect")
        .classed("row-bg", true)
        .attr("data-id", task.id)
        .attr("x", 0).attr("y", i * rowH).attr("width", chartW).attr("height", rowH)
        .attr("fill", task.isSummary ? "#e6f7f6" : (i % 2 === 0 ? "#fafbfc" : "#fff"));
    });

    // Vertical grid — top (stronger)
    topTicks.forEach(d => {
      this.bodySvg.append("line")
        .attr("x1", xScale(d)).attr("x2", xScale(d))
        .attr("y1", 0).attr("y2", fullH)
        .attr("stroke", "#c9cdd4").attr("stroke-width", 1.5);
    });

    // Vertical grid — bottom (lighter)
    botTicks.forEach(d => {
      const isWE = isDayView && (d.getDay() === 0 || d.getDay() === 6);
      if (isWE) {
        const x1   = xScale(d);
        const next = new Date(d); next.setDate(next.getDate() + 1);
        const wid  = Math.max(xScale(next) - x1, 0);
        this.bodySvg.append("rect")
          .attr("x", x1).attr("y", 0).attr("width", wid).attr("height", fullH)
          .attr("fill", "#f1f5f9").attr("opacity", 0.6);
      }
      this.bodySvg.append("line")
        .attr("x1", xScale(d)).attr("x2", xScale(d))
        .attr("y1", 0).attr("y2", fullH)
        .attr("stroke", "#eaecef").attr("stroke-width", 1);
    });

    // Horizontal row dividers
    visibleTasks.forEach((_, i) => {
      this.bodySvg.append("line")
        .attr("x1", 0).attr("x2", chartW)
        .attr("y1", (i + 1) * rowH).attr("y2", (i + 1) * rowH)
        .attr("stroke", "#eaecef").attr("stroke-width", 1);
    });

    // ── Today line + header marker ───────────────────────────────────────────
    if (showToday) {
      const _now  = new Date();
      // Normalize to local midnight — same as all task dates — so xScale position is exact
      const today = new Date(_now.getFullYear(), _now.getMonth(), _now.getDate(), 0, 0, 0, 0);
      // Always render (domain now guaranteed to include today from dataParser)
      if (today >= minDate && today <= maxDate) {
        const tx = xScale(today);

        // Subtle band in body
        this.bodySvg.append("rect")
          .attr("x", tx - this.pxPerDay * 0.5).attr("y", 0)
          .attr("width", Math.max(this.pxPerDay, 2)).attr("height", fullH)
          .attr("fill", cToday).attr("opacity", 0.06);

        // Solid today line in body
        this.bodySvg.append("line")
          .attr("x1", tx).attr("x2", tx).attr("y1", 0).attr("y2", fullH)
          .attr("stroke", cToday).attr("stroke-width", 2)
          .attr("stroke-dasharray", "5,4").attr("opacity", 0.9);

        // Header: dashed line full height
        this.headerSvg.append("line")
          .attr("x1", tx).attr("x2", tx)
          .attr("y1", 0).attr("y2", this.HEADER_H + 2)
          .attr("stroke", cToday).attr("stroke-width", 2)
          .attr("stroke-dasharray", "5,4").attr("opacity", 0.9);

        // Header: "Hoje" pill rendered inside headerSvg (avoids HTML stacking issues)
        const pillW = 36, pillH = 18, pillR = 9;
        const pillX = tx - pillW / 2, pillY = 3;
        this.headerSvg.append("rect")
          .attr("x", pillX).attr("y", pillY)
          .attr("width", pillW).attr("height", pillH)
          .attr("rx", pillR).attr("ry", pillR)
          .attr("fill", cToday);
        this.headerSvg.append("text")
          .attr("x", tx).attr("y", pillY + pillH / 2 + 4)
          .attr("text-anchor", "middle")
          .attr("fill", "#fff").attr("font-size", "9.5px")
          .attr("font-weight", "700")
          .attr("font-family", "Segoe UI, sans-serif")
          .text("Hoje");
      }
    }

    // Hide the HTML pill — now rendered in SVG above
    this.todayPill.style.display = "none";

    // ── Dependencies ──────────────────────────────────────────────────────────
    if (showDeps) this.renderDependencies(visibleTasks, xScale, rowH, cDepLine);

    // ── Bars ─────────────────────────────────────────────────────────────────
    const barsG = this.bodySvg.append("g")
      .classed("bars", true).attr("clip-path", "url(#bars-clip)");

    visibleTasks.forEach((task, i) => {
      const rowY  = i * rowH;
      const barCY = rowY + rowH / 2;
      const barH  = task.isSummary ? Math.round(rowH * 0.30) : Math.round(rowH * 0.44);
      const baseH = Math.max(Math.round(rowH * 0.14), 4);
      const barY  = barCY - barH / 2;
      const x1    = xScale(task.plannedStart);
      // Always derive width from xScale so position + width use the same mapping.
      // Add 1 day to end → single-day tasks get exactly 1 cell width.
      // +1 day using local-time constructor — safe across DST boundaries
      const endPlus1 = new Date(
        task.plannedEnd.getFullYear(),
        task.plannedEnd.getMonth(),
        task.plannedEnd.getDate() + 1
      );
      const barW  = Math.max(xScale(endPlus1) - x1, 4);

      const tg = barsG.append("g")
        .classed("task-group", true).attr("data-id", task.id)
        .style("cursor", "pointer")
        .on("click", (ev: MouseEvent) => {
          ev.stopPropagation();
          this.selectionManager.select(task.selectionId, ev.ctrlKey || ev.metaKey)
            .then(() => this.syncHighlight());
        })
        .on("contextmenu", (ev: MouseEvent) => {
          ev.preventDefault();
          this.selectionManager.showContextMenu(task.selectionId, { x: ev.clientX, y: ev.clientY });
        })
        .on("mouseenter", (ev: MouseEvent) => this.showTooltip(ev, task))
        .on("mouseleave", () => { this.tooltip.style.display = "none"; });

      if (task.isMilestone) {
        // ── MILESTONE: diamond centered on plannedStart ──────────────────
        const ms   = xScale(task.plannedStart);

        // Fixed ratio of rowH — never grows with zoom
        const half = Math.round(rowH * 0.28);

        const diamond = `M${ms},${barCY - half} L${ms + half},${barCY} L${ms},${barCY + half} L${ms - half},${barCY} Z`;

        tg.append("path").attr("d", diamond)
          .attr("fill", cMilestone).attr("opacity", 0.92);
        tg.append("path").attr("d", diamond)
          .attr("fill", "none")
          .attr("stroke", "#fff").attr("stroke-width", 1.5).attr("opacity", 0.5);

        // Baseline: small dot directly below diamond bottom tip
        if (showBaseline && task.baselineStart) {
          const bx   = xScale(task.baselineStart);
          const dotR = Math.max(Math.round(rowH * 0.08), 3);
          const dotY = barCY + half + 2 + dotR;  // just below bottom tip

          tg.append("circle")
            .attr("cx", bx).attr("cy", dotY).attr("r", dotR + 2)
            .attr("fill", cMsBaseline).attr("opacity", 0.28);
          tg.append("circle")
            .attr("cx", bx).attr("cy", dotY).attr("r", dotR)
            .attr("fill", cMsBaseline)
            .attr("stroke", "#fff").attr("stroke-width", 1).attr("opacity", 1);
        }

      } else {
        // ── REGULAR BAR ──────────────────────────────────────────────────────
        // All three layers share the same x/y/w/h so they align perfectly.
        // The outline uses stroke-alignment via inset (half stroke-width).
        const sw = 1.5;  // stroke width
        const progressW = Math.min(Math.max(barW * (task.progress / 100), this.BAR_R * 2), barW);

        // Planned bg (translucent) — lighter tint fixed so barra tem corpo visível
        tg.append("rect")
          .attr("x", x1).attr("y", barY).attr("width", barW).attr("height", barH)
          .attr("rx", this.BAR_R).attr("fill", "#96DDDA").attr("opacity", 0.45);

        // Progress fill — capped at barW so it never overflows
        if (task.progress > 0) {
          tg.append("rect")
            .attr("x", x1).attr("y", barY)
            .attr("width", progressW)
            .attr("height", barH).attr("rx", this.BAR_R)
            .attr("fill", task.isSummary ? cSummary : cProgress);
        }

        // Planned outline — same bounds, stroke centered on edge
        tg.append("rect")
          .attr("x", x1).attr("y", barY)
          .attr("width", barW).attr("height", barH)
          .attr("rx", this.BAR_R).attr("fill", "none")
          .attr("stroke", task.isSummary ? cSummary : cPlanned).attr("stroke-width", sw);

        // % label — inside bar when wide enough, outside (right) when narrow
        // Summary tasks always show outside (bar is thinner)
        {
          const pctText = `${Math.round(task.progress)}%`;
          const minInside = 42;  // px threshold to fit label inside
          const summaryColor = cSummary;

          if (!task.isSummary && barW >= minInside) {
            // ── Inside bar ──
            const lx = task.progress > 0
              ? x1 + Math.min(barW * (task.progress / 100) / 2, barW / 2)
              : x1 + barW / 2;
            tg.append("text")
              .attr("x", lx).attr("y", barCY)
              .attr("dy", "0.35em")
              .attr("text-anchor", "middle")
              .attr("pointer-events", "none")
              .attr("fill", "#fff").attr("font-size", "9.5px").attr("font-weight", "700")
              .attr("font-family", "Segoe UI, sans-serif")
              .attr("paint-order", "stroke")
              .attr("stroke", "rgba(0,0,0,0.25)").attr("stroke-width", "2px")
              .text(pctText);
          } else {
            // ── Outside bar (right side) ──
            tg.append("text")
              .attr("x", x1 + barW + 5).attr("y", barCY)
              .attr("dy", "0.35em")
              .attr("text-anchor", "start")
              .attr("pointer-events", "none")
              .attr("fill", task.isSummary ? summaryColor : "#6b7280")
              .attr("font-size", "9.5px").attr("font-weight", "700")
              .attr("font-family", "Segoe UI, sans-serif")
              .text(pctText);
          }
        }

        // Baseline below planned bar
        if (showBaseline && task.baselineStart && task.baselineEnd) {
          const bx1 = xScale(task.baselineStart);
          const bEndPlus1 = new Date(
            task.baselineEnd!.getFullYear(),
            task.baselineEnd!.getMonth(),
            task.baselineEnd!.getDate() + 1
          );
          const bW  = Math.max(xScale(bEndPlus1) - bx1, 4);
          const bY  = barY + barH + 3;
          tg.append("rect")
            .attr("x", bx1).attr("y", bY).attr("width", bW).attr("height", baseH)
            .attr("rx", 2).attr("fill", cBaseline).attr("opacity", 0.55);
          tg.append("rect")
            .attr("x", bx1 + .5).attr("y", bY + .5)
            .attr("width", bW - 1).attr("height", baseH - 1)
            .attr("rx", 2).attr("fill", "none")
            .attr("stroke", cBaseline).attr("stroke-width", 1).attr("opacity", 0.9);
        }
      }
    });

    // ── Sync header scroll with body after re-render (fixes misalignment on zoom) ──
    this.chartHeaderWrap.scrollLeft = this.chartBodyWrap.scrollLeft;
  }

  // ── Dependencies ─────────────────────────────────────────────────────────
  private renderDependencies(
    tasks: GanttTask[], xScale: d3.ScaleTime<number, number>, rowH: number, color: string
  ): void {
    // Build lookup with multiple keys so deps match by taskId, wbs, or name
    const idxMap = new Map<string, number>();
    tasks.forEach((t, i) => {
      if (t.id)     idxMap.set(t.id, i);
      if (t.taskId && !idxMap.has(t.taskId)) idxMap.set(t.taskId, i);
      if (t.wbs    && !idxMap.has(t.wbs))    idxMap.set(t.wbs, i);
      if (t.name   && !idxMap.has(t.name))   idxMap.set(t.name, i);
    });

    // ── DIAGNOSTIC LOG — remove after debugging ──
    console.log("[GANTT-DEPS] idxMap keys:", Array.from(idxMap.keys()).slice(0, 20));
    console.log("[GANTT-DEPS] sample tasks:", tasks.slice(0, 5).map(t => ({
      id: t.id, taskId: t.taskId, wbs: t.wbs, deps: t.dependencies
    })));
    let matchCount = 0, missCount = 0;
    tasks.forEach(t => t.dependencies.forEach(depId => {
      if (idxMap.has(depId)) matchCount++; else missCount++;
    }));
    console.log(`[GANTT-DEPS] matches: ${matchCount}, misses: ${missCount}`);
    const g = this.bodySvg.append("g").classed("deps", true);
    const R = 4; // corner radius

    tasks.forEach((task, toIdx) => {
      task.dependencies.forEach(depId => {
        const fromIdx = idxMap.get(depId);
        if (fromIdx === undefined) return;
        const from = tasks[fromIdx];

        // End of predecessor bar (+1 day)
        const fromEndPlus1 = new Date(
          from.plannedEnd.getFullYear(),
          from.plannedEnd.getMonth(),
          from.plannedEnd.getDate() + 1
        );
        const x1 = xScale(fromEndPlus1);
        const y1 = fromIdx * rowH + rowH / 2;
        const x2 = xScale(task.plannedStart);
        const y2 = toIdx   * rowH + rowH / 2;

        // Elbow X: horizontal exit then vertical drop
        const elbowX = Math.max(x1 + 10, x2 - 12);
        const dy = y2 - y1;
        const r  = Math.min(R, Math.abs(dy) / 2, Math.abs(elbowX - x1) / 2, Math.abs(x2 - elbowX) / 2);

        let d: string;
        if (Math.abs(dy) < 1) {
          // Same row — straight horizontal
          d = `M${x1} ${y1} L${x2} ${y2}`;
        } else {
          const sY = dy > 0 ? 1 : -1; // vertical direction sign
          d = `M${x1} ${y1}`
            + ` H${elbowX - r}`
            + ` Q${elbowX} ${y1} ${elbowX} ${y1 + sY * r}`
            + ` V${y2 - sY * r}`
            + ` Q${elbowX} ${y2} ${elbowX + r} ${y2}`
            + ` H${x2}`;
        }

        g.append("path")
          .attr("d", d)
          .attr("fill", "none")
          .attr("stroke", color)
          .attr("stroke-width", 1.2)
          .attr("marker-end", "url(#dep-arrow)");
      });
    });
  }

  // ── Highlight ────────────────────────────────────────────────────────────
  private syncHighlight(): void {
    if (!this.viewModel) return;
    const sel = this.selectionManager.getSelectionIds() as powerbi.visuals.ISelectionId[];
    const has = sel.length > 0;
    this.viewModel.tasks.forEach(t => {
      t.isHighlighted = has ? sel.some(id => id.equals(t.selectionId)) : undefined;
    });
    this.bodySvg.selectAll<SVGGElement, unknown>(".task-group").each((_, i, nodes) => {
      const el  = nodes[i] as SVGGElement;
      const dim = this.viewModel?.tasks.find(t => t.id === el.getAttribute("data-id"))?.isHighlighted === false;
      d3.select(el).attr("opacity", dim ? 0.2 : 1);
    });
    // Update row background fills to reflect selection state
    this.bodySvg.selectAll<SVGRectElement, unknown>(".row-bg").each((_, i, nodes) => {
      const el   = nodes[i] as SVGRectElement;
      const task = this.viewModel?.tasks.find(t => t.id === el.getAttribute("data-id"));
      if (!task) return;
      let fill: string;
      if (task.isHighlighted === true) {
        fill = "#b2e8e6";  // selected — teal tint
      } else if (task.isHighlighted === false) {
        fill = task.isSummary ? "#e6f7f6" : "#f5f5f5";  // dimmed
      } else {
        // No selection active — restore normal alternating background
        const visIdx = this.viewModel!.tasks.filter(t => t.isVisible).indexOf(task);
        fill = task.isSummary ? "#e6f7f6" : (visIdx % 2 === 0 ? "#fafbfc" : "#fff");
      }
      d3.select(el).attr("fill", fill);
    });
    this.sideBody.querySelectorAll<HTMLElement>(".sidebar-row").forEach(row => {
      const t = this.viewModel?.tasks.find(x => x.id === row.dataset.id);
      row.classList.toggle("is-dimmed",   t?.isHighlighted === false);
      row.classList.toggle("is-selected", t?.isHighlighted === true);
    });
  }

  // ── Tooltip ───────────────────────────────────────────────────────────────
  private showTooltip(ev: MouseEvent, task: GanttTask): void {
    const fmt = d3.timeFormat("%d/%m/%y");

    // Builds a date block: Planejado + Baseline with pipe + delta (planned − baseline)
    const dateBlock = (
      label: string,
      planned: Date,
      baseline: Date | null
    ): string => {
      const d = baseline
        ? Math.round((planned.getTime() - baseline.getTime()) / 86400000)
        : null;
      const cls = d === null ? "" : d === 0 ? "ontime" : d > 0 ? "late" : "early";
      const val = d === null ? "" : d === 0 ? "0" : d > 0 ? `+${d}` : `${d}`;

      return `
        <div class="tt-section-label">${label}</div>
        <div class="tt-date-group">
          <div class="tt-date-lines">
            <div class="tt-date-row">
              <span class="tt-date-type">Planejado</span>
              <span class="tt-date-val">${fmt(planned)}</span>
            </div>
            ${baseline ? `
            <div class="tt-date-row muted">
              <span class="tt-date-type">Baseline</span>
              <span class="tt-date-val">${fmt(baseline)}</span>
            </div>` : ""}
          </div>
          ${d !== null ? `
          <div class="tt-pipe-col">
            <span class="tt-delta ${cls}">${val}</span>
          </div>` : ""}
        </div>`;
    };

    // Status: compare actual progress vs expected progress based on today
    const statusLabel = (): string => {
      const STATUS_MAP: Record<string, { cls: string; label: string }> = {
        adiantado:    { cls: "early",      label: "Adiantado" },
        concluida:    { cls: "ontime",     label: "Concluída" },
        atrasado:     { cls: "late",       label: "Atrasado" },
        em_andamento: { cls: "inprogress", label: "Em andamento" },
        no_prazo:     { cls: "ontime",     label: "No prazo" },
      };
      const s = this.getTaskStatus(task);
      const m = STATUS_MAP[s] || STATUS_MAP["no_prazo"];
      return `<span class="tt-status-badge ${m.cls}">${m.label}</span>`;
    };

    if (task.isMilestone) {
      this.tooltip.innerHTML = `
        <div class="tt-badge">◆ Marco</div>
        <div class="tt-title">${task.name}</div>
        ${task.wbs ? `<div class="tt-wbs">WBS: ${task.wbs}</div>` : ""}
        <hr class="tt-divider"/>
        ${dateBlock("Data", task.plannedStart, task.baselineStart)}`;
    } else {
      this.tooltip.innerHTML = `
        <div class="tt-title">${task.name}</div>
        ${task.wbs ? `<div class="tt-wbs">WBS: ${task.wbs}</div>` : ""}
        <hr class="tt-divider"/>
        ${dateBlock("Início", task.plannedStart, task.baselineStart)}
        ${dateBlock("Término", task.plannedEnd, task.baselineEnd)}
        <hr class="tt-divider"/>
        <div class="tt-progress-row">
          <span class="tt-progress-label">Progresso: <b>${Math.round(task.progress)}%</b></span>
          ${statusLabel()}
        </div>
        <div class="tt-bar-bg"><div class="tt-bar-fill" style="width:${task.progress}%"></div></div>`;
    }

    const rect = this.root.getBoundingClientRect();
    let tx = ev.clientX - rect.left + 14;
    let ty = ev.clientY - rect.top  + 14;
    if (tx + 300 > rect.width)  tx = ev.clientX - rect.left - 310;
    if (ty + 260 > rect.height) ty = ev.clientY - rect.top  - 270;
    this.tooltip.style.cssText = `left:${tx}px;top:${ty}px;display:block;`;
  }

  private renderEmptyState(): void {
    this.sideBody.innerHTML = "";
    this.bodySvg.selectAll("*").remove();
    this.bodySvg.attr("width", 300).attr("height", 150);
    this.bodySvg.append("text")
      .attr("x", 150).attr("y", 75).attr("text-anchor", "middle")
      .attr("fill", "#9ca3af").attr("font-size", "13px")
      .text("Adicione os campos para exibir o Gantt");
  }

  private clear(): void {
    this.sideBody.innerHTML = "";
    this.headerSvg.selectAll("*").remove();
    this.bodySvg.selectAll("*").remove();
  }
}