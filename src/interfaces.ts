export interface GanttTask {
  id: string;
  taskId: string;
  name: string;
  wbs: string;
  group: string;
  outlineLevel: number;
  isSummary: boolean;
  isGroupHeader: boolean;
  plannedStart: Date;
  plannedEnd: Date;
  baselineStart: Date | null;
  baselineEnd: Date | null;
  progress: number;
  dependencies: string[];
  isMilestone: boolean;
  selectionId: powerbi.visuals.ISelectionId;
  isHighlighted?: boolean;
  isVisible?: boolean;
}

export interface GanttSettings {
  colors: {
    plannedBarColor: string;
    summaryBarColor: string;
    baselineBarColor: string;
    progressColor: string;
    todayLineColor: string;
    milestoneColor: string;
    milestoneBaselineColor: string;
    dependencyLineColor: string;
  };
  layout: {
    rowHeight: number;
    wbsColumnWidth: number;
    taskLabelWidth: number;
    showDependencies: boolean;
    showToday: boolean;
    showBaseline: boolean;
    showStatusLabels: boolean;
    showStatusBar: boolean;
    showWbs: boolean;
  };
}

export const DEFAULT_SETTINGS: GanttSettings = {
  colors: {
    plannedBarColor:        "#0E938E",
    summaryBarColor:        "#0E938E",
    baselineBarColor:       "#A5A5A5",
    progressColor:          "#70AD47",
    todayLineColor:         "#f66a0a",
    milestoneColor:         "#0E938E",
    milestoneBaselineColor: "#F5C400",
    dependencyLineColor:    "#d1d5db",
  },
  layout: {
    rowHeight:        36,
    wbsColumnWidth:   64,
    taskLabelWidth:   220,
    showDependencies: true,
    showToday:        true,
    showBaseline:     true,
    showStatusLabels: true,
    showStatusBar:    true,
    showWbs:          true,
  },
};

export interface GanttViewModel {
  tasks: GanttTask[];
  settings: GanttSettings;
  minDate: Date;
  maxDate: Date;
}