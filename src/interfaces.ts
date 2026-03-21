export interface GanttTask {
  id: string;
  name: string;
  wbs: string;
  outlineLevel: number;
  isSummary: boolean;
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
    baselineBarColor: string;
    progressColor: string;
    todayLineColor: string;
    milestoneColor: string;
    milestoneBaselineColor: string;
  };
  layout: {
    rowHeight: number;
    wbsColumnWidth: number;
    taskLabelWidth: number;
    showDependencies: boolean;
    showToday: boolean;
    showBaseline: boolean;
  };
}

export const DEFAULT_SETTINGS: GanttSettings = {
  colors: {
    plannedBarColor:        "#4472C4",
    baselineBarColor:       "#A5A5A5",
    progressColor:          "#70AD47",
    todayLineColor:         "#f66a0a",
    milestoneColor:         "#2C3E6B",
    milestoneBaselineColor: "#F5C400",
  },
  layout: {
    rowHeight:        36,
    wbsColumnWidth:   64,
    taskLabelWidth:   220,
    showDependencies: true,
    showToday:        true,
    showBaseline:     true,
  },
};

export interface GanttViewModel {
  tasks: GanttTask[];
  settings: GanttSettings;
  minDate: Date;
  maxDate: Date;
}