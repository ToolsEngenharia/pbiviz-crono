import { formattingSettings } from "powerbi-visuals-utils-formattingmodel";

import FormattingSettingsCard  = formattingSettings.SimpleCard;
import FormattingSettingsSlice = formattingSettings.Slice;
import FormattingSettingsModel = formattingSettings.Model;

// ── Cores ─────────────────────────────────────────────────────────────────────
class ColorsCard extends FormattingSettingsCard {
  plannedBarColor = new formattingSettings.ColorPicker({
    name: "plannedBarColor", displayName: "Barra Planejada",
    value: { value: "#0E938E" },
  });
  baselineBarColor = new formattingSettings.ColorPicker({
    name: "baselineBarColor", displayName: "Barra Baseline",
    value: { value: "#A5A5A5" },
  });
  progressColor = new formattingSettings.ColorPicker({
    name: "progressColor", displayName: "Progresso",
    value: { value: "#70AD47" },
  });
  todayLineColor = new formattingSettings.ColorPicker({
    name: "todayLineColor", displayName: "Linha de Hoje",
    value: { value: "#f66a0a" },
  });
  milestoneColor = new formattingSettings.ColorPicker({
    name: "milestoneColor", displayName: "Marco (losango)",
    value: { value: "#0E938E" },
  });
  milestoneBaselineColor = new formattingSettings.ColorPicker({
    name: "milestoneBaselineColor", displayName: "Marco Baseline (ponto)",
    value: { value: "#F5C400" },
  });
  summaryBarColor = new formattingSettings.ColorPicker({
    name: "summaryBarColor", displayName: "Barra Resumo",
    value: { value: "#0E938E" },
  });
  dependencyLineColor = new formattingSettings.ColorPicker({
    name: "dependencyLineColor", displayName: "Linha de Dependência",
    value: { value: "#d1d5db" },
  });

  name = "ganttColors";
  displayName = "Cores";
  slices: FormattingSettingsSlice[] = [
    this.plannedBarColor,
    this.summaryBarColor,
    this.baselineBarColor,
    this.progressColor,
    this.todayLineColor,
    this.milestoneColor,
    this.milestoneBaselineColor,
    this.dependencyLineColor,
  ];
}

// ── Layout ────────────────────────────────────────────────────────────────────
class LayoutCard extends FormattingSettingsCard {
  rowHeight = new formattingSettings.NumUpDown({
    name: "rowHeight", displayName: "Altura da Linha (px)", value: 36,
    options: {
      minValue: { type: powerbi.visuals.ValidatorType.Min, value: 24 },
      maxValue: { type: powerbi.visuals.ValidatorType.Max, value: 100 },
    },
  });
  wbsColumnWidth = new formattingSettings.NumUpDown({
    name: "wbsColumnWidth", displayName: "Largura Coluna WBS (px)", value: 64,
    options: {
      minValue: { type: powerbi.visuals.ValidatorType.Min, value: 0 },
      maxValue: { type: powerbi.visuals.ValidatorType.Max, value: 160 },
    },
  });
  taskLabelWidth = new formattingSettings.NumUpDown({
    name: "taskLabelWidth", displayName: "Largura Label Tarefa (px)", value: 220,
    options: {
      minValue: { type: powerbi.visuals.ValidatorType.Min, value: 80 },
      maxValue: { type: powerbi.visuals.ValidatorType.Max, value: 500 },
    },
  });
  showDependencies = new formattingSettings.ToggleSwitch({
    name: "showDependencies", displayName: "Exibir Dependências", value: true,
  });
  showToday = new formattingSettings.ToggleSwitch({
    name: "showToday", displayName: "Exibir Linha de Hoje", value: true,
  });
  showBaseline = new formattingSettings.ToggleSwitch({
    name: "showBaseline", displayName: "Exibir Baseline", value: true,
  });
  showStatusLabels = new formattingSettings.ToggleSwitch({
    name: "showStatusLabels", displayName: "Exibir Labels de Status", value: true,
  });
  showStatusBar = new formattingSettings.ToggleSwitch({
    name: "showStatusBar", displayName: "Exibir Barra de Status", value: true,
  });
  showWbs = new formattingSettings.ToggleSwitch({
    name: "showWbs", displayName: "Exibir WBS", value: true,
  });

  name = "ganttLayout";
  displayName = "Layout";
  slices: FormattingSettingsSlice[] = [
    this.rowHeight, this.wbsColumnWidth, this.taskLabelWidth,
    this.showDependencies, this.showToday, this.showBaseline,
    this.showStatusLabels, this.showStatusBar, this.showWbs,
  ];
}

export class GanttFormattingSettings extends FormattingSettingsModel {
  colors = new ColorsCard();
  layout = new LayoutCard();
  cards  = [this.colors, this.layout];
}