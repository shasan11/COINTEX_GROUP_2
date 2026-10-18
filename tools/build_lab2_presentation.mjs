import fs from "node:fs/promises";
import path from "node:path";
import { pathToFileURL } from "node:url";

const workspaceDir = "C:\\Users\\HP\\Desktop\\cointex";
const runtimeRoot = "C:\\Users\\HP\\.cache\\codex-runtimes\\codex-primary-runtime\\dependencies";
const runtimeModules = path.join(runtimeRoot, "node", "node_modules");
const runtimePython = path.join(runtimeRoot, "python", "python.exe");
const skillDir = "C:\\Users\\HP\\.codex\\plugins\\cache\\openai-primary-runtime\\presentations\\26.909.12148\\skills\\presentations";
process.env.RUNTIME_NODE = path.join(runtimeRoot, "node", "bin", "node.exe");
process.env.RUNTIME_NODE_MODULES = runtimeModules;
process.env.RUNTIME_BIN_DIR = path.join(runtimeRoot, "bin", "override");
process.env.RUNTIME_PYTHON = runtimePython;
const artifactModule = path.join(runtimeModules, "@oai", "artifact-tool", "dist", "artifact_tool.mjs");
const { Presentation, PresentationFile } = await import(pathToFileURL(artifactModule).href);
const {
  applyPresentationChartFont,
  finalizePresentation,
  makeNativeBulletParagraphs,
} = await import(pathToFileURL(path.join(skillDir, "container_tools", "artifact_tool_utils.mjs")).href);

const buildDir = path.join(workspaceDir, "tmp", "lab2_presentation_build");
const outputDir = path.join(workspaceDir, "output", "lab2");
const finalPath = path.join(outputDir, "COMP6009_Lab2_Presentation_Final.pptx");
await fs.mkdir(buildDir, { recursive: true });
await fs.mkdir(outputDir, { recursive: true });

const FONT = "Arial";
const C = {
  navy: "#17365D",
  blue: "#2F75B5",
  sky: "#DCEAF7",
  green: "#16865C",
  paleGreen: "#E5F4ED",
  orange: "#D97706",
  paleOrange: "#FFF2DB",
  red: "#B42318",
  ink: "#172033",
  gray: "#526173",
  line: "#D6DEE7",
  pale: "#F5F7FA",
  white: "#FFFFFF",
};

const presentation = Presentation.create({ slideSize: { width: 1280, height: 720 } });

function addText(slide, text, position, options = {}) {
  const shape = slide.shapes.add({
    geometry: "textbox",
    position,
    fill: "none",
    line: { fill: "none", width: 0 },
  });
  shape.text = text;
  shape.text.style = {
    typeface: FONT,
    fontSize: options.fontSize ?? 25,
    bold: options.bold ?? false,
    color: options.color ?? C.ink,
    alignment: options.alignment ?? "left",
    verticalAlignment: options.verticalAlignment ?? "middle",
    autoFit: options.autoFit ?? "shrinkText",
  };
  return shape;
}

function addBox(slide, position, fill, line = C.line, radius = "roundRect") {
  return slide.shapes.add({
    geometry: radius,
    position,
    fill,
    line: { style: "solid", fill: line, width: 1.2 },
  });
}

function title(slide, text, subtitle = null) {
  addText(slide, text, { left: 64, top: 38, width: 1152, height: 60 }, {
    fontSize: 36, bold: true, color: C.navy, autoFit: "none",
  });
  if (subtitle) {
    addText(slide, subtitle, { left: 66, top: 96, width: 1120, height: 36 }, {
      fontSize: 17, color: C.gray, autoFit: "none",
    });
  }
}

function footer(slide, number) {
  addText(slide, `COMP6009 Lab 2    ${number}`, { left: 64, top: 683, width: 1152, height: 20 }, {
    fontSize: 11, color: "#778493", alignment: "right", autoFit: "none",
  });
}

function bullets(slide, items, position, options = {}) {
  const shape = slide.shapes.add({
    geometry: "textbox", position, fill: "none", line: { fill: "none", width: 0 },
  });
  shape.text = makeNativeBulletParagraphs(items, {
    marginLeftPoints: options.marginLeftPoints ?? 20,
    hangingPoints: options.hangingPoints ?? 10,
    spaceAfterPoints: options.spaceAfterPoints ?? 12,
  });
  shape.text.style = {
    typeface: FONT,
    fontSize: options.fontSize ?? 23,
    color: options.color ?? C.ink,
    autoFit: "shrinkText",
  };
  return shape;
}

function notes(slide, text) {
  slide.speakerNotes.textFrame.setText(text);
}

function addStyledTable(slide, values, frame, widths) {
  const table = slide.tables.add({
    rows: values.length,
    columns: values[0].length,
    left: frame.left,
    top: frame.top,
    width: frame.width,
    height: frame.height,
    columnWidths: widths,
    values,
  });
  table.borders.assign({ style: "solid", fill: C.line, width: 1 });
  table.cells.block({ row: 0, column: 0, rowCount: 1, columnCount: values[0].length }).assign({
    fill: C.navy,
    textStyle: { typeface: FONT, fontSize: 16, bold: true, color: C.white, alignment: "center" },
    margins: { left: 8, right: 8, top: 5, bottom: 5 },
    anchor: "middle",
  });
  for (let r = 1; r < values.length; r += 1) {
    table.cells.block({ row: r, column: 0, rowCount: 1, columnCount: values[0].length }).assign({
      fill: r % 2 === 0 ? C.pale : C.white,
      textStyle: { typeface: FONT, fontSize: 15, color: C.ink, alignment: "left" },
      margins: { left: 8, right: 8, top: 5, bottom: 5 },
      anchor: "middle",
    });
  }
  return table;
}

// Slide 1
{
  const slide = presentation.slides.add();
  slide.background.fill = C.navy;
  const logo = await fs.readFile(path.join(workspaceDir, "cointex_logo.png"));
  slide.images.add({
    blob: logo, contentType: "image/png", alt: "CoinTex logo", fit: "contain",
    position: { left: 92, top: 86, width: 220, height: 220 },
  });
  addText(slide, "CoinTex Genetic Algorithm Improvement", { left: 372, top: 104, width: 810, height: 150 }, {
    fontSize: 46, bold: true, color: C.white, autoFit: "shrinkText",
  });
  addText(slide, "COMP6009 Lab 2", { left: 376, top: 274, width: 700, height: 54 }, {
    fontSize: 26, color: "#CFE2F3", autoFit: "none",
  });
  addText(slide, "Student name and ID: ____________________\nGroup members: __________________________", { left: 376, top: 402, width: 720, height: 100 }, {
    fontSize: 20, color: C.white,
  });
  addText(slide, "Improved fitness, tuned operators, reproducible evidence", { left: 92, top: 615, width: 1096, height: 42 }, {
    fontSize: 18, color: "#B9D7EA", alignment: "center", autoFit: "none",
  });
  notes(slide, "Suggested time: 20 seconds. Introduce the group and state the objective: improve the genetic agent so it collects all coins more reliably. Assign one member to open and another to explain the algorithm.");
}

// Slide 2
{
  const slide = presentation.slides.add();
  slide.background.fill = C.white;
  title(slide, "Problem and Baseline", "The supplied agent moves toward coins but does not avoid danger effectively");
  const screenshot = await fs.readFile(path.join(workspaceDir, "cointex_media", "04_gameplay.png"));
  slide.images.add({
    blob: screenshot, contentType: "image/png", alt: "CoinTex gameplay with coins, monsters and fire",
    fit: "cover", geometry: "roundRect", borderRadius: "rounded-xl",
    position: { left: 64, top: 160, width: 540, height: 430 },
  });
  bullets(slide, [
    "Chromosome [x, y] represents the next movement target",
    "Blackboard population contains only 10 chromosomes",
    "Random parent selection provides weak selection pressure",
    "Mutation can generate coordinates from 1 to 10, outside the screen",
    "Monster and fire branches add or subtract zero",
  ], { left: 660, top: 168, width: 520, height: 390 }, { fontSize: 22, spaceAfterPoints: 11 });
  addBox(slide, { left: 660, top: 576, width: 520, height: 66 }, C.paleOrange, "#F1C27D");
  addText(slide, "Observed design risk: coin attraction can select a path through a hazard", { left: 680, top: 585, width: 480, height: 48 }, {
    fontSize: 18, bold: true, color: "#8A4B08", alignment: "center",
  });
  footer(slide, 2);
  notes(slide, "Suggested time: 45 seconds. Explain that the original chromosome is valid, but the fitness and operators do not match the game's coordinate bounds or hazards. Mention that swap mutation is weak for only two genes.");
}

// Slide 3
{
  const slide = presentation.slides.add();
  slide.background.fill = C.white;
  title(slide, "Genetic Control Loop", "Repeated local planning, not one complete route");
  const labels = [
    ["1", "Sense", "Read player, coins and hazards"],
    ["2", "Evaluate", "Score candidate [x, y] targets"],
    ["3", "Evolve", "Select, cross and mutate"],
    ["4", "Act", "Move toward the best target"],
    ["5", "Repeat", "Search again after motion"],
  ];
  labels.forEach((item, i) => {
    const left = 58 + i * 244;
    addBox(slide, { left, top: 212, width: 204, height: 250 }, i === 3 ? C.paleGreen : C.sky, i === 3 ? "#7ABF9E" : "#8CB5D8");
    addText(slide, item[0], { left: left + 70, top: 228, width: 64, height: 64 }, {
      fontSize: 31, bold: true, color: i === 3 ? C.green : C.navy, alignment: "center", autoFit: "none",
    });
    addText(slide, item[1], { left: left + 18, top: 312, width: 168, height: 42 }, {
      fontSize: 24, bold: true, color: C.ink, alignment: "center", autoFit: "none",
    });
    addText(slide, item[2], { left: left + 18, top: 366, width: 168, height: 72 }, {
      fontSize: 17, color: C.gray, alignment: "center",
    });
  });
  addText(slide, "The game changes after every action, so all candidates must be scored against a fresh snapshot", { left: 168, top: 520, width: 944, height: 70 }, {
    fontSize: 24, bold: true, color: C.navy, alignment: "center",
  });
  footer(slide, 3);
  notes(slide, "Suggested time: 45 seconds. Explain the receding-target loop from left to right. Emphasize that one chromosome contains a target, not a full path. This limitation becomes important on the hardest level.");
}

// Slide 4
{
  const slide = presentation.slides.add();
  slide.background.fill = C.white;
  title(slide, "Submitted Algorithm", "Fitness and operators address different failure modes");
  addStyledTable(slide, [
    ["Component", "Baseline", "Submitted change"],
    ["Coin objective", "Nearest coin by simple distance", "All coins compete as reachable goals"],
    ["Safety", "Endpoint square; zero effective penalty", "Clearance along the full movement path"],
    ["Time pressure", "None", "Coin pressure increases as time falls"],
    ["Selection", "Random", "Tournament of four"],
    ["Crossover", "Default", "Scattered in PyGAD; blend in benchmark"],
    ["Mutation", "Swap", "Adaptive replacement within bounds"],
    ["Memory", "No explicit elitism", "Best three targets retained"],
  ], { left: 62, top: 150, width: 1156, height: 460 }, [260, 382, 514]);
  addText(slide, "Population 60, parents 15, elites 3, mutation center 0.30", { left: 120, top: 627, width: 1040, height: 38 }, {
    fontSize: 20, bold: true, color: C.green, alignment: "center", autoFit: "none",
  });
  footer(slide, 4);
  notes(slide, "Suggested time: 55 seconds. Divide this slide between two members if useful. One member explains the fitness changes. Another explains tournament selection, elitism, crossover and adaptive mutation.");
}

// Slide 5
{
  const slide = presentation.slides.add();
  slide.background.fill = C.white;
  title(slide, "Experiment Design", "Fixed seeds separate fitness gains from operator gains");
  addStyledTable(slide, [
    ["Configuration", "Fitness", "Operators"],
    ["Baseline", "Original", "Baseline"],
    ["Fitness only", "Improved", "Baseline"],
    ["Operators only", "Original", "Improved"],
    ["Improved", "Improved", "Improved"],
  ], { left: 66, top: 164, width: 710, height: 300 }, [230, 220, 260]);
  addBox(slide, { left: 830, top: 164, width: 350, height: 300 }, C.pale, C.line);
  addText(slide, "320 trials", { left: 865, top: 190, width: 280, height: 60 }, {
    fontSize: 38, bold: true, color: C.navy, alignment: "center", autoFit: "none",
  });
  addText(slide, "4 levels\n20 fixed seeds\n4 configurations", { left: 885, top: 270, width: 240, height: 150 }, {
    fontSize: 23, color: C.ink, alignment: "center",
  });
  bullets(slide, [
    "KPIs: success, elapsed time, remaining health and coin completion",
    "Levels 1, 20, 40 and 60 span the current difficulty curve",
    "Equal 60-member populations make the main comparison fair",
    "Separate nine-setting grid tunes population and mutation",
  ], { left: 80, top: 500, width: 1100, height: 148 }, { fontSize: 20, spaceAfterPoints: 6 });
  footer(slide, 5);
  notes(slide, "Suggested time: 45 seconds. State that the benchmark is a deterministic simulator using current CoinTex level parameters. It supports controlled comparison but does not replace final GUI play testing.");
}

// Slide 6
{
  const slide = presentation.slides.add();
  slide.background.fill = C.white;
  title(slide, "Success Rate Improved by 46.25 Percentage Points", "Average success increased from 25.0% to 71.25% across the four tested levels");
  const chart = slide.charts.add("bar", {
    position: { left: 72, top: 164, width: 820, height: 430 },
    categories: ["Level 1", "Level 20", "Level 40", "Level 60"],
    series: [
      { name: "Baseline", values: [90, 10, 0, 0], fill: "#9AA6B2" },
      { name: "Improved", values: [100, 90, 85, 10], fill: C.green },
    ],
    barOptions: { direction: "column", grouping: "clustered", gapWidth: 60 },
    hasLegend: true,
    legend: { position: "bottom", overlay: false, textStyle: { typeface: FONT, fontSize: 15, fill: C.gray } },
    xAxis: { textStyle: { typeface: FONT, fontSize: 15, fill: C.gray }, line: { style: "solid", fill: C.line, width: 1 } },
    yAxis: { min: 0, max: 100, majorUnit: 20, numberFormatCode: "0\"%\"", textStyle: { typeface: FONT, fontSize: 14, fill: C.gray }, majorGridlines: { style: "solid", fill: "#E6EBF0", width: 1 } },
    dataLabels: { showValue: true, position: "outEnd", textStyle: { typeface: FONT, fontSize: 14, fill: C.ink, bold: true } },
    chartFill: C.white,
    chartLine: { fill: "none", width: 0 },
    plotAreaFill: C.white,
    plotAreaLine: { fill: "none", width: 0 },
  });
  applyPresentationChartFont(chart, { fontFamily: FONT });
  addBox(slide, { left: 930, top: 190, width: 270, height: 150 }, C.paleGreen, "#8BC4A7");
  addText(slide, "+46.25", { left: 955, top: 208, width: 220, height: 64 }, {
    fontSize: 42, bold: true, color: C.green, alignment: "center", autoFit: "none",
  });
  addText(slide, "percentage points\noverall", { left: 960, top: 278, width: 210, height: 48 }, {
    fontSize: 17, color: C.green, alignment: "center",
  });
  addText(slide, "Level 40\n0% to 85%", { left: 930, top: 380, width: 270, height: 110 }, {
    fontSize: 29, bold: true, color: C.navy, alignment: "center",
  });
  addText(slide, "Level 60 remained difficult at 10% success", { left: 930, top: 532, width: 270, height: 60 }, {
    fontSize: 17, bold: true, color: C.red, alignment: "center",
  });
  footer(slide, 6);
  notes(slide, "Suggested time: 55 seconds. Lead with the measured result. Explain that the improved agent matched or exceeded baseline at every level. Do not claim the hardest level is solved: success there was only 10%. Source: output/lab2/benchmark_summary.csv, 20 fixed seeds per configuration and level.");
}

// Slide 7
{
  const slide = presentation.slides.add();
  slide.background.fill = C.white;
  title(slide, "Parameter Tuning Selected Population 60 and Mutation 0.30", "Only one tested setting completed all 16 middle-level tuning trials");
  const chart = slide.charts.add("bar", {
    position: { left: 66, top: 165, width: 790, height: 430 },
    categories: ["60 / 0.15", "60 / 0.30", "60 / 0.50", "120 / 0.15", "120 / 0.30", "120 / 0.50", "240 / 0.15", "240 / 0.30", "240 / 0.50"],
    series: [{
      name: "Success rate",
      values: [62.5, 100, 75, 81.25, 75, 81.25, 87.5, 87.5, 62.5],
      fill: C.blue,
      points: [{ idx: 1, fill: C.green }],
    }],
    barOptions: { direction: "bar", grouping: "clustered", gapWidth: 35 },
    hasLegend: false,
    xAxis: { visible: false, min: 0, max: 100, majorUnit: 20, majorGridlines: { style: "solid", fill: "#E6EBF0", width: 1 } },
    yAxis: { textStyle: { typeface: FONT, fontSize: 13, fill: C.gray }, line: { style: "solid", fill: C.line, width: 1 } },
    dataLabels: { showValue: true, position: "outEnd", textStyle: { typeface: FONT, fontSize: 12, fill: C.ink, bold: true } },
    chartFill: C.white,
    chartLine: { fill: "none", width: 0 },
    plotAreaFill: C.white,
    plotAreaLine: { fill: "none", width: 0 },
  });
  applyPresentationChartFont(chart, { fontFamily: FONT });
  addBox(slide, { left: 900, top: 185, width: 300, height: 300 }, C.paleGreen, "#8BC4A7");
  addText(slide, "Selected setting", { left: 930, top: 210, width: 240, height: 42 }, {
    fontSize: 21, bold: true, color: C.green, alignment: "center", autoFit: "none",
  });
  addText(slide, "Population 60\nMutation 0.30", { left: 930, top: 275, width: 240, height: 100 }, {
    fontSize: 30, bold: true, color: C.navy, alignment: "center",
  });
  addText(slide, "100% success\n57.86% mean health", { left: 930, top: 395, width: 240, height: 66 }, {
    fontSize: 19, color: C.ink, alignment: "center",
  });
  addText(slide, "Labels show population / mutation; bars show success rate (%)", { left: 140, top: 610, width: 680, height: 28 }, {
    fontSize: 15, color: C.gray, alignment: "center", autoFit: "none",
  });
  footer(slide, 7);
  notes(slide, "Suggested time: 45 seconds. Explain that the grid used levels 20 and 40 with eight fixed seeds each. Larger populations sometimes won faster but used more fitness evaluations and did not improve success. Source: output/lab2/tuning_summary.csv.");
}

// Slide 8
{
  const slide = presentation.slides.add();
  slide.background.fill = C.white;
  title(slide, "Hardest-Level Limitation", "A two-gene endpoint cannot describe a safe multi-step route through fast moving hazards");
  addBox(slide, { left: 72, top: 170, width: 470, height: 390 }, C.paleOrange, "#E5B86E");
  addText(slide, "Level 60", { left: 110, top: 202, width: 394, height: 58 }, {
    fontSize: 38, bold: true, color: C.red, alignment: "center", autoFit: "none",
  });
  addText(slide, "10% success\n69.0% mean coin completion\n0.97% mean health", { left: 120, top: 290, width: 374, height: 170 }, {
    fontSize: 27, bold: true, color: C.ink, alignment: "center",
  });
  addText(slide, "The benchmark is controlled simulation evidence, not proof of identical GUI performance", { left: 108, top: 488, width: 398, height: 48 }, {
    fontSize: 16, color: "#8A4B08", alignment: "center",
  });
  addText(slide, "Proposed hybrid", { left: 650, top: 168, width: 470, height: 48 }, {
    fontSize: 29, bold: true, color: C.navy, alignment: "center", autoFit: "none",
  });
  const hybrid = [
    ["GA target", C.sky, C.navy],
    ["Short-horizon safety simulation", C.paleGreen, C.green],
    ["Safe waypoint", C.sky, C.navy],
  ];
  hybrid.forEach((item, i) => {
    const top = 246 + i * 120;
    addBox(slide, { left: 670, top, width: 430, height: 78 }, item[1], item[2]);
    addText(slide, item[0], { left: 690, top: top + 12, width: 390, height: 54 }, {
      fontSize: 22, bold: true, color: item[2], alignment: "center",
    });
  });
  addText(slide, "Predict danger motion before accepting each waypoint", { left: 680, top: 610, width: 410, height: 42 }, {
    fontSize: 18, color: C.gray, alignment: "center",
  });
  footer(slide, 8);
  notes(slide, "Suggested time: 55 seconds. Explain why the agent still fails. The GA sees one snapshot and proposes one endpoint. A model-predictive safety layer would simulate several short steps and reject unsafe targets before movement. This proposal directly addresses the measured collision failure.");
}

// Slide 9
{
  const slide = presentation.slides.add();
  slide.background.fill = C.navy;
  addText(slide, "What We Changed and Learned", { left: 100, top: 72, width: 1080, height: 70 }, {
    fontSize: 41, bold: true, color: C.white, alignment: "center", autoFit: "none",
  });
  const conclusions = [
    ["Fitness", "Full-path safety and all-coin scoring produced the largest gain"],
    ["Operators", "Tournament selection, elitism and adaptive mutation improved speed and stability"],
    ["Evidence", "Success rose from 25.0% to 71.25% across 320 controlled trials"],
    ["Next step", "Predict moving hazards before committing to each waypoint"],
  ];
  conclusions.forEach((item, i) => {
    const top = 180 + i * 100;
    addText(slide, item[0], { left: 180, top, width: 190, height: 58 }, {
      fontSize: 24, bold: true, color: "#9FD5BD", alignment: "right", autoFit: "none",
    });
    addText(slide, item[1], { left: 410, top, width: 700, height: 58 }, {
      fontSize: 22, color: C.white,
    });
  });
  addText(slide, "Questions", { left: 390, top: 612, width: 500, height: 56 }, {
    fontSize: 34, bold: true, color: "#CFE2F3", alignment: "center", autoFit: "none",
  });
  notes(slide, "Suggested time: 30 seconds, leaving time for questions. Each member should be ready to explain one of these four conclusions and to distinguish simulator results from final GUI validation.");
}

const requirements = {
  explicitTotalSlideCount: 9,
  requiredNativeTableOwnerSlides: [4, 5],
  requiredNativeChartOwnerSlides: [6, 7],
  requiredEmbeddedWorkbookChartOwnerSlides: [],
  materializeLiteralChartWorkbooks: true,
  nativeChartTargetApplication: "powerpoint",
};
const fontPolicy = { basis: "design", families: [FONT] };
const expectedSlideSizeEmu = "12192000,6858000";
const stagingDir = path.join(workspaceDir, ".codex-finalizer-lab2");
await fs.mkdir(stagingDir, { recursive: true });
const candidatePath = path.join(stagingDir, "candidate.pptx");
await (await PresentationFile.exportPptx(presentation)).save(candidatePath);

const result = await finalizePresentation({
  ...requirements,
  workspaceDir,
  candidatePath,
  finalPath,
  pythonExecutable: runtimePython,
  integrityValidatorPath: path.join(skillDir, "container_tools", "inspect_presentation_package_integrity.py"),
  layoutValidatorPath: path.join(skillDir, "container_tools", "inspect_presentation_layout_geometry.py"),
  layoutArgs: [
    "--expected-slide-size-emu", expectedSlideSizeEmu,
    "--validate-bullet-geometry",
    "--validate-heading-fit",
    "--require-native-table-slide", "4",
    "--require-native-table-slide", "5",
  ],
  requiredNativeTableOwnerSlides: requirements.requiredNativeTableOwnerSlides,
  requiredNativeChartOwnerSlides: requirements.requiredNativeChartOwnerSlides,
  fontPolicy,
  verifyArtifactToolImport: true,
  materializeLiteralChartWorkbooks: true,
  nativeChartTargetApplication: "powerpoint",
  receiptPath: path.join(stagingDir, "COMP6009_Lab2_Presentation_Final.validation.json"),
});

console.log(JSON.stringify({ finalPath, result }, null, 2));
