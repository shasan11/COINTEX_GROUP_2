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
const { applyPresentationChartFont, finalizePresentation } = await import(
  pathToFileURL(path.join(skillDir, "container_tools", "artifact_tool_utils.mjs")).href,
);

const outputDir = path.join(workspaceDir, "output", "lab2");
const stagingDir = path.join(workspaceDir, ".codex-finalizer-lab2-premium");
const finalPath = path.join(outputDir, "COMP6009_Lab2_Presentation_Premium_v4.pptx");
await fs.mkdir(outputDir, { recursive: true });
await fs.mkdir(stagingDir, { recursive: true });

const FONT = "Arial";
const C = {
  night: "#08111F",
  navy: "#102A43",
  blue: "#2F80ED",
  cyan: "#2EC4B6",
  green: "#159A68",
  yellow: "#F7B500",
  orange: "#F79009",
  red: "#D92D20",
  ink: "#162033",
  slate: "#506176",
  mist: "#EEF3F7",
  line: "#D5DEE8",
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
    fontSize: options.fontSize ?? 24,
    bold: options.bold ?? false,
    color: options.color ?? C.ink,
    alignment: options.alignment ?? "left",
    verticalAlignment: options.verticalAlignment ?? "middle",
    autoFit: options.autoFit ?? "shrinkText",
  };
  return shape;
}

function rect(slide, position, fill, line = fill, radius = "rect") {
  return slide.shapes.add({
    geometry: radius,
    position,
    fill,
    line: { style: "solid", fill: line, width: 1 },
  });
}

function line(slide, x1, y1, x2, y2, color = C.line, width = 2) {
  return slide.shapes.add({
    geometry: "line",
    position: { left: x1, top: y1, width: x2 - x1, height: y2 - y1 },
    fill: "none",
    line: { style: "solid", fill: color, width },
  });
}

function title(slide, heading, subheading = null, dark = false) {
  addText(slide, heading, { left: 70, top: 40, width: 1140, height: 58 }, {
    fontSize: 38, bold: true, color: dark ? C.white : C.navy, autoFit: "none",
  });
  if (subheading) {
    addText(slide, subheading, { left: 72, top: 100, width: 1120, height: 34 }, {
      fontSize: 18, color: dark ? "#C5D4E3" : C.slate, autoFit: "none",
    });
  }
}

function page(slide, number, dark = false) {
  addText(slide, String(number).padStart(2, "0"), { left: 1170, top: 674, width: 45, height: 20 }, {
    fontSize: 11, color: dark ? "#8FA6BC" : "#8090A0", alignment: "right", autoFit: "none",
  });
}

function note(slide, seconds, script, sources = []) {
  const sourceText = sources.length ? `\nSources: ${sources.join("; ")}` : "";
  slide.speakerNotes.textFrame.setText(`Suggested time: ${seconds} seconds. ${script}${sourceText}`);
}

async function addImage(slide, relativePath, position, alt, fit = "cover") {
  const blob = await fs.readFile(path.join(workspaceDir, relativePath));
  return slide.images.add({
    blob,
    contentType: "image/png",
    alt,
    fit,
    position,
  });
}

function stage(slide, x, number, label, detail, accent = C.blue) {
  rect(slide, { left: x, top: 265, width: 190, height: 220 }, C.white, C.line, "roundRect");
  addText(slide, number, { left: x + 18, top: 282, width: 45, height: 42 }, {
    fontSize: 30, bold: true, color: accent, alignment: "center", autoFit: "none",
  });
  addText(slide, label, { left: x + 20, top: 340, width: 150, height: 42 }, {
    fontSize: 23, bold: true, color: C.navy, alignment: "center", autoFit: "none",
  });
  addText(slide, detail, { left: x + 20, top: 395, width: 150, height: 64 }, {
    fontSize: 16, color: C.slate, alignment: "center",
  });
}

function styledTable(slide, values, frame, widths, options = {}) {
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
    fill: options.headerFill ?? C.navy,
    textStyle: { typeface: FONT, fontSize: options.headerSize ?? 16, bold: true, color: C.white, alignment: "center" },
    margins: { left: 8, right: 8, top: 5, bottom: 5 },
    anchor: "middle",
  });
  for (let r = 1; r < values.length; r += 1) {
    table.cells.block({ row: r, column: 0, rowCount: 1, columnCount: values[0].length }).assign({
      fill: r % 2 === 0 ? "#F3F6F9" : C.white,
      textStyle: { typeface: FONT, fontSize: options.bodySize ?? 15, color: C.ink, alignment: options.bodyAlign ?? "center" },
      margins: { left: 8, right: 8, top: 5, bottom: 5 },
      anchor: "middle",
    });
  }
  return table;
}

// Slide 1: cover
{
  const slide = presentation.slides.add();
  slide.background.fill = C.night;
  await addImage(slide, "cointex_media/04_gameplay.png", { left: 510, top: 0, width: 770, height: 720 }, "CoinTex gameplay", "cover");
  rect(slide, { left: 0, top: 0, width: 595, height: 720 }, C.night);
  rect(slide, { left: 560, top: 0, width: 45, height: 720 }, C.yellow);
  addText(slide, "COINTEX", { left: 78, top: 94, width: 400, height: 44 }, {
    fontSize: 21, bold: true, color: C.yellow, autoFit: "none",
  });
  addText(slide, "Safer decisions\nwith a genetic\nalgorithm", { left: 76, top: 150, width: 450, height: 245 }, {
    fontSize: 53, bold: true, color: C.white, autoFit: "shrinkText",
  });
  addText(slide, "COMP6009 Lab 2", { left: 80, top: 430, width: 360, height: 40 }, {
    fontSize: 24, color: "#C5D4E3", autoFit: "none",
  });
  addText(slide, "Student name and ID: ____________________\nGroup members: __________________________", { left: 80, top: 520, width: 420, height: 78 }, {
    fontSize: 17, color: C.white,
  });
  note(slide, 20, "Introduce the group and state the aim: improve the supplied genetic agent so it collects coins more reliably while avoiding hazards.", ["cointex_media/04_gameplay.png"]);
}

// Slide 2: problem
{
  const slide = presentation.slides.add();
  slide.background.fill = C.white;
  title(slide, "The control problem", "Every chromosome proposes the character's next destination");
  await addImage(slide, "cointex_media/07_auto_player.png", { left: 70, top: 160, width: 575, height: 430 }, "CoinTex Auto Player settings", "cover");
  addText(slide, "[ x , y ]", { left: 735, top: 165, width: 380, height: 92 }, {
    fontSize: 58, bold: true, color: C.blue, alignment: "center", autoFit: "none",
  });
  addText(slide, "One chromosome stores one target point", { left: 725, top: 260, width: 400, height: 44 }, {
    fontSize: 22, bold: true, color: C.navy, alignment: "center", autoFit: "none",
  });
  line(slide, 735, 330, 1115, 330, C.line, 2);
  addText(slide, "The game changes after every movement. Coins disappear, hazards move, and the best target changes.", { left: 735, top: 355, width: 390, height: 110 }, {
    fontSize: 23, color: C.ink, alignment: "center",
  });
  addText(slide, "The agent must search again after every action", { left: 735, top: 500, width: 390, height: 56 }, {
    fontSize: 22, bold: true, color: C.green, alignment: "center",
  });
  page(slide, 2);
  note(slide, 35, "Explain that a chromosome is not a full route. It stores one immediate target. This makes the task a repeated local planning problem.", ["cointex_media/07_auto_player.png", "main_updated.py"]);
}

// Slide 3: GA loop
{
  const slide = presentation.slides.add();
  slide.background.fill = C.mist;
  title(slide, "Genetic control loop", "The agent evaluates a fresh game snapshot before every movement");
  stage(slide, 55, "1", "Sense", "Read player, coins and hazards");
  stage(slide, 300, "2", "Score", "Evaluate 60 candidate targets");
  stage(slide, 545, "3", "Evolve", "Select, cross and mutate");
  stage(slide, 790, "4", "Move", "Animate toward the best target", C.green);
  stage(slide, 1035, "5", "Repeat", "Read the changed game state");
  for (const x of [245, 490, 735, 980]) line(slide, x, 375, x + 50, 375, C.blue, 3);
  addText(slide, "This loop continues until all coins are collected, time expires, or health reaches zero", { left: 160, top: 555, width: 960, height: 62 }, {
    fontSize: 24, bold: true, color: C.navy, alignment: "center",
  });
  page(slide, 3);
  note(slide, 45, "Walk from left to right. PyGAD evaluates candidate targets, produces a new population, and sends the best coordinate to Kivy for animation.", ["main_updated.py", "PlayerGA/ga_core.py"]);
}

// Slide 4: baseline failures
{
  const slide = presentation.slides.add();
  slide.background.fill = C.night;
  title(slide, "Why the supplied agent fails", "The original operators and fitness function do not match the game", true);
  const issues = [
    ["10", "chromosomes", "The search covers little of the screen"],
    ["Random", "parent selection", "Weak candidates reproduce too often"],
    ["1 to 10", "mutation range", "Targets can leave the normalized screen"],
    ["0", "hazard effect", "Some monster and fire branches change nothing"],
  ];
  issues.forEach((item, i) => {
    const x = 75 + i * 295;
    addText(slide, item[0], { left: x, top: 190, width: 245, height: 82 }, {
      fontSize: 43, bold: true, color: i === 3 ? C.red : C.yellow, alignment: "center", autoFit: "none",
    });
    addText(slide, item[1], { left: x, top: 270, width: 245, height: 45 }, {
      fontSize: 20, bold: true, color: C.white, alignment: "center", autoFit: "none",
    });
    addText(slide, item[2], { left: x + 10, top: 332, width: 225, height: 76 }, {
      fontSize: 17, color: "#C5D4E3", alignment: "center",
    });
    if (i < 3) line(slide, x + 270, 190, x + 270, 410, "#29445D", 1);
  });
  rect(slide, { left: 175, top: 500, width: 930, height: 96 }, "#132A3D", "#29445D", "roundRect");
  addText(slide, "A safe endpoint can still require a route through danger", { left: 220, top: 517, width: 840, height: 60 }, {
    fontSize: 25, bold: true, color: C.white, alignment: "center",
  });
  page(slide, 4, true);
  note(slide, 45, "Explain the four code-level faults. Finish with the most important fitness problem: endpoint safety does not protect the route between the player and target.", ["Supplied main_updated.py baseline", "PlayerGA/ga_core.py baseline_fitness"]);
}

// Slide 5: fitness redesign
{
  const slide = presentation.slides.add();
  slide.background.fill = C.white;
  title(slide, "Fitness redesign", "The new score balances coin progress with full-path safety");
  await addImage(slide, "cointex_media/05_combat_chasers.png", { left: 650, top: 150, width: 560, height: 440 }, "CoinTex player facing a chasing monster", "cover");
  addText(slide, "Score =", { left: 78, top: 170, width: 145, height: 50 }, {
    fontSize: 31, bold: true, color: C.navy, autoFit: "none",
  });
  const terms = [
    ["Coin progress", "All remaining coins compete as goals", C.blue, "+"],
    ["Path clearance", "Distance from each hazard to the movement line", C.green, "+"],
    ["Escape reward", "Nearby danger rewards movement away", C.orange, "+"],
    ["Risk and distance", "Unsafe paths, off-screen targets and long moves", C.red, "−"],
  ];
  terms.forEach((item, i) => {
    const top = 245 + i * 88;
    addText(slide, item[3], { left: 82, top, width: 48, height: 48 }, {
      fontSize: 34, bold: true, color: item[2], alignment: "center", autoFit: "none",
    });
    addText(slide, item[0], { left: 145, top, width: 210, height: 36 }, {
      fontSize: 22, bold: true, color: C.navy, autoFit: "none",
    });
    addText(slide, item[1], { left: 145, top: top + 35, width: 410, height: 40 }, {
      fontSize: 17, color: C.slate,
    });
  });
  addText(slide, "Safety is measured along the whole route", { left: 700, top: 610, width: 460, height: 40 }, {
    fontSize: 21, bold: true, color: C.green, alignment: "center", autoFit: "none",
  });
  page(slide, 5);
  note(slide, 55, "Describe the four score components. Emphasize point_segment_distance: a coin can be safe at its endpoint while the direct route crosses danger.", ["PlayerGA/ga_core.py improved_fitness", "cointex_media/05_combat_chasers.png"]);
}

// Slide 6: tuning
{
  const slide = presentation.slides.add();
  slide.background.fill = C.mist;
  title(slide, "Parameter tuning", "Nine settings were tested on levels 20 and 40 with fixed seeds");
  styledTable(slide, [
    ["Population", "Mutation 0.15", "Mutation 0.30", "Mutation 0.50"],
    ["60", "62.5%", "100.0%", "75.0%"],
    ["120", "81.25%", "75.0%", "81.25%"],
    ["240", "87.5%", "87.5%", "62.5%"],
  ], { left: 70, top: 185, width: 760, height: 330 }, [190, 190, 190, 190], { headerSize: 17, bodySize: 22 });
  rect(slide, { left: 880, top: 185, width: 320, height: 330 }, C.navy, C.navy, "roundRect");
  addText(slide, "Selected", { left: 930, top: 220, width: 220, height: 38 }, {
    fontSize: 21, bold: true, color: C.cyan, alignment: "center", autoFit: "none",
  });
  addText(slide, "60", { left: 930, top: 270, width: 220, height: 100 }, {
    fontSize: 68, bold: true, color: C.white, alignment: "center", autoFit: "none",
  });
  addText(slide, "population", { left: 930, top: 360, width: 220, height: 34 }, {
    fontSize: 20, color: "#C5D4E3", alignment: "center", autoFit: "none",
  });
  addText(slide, "Mutation 0.30", { left: 930, top: 420, width: 220, height: 44 }, {
    fontSize: 24, bold: true, color: C.yellow, alignment: "center", autoFit: "none",
  });
  addText(slide, "100% tuning success with less computation than larger populations", { left: 185, top: 555, width: 910, height: 52 }, {
    fontSize: 24, bold: true, color: C.green, alignment: "center",
  });
  page(slide, 6);
  note(slide, 40, "Read the tuning table by rows. Population 60 with mutation 0.30 completed all 16 tuning trials and avoided the extra cost of 120 or 240 chromosomes.", ["output/lab2/tuning_summary.csv"]);
}

// Slide 7: experiment design
{
  const slide = presentation.slides.add();
  slide.background.fill = C.white;
  title(slide, "Controlled experiment", "Ablation separates fitness gains from operator gains");
  styledTable(slide, [
    ["Configuration", "Fitness", "Operators"],
    ["Baseline", "Original", "Original"],
    ["Fitness only", "Improved", "Original"],
    ["Operators only", "Original", "Improved"],
    ["Full system", "Improved", "Improved"],
  ], { left: 70, top: 180, width: 720, height: 360 }, [260, 230, 230], { headerSize: 17, bodySize: 18, bodyAlign: "left" });
  addText(slide, "320", { left: 900, top: 180, width: 250, height: 110 }, {
    fontSize: 72, bold: true, color: C.blue, alignment: "center", autoFit: "none",
  });
  addText(slide, "controlled trials", { left: 900, top: 285, width: 250, height: 38 }, {
    fontSize: 23, bold: true, color: C.navy, alignment: "center", autoFit: "none",
  });
  addText(slide, "Levels 1, 20, 40 and 60\n20 fixed seeds for every comparison\nEqual 60-member populations", { left: 850, top: 350, width: 350, height: 140 }, {
    fontSize: 20, color: C.ink, alignment: "center",
  });
  rect(slide, { left: 160, top: 580, width: 960, height: 54 }, "#FFF4D6", "#F2CC63", "roundRect");
  addText(slide, "Measured outcomes: success, completion time, remaining health and coin completion", { left: 190, top: 588, width: 900, height: 38 }, {
    fontSize: 20, bold: true, color: "#7A4A00", alignment: "center", autoFit: "none",
  });
  page(slide, 7);
  note(slide, 45, "Explain the ablation design. Fixed seeds expose every configuration to equivalent conditions. State clearly that this is a deterministic simulator based on CoinTex level parameters.", ["PlayerGA/benchmark.py", "output/lab2/benchmark_metadata.json"]);
}

// Slide 8: results
{
  const slide = presentation.slides.add();
  slide.background.fill = C.white;
  title(slide, "Success rose by 46.25 percentage points", "Average success increased from 25.00% to 71.25% across four tested levels");
  const chart = slide.charts.add("bar", {
    position: { left: 55, top: 170, width: 820, height: 430 },
    categories: ["L1", "L20", "L40", "L60"],
    series: [
      { name: "Baseline", values: [90, 10, 0, 0], fill: "#A5B0BD" },
      { name: "Improved", values: [100, 90, 85, 10], fill: C.green },
    ],
    barOptions: { direction: "column", grouping: "clustered", gapWidth: 55 },
    hasLegend: true,
    legend: { position: "bottom", textStyle: { typeface: FONT, fontSize: 15, fill: C.slate } },
    yAxis: { min: 0, max: 100, majorUnit: 20, numberFormatCode: "0\"%\"", textStyle: { typeface: FONT, fontSize: 13, fill: C.slate }, majorGridlines: { style: "solid", fill: "#E3E8EE", width: 1 } },
    xAxis: { textStyle: { typeface: FONT, fontSize: 14, fill: C.slate }, line: { style: "solid", fill: C.line, width: 1 } },
    dataLabels: { showValue: true, position: "outEnd", textStyle: { typeface: FONT, fontSize: 13, fill: C.ink, bold: true } },
    chartFill: C.white,
    chartLine: { fill: "none", width: 0 },
    plotAreaFill: C.white,
    plotAreaLine: { fill: "none", width: 0 },
  });
  applyPresentationChartFont(chart, { fontFamily: FONT });
  addText(slide, "+46.25", { left: 925, top: 200, width: 275, height: 110 }, {
    fontSize: 58, bold: true, color: C.green, alignment: "center", autoFit: "none",
  });
  addText(slide, "percentage points", { left: 925, top: 305, width: 275, height: 35 }, {
    fontSize: 20, color: C.green, alignment: "center", autoFit: "none",
  });
  line(slide, 950, 380, 1175, 380, C.line, 2);
  addText(slide, "Level 40", { left: 945, top: 405, width: 235, height: 38 }, {
    fontSize: 25, bold: true, color: C.navy, alignment: "center", autoFit: "none",
  });
  addText(slide, "0% to 85%", { left: 945, top: 450, width: 235, height: 48 }, {
    fontSize: 30, bold: true, color: C.blue, alignment: "center", autoFit: "none",
  });
  addText(slide, "Level 60 remained difficult at 10% success", { left: 920, top: 545, width: 290, height: 50 }, {
    fontSize: 18, bold: true, color: C.red, alignment: "center",
  });
  page(slide, 8);
  note(slide, 50, "Lead with the average improvement, then compare the four levels. Level 40 shows the clearest gain. Do not hide Level 60: it remains the main limitation.", ["output/lab2/benchmark_summary.csv"]);
}

// Slide 9: limitation and future work
{
  const slide = presentation.slides.add();
  slide.background.fill = C.night;
  title(slide, "Hardest-level limitation", "One target point cannot describe a safe multi-step route through moving hazards", true);
  await addImage(slide, "cointex_media/02_worldmap.png", { left: 70, top: 215, width: 450, height: 253 }, "CoinTex world map showing the difficulty progression", "cover");
  addText(slide, "Hardest tested level: 60", { left: 105, top: 515, width: 380, height: 38 }, {
    fontSize: 26, bold: true, color: C.yellow, alignment: "center", autoFit: "none",
  });
  addText(slide, "10% success   69.0% coin completion   0.97% health", { left: 595, top: 155, width: 585, height: 52 }, {
    fontSize: 23, bold: true, color: C.white, alignment: "center",
  });
  addText(slide, "Proposed hybrid", { left: 710, top: 245, width: 330, height: 42 }, {
    fontSize: 28, bold: true, color: C.cyan, alignment: "center", autoFit: "none",
  });
  const flow = [
    ["GA selects the goal", C.blue],
    ["Short simulation predicts hazard motion", C.green],
    ["Planner returns a safe waypoint", C.yellow],
  ];
  flow.forEach((item, i) => {
    const top = 320 + i * 100;
    rect(slide, { left: 650, top, width: 450, height: 68 }, "#132A3D", item[1], "roundRect");
    addText(slide, item[0], { left: 680, top: top + 12, width: 390, height: 44 }, {
      fontSize: 21, bold: true, color: C.white, alignment: "center", autoFit: "none",
    });
    if (i < 2) line(slide, 875, top + 68, 875, top + 100, item[1], 3);
  });
  addText(slide, "Simulator evidence requires final confirmation in the Windows GUI", { left: 590, top: 625, width: 620, height: 38 }, {
    fontSize: 17, color: "#C5D4E3", alignment: "center", autoFit: "none",
  });
  page(slide, 9, true);
  note(slide, 45, "Explain why the chromosome is the limiting representation. Propose a hybrid: let the GA choose the destination, then simulate a short horizon before accepting the next waypoint. State the simulator caveat.", ["output/lab2/benchmark_summary.csv", "cointex_media/02_worldmap.png"]);
}

// Slide 10: conclusion
{
  const slide = presentation.slides.add();
  slide.background.fill = C.night;
  addText(slide, "What changed", { left: 80, top: 70, width: 440, height: 60 }, {
    fontSize: 44, bold: true, color: C.white, autoFit: "none",
  });
  const items = [
    ["Fitness", "Every coin competes as a goal, and safety covers the whole route"],
    ["Operators", "Tournament selection, elitism and adaptive mutation stabilize the search"],
    ["Evidence", "Average success increased from 25.00% to 71.25% in 320 trials"],
    ["Next step", "Validate selected levels in the live Windows GUI and record screenshots"],
  ];
  items.forEach((item, i) => {
    const top = 175 + i * 102;
    addText(slide, item[0], { left: 90, top, width: 185, height: 48 }, {
      fontSize: 23, bold: true, color: i === 3 ? C.yellow : C.cyan, autoFit: "none",
    });
    addText(slide, item[1], { left: 295, top, width: 820, height: 58 }, {
      fontSize: 22, color: C.white,
    });
  });
  addText(slide, "Questions", { left: 420, top: 620, width: 440, height: 54 }, {
    fontSize: 38, bold: true, color: C.yellow, alignment: "center", autoFit: "none",
  });
  note(slide, 30, "Summarize the four lines and invite questions. Each group member should be ready to explain one improvement and one limitation.", ["PlayerGA/ga_core.py", "output/lab2/benchmark_summary.csv"]);
}

const requirements = {
  explicitTotalSlideCount: 10,
  requiredNativeTableOwnerSlides: [6, 7],
  requiredNativeChartOwnerSlides: [8],
  requiredEmbeddedWorkbookChartOwnerSlides: [],
  materializeLiteralChartWorkbooks: true,
  nativeChartTargetApplication: "powerpoint",
};
const fontPolicy = { basis: "design", families: [FONT] };
const expectedSlideSizeEmu = "12192000,6858000";
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
    "--require-native-table-slide", "6",
    "--require-native-table-slide", "7",
  ],
  requiredNativeTableOwnerSlides: requirements.requiredNativeTableOwnerSlides,
  requiredNativeChartOwnerSlides: requirements.requiredNativeChartOwnerSlides,
  fontPolicy,
  verifyArtifactToolImport: true,
  materializeLiteralChartWorkbooks: true,
  nativeChartTargetApplication: "powerpoint",
  receiptPath: path.join(stagingDir, "COMP6009_Lab2_Presentation_Premium_v4.validation.json"),
});

console.log(JSON.stringify({ finalPath, result }, null, 2));
