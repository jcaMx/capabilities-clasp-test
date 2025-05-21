
function onOpen() {
  SlidesApp.getUi()
    .createMenu("Automation Tools")
    .addItem("Create Slides from Inline JSON", "createSlidesFromJson")
    .addItem("Create Slides from capabilities.json File", "createSlidesFromJsonFile")
    .addItem("Create Slides from Content Sheet", "createSlidesFromSheet")
    .addSeparator()
    .addItem("Log All Layout Names", "logAllLayoutNames")
    .addItem("Label Placeholders in Slide", "debugPlaceholdersInTargetLayout")
    .addItem("Create Test Slides for All Layouts", "createTestSlidesForLayouts")
    .addToUi();
}

function createSlidesFromJson() {
  const data = [
    {
      "capability": "Inform",
      "scenario": "Nancy is preparing for a workshop on online safety for seniors and wants to include the latest scam trends targeting older adults.",
      "solution": "She prompts her AI assistant to summarize recent scams and phishing tactics affecting older adults, gathering key points and sources to include in her presentation."
    },
    {
      "capability": "Create & Edit",
      "scenario": "Nancy needs a clear, easy-to-understand guide for using iPhone emergency features for her clients.",
      "solution": "She drafts the guide using plain language, then uses the AI assistant to simplify technical terms and improve flow, ensuring accessibility for her senior audience."
    },
    {
      "capability": "Organize",
      "scenario": "Nancy has notes from several client sessions and wants to track common issues by device type.",
      "solution": "She inputs the notes into the AI assistant, which categorizes them by device and issue type, creating a useful reference for future training materials."
    },
    {
      "capability": "Transform",
      "scenario": "Nancy wants to turn her written instructions for using Zoom into a visual one-page handout.",
      "solution": "She asks the AI assistant to extract step-by-step actions and convert them into a simplified instructional layout with headings and visuals."
    },
    {
      "capability": "Analyze",
      "scenario": "Nancy is deciding whether to expand her services to include group Zoom classes for assisted-living centers.",
      "solution": "She uses the AI assistant to weigh pros and cons, compare costs and potential benefits, and help her identify key success factors."
    },
    {
      "capability": "Personify or Simulate",
      "scenario": "Nancy wants to practice handling a reluctant client who is nervous about using email for the first time.",
      "solution": "She asks the AI assistant to role-play as a skeptical senior, allowing her to rehearse compassionate and clear teaching responses."
    },
    {
      "capability": "Explore & Guide",
      "scenario": "Nancy is considering partnering with a local senior living community to offer tech workshops but isn’t sure how to start.",
      "solution": "She asks the AI assistant to outline a partnership proposal and suggest key benefits to highlight in her pitch to facility managers."
    }
  ];

  const presentation = SlidesApp.getActivePresentation();
  const layout = presentation.getLayouts().find(l => l.getLayoutName() === "CUSTOM_1_2_1_2");

  if (!layout) {
    Logger.log("Layout not found");
    return;
  }

  data.forEach(entry => {
    const slide = presentation.appendSlide(layout);
    const textShapes = slide.getShapes()
      .filter(shape => shape.getShapeType() === SlidesApp.ShapeType.TEXT_BOX)
      .map(shape => ({ shape: shape, left: shape.getLeft() }))
      .sort((a, b) => a.left - b.left);

    if (textShapes.length < 3) return;

    const titleShape = textShapes[0].shape;
    const leftBodyShape = textShapes[1].shape;
    const rightBodyShape = textShapes[2].shape;

    titleShape.getText().setText(entry.capability);
    leftBodyShape.getText().setText(entry.scenario);
    rightBodyShape.getText().setText(entry.solution);
  });
}

function createSlidesFromJsonFile() {
  const presentation = SlidesApp.getActivePresentation();
  const presentationId = presentation.getId();
  const presentationFile = DriveApp.getFileById(presentationId);
  const parentFolder = presentationFile.getParents().next();

  const files = parentFolder.getFilesByName("capabilities.json");
  if (!files.hasNext()) {
    Logger.log("capabilities.json not found in the same folder as the presentation.");
    return;
  }

  const jsonFile = files.next();
  const jsonText = jsonFile.getBlob().getDataAsString();
  let data;
  try {
    data = JSON.parse(jsonText);
  } catch (e) {
    Logger.log("Error parsing capabilities.json: " + e);
    return;
  }

  const layout = presentation.getLayouts().find(l => l.getLayoutName() === "CUSTOM_1_2_1_2");
  if (!layout) {
    Logger.log("Layout not found");
    return;
  }

  data.forEach(entry => {
    const slide = presentation.appendSlide(layout);
    const textShapes = slide.getShapes()
      .filter(shape => shape.getShapeType() === SlidesApp.ShapeType.TEXT_BOX)
      .map(shape => ({ shape: shape, left: shape.getLeft() }))
      .sort((a, b) => a.left - b.left);

    if (textShapes.length < 3) return;

    const titleShape = textShapes[0].shape;
    const leftBodyShape = textShapes[1].shape;
    const rightBodyShape = textShapes[2].shape;

    titleShape.getText().setText(entry.capability);
    leftBodyShape.getText().setText(entry.scenario);
    rightBodyShape.getText().setText(entry.solution);
  });
}

function createSlidesFromSheet() {
  const presentation = SlidesApp.getActivePresentation();
  const presentationId = presentation.getId();
  const presentationFile = DriveApp.getFileById(presentationId);
  const parentFolder = presentationFile.getParents().next();

  const files = parentFolder.getFilesByName("Content");
  if (!files.hasNext()) {
    Logger.log("Content spreadsheet not found in the same folder.");
    return;
  }

  const sheetFile = files.next();
  const spreadsheet = SpreadsheetApp.open(sheetFile);
  const sheet = spreadsheet.getSheetByName("Capabilities");

  if (!sheet) {
    Logger.log("Sheet named 'Capabilities' not found in the spreadsheet.");
    return;
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1);

  const scenarioIndex = headers.indexOf("scenario");
  const capabilityIndex = headers.indexOf("capability");
  const solutionIndex = headers.indexOf("solution");

  if (scenarioIndex === -1 || capabilityIndex === -1 || solutionIndex === -1) {
    Logger.log("Required columns (scenario, capability, solution) not found.");
    return;
  }

  const layout = presentation.getLayouts().find(l => l.getLayoutName() === "CUSTOM_1_2_1_2");
  if (!layout) {
    Logger.log("Layout not found");
    return;
  }

  rows.forEach(row => {
    const scenario = row[scenarioIndex];
    const capability = row[capabilityIndex];
    const solution = row[solutionIndex];

    const slide = presentation.appendSlide(layout);
    const textShapes = slide.getShapes()
      .filter(shape => shape.getShapeType() === SlidesApp.ShapeType.TEXT_BOX)
      .map(shape => ({ shape: shape, left: shape.getLeft() }))
      .sort((a, b) => a.left - b.left);

    if (textShapes.length < 3) return;

    const titleShape = textShapes[0].shape;
    const leftBodyShape = textShapes[1].shape;
    const rightBodyShape = textShapes[2].shape;

    titleShape.getText().setText(capability);
    leftBodyShape.getText().setText(scenario);
    rightBodyShape.getText().setText(solution);
  });
}

function logAllLayoutNames() {
  const presentation = SlidesApp.getActivePresentation();
  const layouts = presentation.getLayouts();

  layouts.forEach(layout => {
    Logger.log("Layout name: " + layout.getLayoutName());
  });
}

function debugPlaceholdersInTargetLayout() {
  const presentation = SlidesApp.getActivePresentation();
  const layout = presentation.getLayouts().find(l => l.getLayoutName() === "CUSTOM_1_2_1_2");

  const slide = presentation.appendSlide(layout);
  const shapes = slide.getShapes();

  shapes.forEach((shape, index) => {
    const box = slide.insertTextBox(
      `Shape #${index}\nType: ${shape.getShapeType()}\nLeft: ${shape.getLeft()}\nTop: ${shape.getTop()}`
    );
    box.setLeft(shape.getLeft());
    box.setTop(shape.getTop());

    try {
      const placeholder = shape.getPlaceholder();
      if (placeholder) {
        box.getText().appendText(`\nPlaceholderType: ${placeholder.getType()}`);
      }
    } catch (e) {}

    box.getText().getTextStyle().setFontSize(10);
    box.getText().getTextStyle().setForegroundColor('#ff0000');
  });
}

function createTestSlidesForLayouts() {
  const presentation = SlidesApp.getActivePresentation();
  const layouts = presentation.getLayouts();

  layouts.forEach(layout => {
    const slide = presentation.appendSlide(layout);
    const title = slide.insertTextBox("Layout: " + layout.getLayoutName(), 50, 50, 400, 50);
    title.getText().getTextStyle().setFontSize(14).setBold(true);
  });
}
