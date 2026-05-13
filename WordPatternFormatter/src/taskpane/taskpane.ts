/* global Office, Word, document, HTMLInputElement, HTMLSelectElement, localStorage, console */

type AlignmentValue = "Left" | "Centered" | "Right" | "Justified";

type TextPattern = {
  name: string;
  fontName: string;
  fontSize: number;
  alignment: AlignmentValue;
  lineSpacing: number;
  bold: boolean;
  italic: boolean;
};

const STORAGE_KEY = "word_text_patterns";
let editingPatternName: string | null = null;

Office.onReady(() => {
  renderSavedPatterns();
  setupTabs();

  document.getElementById("newPatternTab")?.addEventListener("click", () => switchTab("new"));
  document.getElementById("savedPatternsTab")?.addEventListener("click", () => switchTab("saved"));
  document
    .getElementById("readSelectionButton")
    ?.addEventListener("click", readFormattingFromSelection);
  document.getElementById("savePatternButton")?.addEventListener("click", savePattern);
  document.getElementById("applyPatternButton")?.addEventListener("click", applySelectedPattern);
  document.getElementById("deletePatternButton")?.addEventListener("click", deleteSelectedPattern);
  document
    .getElementById("editPatternButton")
    ?.addEventListener("click", loadSelectedPatternForEditing);
});

function setupTabs(): void {
  switchTab("new");
}

function switchTab(tab: "new" | "saved"): void {
  const newTab = document.getElementById("newPatternTab");
  const savedTab = document.getElementById("savedPatternsTab");
  const newPanel = document.getElementById("newPatternPanel");
  const savedPanel = document.getElementById("savedPatternsPanel");
  const isNew = tab === "new";

  newTab?.classList.toggle("active", isNew);
  savedTab?.classList.toggle("active", !isNew);
  newTab?.setAttribute("aria-selected", String(isNew));
  savedTab?.setAttribute("aria-selected", String(!isNew));
  newPanel?.classList.toggle("active", isNew);
  savedPanel?.classList.toggle("active", !isNew);
  newPanel?.toggleAttribute("hidden", !isNew);
  savedPanel?.toggleAttribute("hidden", isNew);
}

function loadSelectedPatternForEditing(): void {
  const pattern = getSelectedPattern();

  if (!pattern) {
    showMessage("Choose a pattern to edit.");
    return;
  }

  editingPatternName = pattern.name;

  setInputValue("patternName", pattern.name);
  setInputValue("fontName", pattern.fontName);
  setInputValue("fontSize", String(pattern.fontSize));
  setInputValue("alignment", pattern.alignment);
  setInputValue("lineSpacing", String(convertWordPointsToLineSpacing(pattern.lineSpacing)));
  setCheckboxValue("bold", pattern.bold);
  setCheckboxValue("italic", pattern.italic);

  switchTab("new");
  showMessage("The pattern was loaded for editing.");
}

function getInputValue(id: string): string {
  const element = document.getElementById(id) as HTMLInputElement | HTMLSelectElement;
  return element.value;
}

function setInputValue(id: string, value: string): void {
  const element = document.getElementById(id) as HTMLInputElement | HTMLSelectElement;

  if (element) {
    element.value = value;
  }
}

function getCheckboxValue(id: string): boolean {
  const element = document.getElementById(id) as HTMLInputElement;
  return element.checked;
}

function setCheckboxValue(id: string, value: boolean): void {
  const element = document.getElementById(id) as HTMLInputElement;

  if (element) {
    element.checked = value;
  }
}

function showMessage(message: string): void {
  const messageElement = document.getElementById("message");

  if (messageElement) {
    messageElement.textContent = message;
  }
}

function getPatterns(): TextPattern[] {
  const data = localStorage.getItem(STORAGE_KEY);

  if (!data) {
    return [];
  }

  try {
    return JSON.parse(data) as TextPattern[];
  } catch {
    return [];
  }
}

function savePatterns(patterns: TextPattern[]): void {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(patterns));
}

function buildPatternFromForm(): TextPattern | null {
  const name = getInputValue("patternName").trim();

  if (!name) {
    showMessage("Enter a pattern name.");
    return null;
  }

  return {
    name: name,
    fontName: getInputValue("fontName"),
    fontSize: Number(getInputValue("fontSize")),
    alignment: getInputValue("alignment") as AlignmentValue,
    lineSpacing: Number(getInputValue("lineSpacing")),
    bold: getCheckboxValue("bold"),
    italic: getCheckboxValue("italic"),
  };
}

function savePattern(): void {
  const pattern = buildPatternFromForm();

  if (!pattern) {
    return;
  }

  const patterns = getPatterns();

  if (editingPatternName) {
    const index = patterns.findIndex((p) => p.name === editingPatternName);

    if (index >= 0) {
      patterns[index] = pattern;
    } else {
      patterns.push(pattern);
    }

    editingPatternName = null;
  } else {
    const existingIndex = patterns.findIndex((p) => p.name === pattern.name);

    if (existingIndex >= 0) {
      patterns[existingIndex] = pattern;
    } else {
      patterns.push(pattern);
    }
  }

  savePatterns(patterns);
  renderSavedPatterns();
  setInputValue("savedPatterns", pattern.name);
  switchTab("saved");

  showMessage("The pattern was saved.");
}

function renderSavedPatterns(): void {
  const select = document.getElementById("savedPatterns") as HTMLSelectElement;

  if (!select) {
    return;
  }

  select.innerHTML = "";

  const patterns = getPatterns();

  if (patterns.length === 0) {
    const option = document.createElement("option");
    option.textContent = "No saved patterns";
    option.value = "";
    select.appendChild(option);
    return;
  }

  patterns.forEach((pattern) => {
    const option = document.createElement("option");
    option.value = pattern.name;
    option.textContent = pattern.name;
    select.appendChild(option);
  });
}

function getSelectedPattern(): TextPattern | null {
  const selectedName = getInputValue("savedPatterns");

  if (!selectedName) {
    return null;
  }

  const patterns = getPatterns();
  return patterns.find((pattern) => pattern.name === selectedName) ?? null;
}

async function readFormattingFromSelection(): Promise<void> {
  try {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();

      // eslint-disable-next-line office-addins/no-navigational-load
      selection.load("font/name,font/size,font/bold,font/italic,paragraphs");

      await context.sync();

      const paragraphs = selection.paragraphs;
      paragraphs.load("items/alignment,items/lineSpacing");

      await context.sync();

      let alignment: AlignmentValue = "Left";
      let lineSpacing = 12;

      if (paragraphs.items.length > 0) {
        const firstParagraph = paragraphs.items[0];
        alignment = firstParagraph.alignment as AlignmentValue;
        lineSpacing = convertWordPointsToLineSpacing(firstParagraph.lineSpacing || 12);
      }

      const fontName = selection.font.name || "Times New Roman";
      const fontSize = selection.font.size || 12;
      const bold = selection.font.bold === true;
      const italic = selection.font.italic === true;

      if (!getInputValue("patternName").trim()) {
        setInputValue("patternName", "Pattern from selection");
      }

      setInputValue("fontName", fontName);
      setInputValue("fontSize", String(fontSize));
      setInputValue("alignment", alignment);
      setInputValue("lineSpacing", String(lineSpacing));
      setCheckboxValue("bold", bold);
      setCheckboxValue("italic", italic);

      showMessage("Formatting was read from the selected text.");
    });
  } catch (error) {
    showMessage("Could not read formatting from the selection.");
    console.error(error);
  }
}

async function applySelectedPattern(): Promise<void> {
  const pattern = getSelectedPattern();

  if (!pattern) {
    showMessage("Choose a saved pattern.");
    return;
  }

  try {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();

      selection.font.name = pattern.fontName;
      selection.font.size = pattern.fontSize;
      selection.font.bold = pattern.bold;
      selection.font.italic = pattern.italic;
      // eslint-disable-next-line office-addins/no-navigational-load
      selection.load("paragraphs");

      await context.sync();

      const paragraphs = selection.paragraphs;
      paragraphs.load("items");

      await context.sync();

      paragraphs.items.forEach((paragraph) => {
        paragraph.alignment = pattern.alignment as Word.Alignment;
        paragraph.lineSpacing = convertLineSpacingToWordPoints(pattern.lineSpacing);
      });

      await context.sync();
    });

    showMessage("The pattern was applied to the selected text.");
  } catch (error) {
    showMessage("An error occurred while applying the pattern.");
    console.error(error);
  }
}

function convertLineSpacingToWordPoints(value: number): number {
  if (!value || Number.isNaN(value)) {
    return 12;
  }

  if (value > 5) {
    return value;
  }

  return value * 12;
}

function convertWordPointsToLineSpacing(value: number): number {
  if (!value || Number.isNaN(value)) {
    return 1;
  }

  if (value > 5) {
    return Number((value / 12).toFixed(1));
  }

  return value;
}

function deleteSelectedPattern(): void {
  const selectedName = getInputValue("savedPatterns");

  if (!selectedName) {
    showMessage("Choose a pattern to delete.");
    return;
  }

  const patterns = getPatterns();
  const updatedPatterns = patterns.filter((pattern) => pattern.name !== selectedName);

  savePatterns(updatedPatterns);
  renderSavedPatterns();
  editingPatternName = null;

  showMessage("The pattern was deleted.");
}
