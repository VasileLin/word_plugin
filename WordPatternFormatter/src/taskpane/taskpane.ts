/* global Office, Word, document, HTMLInputElement, HTMLSelectElement, HTMLElement, localStorage, console, Blob, URL, FileReader, setTimeout, Event */

type AlignmentValue = "Left" | "Centered" | "Right" | "Justified";

type TextPattern = {
  name: string;
  fontName: string;
  fontSize: number;
  alignment: AlignmentValue;
  lineSpacing: number;
  bold: boolean;
  italic: boolean;
  underline: boolean;
  textColor: string;
  highlightColor: string;
  spaceBefore: number;
  spaceAfter: number;
  leftIndent: number;
  rightIndent: number;
  firstLineIndent: number;
};

const STORAGE_KEY = "word_text_patterns";
const LAST_TAB_KEY = "word_text_patterns_last_tab";
const LAST_SCOPE_KEY = "word_text_patterns_last_scope";
const SORT_KEY = "word_text_patterns_sort";
let editingPatternName: string | null = null;
let lastDeleteActionAt = 0;

Office.onReady(() => {
  initializeDefaultPatterns();
  renderSavedPatterns();
  setupTabs();

  document.getElementById("newPatternTab")?.addEventListener("click", () => switchTab("new"));
  document.getElementById("savedPatternsTab")?.addEventListener("click", () => switchTab("saved"));
  document
    .getElementById("readSelectionButton")
    ?.addEventListener("click", readFormattingFromSelection);
  document.getElementById("clearFormButton")?.addEventListener("click", clearForm);
  document.getElementById("savePatternButton")?.addEventListener("click", savePattern);
  document.getElementById("applyPatternButton")?.addEventListener("click", applySelectedPattern);
  document
    .getElementById("deletePatternButton")
    ?.addEventListener("click", requestDeleteSelectedPattern);
  document
    .getElementById("deletePatternButton")
    ?.addEventListener("pointerup", requestDeleteSelectedPattern);
  document.getElementById("exportPatternsButton")?.addEventListener("click", exportPatterns);
  document.getElementById("importPatternsButton")?.addEventListener("click", openImportDialog);
  document.getElementById("resetDefaultsButton")?.addEventListener("click", resetDefaultPatterns);
  document.getElementById("importPatternsFile")?.addEventListener("change", importPatternsFromFile);
  document
    .getElementById("duplicatePatternButton")
    ?.addEventListener("click", duplicateSelectedPattern);
  document.getElementById("patternSearch")?.addEventListener("input", renderSavedPatterns);
  document.getElementById("patternSort")?.addEventListener("change", handlePatternSortChange);
  document.getElementById("applyScope")?.addEventListener("change", persistApplyScope);
  document.getElementById("savedPatterns")?.addEventListener("change", updatePatternPreview);
  document
    .getElementById("editPatternButton")
    ?.addEventListener("click", loadSelectedPatternForEditing);
});

function setupTabs(): void {
  const savedTab = localStorage.getItem(LAST_TAB_KEY);
  const savedScope = localStorage.getItem(LAST_SCOPE_KEY);
  const savedSort = localStorage.getItem(SORT_KEY);

  setInputValue(
    "applyScope",
    savedScope === "paragraph" || savedScope === "document" ? savedScope : "selection"
  );
  setInputValue("patternSort", savedSort === "za" ? "za" : "az");
  switchTab(savedTab === "saved" ? "saved" : "new");
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
  localStorage.setItem(LAST_TAB_KEY, tab);
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
  setCheckboxValue("underline", pattern.underline);
  setInputValue("textColor", pattern.textColor);
  setInputValue("highlightColor", pattern.highlightColor);
  setInputValue("spaceBefore", String(pattern.spaceBefore));
  setInputValue("spaceAfter", String(pattern.spaceAfter));
  setInputValue("leftIndent", String(pattern.leftIndent));
  setInputValue("rightIndent", String(pattern.rightIndent));
  setInputValue("firstLineIndent", String(pattern.firstLineIndent));

  switchTab("new");
  showMessage("The pattern was loaded for editing.");
}

function loadPatternIntoForm(pattern: TextPattern, name: string): void {
  setInputValue("patternName", name);
  setInputValue("fontName", pattern.fontName);
  setInputValue("fontSize", String(pattern.fontSize));
  setInputValue("alignment", pattern.alignment);
  setInputValue("lineSpacing", String(convertWordPointsToLineSpacing(pattern.lineSpacing)));
  setCheckboxValue("bold", pattern.bold);
  setCheckboxValue("italic", pattern.italic);
  setCheckboxValue("underline", pattern.underline);
  setInputValue("textColor", pattern.textColor);
  setInputValue("highlightColor", pattern.highlightColor);
  setInputValue("spaceBefore", String(pattern.spaceBefore));
  setInputValue("spaceAfter", String(pattern.spaceAfter));
  setInputValue("leftIndent", String(pattern.leftIndent));
  setInputValue("rightIndent", String(pattern.rightIndent));
  setInputValue("firstLineIndent", String(pattern.firstLineIndent));
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

  showToast(message);
}

function showToast(message: string): void {
  const toastContainer = document.getElementById("toastContainer");

  if (!toastContainer) {
    return;
  }

  const toast = document.createElement("div");
  toast.className = "toast";
  toast.textContent = message;
  toastContainer.appendChild(toast);

  setTimeout(() => {
    toast.classList.add("leaving");

    setTimeout(() => {
      toast.remove();
    }, 180);
  }, 2600);
}

function getPatterns(): TextPattern[] {
  const data = localStorage.getItem(STORAGE_KEY);

  if (!data) {
    return [];
  }

  try {
    const patterns = JSON.parse(data) as TextPattern[];
    return patterns.map(normalizePattern);
  } catch {
    return [];
  }
}

function savePatterns(patterns: TextPattern[]): void {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(patterns.map(normalizePattern)));
}

function initializeDefaultPatterns(): void {
  if (localStorage.getItem(STORAGE_KEY)) {
    return;
  }

  savePatterns(getDefaultPatterns());
}

function getDefaultPatterns(): TextPattern[] {
  return [
    createPattern({
      name: "Academic Body",
      fontName: "Times New Roman",
      fontSize: 12,
      alignment: "Justified",
      lineSpacing: 1.5,
      spaceAfter: 6,
    }),
    createPattern({
      name: "Heading",
      fontName: "Calibri",
      fontSize: 18,
      alignment: "Left",
      lineSpacing: 1.15,
      bold: true,
      textColor: "#176b5d",
      spaceBefore: 12,
      spaceAfter: 6,
    }),
    createPattern({
      name: "Quote",
      fontName: "Georgia",
      fontSize: 12,
      alignment: "Left",
      lineSpacing: 1.2,
      italic: true,
      leftIndent: 24,
      rightIndent: 12,
    }),
    createPattern({
      name: "Compact Paragraph",
      fontName: "Calibri",
      fontSize: 11,
      alignment: "Left",
      lineSpacing: 1,
    }),
  ];
}

function resetDefaultPatterns(): void {
  savePatterns(mergePatterns(getPatterns(), getDefaultPatterns()));
  setInputValue("patternSearch", "");
  setInputValue("patternSort", "az");
  localStorage.setItem(SORT_KEY, "az");
  renderSavedPatterns();
  setInputValue("savedPatterns", "Academic Body");
  updatePatternPreview();
  switchTab("saved");
  showMessage("Default presets were restored.");
}

function createPattern(pattern: Partial<TextPattern> & { name: string }): TextPattern {
  return normalizePattern({
    name: pattern.name,
    fontName: pattern.fontName ?? "Times New Roman",
    fontSize: pattern.fontSize ?? 12,
    alignment: pattern.alignment ?? "Left",
    lineSpacing: pattern.lineSpacing ?? 1.5,
    bold: pattern.bold ?? false,
    italic: pattern.italic ?? false,
    underline: pattern.underline ?? false,
    textColor: pattern.textColor ?? "#000000",
    highlightColor: pattern.highlightColor ?? "",
    spaceBefore: pattern.spaceBefore ?? 0,
    spaceAfter: pattern.spaceAfter ?? 0,
    leftIndent: pattern.leftIndent ?? 0,
    rightIndent: pattern.rightIndent ?? 0,
    firstLineIndent: pattern.firstLineIndent ?? 0,
  });
}

function normalizePattern(pattern: TextPattern): TextPattern {
  return {
    name: pattern.name,
    fontName: pattern.fontName || "Times New Roman",
    fontSize: Number(pattern.fontSize) || 12,
    alignment: pattern.alignment || "Left",
    lineSpacing: Number(pattern.lineSpacing) || 1.5,
    bold: pattern.bold === true,
    italic: pattern.italic === true,
    underline: pattern.underline === true,
    textColor: pattern.textColor || "#000000",
    highlightColor: pattern.highlightColor || "",
    spaceBefore: Number(pattern.spaceBefore) || 0,
    spaceAfter: Number(pattern.spaceAfter) || 0,
    leftIndent: Number(pattern.leftIndent) || 0,
    rightIndent: Number(pattern.rightIndent) || 0,
    firstLineIndent: Number(pattern.firstLineIndent) || 0,
  };
}

function clearForm(): void {
  editingPatternName = null;
  setInputValue("patternName", "");
  setInputValue("fontName", "Times New Roman");
  setInputValue("fontSize", "12");
  setInputValue("alignment", "Left");
  setInputValue("lineSpacing", "1.5");
  setCheckboxValue("bold", false);
  setCheckboxValue("italic", false);
  setCheckboxValue("underline", false);
  setInputValue("textColor", "#000000");
  setInputValue("highlightColor", "");
  setInputValue("spaceBefore", "0");
  setInputValue("spaceAfter", "0");
  setInputValue("leftIndent", "0");
  setInputValue("rightIndent", "0");
  setInputValue("firstLineIndent", "0");
  showMessage("The form was reset.");
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
    underline: getCheckboxValue("underline"),
    textColor: getInputValue("textColor"),
    highlightColor: getInputValue("highlightColor"),
    spaceBefore: Number(getInputValue("spaceBefore")),
    spaceAfter: Number(getInputValue("spaceAfter")),
    leftIndent: Number(getInputValue("leftIndent")),
    rightIndent: Number(getInputValue("rightIndent")),
    firstLineIndent: Number(getInputValue("firstLineIndent")),
  };
}

function savePattern(): void {
  const pattern = buildPatternFromForm();

  if (!pattern) {
    return;
  }

  const patterns = getPatterns();
  let message = "The pattern was added.";

  if (editingPatternName) {
    const index = patterns.findIndex((p) => p.name === editingPatternName);

    if (index >= 0) {
      patterns[index] = pattern;
    } else {
      patterns.push(pattern);
    }

    editingPatternName = null;
    message = "The pattern was updated.";
  } else {
    const existingIndex = patterns.findIndex((p) => p.name === pattern.name);

    if (existingIndex >= 0) {
      patterns[existingIndex] = pattern;
      message = "The existing pattern was updated.";
    } else {
      patterns.push(pattern);
    }
  }

  savePatterns(patterns);
  renderSavedPatterns();
  setInputValue("savedPatterns", pattern.name);
  updatePatternPreview();
  switchTab("saved");

  showMessage(message);
}

function renderSavedPatterns(): void {
  const select = document.getElementById("savedPatterns") as HTMLSelectElement;

  if (!select) {
    return;
  }

  select.innerHTML = "";

  const patterns = getFilteredPatterns();

  if (patterns.length === 0) {
    const option = document.createElement("option");
    option.textContent = getPatterns().length === 0 ? "No saved patterns" : "No matching patterns";
    option.value = "";
    select.appendChild(option);
    updatePatternPreview();
    return;
  }

  patterns.forEach((pattern) => {
    const option = document.createElement("option");
    option.value = pattern.name;
    option.textContent = pattern.name;
    select.appendChild(option);
  });

  updatePatternPreview();
}

function getFilteredPatterns(): TextPattern[] {
  const query = getInputValue("patternSearch").trim().toLowerCase();
  const patterns = getSortedPatterns(getPatterns());

  if (!query) {
    return patterns;
  }

  return patterns.filter((pattern) =>
    [
      pattern.name,
      pattern.fontName,
      pattern.alignment,
      pattern.bold ? "bold" : "regular",
      pattern.italic ? "italic" : "normal",
      pattern.underline ? "underline" : "",
    ]
      .join(" ")
      .toLowerCase()
      .includes(query)
  );
}

function getSortedPatterns(patterns: TextPattern[]): TextPattern[] {
  const sort = getInputValue("patternSort");

  return [...patterns].sort((first, second) => {
    const comparison = first.name.localeCompare(second.name);
    return sort === "za" ? -comparison : comparison;
  });
}

function handlePatternSortChange(): void {
  localStorage.setItem(SORT_KEY, getInputValue("patternSort"));
  renderSavedPatterns();
}

function persistApplyScope(): void {
  localStorage.setItem(LAST_SCOPE_KEY, getInputValue("applyScope"));
}

function getSelectedPattern(): TextPattern | null {
  const selectedName = getSelectedPatternName();

  if (!selectedName) {
    return null;
  }

  const patterns = getPatterns();
  return patterns.find((pattern) => pattern.name === selectedName) ?? null;
}

function getSelectedPatternName(): string {
  const select = document.getElementById("savedPatterns") as HTMLSelectElement;

  if (!select || select.selectedIndex < 0) {
    return "";
  }

  const selectedOption = select.options[select.selectedIndex];
  return selectedOption?.value || select.value || "";
}

function updatePatternPreview(): void {
  renderPatternPreview(getSelectedPattern());
}

function renderPatternPreview(pattern: TextPattern | null): void {
  const preview = document.getElementById("patternPreview");

  if (!preview) {
    return;
  }

  preview.innerHTML = "";

  if (!pattern) {
    const empty = document.createElement("p");
    empty.className = "preview-empty";
    empty.textContent = "No pattern selected.";
    preview.appendChild(empty);
    return;
  }

  const sample = document.createElement("p") as HTMLElement;
  sample.className = "preview-sample";
  sample.textContent = "The quick brown fox jumps over the lazy dog.";
  sample.style.fontFamily = pattern.fontName;
  sample.style.fontSize = `${pattern.fontSize}px`;
  sample.style.fontWeight = pattern.bold ? "700" : "400";
  sample.style.fontStyle = pattern.italic ? "italic" : "normal";
  sample.style.textDecoration = pattern.underline ? "underline" : "none";
  sample.style.color = pattern.textColor;
  sample.style.backgroundColor = pattern.highlightColor || "transparent";
  sample.style.padding = pattern.highlightColor ? "2px 4px" : "0";
  sample.style.textAlign = getCssAlignment(pattern.alignment);
  sample.style.lineHeight = String(convertWordPointsToLineSpacing(pattern.lineSpacing));

  const meta = document.createElement("div");
  meta.className = "preview-meta";

  [
    pattern.fontName,
    `${pattern.fontSize}px`,
    pattern.alignment,
    `${convertWordPointsToLineSpacing(pattern.lineSpacing)} line`,
    pattern.bold ? "Bold" : "Regular",
    pattern.italic ? "Italic" : "Normal",
    pattern.underline ? "Underline" : "No underline",
    `Before ${pattern.spaceBefore}pt`,
    `After ${pattern.spaceAfter}pt`,
    `Left ${pattern.leftIndent}pt`,
    `Right ${pattern.rightIndent}pt`,
    `First ${pattern.firstLineIndent}pt`,
  ].forEach((label) => {
    const chip = document.createElement("span");
    chip.className = "preview-chip";
    chip.textContent = label;
    meta.appendChild(chip);
  });

  preview.appendChild(sample);
  preview.appendChild(meta);
}

function getCssAlignment(alignment: AlignmentValue): string {
  if (alignment === "Centered") {
    return "center";
  }

  if (alignment === "Right") {
    return "right";
  }

  if (alignment === "Justified") {
    return "justify";
  }

  return "left";
}

async function readFormattingFromSelection(): Promise<void> {
  try {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();

      // eslint-disable-next-line office-addins/no-navigational-load
      selection.load(
        "font/name,font/size,font/bold,font/italic,font/color,font/underline,font/highlightColor,paragraphs"
      );

      await context.sync();

      const paragraphs = selection.paragraphs;
      paragraphs.load(
        "items/alignment,items/lineSpacing,items/spaceBefore,items/spaceAfter,items/leftIndent,items/rightIndent,items/firstLineIndent"
      );

      await context.sync();

      let alignment: AlignmentValue = "Left";
      let lineSpacing = 12;

      if (paragraphs.items.length > 0) {
        const firstParagraph = paragraphs.items[0];
        alignment = firstParagraph.alignment as AlignmentValue;
        lineSpacing = convertWordPointsToLineSpacing(firstParagraph.lineSpacing || 12);
        setInputValue("spaceBefore", String(firstParagraph.spaceBefore || 0));
        setInputValue("spaceAfter", String(firstParagraph.spaceAfter || 0));
        setInputValue("leftIndent", String(firstParagraph.leftIndent || 0));
        setInputValue("rightIndent", String(firstParagraph.rightIndent || 0));
        setInputValue("firstLineIndent", String(firstParagraph.firstLineIndent || 0));
      }

      const fontName = selection.font.name || "Times New Roman";
      const fontSize = selection.font.size || 12;
      const bold = selection.font.bold === true;
      const italic = selection.font.italic === true;
      const underline = String(selection.font.underline || "None") !== "None";
      const textColor = normalizeColor(selection.font.color, "#000000");
      const highlightColor = normalizeHighlightColor(selection.font.highlightColor);

      if (!getInputValue("patternName").trim()) {
        setInputValue("patternName", "Pattern from selection");
      }

      setInputValue("fontName", fontName);
      setInputValue("fontSize", String(fontSize));
      setInputValue("alignment", alignment);
      setInputValue("lineSpacing", String(lineSpacing));
      setCheckboxValue("bold", bold);
      setCheckboxValue("italic", italic);
      setCheckboxValue("underline", underline);
      setInputValue("textColor", textColor);
      setInputValue("highlightColor", highlightColor);

      showMessage("Formatting was read from the selected text.");
    });
  } catch (error) {
    showMessage("Could not read formatting from the selection.");
    console.error(error);
  }
}

async function applySelectedPattern(): Promise<void> {
  const pattern = getSelectedPattern();
  const scope = getInputValue("applyScope");

  if (!pattern) {
    showMessage("Choose a saved pattern.");
    return;
  }

  try {
    await Word.run(async (context) => {
      if (scope === "document") {
        await applyPatternToDocument(context, pattern);
        return;
      }

      if (scope === "paragraph") {
        await applyPatternToCurrentParagraph(context, pattern);
        return;
      }

      const selection = context.document.getSelection();

      applyFontSettings(selection, pattern);
      // eslint-disable-next-line office-addins/no-navigational-load
      selection.load("paragraphs");

      await context.sync();

      const paragraphs = selection.paragraphs;
      paragraphs.load("items");

      await context.sync();

      paragraphs.items.forEach((paragraph) => {
        applyParagraphSettings(paragraph, pattern);
      });

      await context.sync();
    });

    showMessage(
      scope === "document"
        ? "The pattern was applied to the whole document."
        : scope === "paragraph"
          ? "The pattern was applied to the current paragraph."
          : "The pattern was applied to the selected text."
    );
  } catch (error) {
    showMessage("An error occurred while applying the pattern.");
    console.error(error);
  }
}

async function applyPatternToCurrentParagraph(
  context: Word.RequestContext,
  pattern: TextPattern
): Promise<void> {
  const selection = context.document.getSelection();

  // eslint-disable-next-line office-addins/no-navigational-load
  selection.load("paragraphs");

  await context.sync();

  const paragraphs = selection.paragraphs;
  paragraphs.load("items");

  await context.sync();

  if (paragraphs.items.length === 0) {
    return;
  }

  const paragraph = paragraphs.items[0];
  const range = paragraph.getRange();

  applyFontSettings(range, pattern);
  applyParagraphSettings(paragraph, pattern);

  await context.sync();
}

async function applyPatternToDocument(
  context: Word.RequestContext,
  pattern: TextPattern
): Promise<void> {
  const body = context.document.body;

  applyFontSettings(body, pattern);
  body.load("paragraphs");

  await context.sync();

  const paragraphs = body.paragraphs;
  paragraphs.load("items");

  await context.sync();

  paragraphs.items.forEach((paragraph) => {
    applyParagraphSettings(paragraph, pattern);
  });

  await context.sync();
}

function applyFontSettings(target: Word.Range | Word.Body, pattern: TextPattern): void {
  target.font.name = pattern.fontName;
  target.font.size = pattern.fontSize;
  target.font.bold = pattern.bold;
  target.font.italic = pattern.italic;
  target.font.underline = pattern.underline ? "Single" : "None";
  target.font.color = pattern.textColor;

  if (pattern.highlightColor) {
    target.font.highlightColor = pattern.highlightColor;
  }
}

function applyParagraphSettings(paragraph: Word.Paragraph, pattern: TextPattern): void {
  paragraph.alignment = pattern.alignment as Word.Alignment;
  paragraph.lineSpacing = convertLineSpacingToWordPoints(pattern.lineSpacing);
  paragraph.spaceBefore = pattern.spaceBefore;
  paragraph.spaceAfter = pattern.spaceAfter;
  paragraph.leftIndent = pattern.leftIndent;
  paragraph.rightIndent = pattern.rightIndent;
  paragraph.firstLineIndent = pattern.firstLineIndent;
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

function normalizeColor(value: string | undefined, fallback: string): string {
  if (!value || value === "Mixed") {
    return fallback;
  }

  return value;
}

function normalizeHighlightColor(value: string | undefined): string {
  if (!value || value === "Mixed" || value.toLowerCase() === "#ffffff") {
    return "";
  }

  return value;
}

function requestDeleteSelectedPattern(event?: Event): void {
  event?.preventDefault();

  const now = Date.now();

  if (now - lastDeleteActionAt < 400) {
    return;
  }

  lastDeleteActionAt = now;
  deleteSelectedPattern();
}

function deleteSelectedPattern(): void {
  const pattern = getSelectedPattern();

  if (!pattern) {
    showMessage("Choose a pattern to delete.");
    return;
  }

  const patterns = getPatterns();
  const updatedPatterns = patterns.filter((savedPattern) => savedPattern.name !== pattern.name);

  if (updatedPatterns.length === patterns.length) {
    showMessage(`Could not find "${pattern.name}" in saved patterns.`);
    return;
  }

  savePatterns(updatedPatterns);
  renderSavedPatterns();
  editingPatternName = editingPatternName === pattern.name ? null : editingPatternName;

  showMessage(`"${pattern.name}" was deleted.`);
}

function exportPatterns(): void {
  const patterns = getPatterns();

  if (patterns.length === 0) {
    showMessage("There are no patterns to export.");
    return;
  }

  const exportData = {
    version: 1,
    exportedAt: new Date().toISOString(),
    patterns: patterns,
  };
  const blob = new Blob([JSON.stringify(exportData, null, 2)], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");

  link.href = url;
  link.download = "word-patterns.json";
  link.click();
  URL.revokeObjectURL(url);
  showMessage("Patterns were exported.");
}

function openImportDialog(): void {
  const input = document.getElementById("importPatternsFile") as HTMLInputElement;

  if (input) {
    input.value = "";
    input.click();
  }
}

function importPatternsFromFile(): void {
  const input = document.getElementById("importPatternsFile") as HTMLInputElement;
  const file = input.files?.[0];

  if (!file) {
    return;
  }

  const reader = new FileReader();

  reader.onload = () => {
    try {
      const importedPatterns = parseImportedPatterns(String(reader.result ?? ""));

      if (importedPatterns.length === 0) {
        showMessage("No valid patterns were found in the file.");
        return;
      }

      const mergedPatterns = mergePatterns(getPatterns(), importedPatterns);
      savePatterns(mergedPatterns);
      renderSavedPatterns();
      switchTab("saved");
      showMessage(`${importedPatterns.length} pattern(s) were imported.`);
    } catch (error) {
      showMessage("Could not import patterns from that JSON file.");
      console.error(error);
    }
  };

  reader.readAsText(file);
}

function parseImportedPatterns(contents: string): TextPattern[] {
  const data = JSON.parse(contents);
  const source = Array.isArray(data) ? data : data.patterns;

  if (!Array.isArray(source)) {
    return [];
  }

  return source.filter(isValidPattern).map((pattern) => ({
    name: pattern.name,
    fontName: pattern.fontName,
    fontSize: Number(pattern.fontSize),
    alignment: pattern.alignment,
    lineSpacing: Number(pattern.lineSpacing),
    bold: pattern.bold === true,
    italic: pattern.italic === true,
    underline: pattern.underline === true,
    textColor: pattern.textColor || "#000000",
    highlightColor: pattern.highlightColor || "",
    spaceBefore: Number(pattern.spaceBefore) || 0,
    spaceAfter: Number(pattern.spaceAfter) || 0,
    leftIndent: Number(pattern.leftIndent) || 0,
    rightIndent: Number(pattern.rightIndent) || 0,
    firstLineIndent: Number(pattern.firstLineIndent) || 0,
  }));
}

function isValidPattern(pattern: TextPattern): boolean {
  return (
    pattern &&
    typeof pattern.name === "string" &&
    pattern.name.trim().length > 0 &&
    typeof pattern.fontName === "string" &&
    Number.isFinite(Number(pattern.fontSize)) &&
    ["Left", "Centered", "Right", "Justified"].includes(pattern.alignment) &&
    Number.isFinite(Number(pattern.lineSpacing))
  );
}

function mergePatterns(
  currentPatterns: TextPattern[],
  importedPatterns: TextPattern[]
): TextPattern[] {
  const mergedPatterns = [...currentPatterns];

  importedPatterns.forEach((importedPattern) => {
    const index = mergedPatterns.findIndex((pattern) => pattern.name === importedPattern.name);

    if (index >= 0) {
      mergedPatterns[index] = importedPattern;
    } else {
      mergedPatterns.push(importedPattern);
    }
  });

  return mergedPatterns;
}

function duplicateSelectedPattern(): void {
  const pattern = getSelectedPattern();

  if (!pattern) {
    showMessage("Choose a pattern to duplicate.");
    return;
  }

  editingPatternName = null;
  loadPatternIntoForm(pattern, getDuplicatePatternName(pattern.name));
  switchTab("new");
  showMessage("The pattern was copied into a new draft.");
}

function getDuplicatePatternName(name: string): string {
  const patterns = getPatterns();
  let copyNumber = 1;
  let candidate = `${name} Copy`;

  while (patterns.some((pattern) => pattern.name === candidate)) {
    copyNumber += 1;
    candidate = `${name} Copy ${copyNumber}`;
  }

  return candidate;
}
