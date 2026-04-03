import * as XLSX from "xlsx";

type VariableStatus = "not-inserted" | "up-to-date" | "outdated";

type VariableItem = {
  name: string;
  ref: string;
  displayValue: string;
  matrix: string[][];
  isBlock: boolean;
  usageCount: number;
  status: VariableStatus;
};

type SavedFileMeta = {
  fileName: string;
  lastLoadedAt: string;
};

type ParsedControlMeta = {
  variableName: string;
  mode: "inline" | "block";
  prefix: string;
  suffix: string;
};

let loadedVariables: VariableItem[] = [];
let isHighlightActive = false;

Office.onReady(() => {
  const loadButton = document.getElementById("loadExcelButton") as HTMLButtonElement | null;
  const updateAllButton = document.getElementById("updateAllButton") as HTMLButtonElement | null;
  const toggleHighlightButton = document.getElementById("toggleHighlightButton") as HTMLButtonElement | null;

  if (loadButton) loadButton.onclick = handleExcelLoad;
  if (updateAllButton) updateAllButton.onclick = handleUpdateAll;
  if (toggleHighlightButton) toggleHighlightButton.onclick = toggleHighlight;

  renderSavedFileInfo();
  renderHighlightButtonState();
});

async function handleExcelLoad(): Promise<void> {
  const fileInput = document.getElementById("excelFile") as HTMLInputElement | null;
  const variablesList = document.getElementById("variablesList") as HTMLDivElement | null;

  if (!fileInput || !variablesList) return;

  const file = fileInput.files?.[0];
  if (!file) {
    setStatus("Bitte zuerst eine Excel-Datei auswählen.");
    return;
  }

  try {
    const arrayBuffer = await file.arrayBuffer();

    const workbook = XLSX.read(arrayBuffer, {
      type: "array",
      cellNF: true,
      cellText: true,
      cellDates: true,
    });

    const result: VariableItem[] = [];
    const names = workbook.Workbook?.Names ?? [];

    for (const namedItem of names as any[]) {
      const variableName = String(namedItem.Name ?? "").trim();
      const ref = String(namedItem.Ref ?? "").trim();

      if (!variableName || !ref) continue;

      const resolved = resolveNamedReferenceData(workbook, ref);

      result.push({
        name: variableName,
        ref,
        displayValue: resolved.displayValue,
        matrix: resolved.matrix,
        isBlock: resolved.isBlock,
        usageCount: 0,
        status: "not-inserted",
      });
    }

    loadedVariables = result;

    await saveLinkedFileMeta({
      fileName: file.name,
      lastLoadedAt: new Date().toLocaleString("de-DE"),
    });

    renderSavedFileInfo();

    if (result.length === 0) {
      variablesList.innerHTML = "<p>Keine benannten Variablen oder Bereiche gefunden.</p>";
      setStatus("Es wurden keine benannten Variablen gefunden.");
      return;
    }

    await refreshUsageStatus();
    setStatus(`${result.length} Variable(n) geladen.`);
  } catch (error) {
    console.error(error);
    setStatus("Fehler beim Lesen der Excel-Datei.");
  }
}

async function handleUpdateAll(): Promise<void> {
  if (loadedVariables.length === 0) {
    setStatus("Bitte zuerst eine Excel-Datei laden.");
    return;
  }

  try {
    let updatedCount = 0;

    await Word.run(async (context) => {
      const controls = context.document.contentControls;
      controls.load("items/tag,title,text");
      await context.sync();

      for (const control of controls.items) {
        const meta = parseControlMeta(control.tag, control.title);
        if (!meta) continue;

        const variable = loadedVariables.find((v) => v.name === meta.variableName);
        if (!variable) continue;

        const newText = buildRenderedText(variable, meta.mode, meta.prefix, meta.suffix);
        control.insertText(newText, Word.InsertLocation.replace);
        updatedCount++;
      }

      await context.sync();
    });

    await refreshUsageStatus();

    if (isHighlightActive) {
      await applyHighlightToAllVariables(true);
    }

    setStatus(`${updatedCount} Feld(er) im Dokument aktualisiert.`);
  } catch (error) {
    console.error(error);
    setStatus("Fehler beim Aktualisieren der Variablen im Word-Dokument.");
  }
}

async function updateSingleVariable(variableName: string): Promise<void> {
  const variable = loadedVariables.find((v) => v.name === variableName);
  if (!variable) return;

  try {
    let updatedCount = 0;

    await Word.run(async (context) => {
      const controls = context.document.contentControls;
      controls.load("items/tag,title,text");
      await context.sync();

      for (const control of controls.items) {
        const meta = parseControlMeta(control.tag, control.title);
        if (!meta) continue;
        if (meta.variableName !== variableName) continue;

        const newText = buildRenderedText(variable, meta.mode, meta.prefix, meta.suffix);
        control.insertText(newText, Word.InsertLocation.replace);
        updatedCount++;
      }

      await context.sync();
    });

    await refreshUsageStatus();

    if (isHighlightActive) {
      await applyHighlightToAllVariables(true);
    }

    setStatus(`${updatedCount} Vorkommen von "${variableName}" aktualisiert.`);
  } catch (error) {
    console.error(error);
    setStatus(`Fehler beim Aktualisieren von "${variableName}".`);
  }
}

async function refreshUsageStatus(): Promise<void> {
  if (loadedVariables.length === 0) {
    renderVariables(loadedVariables);
    return;
  }

  try {
    for (const variable of loadedVariables) {
      variable.usageCount = 0;
      variable.status = "not-inserted";
    }

    const variableMap = new Map<string, VariableItem>();
    loadedVariables.forEach((v) => variableMap.set(v.name, v));

    await Word.run(async (context) => {
      const controls = context.document.contentControls;
      controls.load("items/tag,title,text");
      await context.sync();

      for (const control of controls.items) {
        const meta = parseControlMeta(control.tag, control.title);
        if (!meta) continue;

        const variable = variableMap.get(meta.variableName);
        if (!variable) continue;

        variable.usageCount += 1;

        const expectedText = buildRenderedText(variable, meta.mode, meta.prefix, meta.suffix);
        const currentText = control.text ?? "";

        const matches =
          meta.mode === "block"
            ? textsEqualForBlockStatus(currentText, expectedText)
            : textsEqualForInlineStatus(currentText, expectedText);

        if (variable.status !== "outdated") {
          variable.status = matches ? "up-to-date" : "outdated";
        }
      }

      for (const variable of loadedVariables) {
        if (variable.usageCount === 0) {
          variable.status = "not-inserted";
        }
      }

      await context.sync();
    });

    renderVariables(loadedVariables);
  } catch (error) {
    console.error(error);
    setStatus("Fehler beim Prüfen des Dokumentstatus.");
  }
}

async function toggleHighlight(): Promise<void> {
  try {
    isHighlightActive = !isHighlightActive;
    await applyHighlightToAllVariables(isHighlightActive);
    renderHighlightButtonState();

    setStatus(
      isHighlightActive
        ? "Alle Variablen im Dokument sind hervorgehoben."
        : "Die Hervorhebung der Variablen wurde entfernt."
    );
  } catch (error) {
    console.error(error);
    isHighlightActive = !isHighlightActive;
    renderHighlightButtonState();
    setStatus("Fehler beim Umschalten der Hervorhebung.");
  }
}

async function applyHighlightToAllVariables(enable: boolean): Promise<void> {
  await Word.run(async (context) => {
    const controls = context.document.contentControls;
    controls.load("items/tag,title");
    await context.sync();

    for (const control of controls.items) {
      const meta = parseControlMeta(control.tag, control.title);
      if (!meta) continue;

      const range = control.getRange();
      range.font.highlightColor = enable ? "#FFF59D" : null;
    }

    await context.sync();
  });
}

function renderHighlightButtonState(): void {
  const button = document.getElementById("toggleHighlightButton");
  const stateInfo = document.getElementById("highlightStateInfo");

  if (button) {
    button.textContent = isHighlightActive
      ? "Hervorhebung ausblenden"
      : "Variablen hervorheben";
  }

  if (stateInfo) {
    stateInfo.textContent = isHighlightActive
      ? "Hervorhebung: aktiv"
      : "Hervorhebung: aus";
  }
}

function renderVariables(variables: VariableItem[]): void {
  const variablesList = document.getElementById("variablesList") as HTMLDivElement | null;
  if (!variablesList) return;

  if (variables.length === 0) {
    variablesList.innerHTML = "Noch keine Variablen geladen.";
    return;
  }

  variablesList.innerHTML = variables
    .map((item, index) => {
      const statusLabel =
        item.status === "up-to-date"
          ? "Aktuell"
          : item.status === "outdated"
            ? "Veraltet"
            : "Neu";

      const statusColor =
        item.status === "up-to-date"
          ? "#107c10"
          : item.status === "outdated"
            ? "#d83b01"
            : "#666";

      const preview = item.isBlock
        ? escapeHtml(item.displayValue).replace(/\n/g, "<br/>")
        : escapeHtml(item.displayValue);

      return `
        <div style="margin-bottom:12px; padding:10px; border:1px solid #ccc; border-radius:4px;">
          <div style="font-size:13px; line-height:1.4;">
            <strong>${escapeHtml(item.name)}</strong>
            &nbsp;·&nbsp;Verwendet: ${item.usageCount}
            &nbsp;·&nbsp;${escapeHtml(item.ref)}
            &nbsp;·&nbsp;<span style="color:${statusColor}; font-weight:600;">${statusLabel}</span>
          </div>

          <div style="margin-top:8px; padding:8px; background:#f7f7f7; border-radius:4px; white-space:pre-wrap;">${preview}</div>

          <div style="margin-top:10px;">
            <label style="font-size:12px; color:#444;">Einfügemodus:</label>
            <select id="mode-${index}" style="width:100%; margin-top:4px;">
              <option value="inline">Inline</option>
              <option value="block"${item.isBlock ? " selected" : ""}>Block</option>
            </select>
          </div>

          <div style="display:flex; gap:8px; margin-top:10px; flex-wrap:wrap;">
            <button data-index="${index}" class="insertVariableButton ms-Button ms-Button--primary">
              <span class="ms-Button-label">Einfügen</span>
            </button>

            <button data-index="${index}" class="updateSingleVariableButton ms-Button">
              <span class="ms-Button-label">Nur diese aktualisieren</span>
            </button>
          </div>
        </div>
      `;
    })
    .join("");

  const insertButtons = document.querySelectorAll(".insertVariableButton");
  insertButtons.forEach((button) => {
    button.addEventListener("click", async (event) => {
      const target = event.currentTarget as HTMLButtonElement;
      const index = Number(target.getAttribute("data-index"));
      const variable = loadedVariables[index];
      if (!variable) return;

      const modeSelect = document.getElementById(`mode-${index}`) as HTMLSelectElement | null;
      const mode = (modeSelect?.value as "inline" | "block") || "inline";

      await insertVariableAsContentControl(variable, mode);
    });
  });

  const updateButtons = document.querySelectorAll(".updateSingleVariableButton");
  updateButtons.forEach((button) => {
    button.addEventListener("click", async (event) => {
      const target = event.currentTarget as HTMLButtonElement;
      const index = Number(target.getAttribute("data-index"));
      const variable = loadedVariables[index];
      if (!variable) return;

      await updateSingleVariable(variable.name);
    });
  });
}

async function insertVariableAsContentControl(variable: VariableItem, mode: "inline" | "block"): Promise<void> {
  try {
    await Word.run(async (context) => {
      const range = context.document.getSelection();
      range.load("text");
      await context.sync();

      const selectedText = range.text ?? "";
      const prefix = selectedText.match(/^\s*/)?.[0] ?? "";
      const suffix = selectedText.match(/\s*$/)?.[0] ?? "";

      const contentControl = range.insertContentControl();

      contentControl.tag = buildRichTag(variable.name, mode, prefix, suffix);
      contentControl.title = `Excel-Variable: ${variable.name}`;
      contentControl.appearance = "BoundingBox";
      contentControl.cannotDelete = false;
      contentControl.cannotEdit = false;

      const renderedText = buildRenderedText(variable, mode, prefix, suffix);
      contentControl.insertText(renderedText, Word.InsertLocation.replace);

      await context.sync();
    });

    await refreshUsageStatus();

    if (isHighlightActive) {
      await applyHighlightToAllVariables(true);
    }

    setStatus(`Variable "${variable.name}" wurde eingefügt.`);
  } catch (error) {
    console.error(error);
    setStatus(`Fehler beim Einfügen der Variable "${variable.name}".`);
  }
}

function parseControlMeta(tag: string | undefined, title: string | undefined): ParsedControlMeta | null {
  if (!tag) return null;

  if (tag.startsWith("excelvar|")) {
    const parts = tag.split("|");
    if (parts.length >= 5) {
      return {
        variableName: decodeSafe(parts[1]),
        mode: decodeSafe(parts[2]) === "block" ? "block" : "inline",
        prefix: decodeSafe(parts[3]),
        suffix: decodeSafe(parts[4]),
      };
    }
  }

  if (tag.startsWith("excelvar:")) {
    const variableName = tag.substring("excelvar:".length);
    return {
      variableName,
      mode: inferModeFromTitleOrDefault(title),
      prefix: "",
      suffix: "",
    };
  }

  return null;
}

function inferModeFromTitleOrDefault(title: string | undefined): "inline" | "block" {
  if (!title) return "inline";
  const lower = title.toLowerCase();
  if (lower.includes("block")) return "block";
  return "inline";
}

function buildRichTag(
  variableName: string,
  mode: "inline" | "block",
  prefix: string,
  suffix: string
): string {
  return [
    "excelvar",
    encodeURIComponent(variableName),
    encodeURIComponent(mode),
    encodeURIComponent(prefix),
    encodeURIComponent(suffix),
  ].join("|");
}

function decodeSafe(value: string): string {
  try {
    return decodeURIComponent(value);
  } catch {
    return value;
  }
}

function textsEqualForInlineStatus(currentText: string, expectedText: string): boolean {
  return normalizeInlineText(currentText) === normalizeInlineText(expectedText);
}

function textsEqualForBlockStatus(currentText: string, expectedText: string): boolean {
  return normalizeBlockText(currentText) === normalizeBlockText(expectedText);
}

function normalizeInlineText(value: string): string {
  return value
    .replace(/\u00A0/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function normalizeBlockText(value: string): string {
  return value
    .replace(/\r\n/g, "\n")
    .replace(/\r/g, "\n")
    .replace(/\u00A0/g, " ")
    .replace(/\t/g, " ")
    .split("\n")
    .map((line) => line.replace(/\s+/g, " ").trim())
    .filter((line) => line.length > 0)
    .join("\n");
}

function resolveNamedReferenceData(
  workbook: XLSX.WorkBook,
  ref: string
): { displayValue: string; matrix: string[][]; isBlock: boolean } {
  const cleanedRef = ref.replace(/^=/, "");
  const match = cleanedRef.match(/^(?:'([^']+)'|([^!]+))!(.+)$/);

  if (!match) {
    return {
      displayValue: ref,
      matrix: [[ref]],
      isBlock: false,
    };
  }

  const sheetName = (match[1] || match[2] || "").trim();
  const rangeAddress = match[3];
  const worksheet = workbook.Sheets[sheetName];

  if (!worksheet) {
    return {
      displayValue: ref,
      matrix: [[ref]],
      isBlock: false,
    };
  }

  try {
    const range = XLSX.utils.decode_range(rangeAddress);
    const matrix: string[][] = [];

    for (let r = range.s.r; r <= range.e.r; r++) {
      const row: string[] = [];
      for (let c = range.s.c; c <= range.e.c; c++) {
        const cellAddress = XLSX.utils.encode_cell({ r, c });
        const cell = worksheet[cellAddress];
        row.push(getCellDisplayValue(cell));
      }
      matrix.push(row);
    }

    const isBlock = matrix.length > 1 || (matrix[0] && matrix[0].length > 1);
    const displayValue = isBlock
      ? matrix.map((row) => row.join(" | ")).join("\n")
      : (matrix[0]?.[0] ?? "");

    return {
      displayValue,
      matrix,
      isBlock,
    };
  } catch {
    return {
      displayValue: ref,
      matrix: [[ref]],
      isBlock: false,
    };
  }
}

function getCellDisplayValue(cell: XLSX.CellObject | undefined): string {
  if (!cell) return "";

  const anyCell = cell as any;

  if (anyCell.t === "d" && anyCell.v instanceof Date) {
    return formatGermanDate(anyCell.v);
  }

  if (typeof anyCell.v === "number" && typeof anyCell.z === "string") {
    const formatString = anyCell.z.toLowerCase();

    if (formatString.includes("%")) {
      return new Intl.NumberFormat("de-DE", {
        style: "percent",
        minimumFractionDigits: guessFractionDigitsFromFormat(formatString),
        maximumFractionDigits: guessFractionDigitsFromFormat(formatString),
      }).format(anyCell.v);
    }

    if (looksLikeDateFormat(formatString)) {
      const parsedDate = XLSX.SSF.parse_date_code(anyCell.v);
      if (parsedDate) {
        const jsDate = new Date(parsedDate.y, parsedDate.m - 1, parsedDate.d);
        return formatGermanDate(jsDate);
      }
    }
  }

  if (typeof anyCell.w === "string" && anyCell.w.length > 0) {
    const maybeDate = tryNormalizeEnglishDateToGerman(anyCell.w);
    if (maybeDate) return maybeDate;

    const maybePercent = tryNormalizePercentToGerman(anyCell.w);
    if (maybePercent) return maybePercent;

    return String(anyCell.w);
  }

  if (anyCell.v === undefined || anyCell.v === null) return "";
  return String(anyCell.v);
}

function looksLikeDateFormat(formatString: string): boolean {
  return /[dmy]/i.test(formatString);
}

function guessFractionDigitsFromFormat(formatString: string): number {
  const match = formatString.match(/0\.([0#]+)/);
  return match ? match[1].length : 0;
}

function formatGermanDate(date: Date): string {
  return new Intl.DateTimeFormat("de-DE").format(date);
}

function tryNormalizeEnglishDateToGerman(value: string): string | null {
  const trimmed = value.trim();

  const isoMatch = trimmed.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (isoMatch) {
    const year = Number(isoMatch[1]);
    const month = Number(isoMatch[2]) - 1;
    const day = Number(isoMatch[3]);
    return formatGermanDate(new Date(year, month, day));
  }

  const usMatch = trimmed.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (usMatch) {
    const month = Number(usMatch[1]) - 1;
    const day = Number(usMatch[2]);
    const year = Number(usMatch[3]);
    return formatGermanDate(new Date(year, month, day));
  }

  return null;
}

function tryNormalizePercentToGerman(value: string): string | null {
  const trimmed = value.trim();
  const match = trimmed.match(/^(-?\d+(?:[.,]\d+)?)\s?%$/);
  if (!match) return null;

  const num = Number(match[1].replace(",", "."));
  if (Number.isNaN(num)) return null;

  return `${num.toLocaleString("de-DE")} %`;
}

function buildRenderedText(
  variable: VariableItem,
  mode: "inline" | "block",
  prefix = "",
  suffix = ""
): string {
  let core = "";

  if (mode === "block") {
    core = variable.matrix.map((row) => row.join("\t")).join("\n");
  } else if (variable.isBlock) {
    core = variable.matrix.map((row) => row.join(" ")).join(" ");
  } else {
    core = variable.displayValue;
  }

  return `${prefix}${core}${suffix}`;
}

async function saveLinkedFileMeta(meta: SavedFileMeta): Promise<void> {
  Office.context.document.settings.set("excelWordLinkedFileMeta", meta);

  await new Promise<void>((resolve, reject) => {
    Office.context.document.settings.saveAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve();
      } else {
        reject(result.error);
      }
    });
  });
}

function renderSavedFileInfo(): void {
  const linkedFileInfo = document.getElementById("linkedFileInfo");
  if (!linkedFileInfo) return;

  const meta = Office.context.document.settings.get("excelWordLinkedFileMeta") as SavedFileMeta | undefined;

  if (!meta) {
    linkedFileInfo.innerHTML = "Noch keine Excel-Datei diesem Word-Dokument zugeordnet.";
    return;
  }

  linkedFileInfo.innerHTML =
    `<strong>Zugeordnete Excel-Datei:</strong> ${escapeHtml(meta.fileName)}<br/>` +
    `<strong>Letzter Import:</strong> ${escapeHtml(meta.lastLoadedAt)}`;
}

function setStatus(message: string): void {
  const status = document.getElementById("statusMessage");
  if (status) status.textContent = message;
}

function escapeHtml(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}