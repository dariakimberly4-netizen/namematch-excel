(function () {
  "use strict";

  const state = {
    a: null,
    b: null,
    matches: [],
    uniqueMatches: 0,
  };

  const els = {
    fileA: document.getElementById("file-a"),
    fileB: document.getElementById("file-b"),
    cardA: document.getElementById("card-a"),
    cardB: document.getElementById("card-b"),
    statusA: document.getElementById("status-a"),
    statusB: document.getElementById("status-b"),
    columnA: document.getElementById("column-a"),
    columnB: document.getElementById("column-b"),
    columnWrapA: document.getElementById("column-wrap-a"),
    columnWrapB: document.getElementById("column-wrap-b"),
    ignoreOrder: document.getElementById("ignore-order"),
    compare: document.getElementById("compare"),
    results: document.getElementById("results"),
    matchCount: document.getElementById("match-count"),
    rowsA: document.getElementById("rows-a"),
    rowsB: document.getElementById("rows-b"),
    uniqueCount: document.getElementById("unique-count"),
    previewBody: document.getElementById("preview-body"),
    tableWrap: document.getElementById("table-wrap"),
    noMatches: document.getElementById("no-matches"),
    download: document.getElementById("download"),
    reset: document.getElementById("reset"),
    error: document.getElementById("error"),
    errorMessage: document.getElementById("error-message"),
  };

  const nameHeaderHints = [
    "full name", "fullname", "name", "pangalan", "beneficiary name", "senior name",
    "member name", "client name", "resident name", "applicant name", "employee name",
  ];

  function cleanHeader(value, fallbackIndex) {
    const header = String(value == null ? "" : value).trim();
    return header || `Column ${fallbackIndex + 1}`;
  }

  function uniqueHeaders(rawHeaders) {
    const seen = new Map();
    return rawHeaders.map((value, index) => {
      const base = cleanHeader(value, index);
      const count = (seen.get(base.toLowerCase()) || 0) + 1;
      seen.set(base.toLowerCase(), count);
      return count === 1 ? base : `${base} (${count})`;
    });
  }

  function chooseNameColumn(headers) {
    let bestIndex = 0;
    let bestScore = -1;
    headers.forEach((header, index) => {
      const normalized = header.toLowerCase().replace(/[_-]+/g, " ").replace(/\s+/g, " ").trim();
      let score = 0;
      nameHeaderHints.forEach((hint, hintIndex) => {
        if (normalized === hint) score = Math.max(score, 100 - hintIndex);
        else if (normalized.includes(hint)) score = Math.max(score, 50 - hintIndex);
      });
      if (/first|middle|last|surname/.test(normalized) && !/full/.test(normalized)) score = Math.max(score, 5);
      if (score > bestScore) {
        bestScore = score;
        bestIndex = index;
      }
    });
    return bestIndex;
  }

  function normalizeName(value, ignoreOrder) {
    let name = String(value == null ? "" : value)
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .toLowerCase()
      .replace(/[’']/g, "")
      .replace(/[^a-z0-9,\s-]/g, " ")
      .replace(/[-]+/g, " ")
      .replace(/\s+/g, " ")
      .trim();

    if (!name) return "";

    if (ignoreOrder) {
      name = name.replace(/,/g, " ").split(/\s+/).filter(Boolean).sort().join(" ");
    } else {
      name = name.replace(/,/g, " ").replace(/\s+/g, " ").trim();
    }
    return name;
  }

  function displayName(value) {
    return String(value == null ? "" : value).trim();
  }

  function escapeCellText(value) {
    if (value instanceof Date) return value;
    if (typeof value === "number" || typeof value === "boolean") return value;
    const text = String(value == null ? "" : value);
    return /^[=+\-@]/.test(text) ? `'${text}` : text;
  }

  function populateColumnSelect(select, headers, selectedIndex) {
    select.textContent = "";
    headers.forEach((header, index) => {
      const option = document.createElement("option");
      option.value = String(index);
      option.textContent = header;
      option.selected = index === selectedIndex;
      select.appendChild(option);
    });
  }

  function setFileStatus(side, fileData) {
    const card = side === "a" ? els.cardA : els.cardB;
    const status = side === "a" ? els.statusA : els.statusB;
    const wrap = side === "a" ? els.columnWrapA : els.columnWrapB;
    const select = side === "a" ? els.columnA : els.columnB;

    card.classList.add("has-file");
    status.innerHTML = "";
    const name = document.createElement("strong");
    name.textContent = fileData.fileName;
    const detail = document.createElement("span");
    detail.textContent = `${fileData.rows.length.toLocaleString()} data rows • ${fileData.sheetName}`;
    status.append(name, detail);
    populateColumnSelect(select, fileData.headers, fileData.nameColumnIndex);
    wrap.classList.remove("is-hidden");
  }

  function showError(message) {
    els.errorMessage.textContent = message;
    els.error.classList.remove("is-hidden");
    els.error.scrollIntoView({ behavior: "auto", block: "center" });
  }

  function hideError() {
    els.error.classList.add("is-hidden");
  }

  async function readSpreadsheet(file) {
    if (!window.XLSX) {
      throw new Error("The Excel reader did not load. Please check your internet connection and reload the page.");
    }
    if (!file) throw new Error("No file was selected.");
    if (file.size > 50 * 1024 * 1024) throw new Error("The file is larger than 50 MB. Please use a smaller file.");

    const bytes = await file.arrayBuffer();
    const workbook = XLSX.read(bytes, { type: "array", cellDates: true, raw: true });
    const sheetName = workbook.SheetNames[0];
    if (!sheetName) throw new Error("The workbook does not contain a worksheet.");

    const matrix = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {
      header: 1,
      defval: "",
      raw: true,
      blankrows: false,
    });

    if (!matrix.length) throw new Error("The first worksheet is empty.");
    const maxColumns = Math.max(...matrix.map((row) => row.length));
    const headers = uniqueHeaders(Array.from({ length: maxColumns }, (_, i) => matrix[0][i]));
    const rows = matrix.slice(1)
      .filter((row) => row.some((value) => String(value == null ? "" : value).trim() !== ""))
      .map((row, rowIndex) => ({
        excelRow: rowIndex + 2,
        values: Array.from({ length: maxColumns }, (_, i) => row[i] == null ? "" : row[i]),
      }));

    return {
      fileName: file.name,
      sheetName,
      headers,
      rows,
      nameColumnIndex: chooseNameColumn(headers),
    };
  }

  async function handleFile(side, file) {
    hideError();
    els.results.classList.add("is-hidden");
    try {
      const data = await readSpreadsheet(file);
      state[side] = data;
      setFileStatus(side, data);
      els.compare.disabled = !(state.a && state.b);
    } catch (error) {
      state[side] = null;
      els.compare.disabled = true;
      showError(error.message || "Please select a valid Excel or CSV file and try again.");
    }
  }

  function groupRows(fileData, columnIndex, ignoreOrder) {
    const groups = new Map();
    fileData.rows.forEach((row) => {
      const rawName = row.values[columnIndex];
      const key = normalizeName(rawName, ignoreOrder);
      if (!key) return;
      if (!groups.has(key)) groups.set(key, []);
      groups.get(key).push({ ...row, rawName });
    });
    return groups;
  }

  function compareFiles() {
    hideError();
    if (!state.a || !state.b) return;

    const columnA = Number(els.columnA.value);
    const columnB = Number(els.columnB.value);
    const ignoreOrder = els.ignoreOrder.checked;
    state.a.nameColumnIndex = columnA;
    state.b.nameColumnIndex = columnB;

    const groupsA = groupRows(state.a, columnA, ignoreOrder);
    const groupsB = groupRows(state.b, columnB, ignoreOrder);
    const matches = [];
    let uniqueMatches = 0;

    groupsA.forEach((rowsFromA, key) => {
      const rowsFromB = groupsB.get(key);
      if (!rowsFromB) return;
      uniqueMatches += 1;
      const occurrences = Math.max(rowsFromA.length, rowsFromB.length);
      for (let i = 0; i < occurrences; i += 1) {
        const aRow = rowsFromA[i] || null;
        const bRow = rowsFromB[i] || null;
        matches.push({
          key,
          display: displayName(aRow ? aRow.rawName : bRow.rawName),
          aRow,
          bRow,
          countA: rowsFromA.length,
          countB: rowsFromB.length,
        });
      }
    });

    matches.sort((left, right) => left.display.localeCompare(right.display, undefined, { sensitivity: "base" }));
    state.matches = matches;
    state.uniqueMatches = uniqueMatches;
    renderResults();
  }

  function renderResults() {
    els.matchCount.textContent = state.matches.length.toLocaleString();
    els.rowsA.textContent = state.a.rows.length.toLocaleString();
    els.rowsB.textContent = state.b.rows.length.toLocaleString();
    els.uniqueCount.textContent = state.uniqueMatches.toLocaleString();
    els.previewBody.textContent = "";

    state.matches.slice(0, 100).forEach((match, index) => {
      const tr = document.createElement("tr");
      const values = [
        index + 1,
        match.display,
        match.aRow ? match.aRow.excelRow : "—",
        match.bRow ? match.bRow.excelRow : "—",
      ];
      values.forEach((value) => {
        const td = document.createElement("td");
        td.textContent = String(value);
        tr.appendChild(td);
      });
      els.previewBody.appendChild(tr);
    });

    const hasMatches = state.matches.length > 0;
    els.tableWrap.classList.toggle("is-hidden", !hasMatches);
    els.noMatches.classList.toggle("is-hidden", hasMatches);
    els.download.classList.toggle("is-hidden", !hasMatches);
    els.results.classList.remove("is-hidden");
    els.results.scrollIntoView({ behavior: "auto", block: "start" });
  }

  function safeSheetValue(value) {
    if (value instanceof Date) return value;
    return escapeCellText(value);
  }

  function buildOutputRows() {
    const aHeaders = state.a.headers.map((header) => `File 1 - ${header}`);
    const bHeaders = state.b.headers.map((header) => `File 2 - ${header}`);
    const headers = [
      "Matching Name",
      "File 1 Row",
      "File 2 Row",
      "Occurrences in File 1",
      "Occurrences in File 2",
      ...aHeaders,
      ...bHeaders,
    ];

    const rows = state.matches.map((match) => [
      safeSheetValue(match.display),
      match.aRow ? match.aRow.excelRow : "",
      match.bRow ? match.bRow.excelRow : "",
      match.countA,
      match.countB,
      ...(match.aRow ? match.aRow.values.map(safeSheetValue) : state.a.headers.map(() => "")),
      ...(match.bRow ? match.bRow.values.map(safeSheetValue) : state.b.headers.map(() => "")),
    ]);

    return [headers, ...rows];
  }

  function downloadWorkbook() {
    try {
      if (!window.XLSX || !state.matches.length) return;
      const workbook = XLSX.utils.book_new();
      const matchSheet = XLSX.utils.aoa_to_sheet(buildOutputRows());
      matchSheet["!freeze"] = { xSplit: 0, ySplit: 1 };
      matchSheet["!autofilter"] = { ref: matchSheet["!ref"] };
      matchSheet["!cols"] = buildOutputRows()[0].map((header, index) => ({
        wch: Math.min(index === 0 ? 32 : Math.max(12, String(header).length + 2), 34),
      }));

      const summaryRows = [
        ["NameMatch Excel - Comparison Summary", ""],
        ["First File", safeSheetValue(state.a.fileName)],
        ["Second File", safeSheetValue(state.b.fileName)],
        ["Name Column - File 1", safeSheetValue(state.a.headers[state.a.nameColumnIndex])],
        ["Name Column - File 2", safeSheetValue(state.b.headers[state.b.nameColumnIndex])],
        ["Rows in File 1", state.a.rows.length],
        ["Rows in File 2", state.b.rows.length],
        ["Unique Matching Names", state.uniqueMatches],
        ["Matching Output Rows", state.matches.length],
        ["Comparison Rule", els.ignoreOrder.checked ? "Capitalization, punctuation, spacing, and name order ignored" : "Capitalization, punctuation, and spacing ignored"],
        ["Created", new Date()],
      ];
      const summarySheet = XLSX.utils.aoa_to_sheet(summaryRows);
      summarySheet["!cols"] = [{ wch: 28 }, { wch: 58 }];

      XLSX.utils.book_append_sheet(workbook, matchSheet, "Matching Names");
      XLSX.utils.book_append_sheet(workbook, summarySheet, "Summary");
      const stamp = new Date().toISOString().slice(0, 10);
      XLSX.writeFile(workbook, `Matching-Names-${stamp}.xlsx`, { compression: true });
    } catch (error) {
      showError(error.message || "The Excel report could not be created. Please try again.");
    }
  }

  function resetApp() {
    state.a = null;
    state.b = null;
    state.matches = [];
    state.uniqueMatches = 0;
    els.fileA.value = "";
    els.fileB.value = "";
    els.cardA.classList.remove("has-file");
    els.cardB.classList.remove("has-file");
    els.statusA.innerHTML = "<strong>No file selected</strong><span>Excel or CSV</span>";
    els.statusB.innerHTML = "<strong>No file selected</strong><span>Excel or CSV</span>";
    els.columnWrapA.classList.add("is-hidden");
    els.columnWrapB.classList.add("is-hidden");
    els.compare.disabled = true;
    els.results.classList.add("is-hidden");
    hideError();
    window.scrollTo({ top: 0, behavior: "auto" });
  }

  els.fileA.addEventListener("change", (event) => handleFile("a", event.target.files[0]));
  els.fileB.addEventListener("change", (event) => handleFile("b", event.target.files[0]));
  els.compare.addEventListener("click", compareFiles);
  els.download.addEventListener("click", downloadWorkbook);
  els.reset.addEventListener("click", resetApp);
})();
