/**
 * Логика веб-приложения учета школьной посещаемости (Чистый JavaScript + ExcelJS).
 * Работает 100% на стороне браузера без серверов и без Python.
 */

// Эталонный список из 29 классов
const STANDARD_CLASSES = [
  "1а", "1б", "2а", "2б", "2в",
  "3а", "3б", "3в",
  "4а", "4б", "4в",
  "5а", "5б", "5в",
  "6а", "6б", "6в",
  "7а", "7б", "7в",
  "8а", "8б", "8в", "8г",
  "9а", "9б", "9в",
  "10", "11",
];

const GRADES_1_4 = new Set(STANDARD_CLASSES.slice(0, 11));   // 11 классов
const GRADES_5_9 = new Set(STANDARD_CLASSES.slice(11, 27));  // 16 классов
const GRADES_10_11 = new Set(STANDARD_CLASSES.slice(27, 29)); // 2 класса

// Замена похожих символов латиницы на кириллицу
const HOMOGLYPHS = {
  a: "а", b: "б", c: "с", e: "е", k: "к",
  m: "м", o: "о", p: "р", t: "т", x: "х",
};

const MONTH_NAMES_RU = {
  1: "Январь", 2: "Февраль", 3: "Март", 4: "Апрель",
  5: "Май", 6: "Июнь", 7: "Июль", 8: "Август",
  9: "Сентябрь", 10: "Октябрь", 11: "Ноябрь", 12: "Декабрь",
};

/**
 * Нормализует название класса (удаление пробелов, нижний регистр, латиница -> кириллица)
 */
function normalizeClassName(name) {
  if (name === null || name === undefined) return "";
  let s = String(name).trim().toLowerCase();
  let res = "";
  for (let ch of s) {
    res += HOMOGLYPHS[ch] || ch;
  }
  return res.replace(/\s+/g, "");
}

/**
 * Парсит дату из различных форматов в объект { dateObj, iso, display }
 */
function parseDateValue(val) {
  if (!val) return null;

  if (val instanceof Date) {
    if (isNaN(val.getTime())) return null;
    const y = val.getFullYear();
    const m = String(val.getMonth() + 1).padStart(2, "0");
    const d = String(val.getDate()).padStart(2, "0");
    return {
      dateObj: new Date(y, val.getMonth(), val.getDate()),
      iso: `${y}-${m}-${d}`,
      display: `${d}.${m}.${y}`,
      year: y,
      month: val.getMonth() + 1,
    };
  }

  const s = String(val).trim();
  if (!s) return null;

  // YYYY-MM-DD
  let match = s.match(/(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (match) {
    const y = parseInt(match[1], 10);
    const m = parseInt(match[2], 10);
    const d = parseInt(match[3], 10);
    return {
      dateObj: new Date(y, m - 1, d),
      iso: `${y}-${String(m).padStart(2, "0")}-${String(d).padStart(2, "0")}`,
      display: `${String(d).padStart(2, "0")}.${String(m).padStart(2, "0")}.${y}`,
      year: y,
      month: m,
    };
  }

  // DD.MM.YYYY
  match = s.match(/(\d{1,2})\.(\d{1,2})\.(\d{4})/);
  if (match) {
    const d = parseInt(match[1], 10);
    const m = parseInt(match[2], 10);
    const y = parseInt(match[3], 10);
    return {
      dateObj: new Date(y, m - 1, d),
      iso: `${y}-${String(m).padStart(2, "0")}-${String(d).padStart(2, "0")}`,
      display: `${String(d).padStart(2, "0")}.${String(m).padStart(2, "0")}.${y}`,
      year: y,
      month: m,
    };
  }

  return null;
}

function parseIntSafe(val, def = 0) {
  if (val === null || val === undefined) return def;
  const n = parseInt(String(val).replace(",", ".").trim(), 10);
  return isNaN(n) || n < 0 ? def : n;
}

function parseStrSafe(val) {
  if (val === null || val === undefined) return null;
  const s = String(val).trim();
  return s.length > 0 ? s : null;
}

/**
 * Читает и анализирует загруженный файл выгрузки формы через ExcelJS
 */
async function loadFormResponses(arrayBuffer) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(arrayBuffer);
  const ws = wb.worksheets[0];
  if (!ws) throw new Error("В файле нет листов");

  // Поиск индексов колонок по заголовкам первой строки
  const headerRow = ws.getRow(1);
  const headerIndices = {};

  headerRow.eachCell({ includeEmpty: false }, (cell, colNumber) => {
    const h = String(cell.value || "").trim().toLowerCase();
    if (h === "id") headerIndices.id = colNumber;
    else if (h.includes("время создания")) headerIndices.created_at = colNumber;
    else if (h.includes("класс")) headerIndices.class = colNumber;
    else if (h.includes("дата") || h.includes("дату")) headerIndices.date = colNumber;
    else if (h.includes("болезни") && h.includes("сколько")) headerIndices.sick_count = colNumber;
    else if (h.includes("болезни") && (h.includes("фио") || h.includes("введите"))) headerIndices.sick_names = colNumber;
    else if (h.includes("уважительной") && h.includes("сколько")) headerIndices.excused_count = colNumber;
    else if (h.includes("уважительной") && (h.includes("фио") || h.includes("введите"))) headerIndices.excused_names = colNumber;
  });

  // Запасные позиции (1-indexed в ExcelJS)
  const colId = headerIndices.id || 1;
  const colCreatedAt = headerIndices.created_at || 3;
  const colClass = headerIndices.class || 5;
  const colDate = headerIndices.date || 6;
  const colSickCount = headerIndices.sick_count || 7;
  const colSickNames = headerIndices.sick_names || 8;
  const colExcCount = headerIndices.excused_count || 9;
  const colExcNames = headerIndices.excused_names || 10;

  const recordsByIso = {};
  const dateMetaMap = {};

  for (let r = 2; r <= ws.rowCount; r++) {
    const row = ws.getRow(r);
    const rawClass = row.getCell(colClass).value;
    const normClass = normalizeClassName(rawClass);
    if (!normClass) continue;

    const rawDate = row.getCell(colDate).value;
    let parsedDate = parseDateValue(rawDate);
    if (!parsedDate) {
      const rawCreated = row.getCell(colCreatedAt).value;
      parsedDate = parseDateValue(rawCreated);
    }
    if (!parsedDate) continue;

    const iso = parsedDate.iso;
    if (!recordsByIso[iso]) {
      recordsByIso[iso] = [];
      dateMetaMap[iso] = parsedDate;
    }

    recordsByIso[iso].append = recordsByIso[iso].push({
      rowIdx: r,
      id: String(row.getCell(colId).value || r),
      createdAt: String(row.getCell(colCreatedAt).value || ""),
      class: normClass,
      dateInfo: parsedDate,
      sickCount: parseIntSafe(row.getCell(colSickCount).value, 0),
      sickNames: parseStrSafe(row.getCell(colSickNames).value),
      excusedCount: parseIntSafe(row.getCell(colExcCount).value, 0),
      excusedNames: parseStrSafe(row.getCell(colExcNames).value),
    });
  }

  const sortedIsos = Object.keys(recordsByIso).sort();
  if (sortedIsos.length === 0) {
    throw new Error("В файле не найдено строк с датами и классами. Проверьте структуру файла.");
  }

  // Разрешение дубликатов внутри каждого дня
  const dedupedByIso = {};
  const warningsByIso = {};

  for (let iso of sortedIsos) {
    const items = recordsByIso[iso];
    const classMap = {};
    const warns = [];

    for (let item of items) {
      const c = item.class;
      if (classMap[c]) {
        warns.push(`Класс ${c}: заполнили форму несколько раз. Использован последний ответ на форму.`);
      }
      classMap[c] = item;
    }
    dedupedByIso[iso] = classMap;
    warningsByIso[iso] = warns;
  }

  // Группировка по месяцам
  const monthsMap = {};
  for (let iso of sortedIsos) {
    const meta = dateMetaMap[iso];
    const mKey = `${MONTH_NAMES_RU[meta.month]} ${meta.year}`;
    if (!monthsMap[mKey]) monthsMap[mKey] = [];
    monthsMap[mKey].push(meta.display);
  }

  const datesDisplay = sortedIsos.map((iso) => dateMetaMap[iso].display);
  const displayToIso = {};
  sortedIsos.forEach((iso) => {
    displayToIso[dateMetaMap[iso].display] = iso;
  });

  return {
    sortedIsos,
    dates: datesDisplay,
    datesIso: sortedIsos,
    dateMetaMap,
    displayToIso,
    months: Object.keys(monthsMap),
    monthsMap,
    dateCounts: sortedIsos.reduce((acc, iso) => {
      acc[dateMetaMap[iso].display] = recordsByIso[iso].length;
      return acc;
    }, {}),
    totalRecords: Object.values(recordsByIso).reduce((acc, arr) => acc + arr.length, 0),
    dedupedByIso,
    warningsByIso,
  };
}

/**
 * Рассчитывает статистику посещаемости за конкретную дату
 */
function computeDayStats(classData) {
  const submittedClasses = [];
  const missingClasses = [];

  let sick14 = 0, exc14 = 0;
  let sick59 = 0, exc59 = 0;
  let sick1011 = 0, exc1011 = 0;

  const tableRows = [];

  for (let c of STANDARD_CLASSES) {
    let sCount = 0, sNames = "", eCount = 0, eNames = "";
    const isSubmitted = !!classData[c];

    if (isSubmitted) {
      submittedClasses.push(c);
      const rec = classData[c];
      sCount = rec.sickCount;
      sNames = rec.sickNames || "";
      eCount = rec.excusedCount;
      eNames = rec.excusedNames || "";
    } else {
      missingClasses.push(c);
    }

    const totalAbsent = sCount + eCount;

    if (GRADES_1_4.has(c)) {
      sick14 += sCount;
      exc14 += eCount;
    } else if (GRADES_5_9.has(c)) {
      sick59 += sCount;
      exc59 += eCount;
    } else if (GRADES_10_11.has(c)) {
      sick1011 += sCount;
      exc1011 += eCount;
    }

    tableRows.push({
      class: c,
      total: totalAbsent,
      sickCount: sCount,
      sickNames: sNames,
      excusedCount: eCount,
      excusedNames: eNames,
      submitted: isSubmitted,
    });
  }

  const total14 = sick14 + exc14;
  const total59 = sick59 + exc59;
  const total1011 = sick1011 + exc1011;
  const totalAll = total14 + total59 + total1011;
  const totalSick = sick14 + sick59 + sick1011;
  const totalExc = exc14 + exc59 + exc1011;

  return {
    table: tableRows,
    submittedClasses,
    missingClasses,
    submittedCount: submittedClasses.length,
    missingCount: missingClasses.length,
    stats: {
      totalAbsent: totalAll,
      totalSick,
      totalExcused: totalExc,
      grades14: { total: total14, sick: sick14, excused: exc14 },
      grades59: { total: total59, sick: sick59, excused: exc59 },
      grades1011: { total: total1011, sick: sick1011, excused: exc1011 },
    },
  };
}

/**
 * Формирует текстовую сводку для мессенджера
 */
function generateTextSummary(dateDisplay, dayStats) {
  const s = dayStats.stats;
  const m = dayStats.missingClasses;

  const lines = [
    `📊 Сводка посещаемости на ${dateDisplay}:`,
    `Всего отсутствуют: ${s.totalAbsent} (по болезни: ${s.totalSick}, по ув. причине: ${s.totalExcused})`,
    "",
    `• 1–4 классы: ${s.grades14.total} (болезнь: ${s.grades14.sick}, ув.: ${s.grades14.excused})`,
    `• 5–9 классы: ${s.grades59.total} (болезнь: ${s.grades59.sick}, ув.: ${s.grades59.excused})`,
    `• 10–11 классы: ${s.grades1011.total} (болезнь: ${s.grades1011.sick}, ув.: ${s.grades1011.excused})`,
  ];

  if (m.length > 0) {
    lines.push("");
    lines.push(`❗ Не сдали отчет (${m.length} классов): ${m.join(", ")}`);
  } else {
    lines.push("");
    lines.push("✅ Все классы сдали отчет вовремя.");
  }

  return lines.join("\n");
}

/**
 * Генерирует многостраничный Excel (.xlsx) за выбранный период дат
 */
async function generatePeriodWorkbook(parsedInfo, selectedDisplayDates) {
  const wb = new ExcelJS.Workbook();
  wb.calcProperties.fullCalcOnLoad = true;

  const fontHeader = { name: "Arial", size: 8, bold: true };
  const fontBold = { name: "Arial", size: 10, bold: true };
  const fontStatTitle = { name: "Arial", size: 14, bold: true };
  const fontRegular = { name: "Arial", size: 10, bold: false };

  const thinBorder = {
    top: { style: "thin", color: { argb: "FF000000" } },
    left: { style: "thin", color: { argb: "FF000000" } },
    bottom: { style: "thin", color: { argb: "FF000000" } },
    right: { style: "thin", color: { argb: "FF000000" } },
  };

  const alignCenter = { vertical: "middle", horizontal: "center", wrapText: true };
  const alignLeft = { vertical: "middle", horizontal: "left" };

  // Сортировка дат строго по хронологии
  const sortedDates = [...selectedDisplayDates].sort((a, b) => {
    const isoA = parsedInfo.displayToIso[a] || "";
    const isoB = parsedInfo.displayToIso[b] || "";
    return isoA.localeCompare(isoB);
  });

  for (let dateStr of sortedDates) {
    const iso = parsedInfo.displayToIso[dateStr];
    const classData = parsedInfo.dedupedByIso[iso] || {};
    const dayStats = computeDayStats(classData);
    const s = dayStats.stats;

    const ws = wb.addWorksheet(dateStr, { views: [{ showGridLines: true }] });

    // Ширина столбцов
    ws.columns = [
      { key: "A", width: 13.0 },
      { key: "B", width: 17.5 },
      { key: "C", width: 13.0 },
      { key: "D", width: 22.0 },
      { key: "E", width: 13.0 },
      { key: "F", width: 45.0 },
      { key: "G", width: 5.0 },
      { key: "H", width: 25.0 },
      { key: "I", width: 14.0 },
      { key: "J", width: 14.0 },
      { key: "K", width: 14.0 },
      { key: "L", width: 14.0 },
    ];

    // Заголовки строки 1
    const headers = [
      ["A1", "классы "],
      ["B1", "Кол-во отсутствующих (всего)"],
      ["C1", "кол-во по болезни"],
      ["D1", "ФИ отсутствующих по болезни"],
      ["E1", "кол-во по ув. причине"],
      ["F1", "Ф.И. отсутствующих по ув. причине"],
    ];
    for (let [cellRef, text] of headers) {
      const cell = ws.getCell(cellRef);
      cell.value = text;
      cell.font = fontHeader;
      cell.alignment = alignCenter;
      cell.border = thinBorder;
    }

    // Блок «СТАТИСТИКА»
    ws.mergeCells("H1:L1");
    const h1 = ws.getCell("H1");
    h1.value = "СТАТИСТИКА";
    h1.font = fontStatTitle;
    h1.alignment = alignCenter;
    ["H1", "I1", "J1", "K1", "L1"].forEach((ref) => {
      ws.getCell(ref).border = thinBorder;
    });

    // Строка 2 статистики
    const statCols = [
      ["I2", "1-4 классы"],
      ["J2", "5-9 классы"],
      ["K2", "10-11 классы"],
      ["L2", "1-11 классы"],
    ];
    for (let [ref, text] of statCols) {
      const cell = ws.getCell(ref);
      cell.value = text;
      cell.font = fontBold;
      cell.alignment = alignCenter;
      cell.border = thinBorder;
    }

    // Строки 3..5 статистики (с живыми формулами и предварительно вычисленными результатами)
    const statRows = [
      {
        row: 3,
        label: "Отсутствует всего",
        formulas: [
          ["I3", "SUM(B2:B12)", s.grades14.total],
          ["J3", "SUM(B13:B28)", s.grades59.total],
          ["K3", "SUM(B29:B30)", s.grades1011.total],
          ["L3", "SUM(I3:K3)", s.totalAbsent],
        ],
      },
      {
        row: 4,
        label: "Отсутствует по ув",
        formulas: [
          ["I4", "SUM(E2:E12)", s.grades14.excused],
          ["J4", "SUM(E13:E28)", s.grades59.excused],
          ["K4", "SUM(E29:E30)", s.grades1011.excused],
          ["L4", "SUM(I4:K4)", s.totalExcused],
        ],
      },
      {
        row: 5,
        label: "Отсутствует по болезни",
        formulas: [
          ["I5", "SUM(C2:C12)", s.grades14.sick],
          ["J5", "SUM(C13:C28)", s.grades59.sick],
          ["K5", "SUM(C29:C30)", s.grades1011.sick],
          ["L5", "SUM(I5:K5)", s.totalSick],
        ],
      },
    ];

    for (let sr of statRows) {
      const hCell = ws.getCell(`H${sr.row}`);
      hCell.value = sr.label;
      hCell.font = fontRegular;
      hCell.alignment = alignLeft;
      hCell.border = thinBorder;

      for (let [cellRef, formulaStr, resultVal] of sr.formulas) {
        const cCell = ws.getCell(cellRef);
        cCell.value = { formula: formulaStr, result: resultVal };
        cCell.font = fontBold;
        cCell.alignment = alignCenter;
        cCell.border = thinBorder;
      }
    }

    // Заполнение 29 классов (строки 2..30)
    for (let idx = 0; idx < STANDARD_CLASSES.length; idx++) {
      const r = idx + 2;
      const c = STANDARD_CLASSES[idx];
      const rec = classData[c];

      const sCount = rec ? rec.sickCount : 0;
      const sNames = rec ? rec.sickNames : null;
      const eCount = rec ? rec.excusedCount : 0;
      const eNames = rec ? rec.excusedNames : null;
      const totalClass = sCount + eCount;

      // Ячейка A: класс
      const cellA = ws.getCell(`A${r}`);
      cellA.value = c;
      cellA.font = fontBold;
      cellA.alignment = alignCenter;
      cellA.border = thinBorder;

      // Ячейка B: формула суммы
      const cellB = ws.getCell(`B${r}`);
      cellB.value = { formula: `C${r}+E${r}`, result: totalClass };
      cellB.font = fontBold;
      cellB.alignment = alignCenter;
      cellB.border = thinBorder;

      // Ячейка C: по болезни
      const cellC = ws.getCell(`C${r}`);
      cellC.value = sCount;
      cellC.font = fontRegular;
      cellC.alignment = alignCenter;
      cellC.border = thinBorder;

      // Ячейка D: ФИО по болезни
      const cellD = ws.getCell(`D${r}`);
      cellD.value = sNames;
      cellD.font = fontRegular;
      cellD.alignment = alignLeft;
      cellD.border = thinBorder;

      // Ячейка E: по ув. причине
      const cellE = ws.getCell(`E${r}`);
      cellE.value = eCount;
      cellE.font = fontRegular;
      cellE.alignment = alignCenter;
      cellE.border = thinBorder;

      // Ячейка F: ФИО по ув. причине
      const cellF = ws.getCell(`F${r}`);
      cellF.value = eNames;
      cellF.font = fontRegular;
      cellF.alignment = alignLeft;
      cellF.border = thinBorder;
    }
  }

  return await wb.xlsx.writeBuffer();
}

// ====================================================
// Инициализация интерфейса (UI) в браузере
// ====================================================
document.addEventListener("DOMContentLoaded", () => {
  const dropZone = document.getElementById("dropZone");
  const fileInput = document.getElementById("fileInput");
  const browseBtn = document.getElementById("browseBtn");
  const fileInfoBar = document.getElementById("fileInfoBar");
  const fileNameEl = document.getElementById("fileName");
  const fileMetaEl = document.getElementById("fileMeta");
  const changeFileBtn = document.getElementById("changeFileBtn");

  const periodSection = document.getElementById("periodSection");
  const previewSection = document.getElementById("previewSection");

  const periodSummaryBadge = document.getElementById("periodSummaryBadge");
  const monthControl = document.getElementById("monthControl");
  const monthSelect = document.getElementById("monthSelect");
  const rangeControl = document.getElementById("rangeControl");
  const dateFrom = document.getElementById("dateFrom");
  const dateTo = document.getElementById("dateTo");
  const singleControl = document.getElementById("singleControl");
  const singleDateSelect = document.getElementById("singleDateSelect");
  const downloadPeriodBtn = document.getElementById("downloadPeriodBtn");
  const downloadSubtext = document.getElementById("downloadSubtext");

  const previewDateSelect = document.getElementById("previewDateSelect");
  const alertsContainer = document.getElementById("alertsContainer");
  const statTotalAbsent = document.getElementById("statTotalAbsent");
  const statTotalSick = document.getElementById("statTotalSick");
  const statTotalExcused = document.getElementById("statTotalExcused");
  const statGrades14 = document.getElementById("statGrades14");
  const statSick14 = document.getElementById("statSick14");
  const statExcused14 = document.getElementById("statExcused14");
  const statGrades59 = document.getElementById("statGrades59");
  const statSick59 = document.getElementById("statSick59");
  const statExcused59 = document.getElementById("statExcused59");
  const statGrades1011 = document.getElementById("statGrades1011");
  const statSick1011 = document.getElementById("statSick1011");
  const statExcused1011 = document.getElementById("statExcused1011");

  const copySummaryBtn = document.getElementById("copySummaryBtn");
  const copySummaryFloatBtn = document.getElementById("copySummaryFloatBtn");
  const copyMissingBtn = document.getElementById("copyMissingBtn");
  const missingBtnCount = document.getElementById("missingBtnCount");

  const classesTableBody = document.getElementById("classesTableBody");
  const summaryTextContent = document.getElementById("summaryTextContent");
  const missingTagsContainer = document.getElementById("missingTagsContainer");
  const tabMissingBadge = document.getElementById("tabMissingBadge");
  const toast = document.getElementById("toast");

  // Текущее состояние
  let parsedInfo = null;
  let currentFile = null;
  let currentPreviewDate = null;
  let currentPreviewData = null;
  let currentTableFilter = "all";

  // Обработка выбора файла
  browseBtn.addEventListener("click", () => fileInput.click());
  changeFileBtn.addEventListener("click", () => fileInput.click());

  fileInput.addEventListener("change", (e) => {
    if (e.target.files.length > 0) processFile(e.target.files[0]);
  });

  dropZone.addEventListener("dragover", (e) => {
    e.preventDefault();
    dropZone.classList.add("dragover");
  });

  dropZone.addEventListener("dragleave", () => {
    dropZone.classList.remove("dragover");
  });

  dropZone.addEventListener("drop", (e) => {
    e.preventDefault();
    dropZone.classList.remove("dragover");
    if (e.dataTransfer.files.length > 0) processFile(e.dataTransfer.files[0]);
  });

  async function processFile(file) {
    if (!file.name.match(/\.(xlsx|xlsm|xltx)$/i)) {
      showToast("Пожалуйста, выберите файл Excel (.xlsx)", "error");
      return;
    }

    currentFile = file;
    showToast("Чтение и обработка файла...", "info");

    try {
      const buffer = await file.arrayBuffer();
      parsedInfo = await loadFormResponses(buffer);

      fileNameEl.textContent = file.name;
      fileMetaEl.textContent = `Записей: ${parsedInfo.totalRecords} • Учебных дней: ${parsedInfo.dates.length}`;
      dropZone.classList.add("hidden");
      fileInfoBar.classList.remove("hidden");

      setupPeriodControls();

      // Отображаем предпросмотр первой даты
      currentPreviewDate = parsedInfo.dates[0];
      previewDateSelect.value = currentPreviewDate;
      renderPreviewForDate(currentPreviewDate);

      periodSection.classList.remove("hidden");
      previewSection.classList.remove("hidden");

      showToast("Файл успешно прочитан!", "success");
    } catch (err) {
      showToast(err.message || "Ошибка при чтении файла", "error");
    }
  }

  function setupPeriodControls() {
    // 1. Месяцы
    monthSelect.innerHTML = "";
    parsedInfo.months.forEach((m) => {
      const opt = document.createElement("option");
      opt.value = m;
      opt.textContent = `${m} (${parsedInfo.monthsMap[m].length} дн.)`;
      monthSelect.appendChild(opt);
    });

    // 2. Диапазон дат
    if (parsedInfo.datesIso.length > 0) {
      dateFrom.min = parsedInfo.datesIso[0];
      dateFrom.max = parsedInfo.datesIso[parsedInfo.datesIso.length - 1];
      dateFrom.value = parsedInfo.datesIso[0];

      dateTo.min = parsedInfo.datesIso[0];
      dateTo.max = parsedInfo.datesIso[parsedInfo.datesIso.length - 1];
      dateTo.value = parsedInfo.datesIso[parsedInfo.datesIso.length - 1];
    }

    // 3. Выбор одного дня
    singleDateSelect.innerHTML = "";
    previewDateSelect.innerHTML = "";
    parsedInfo.dates.forEach((d) => {
      const opt1 = document.createElement("option");
      opt1.value = d;
      opt1.textContent = d;
      singleDateSelect.appendChild(opt1);

      const opt2 = document.createElement("option");
      opt2.value = d;
      opt2.textContent = d;
      previewDateSelect.appendChild(opt2);
    });

    updatePeriodSummary();
  }

  // Переключение режимов периода
  const periodRadios = document.querySelectorAll('input[name="periodMode"]');
  periodRadios.forEach((r) => {
    r.addEventListener("change", (e) => {
      document.querySelectorAll(".mode-pill").forEach((pill) => pill.classList.remove("active"));
      e.target.closest(".mode-pill").classList.add("active");

      const mode = e.target.value;
      monthControl.classList.toggle("hidden", mode !== "month");
      rangeControl.classList.toggle("hidden", mode !== "range");
      singleControl.classList.toggle("hidden", mode !== "single");

      updatePeriodSummary();
    });
  });

  monthSelect.addEventListener("change", updatePeriodSummary);
  dateFrom.addEventListener("change", updatePeriodSummary);
  dateTo.addEventListener("change", updatePeriodSummary);
  singleDateSelect.addEventListener("change", (e) => {
    updatePeriodSummary();
    previewDateSelect.value = e.target.value;
    renderPreviewForDate(e.target.value);
  });

  function getSelectedDatesList() {
    if (!parsedInfo) return [];
    const mode = document.querySelector('input[name="periodMode"]:checked').value;

    if (mode === "all") {
      return parsedInfo.dates;
    } else if (mode === "month") {
      const m = monthSelect.value;
      return parsedInfo.monthsMap[m] || [];
    } else if (mode === "range") {
      const from = dateFrom.value;
      const to = dateTo.value;
      const filtered = [];
      for (let i = 0; i < parsedInfo.datesIso.length; i++) {
        const iso = parsedInfo.datesIso[i];
        if ((!from || iso >= from) && (!to || iso <= to)) {
          filtered.push(parsedInfo.dates[i]);
        }
      }
      return filtered;
    } else if (mode === "single") {
      return [singleDateSelect.value];
    }
    return parsedInfo.dates;
  }

  function updatePeriodSummary() {
    const list = getSelectedDatesList();
    const count = list.length;
    periodSummaryBadge.textContent = `Листов в отчете: ${count}`;

    if (count === 1) {
      downloadSubtext.textContent = `(1 лист: ${list[0]})`;
    } else if (count > 1) {
      downloadSubtext.textContent = `(${count} листов по датам: ${list[0]} — ${list[list.length - 1]})`;
    } else {
      downloadSubtext.textContent = "(нет дат в выбранном диапазоне)";
    }
  }

  // Предпросмотр
  previewDateSelect.addEventListener("change", (e) => {
    renderPreviewForDate(e.target.value);
  });

  function renderPreviewForDate(dateStr) {
    if (!parsedInfo) return;
    currentPreviewDate = dateStr;
    const iso = parsedInfo.displayToIso[dateStr];
    const classData = parsedInfo.dedupedByIso[iso] || {};
    const dayStats = computeDayStats(classData);
    const textSummary = generateTextSummary(dateStr, dayStats);
    const warnings = parsedInfo.warningsByIso[iso] || [];

    currentPreviewData = {
      date: dateStr,
      dayStats,
      textSummary,
      warnings,
    };

    // Статистика
    const s = dayStats.stats;
    statTotalAbsent.textContent = s.totalAbsent;
    statTotalSick.textContent = s.totalSick;
    statTotalExcused.textContent = s.totalExcused;

    statGrades14.textContent = s.grades14.total;
    statSick14.textContent = s.grades14.sick;
    statExcused14.textContent = s.grades14.excused;

    statGrades59.textContent = s.grades59.total;
    statSick59.textContent = s.grades59.sick;
    statExcused59.textContent = s.grades59.excused;

    statGrades1011.textContent = s.grades1011.total;
    statSick1011.textContent = s.grades1011.sick;
    statExcused1011.textContent = s.grades1011.excused;

    missingBtnCount.textContent = dayStats.missingCount;
    tabMissingBadge.textContent = dayStats.missingCount;

    // Предупреждения
    alertsContainer.innerHTML = "";
    if (warnings.length > 0) {
      alertsContainer.classList.remove("hidden");
      warnings.forEach((w) => {
        const div = document.createElement("div");
        div.className = "alert-item";
        div.textContent = `⚠️ ${w}`;
        alertsContainer.appendChild(div);
      });
    } else {
      alertsContainer.classList.add("hidden");
    }

    // Текстовая сводка
    summaryTextContent.textContent = textSummary;

    // Теги не сдавших классов
    missingTagsContainer.innerHTML = "";
    if (dayStats.missingClasses.length > 0) {
      dayStats.missingClasses.forEach((c) => {
        const span = document.createElement("span");
        span.className = "class-tag";
        span.textContent = c;
        missingTagsContainer.appendChild(span);
      });
    } else {
      missingTagsContainer.innerHTML = "<p>🎉 Все 29 классов сдали отчет за этот день!</p>";
    }

    renderClassesTable();
  }

  function renderClassesTable() {
    if (!currentPreviewData) return;
    classesTableBody.innerHTML = "";

    const rows = currentPreviewData.dayStats.table;
    rows.forEach((r) => {
      if (currentTableFilter === "absent" && r.total === 0) return;
      if (currentTableFilter === "missing" && r.submitted) return;

      const tr = document.createElement("tr");
      if (r.total > 0) tr.classList.add("has-absent");
      if (!r.submitted) tr.classList.add("missing-row");

      const statusTag = r.submitted
        ? '<span class="status-tag status-ok">Сдан</span>'
        : '<span class="status-tag status-missing">Нет отчета</span>';

      tr.innerHTML = `
        <td><strong>${r.class}</strong></td>
        <td><strong>${r.total}</strong></td>
        <td>${r.sickCount}</td>
        <td>${r.sickNames || "—"}</td>
        <td>${r.excusedCount}</td>
        <td>${r.excusedNames || "—"}</td>
        <td>${statusTag}</td>
      `;
      classesTableBody.appendChild(tr);
    });
  }

  // Фильтры таблицы
  const filterChips = document.querySelectorAll(".filter-chip");
  filterChips.forEach((chip) => {
    chip.addEventListener("click", (e) => {
      filterChips.forEach((c) => c.classList.remove("active"));
      e.target.classList.add("active");
      currentTableFilter = e.target.dataset.filter;
      renderClassesTable();
    });
  });

  // Вкладки
  const tabBtns = document.querySelectorAll(".tab-btn");
  tabBtns.forEach((btn) => {
    btn.addEventListener("click", (e) => {
      tabBtns.forEach((b) => b.classList.remove("active"));
      document.querySelectorAll(".tab-pane").forEach((p) => p.classList.remove("active"));

      e.target.closest(".tab-btn").classList.add("active");
      const targetPane = document.getElementById(e.target.closest(".tab-btn").dataset.tab);
      if (targetPane) targetPane.classList.add("active");
    });
  });

  // Копирование в буфер
  copySummaryBtn.addEventListener("click", copySummaryToClipboard);
  copySummaryFloatBtn.addEventListener("click", copySummaryToClipboard);

  function copySummaryToClipboard() {
    if (!currentPreviewData) return;
    navigator.clipboard
      .writeText(currentPreviewData.textSummary)
      .then(() => showToast("Текстовая сводка скопирована в буфер обмена!", "success"))
      .catch(() => showToast("Не удалось скопировать", "error"));
  }

  copyMissingBtn.addEventListener("click", () => {
    if (!currentPreviewData) return;
    const missing = currentPreviewData.dayStats.missingClasses;
    if (missing.length === 0) {
      showToast("Все классы сдали отчет!", "info");
      return;
    }
    const text = `❗ Не сдали отчет по посещаемости на ${currentPreviewDate} (${missing.length} классов):\n${missing.join(", ")}`;
    navigator.clipboard
      .writeText(text)
      .then(() => showToast("Список классов-должников скопирован!", "success"))
      .catch(() => showToast("Не удалось скопировать", "error"));
  });

  // ----------------------------------------------------
  // Скачивание сформированного Excel прямо в браузере!
  // ----------------------------------------------------
  downloadPeriodBtn.addEventListener("click", async () => {
    if (!parsedInfo) return;

    const selectedDates = getSelectedDatesList();
    if (selectedDates.length === 0) {
      showToast("Не выбрано ни одной даты для формирования отчета", "error");
      return;
    }

    showToast("Формирование многостраничной таблицы Excel...", "info");

    try {
      const buffer = await generatePeriodWorkbook(parsedInfo, selectedDates);

      // Имя файла
      const mode = document.querySelector('input[name="periodMode"]:checked').value;
      let filename = "Посещаемость.xlsx";
      if (selectedDates.length === 1) {
        filename = `Посещаемость_${selectedDates[0]}.xlsx`;
      } else if (mode === "month") {
        const cleanMonth = monthSelect.value.replace(/\s+/g, "_");
        filename = `Посещаемость_${cleanMonth}.xlsx`;
      } else {
        filename = `Посещаемость_${selectedDates[0]}-${selectedDates[selectedDates.length - 1]}.xlsx`;
      }

      // Браузерное скачивание
      const blob = new Blob([buffer], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      a.remove();
      window.URL.revokeObjectURL(url);

      showToast(`Файл "${filename}" успешно скачан!`, "success");
    } catch (err) {
      showToast("Ошибка при формировании Excel: " + (err.message || err), "error");
    }
  });

  // Всплывающие уведомления
  let toastTimer = null;
  function showToast(msg, type = "info") {
    clearTimeout(toastTimer);
    toast.textContent = msg;
    toast.className = "toast show";
    if (type === "success") toast.classList.add("toast-success");
    if (type === "error") toast.classList.add("toast-error");

    toastTimer = setTimeout(() => {
      toast.className = "toast";
    }, 3500);
  }

  // Модальная фотогалерея школы (Lightbox с перелистыванием)
  const SCHOOL_PHOTOS = [
    {
      src: "images/school-image.jpg",
      title: "МОУ СОШ №12",
      subtitle: "Главный вход и празднично украшенный фасад здания",
    },
    {
      src: "images/school-image2.jpg",
      title: "Фойе первого этажа",
      subtitle: "Пространство «Вместе создаём будущее!» и стенды Новоалександровского округа",
    },
    {
      src: "images/school-image3.jpg",
      title: "Кабинет искусства и музыки",
      subtitle: "Светлый учебный класс с новыми партами и оформлением",
    },
    {
      src: "images/school-image4.jpg",
      title: "Праздничный учебный класс",
      subtitle: "Кабинет с интерактивным оборудованием и оформлением",
    },
    {
      src: "images/school-image5.jpg",
      title: "Школьный холл и зона отдыха",
      subtitle: "Просторная светлая рекреация с мягкими диванами и панорамными окнами",
    },
  ];

  let currentPhotoIndex = 0;

  const schoolLogoThumb = document.getElementById("schoolLogoThumb");
  const schoolPhotoModal = document.getElementById("schoolPhotoModal");
  const closePhotoModalBtn = document.getElementById("closePhotoModalBtn");
  const galleryPrevBtn = document.getElementById("galleryPrevBtn");
  const galleryNextBtn = document.getElementById("galleryNextBtn");
  const modalSchoolImg = document.getElementById("modalSchoolImg");
  const galleryTitle = document.getElementById("galleryTitle");
  const gallerySubtitle = document.getElementById("gallerySubtitle");
  const galleryCounter = document.getElementById("galleryCounter");
  const galleryDots = document.getElementById("galleryDots");

  // Предзагрузка изображений для плавного переключения
  try {
    SCHOOL_PHOTOS.forEach((item) => {
      const preImg = new Image();
      preImg.src = item.src;
    });
  } catch (_) {}

  function renderPhoto(index) {
    if (index < 0) index = SCHOOL_PHOTOS.length - 1;
    if (index >= SCHOOL_PHOTOS.length) index = 0;
    currentPhotoIndex = index;
    const photo = SCHOOL_PHOTOS[currentPhotoIndex];

    if (modalSchoolImg) {
      modalSchoolImg.style.opacity = "0.4";
      modalSchoolImg.src = photo.src;
      modalSchoolImg.alt = photo.title;
      modalSchoolImg.onload = () => {
        modalSchoolImg.style.opacity = "1";
      };
    }
    if (galleryTitle) galleryTitle.textContent = photo.title;
    if (gallerySubtitle) gallerySubtitle.textContent = photo.subtitle;
    if (galleryCounter) galleryCounter.textContent = `${currentPhotoIndex + 1} / ${SCHOOL_PHOTOS.length}`;

    if (galleryDots) {
      galleryDots.innerHTML = "";
      SCHOOL_PHOTOS.forEach((_, i) => {
        const dot = document.createElement("button");
        dot.type = "button";
        dot.className = `gallery-dot ${i === currentPhotoIndex ? "active" : ""}`;
        dot.setAttribute("aria-label", `Перейти к фото ${i + 1}`);
        dot.addEventListener("click", (e) => {
          e.stopPropagation();
          renderPhoto(i);
        });
        galleryDots.appendChild(dot);
      });
    }
  }

  function prevPhoto() {
    renderPhoto(currentPhotoIndex - 1);
  }

  function nextPhoto() {
    renderPhoto(currentPhotoIndex + 1);
  }

  function openPhotoModal(startIndex = 0) {
    if (!schoolPhotoModal) return;
    renderPhoto(startIndex);
    schoolPhotoModal.classList.remove("hidden");
    requestAnimationFrame(() => {
      schoolPhotoModal.classList.add("show");
    });
    document.body.style.overflow = "hidden";
  }

  function closePhotoModal() {
    if (!schoolPhotoModal) return;
    schoolPhotoModal.classList.remove("show");
    setTimeout(() => {
      schoolPhotoModal.classList.add("hidden");
      document.body.style.overflow = "";
    }, 250);
  }

  if (schoolLogoThumb) {
    schoolLogoThumb.addEventListener("click", () => openPhotoModal(0));
    schoolLogoThumb.addEventListener("keydown", (e) => {
      if (e.key === "Enter" || e.key === " ") {
        e.preventDefault();
        openPhotoModal(0);
      }
    });
  }

  if (closePhotoModalBtn) {
    closePhotoModalBtn.addEventListener("click", closePhotoModal);
  }

  if (galleryPrevBtn) {
    galleryPrevBtn.addEventListener("click", (e) => {
      e.stopPropagation();
      prevPhoto();
    });
  }

  if (galleryNextBtn) {
    galleryNextBtn.addEventListener("click", (e) => {
      e.stopPropagation();
      nextPhoto();
    });
  }

  if (schoolPhotoModal) {
    schoolPhotoModal.addEventListener("click", (e) => {
      if (e.target === schoolPhotoModal) {
        closePhotoModal();
      }
    });
  }

  // Управление с клавиатуры
  document.addEventListener("keydown", (e) => {
    if (!schoolPhotoModal || !schoolPhotoModal.classList.contains("show")) return;
    if (e.key === "Escape") {
      closePhotoModal();
    } else if (e.key === "ArrowLeft") {
      prevPhoto();
    } else if (e.key === "ArrowRight") {
      nextPhoto();
    }
  });

  // Поддержка жестов свайпа на сенсорных экранах
  let touchStartX = 0;
  let touchEndX = 0;
  if (modalSchoolImg) {
    modalSchoolImg.addEventListener(
      "touchstart",
      (e) => {
        touchStartX = e.changedTouches[0].screenX;
      },
      { passive: true }
    );
    modalSchoolImg.addEventListener(
      "touchend",
      (e) => {
        touchEndX = e.changedTouches[0].screenX;
        if (touchEndX < touchStartX - 40) {
          nextPhoto();
        } else if (touchEndX > touchStartX + 40) {
          prevPhoto();
        }
      },
      { passive: true }
    );
  }
});
