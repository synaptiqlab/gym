/**
 * =========================================================================
 * 🏋️ GYM COACH API — Синхронизация тренировок с вкладкой "New GYM2"
 * 
 * Логика плана и автосохранения:
 * 1. Сохранение КАЖДОГО подхода на лету (action: "save_set"):
 *    - При каждом клике на галочку (✔) дата записывается в столбец,
 *      а значение подхода сохраняется в соответствующую ячейку (с мягким зелёным фоном).
 * 
 * 2. Умное распознавание сессии:
 *    - Если в последнем столбце дата совпадает с СЕГОДНЯ:
 *      -> Это ТЕКУЩАЯ ТРЕНИРОВКА! Приложение не создаёт новый столбец и не
 *         добавляет лишние +2 повтора. Клиент продолжает тренировку в том же столбце.
 *    - Если даты нет:
 *      -> Это ПЛАН ТРЕНЕРА! Берём значения тренера.
 *    - Если дата из прошлого дня и нового столбца нет:
 *      -> Включаем АВТО-ПРОГРЕССИЮ (+2 повтора к двум последним подходам)
 *         в следующем новом столбце.
 * =========================================================================
 */

const CONFIG = {
  SHEET_NAME: "New GYM2",        // Имя вкладки
  HEADER_COLOR: "#2ea043",      // Зеленый цвет ячейки даты завершенной тренировки
  HEADER_TEXT_COLOR: "#ffffff",
  COMPLETED_CELL_BG: "#e6ffed"  // Мягкий зеленый фон выполненного подхода
};

function doGet(e) {
  try {
    const sheet = getWorkSheet();
    const blockNum = (e && e.parameter && e.parameter.block) ? parseInt(e.parameter.block) : 1;
    const data = parseWorkoutBlock(sheet, blockNum);
    return createJsonResponse({ success: true, sheetName: sheet.getName(), block: blockNum, data: data });
  } catch (err) {
    return createJsonResponse({ success: false, error: err.toString() });
  }
}

function doPost(e) {
  try {
    const sheet = getWorkSheet();
    let payload = e && e.postData && e.postData.contents ? JSON.parse(e.postData.contents) : (e.parameter || {});
    
    // Сохранение отдельного подхода при нажатии галочки (Auto-Save on check)
    if (payload.action === "save_set") {
      const singleResult = saveSingleSet(sheet, payload);
      return createJsonResponse({ success: true, message: "Подход сохранен", result: singleResult });
    }

    // Сохранение всей тренировки при нажатии кнопки "Завершить тренировку"
    const saveResult = saveWorkoutToSheet(sheet, payload);
    return createJsonResponse({ success: true, message: "Тренировка успешно сохранена!", result: saveResult });
  } catch (err) {
    return createJsonResponse({ success: false, error: err.toString() });
  }
}

function createJsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}

function getWorkSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (sheet) return sheet;

  const sheets = ss.getSheets();
  for (let i = 0; i < sheets.length; i++) {
    const clean = sheets[i].getName().replace(/\s+/g, '').toLowerCase();
    if (clean === "newgym2") return sheets[i];
  }
  for (let i = 0; i < sheets.length; i++) {
    const clean = sheets[i].getName().replace(/\s+/g, '').toLowerCase();
    if (clean === "newgym") return sheets[i];
  }
  return sheets[0];
}

function parseWorkoutBlock(sheet, blockNum) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  const range = sheet.getRange(1, 1, lastRow, lastCol);
  const values = range.getValues();
  const backgrounds = range.getBackgrounds();

  let block1Row = -1, block2Row = -1;
  for (let r = 0; r < values.length; r++) {
    const cellVal = String(values[r][0] || "").trim().toLowerCase();
    if (cellVal.indexOf("блок 1") !== -1 || cellVal.indexOf("блок1") !== -1) block1Row = r;
    else if (cellVal.indexOf("блок 2") !== -1 || cellVal.indexOf("блок2") !== -1) block2Row = r;
  }
  if (block1Row === -1) block1Row = 0;

  const startRow = (blockNum === 2 && block2Row !== -1) ? block2Row : block1Row;
  const endRow = (blockNum === 1 && block2Row !== -1) ? block2Row : values.length;

  // Ищем строку с надписью "Дата" в столбце A
  let dateRowIdx = -1;
  for (let r = startRow; r < Math.min(startRow + 5, endRow); r++) {
    const cellA = String(values[r][0] || "").trim().toLowerCase();
    if (cellA.indexOf("дата") !== -1) {
      dateRowIdx = r;
      break;
    }
  }
  if (dateRowIdx === -1) dateRowIdx = startRow + 1;

  // Находим все столбцы с данными (столбцы с индекса 1 / col B и далее)
  const colsWithData = [];
  for (let c = 1; c < lastCol; c++) {
    let hasData = false;
    for (let r = dateRowIdx; r < endRow; r++) {
      if (String(values[r][c] || "").trim() !== "") {
        hasData = true;
        break;
      }
    }
    if (hasData) colsWithData.push(c);
  }

  if (colsWithData.length === 0) {
    colsWithData.push(1);
  }

  const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "GMT+3", "dd.MM.yyyy");
  const lastDataCol = colsWithData[colsWithData.length - 1];
  const lastColDateRaw = (dateRowIdx >= 0) ? values[dateRowIdx][lastDataCol] : "";
  const lastColDateStr = formatDateValue(lastColDateRaw);
  const isSameAsToday = isSameDate(lastColDateRaw, todayStr);

  let prevColIdx = -1;
  let targetColIdx = -1;
  let isAutoGenerated = false;
  let isTodaySession = false;

  if (isSameAsToday) {
    // 🟢 СЛУЧАЙ 1: Сегодняшняя тренировка уже начата или сохранена сегодня!
    // Остаёмся в этом же столбце, не создаём новый столбец и не добавляем лишние +2!
    isTodaySession = true;
    targetColIdx = lastDataCol;
    prevColIdx = (colsWithData.length >= 2) ? colsWithData[colsWithData.length - 2] : lastDataCol;
    isAutoGenerated = false;
  } else if (!lastColDateStr && colsWithData.length >= 2) {
    // 🎯 СЛУЧАЙ 2: В последнем столбце дата НЕ стоит!
    // Это значит, что тренер расписал новый план!
    isTodaySession = false;
    isAutoGenerated = false;
    targetColIdx = lastDataCol;                           // План тренера
    prevColIdx = colsWithData[colsWithData.length - 2];  // Прошлая тренировка с датой
  } else {
    // ⚡ СЛУЧАЙ 3: Прошлая тренировка была в другой день, а нового столбца от тренера нет
    // Включаем авто-прогрессию (+2 повтора к 2 последним подходам)
    isTodaySession = false;
    isAutoGenerated = true;
    prevColIdx = lastDataCol;                            // Прошлая тренировка
    targetColIdx = lastDataCol + 1;                      // Новая тренировка в следующем столбце!
  }

  const prevDateVal = (prevColIdx >= 0 && dateRowIdx >= 0) ? values[dateRowIdx][prevColIdx] : "";
  const prevDateStr = formatDateValue(prevDateVal) || "Прошлая";

  // Парсим упражнения
  const exercises = [];
  let currentExercise = null;

  for (let r = dateRowIdx + 1; r < endRow; r++) {
    const rawName = String(values[r][0] || "").trim();

    if (rawName && !rawName.toLowerCase().startsWith("блок") && !rawName.toLowerCase().startsWith("дата")) {
      if (currentExercise && currentExercise.sets.length > 0) exercises.push(currentExercise);
      currentExercise = { name: rawName.replace(/\n/g, " "), sets: [] };
    }

    if (currentExercise) {
      const prevVal = prevColIdx >= 0 ? String(values[r][prevColIdx] || "").trim() : "";
      let targetVal = (targetColIdx < lastCol) ? String(values[r][targetColIdx] || "").trim() : "";

      if (prevVal || targetVal) {
        let isDone = false;
        if (isTodaySession && targetVal !== "") {
          isDone = true;
        }

        currentExercise.sets.push({
          row: r + 1,
          setNum: currentExercise.sets.length + 1,
          prev: prevVal || targetVal || "—",
          target: targetVal || prevVal || "",
          done: isDone
        });
      }
    }
  }
  if (currentExercise && currentExercise.sets.length > 0) exercises.push(currentExercise);

  // Если сработала автогенерация (+2 повтора в 2 последних подходах)
  if (isAutoGenerated) {
    exercises.forEach(ex => {
      const total = ex.sets.length;
      ex.sets.forEach((set, idx) => {
        if (idx >= total - 2) {
          set.target = addTwoReps(set.prev);
        } else {
          set.target = set.prev;
        }
      });
    });
  } else if (isTodaySession) {
    // Если сегодняшняя тренировка была начата без отдельного плана тренера (автоматически)
    // то для еще не выполненных подходов подставляем авто-план
    exercises.forEach(ex => {
      const total = ex.sets.length;
      ex.sets.forEach((set, idx) => {
        if (!set.done && (!set.target || set.target === set.prev)) {
          if (idx >= total - 2) {
            set.target = addTwoReps(set.prev);
          } else {
            set.target = set.prev;
          }
        }
      });
    });
  }

  return {
    block: blockNum,
    dateRow: dateRowIdx + 1,
    prevCol: prevColIdx + 1,
    targetCol: targetColIdx + 1,
    prevDate: prevDateStr,
    targetDate: todayStr,
    isAutoGenerated: isAutoGenerated,
    isTodaySession: isTodaySession,
    exercises: exercises
  };
}

/**
 * Логика прибавления +2 повторений:
 * 15х12 -> 15х14
 * 20x10 -> 20x12
 * 35хМАХ -> 35хМАХ
 */
function addTwoReps(val) {
  if (!val) return "";
  val = String(val).trim();
  const regex = /^(.*?[xхXХ]\s*)(\d+)(.*)$/;
  const match = val.match(regex);
  if (match) {
    const prefix = match[1];
    const reps = parseInt(match[2], 10);
    const suffix = match[3];
    return prefix + (reps + 2) + suffix;
  }
  return val;
}

function formatDateValue(val) {
  if (!val) return "";
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone() || "GMT+3", "dd.MM.yyyy");
  }
  return String(val).trim().replace(/\[.*?\]/g, "").trim();
}

function isSameDate(dateVal, todayStr) {
  if (!dateVal || !todayStr) return false;
  const s = formatDateValue(dateVal);
  if (!s) return false;
  if (s === todayStr) return true;
  const p1 = s.split(".");
  const p2 = todayStr.split(".");
  if (p1.length >= 2 && p2.length >= 2) {
    if (parseInt(p1[0], 10) === parseInt(p2[0], 10) && parseInt(p1[1], 10) === parseInt(p2[1], 10)) {
      if (p1.length === 2) return true;
      const y1 = p1[2].length === 2 ? "20" + p1[2] : p1[2];
      const y2 = p2[2].length === 2 ? "20" + p2[2] : p2[2];
      return y1 === y2;
    }
  }
  return false;
}

/**
 * Сохранение одного подхода при клике по галочке (Auto-Save on check)
 */
function saveSingleSet(sheet, payload) {
  let targetCol = parseInt(payload.targetCol || 0);
  const dateRow = parseInt(payload.dateRow || 2);
  const dateStr = payload.date || Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "GMT+3", "dd.MM.yyyy");
  const row = parseInt(payload.row || 0);
  const val = String(payload.value || "").trim();
  const completed = Boolean(payload.completed);

  if (targetCol <= 0) targetCol = sheet.getLastColumn() + 1;
  while (targetCol > sheet.getMaxColumns()) {
    sheet.insertColumnAfter(sheet.getMaxColumns());
  }

  // 1. Убеждаемся, что в строке даты стоит сегодняшняя дата
  const dateCell = sheet.getRange(dateRow, targetCol);
  const curDateVal = String(dateCell.getValue() || "").trim();
  if (!curDateVal) {
    dateCell.setValue(dateStr);
    dateCell.setBackground(CONFIG.HEADER_COLOR);
    dateCell.setFontColor(CONFIG.HEADER_TEXT_COLOR);
    dateCell.setFontWeight("bold");
    dateCell.setHorizontalAlignment("center");
  }

  // 2. Записываем подход
  if (row > 0) {
    const cell = sheet.getRange(row, targetCol);
    if (completed) {
      cell.setValue(val);
      cell.setHorizontalAlignment("center");
      cell.setBackground(CONFIG.COMPLETED_CELL_BG);
    } else {
      cell.setBackground("#ffffff");
      cell.clearContent();
    }
  }

  return { success: true, saved: true, row: row, column: targetCol, value: val, completed: completed };
}

/**
 * Сохранение всей тренировки
 */
function saveWorkoutToSheet(sheet, payload) {
  let targetCol = parseInt(payload.targetCol || 0);
  const dateRow = parseInt(payload.dateRow || 2);
  const dateStr = payload.date || Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "GMT+3", "dd.MM.yyyy");
  const results = payload.results || [];

  if (targetCol <= 0) targetCol = sheet.getLastColumn() + 1;
  while (targetCol > sheet.getMaxColumns()) {
    sheet.insertColumnAfter(sheet.getMaxColumns());
  }

  // Проставляем дату тренировки в строку даты целевого столбца с зелёным бейджем
  const dateCell = sheet.getRange(dateRow, targetCol);
  dateCell.setValue(dateStr);
  dateCell.setBackground(CONFIG.HEADER_COLOR);
  dateCell.setFontColor(CONFIG.HEADER_TEXT_COLOR);
  dateCell.setFontWeight("bold");
  dateCell.setHorizontalAlignment("center");

  let savedCount = 0;
  results.forEach(item => {
    if (item.row && item.value && item.completed) {
      const cell = sheet.getRange(item.row, targetCol);
      cell.setValue(item.value);
      cell.setHorizontalAlignment("center");
      cell.setBackground(CONFIG.COMPLETED_CELL_BG);
      savedCount++;
    }
  });

  return { savedSets: savedCount, column: targetCol, dateRow: dateRow, date: dateStr };
}
