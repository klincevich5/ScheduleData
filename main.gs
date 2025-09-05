/*******************************
 * Google Apps Script
 * Build stacked DAY/NIGHT sheets from multiple source spreadsheets
 * - Copies D1:ACx from each source sheet and stacks blocks one under another
 * - VIP is cut strictly at "shift: 4" (fallback → first "shift:")
 * - Logs to Logger and (optionally) to a "_build_log" sheet
 * - Colors cells in D:AC according to COLOR_MAP (exact match by text, case-insensitive)
 * - CLEANUP: keep exactly one "Dealer Name" row, move it to the very top;
 *   delete other keyword rows; then sort the remaining rows A→Z by col D
 * - After cleanup, per dealer (col D) merges duplicate rows across sources by per-slot priority (Tier1>Tier2>Tier3)
 * - After merge, per dealer computes worked hours from E:AB and writes A,B,C = date, start, end
 * - NEW: Move each finished sheet to its target tab index immediately:
 *        DAY 1 → index 2, NIGHT 1 → index 3, then DAY 2 → 4, NIGHT 2 → 5, etc.
 * - NEW: All inserted text is normalized (whitespace cleanup + homoglyphs → ASCII Latin)
 *
 * Interval rules (final):
 * • Non-overlapping → separate intervals
 * • Touching borders → merge into one
 * • Overlap/nesting → take intersection → one interval
 * • Only LATE/HOME break an interval; empty cells = activity (on shift)
 *******************************/

/** === CONFIG (edit here) === */
const TARGET_DOC_ID = '1WFKvhq3FMvKsUaTX5_jMvmB4gwLkJIQzuK3YEqaH5XA';
const TEMPLATE_SHEET_NAME = 'template';

// Период для дат в колонке A
const PERIOD_YEAR  = 2025;  // ← укажи нужный год
const PERIOD_MONTH = 8;     // ← укажи нужный месяц (1..12)

// Временные интервалы смен (используются только для базового времени; фактические B/C считаются по слотам)
const DAY_START   = '09:00';
const DAY_END     = '21:00';
const NIGHT_START = '21:00';
const NIGHT_END   = '09:00';

const TZ = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
выд
const SOURCES = [
  // { key: 'shuffle',    id: '1v0bUnhnYi7TKTzFE4YZJoRwIvGBlm388fmyTIUqnzn8' },
  // { key: 'VIP',        id: '1eJ_sygy-UcHJiM1p9xj3uiht49I_DkP3dCMryGQzJR8' },
  { key: 'GENERIC',    id: '17jKk3zKZzNEGUI4qE8Y0zKnNBGEBSueD8602VZJcV0M' },
  // { key: 'GSBJ',       id: '1qNFOKSmBpxR--QCZpaqdLG0VVGncrVgQxEtBajjqJaA' },
  // { key: 'LEGENDZ',    id: '1Wr-Re89hHeOiDyue5M7Z_b0OqhSDUDQB5mNcUUxqsZA' },
  // { key: 'Game Shows', id: '1Uh28efck0YVR2bGtTr0Yf-1VXd_T1dQUq-TrcAmLNEU' },
];

// Диапазон источника и позиция вставки в целевом листе
const SRC_START_COL = 4;   // D
const SRC_END_COL   = 29;  // AC
const DST_START_COL = 4;   // вставка также с колонки D

// Интервалы/фильтры обработки
const DAY_RANGE = {from: 1, to: 1 };
const INCLUDE_DAY   = true;
const INCLUDE_NIGHT = /*true*/ false;

// Вставлять пустую строку между блоками разных источников?
const INSERT_EMPTY_ROW_BETWEEN_BLOCKS = false;

// === ЛОГИРОВАНИЕ ===
const LOG_TO_SHEET = true;
const LOG_SHEET_NAME = '_build_log';
const LOG_FLUSH_BATCH = 100;
const LOG_DATE_FMT_LOCALE = 'en-GB';

// === ОКРАСКА ===
const ENABLE_COLORING = true;
const COLOR_MAP = {
  "x": {"fg": "#000000", "bg": "#00ff00"},
  "X": {"fg": "#000000", "bg": "#00ff00"},
  "х": {"fg": "#000000", "bg": "#00ff00"},
  "Х": {"fg": "#000000", "bg": "#00ff00"},
  "SC": {"fg": "#000000", "bg": "#00ffff"},
  "TC": {"fg": "#000000", "bg": "#8179c7"},
  "FC": {"fg": "#000000", "bg": "#e6cd74"},
  "HOME": {"fg": "#000000", "bg": "#00ffff"},
  "FLOOR": {"fg": "#ffffff", "bg": "#11499e"},
  "LATE": {"fg": "#000000", "bg": "#ff0000"},
  "Shuffle": {"fg": "#000000", "bg": "#8e7cc3"},
  "Game Shows": {"fg": "#ffff00", "bg": "#073763"},
  "TritonRL": {"fg": "#bfe0f6", "bg": "#0a53a8"},
  "RRR": {"fg": "#000000", "bg": "#674ea7"},
  "TRISTAR": {"fg": "#b77a30", "bg": "#434343"},
  "AB": {"fg": "#d4edbc", "bg": "#11734b"},
  "L7": {"fg": "#ffffff", "bg": "#2b9de8"},
  "DT": {"fg": "#bfe0f6", "bg": "#0a53a8"},
  "TP": {"fg": "#ffcfc9", "bg": "#b10202"},
  "DTL": {"fg": "#e5cff2", "bg": "#5a3286"},
  "VIP": {"fg": "#000000", "bg": "#ffff00"},
  "vBJ2": {"fg": "#000000", "bg": "#e6cd74"},
  "vBJ3": {"fg": "#000000", "bg": "#21cbab"},
  "gBC1": {"fg": "#000000", "bg": "#d5a6bd"},
  "vBC3": {"fg": "#000000", "bg": "#a160f3"},
  "vBC4": {"fg": "#000000", "bg": "#e06666"},
  "vHSB1": {"fg": "#000000", "bg": "#ff50e8"},
  "vDT1": {"fg": "#000000", "bg": "#e91a1a"},
  "gsRL1": {"fg": "#e5cff2", "bg": "#5a3286"},
  "swBC1": {"fg": "#ffffff", "bg": "#11734b"},
  "swRL1": {"fg": "#000000", "bg": "#ffff00"},
  "chBJ1": {"fg": "#b10202", "bg": "#ffcfc9"},
  "chBJ2": {"fg": "#473821", "bg": "#ffe5a0"},
  "chBJ3": {"fg": "#0a53a8", "bg": "#bfe1f6"},
  "TURKISH": {"fg": "#000000", "bg": "#f11d52"},
  "tBJ1": {"fg": "#ffffff", "bg": "#6633cc"},
  "tBJ2": {"fg": "#000000", "bg": "#3d86f8"},
  "tRL1": {"fg": "#000000", "bg": "#f11d52"},
  "GENERIC": {"fg": "#000000", "bg": "#ff9900"},
  "gBJ1": {"fg": "#000000", "bg": "#00ffff"},
  "gBJ3": {"fg": "#000000", "bg": "#ffe599"},
  "gBJ4": {"fg": "#000000", "bg": "#a64d79"},
  "gBJ5": {"fg": "#000000", "bg": "#cc0000"},
  "gBC2": {"fg": "#000000", "bg": "#fbbc04"},
  "gBC3": {"fg": "#000000", "bg": "#3c78d8"},
  "gBC4": {"fg": "#000000", "bg": "#e69138"},
  "gBC5": {"fg": "#000000", "bg": "#ffff00"},
  "gBC6": {"fg": "#000000", "bg": "#6aa84f"},
  "gRL1": {"fg": "#000000", "bg": "#ff6d01"},
  "GSBJ": {"fg": "#000000", "bg": "#a64d79"},
  "LEGENDZ": {"fg": "#000000", "bg": "#34a853"},
  "OVER": {"fg": "#000000", "bg": "#ff6d01"}
};

// === TIER PRIORITY (for per-slot merge across rows of the same dealer) ===
const TIER1_NON_WORK_SET = new Set(['late', 'home']); // highest priority, also break intervals
const TIER3_FLOOR_SET = new Set(['gsh','generic','gsbj','legendz','floor','vip']); // lowest priority

/** === ENTRY POINT === */
function buildRotations() {
  const startedAt = Date.now();
  const runId = makeRunId_();

  const targetSS = SpreadsheetApp.openById(TARGET_DOC_ID);
  const template = targetSS.getSheetByName(TEMPLATE_SHEET_NAME);
  if (!template) throw new Error(`Нет шаблона "${TEMPLATE_SHEET_NAME}" в целевом документе.`);

  if (LOG_TO_SHEET) ensureLogSheetExists_(targetSS, LOG_SHEET_NAME);
  const logBuf = new LogBuffer_(targetSS, LOG_TO_SHEET ? LOG_SHEET_NAME : null, runId);

  // зафиксируем служебные вкладки на первых позициях
  positionFixedTabs_(targetSS, logBuf);

  try {
    logBuf.info('INIT', `Starting; template="${TEMPLATE_SHEET_NAME}"`);

    // Список дней просто из диапазона
    const days = [];
    for (let d = DAY_RANGE.from; d <= DAY_RANGE.to; d++) days.push(d);
    logBuf.info('INIT', `Days: [${days.join(', ')}]; include DAY=${INCLUDE_DAY}, NIGHT=${INCLUDE_NIGHT}`);

    const prefixes = [];
    if (INCLUDE_DAY)   prefixes.push('DAY');
    if (INCLUDE_NIGHT) prefixes.push('NIGHT');

    const colorMapNorm = buildNormalizedColorMap_(COLOR_MAP);

    let createdSheets = 0, copiedBlocks = 0, skippedNoData = 0, coloredCellsTotal = 0, cleanedRowsTotal = 0, mergedDealersTotal = 0;

    for (const n of days) {
      for (const prefix of prefixes) {
        const sheetName = `${prefix} ${n}`;
        const t0 = Date.now();

        logBuf.info('SHEET_START', `Preparing "${sheetName}"`);
        const targetSheet = ensureFreshSheetFromTemplate_(targetSS, template, sheetName, logBuf);
        createdSheets += 1;

        // Сборка блоков из всех источников (с нормализацией текста)
        const res = stackBlocksForSheetName_NoMissingChecks_(sheetName, targetSheet, logBuf);
        copiedBlocks += res.copiedBlocks;
        skippedNoData += res.skippedNoData;

        if (res.lastDataRow > 0) {
          // Очистка/сортировка с переносом "Dealer Name" в шапку
          const cleanedRows = cleanupAndSortWithHeader_(targetSheet, res.lastDataRow, logBuf);
          cleanedRowsTotal += cleanedRows;

          // === NEW: СКЛЕЙКА дублей дилера по приоритетам Tier1 > Tier2 > Tier3 ===
          const mergedCount = mergeDuplicateDealerRows_(targetSheet, logBuf);
          mergedDealersTotal += mergedCount;

          // Аннотация A/B/C и расчёт фактических времён + разбиение интервалов
          annotateDateTimeAndComputeHours_(targetSheet, sheetName, logBuf);
          syncWithScheduleData_(targetSheet, 'ScheduleData', logBuf);

          // Покраска D:AC
          const finalLastRow = targetSheet.getLastRow();
          if (ENABLE_COLORING && finalLastRow > 0) {
            const colored = applyColorMapToArea_(targetSheet, 1, DST_START_COL, finalLastRow, SRC_END_COL - SRC_START_COL + 1, colorMapNorm, logBuf);
            coloredCellsTotal += colored;
            logBuf.info('COLORING', `"${sheetName}": colored_cells=${colored} in D1:AC${finalLastRow}`);
          }
        }

        // === NEW: сразу ставим вкладку на нужный индекс ===
        const desiredIndex = desiredTabIndexFor_(n, prefix);
        moveSheetToIndex_(targetSS, sheetName, desiredIndex, logBuf);

        logBuf.info('SHEET_DONE', `"${sheetName}" done in ${formatDuration_(Date.now() - t0)}; copied=${res.copiedBlocks}`);
      }
    }

    logBuf.info(
      'SUMMARY',
      `Completed in ${formatDuration_(Date.now() - startedAt)}; ` +
      `sheets_created=${createdSheets}, blocks_copied=${copiedBlocks}, skipped_no_data=${skippedNoData}, ` +
      `cleaned_rows=${cleanedRowsTotal}, merged_dealers=${mergedDealersTotal}, colored_cells=${coloredCellsTotal}`
    );
  } catch (e) {
    const msg = e && e.stack ? e.stack : String(e);
    logBuf.error('FATAL', `Unhandled error: ${msg}`);
    throw e;
  } finally {
    logBuf.flush(true);
  }
}



/** === CLEANUP: keep one "Dealer Name" on top; delete others; sort by D (rows 2..n) === */
function cleanupAndSortWithHeader_(sheet, lastDataRow, logBuf) {
  if (!lastDataRow || lastDataRow < 1) return 0;

  const colDVals = sheet.getRange(1, 4, lastDataRow, 1).getDisplayValues().map(r => (r[0] || '').toString());
  const dealerNameRows = [];
  for (let i = 0; i < colDVals.length; i++) {
    if (normalizeText_(colDVals[i]).toLowerCase() === 'dealer name') dealerNameRows.push(i + 1);
  }
  const keepHeaderRow = dealerNameRows.length ? dealerNameRows[0] : -1;

  // Ключевые слова на удаление в D (кроме одного «Dealer Name», который оставляем)
  const keywords = ['Replacements', 'shift', 'Game Shows', 'GENERIC', 'GSBJ', 'LEGENDZ', 'Shuffle', 'VIP', 'Dealer Name']
    .map(s => s.toLowerCase());

  const rowsToDelete = [];
  for (let i = 0; i < colDVals.length; i++) {
    const rowIndex = i + 1;
    const v = normalizeText_(colDVals[i]).toLowerCase();
    const isDealerName = (v === 'dealer name');
    const matchesKeyword = keywords.some(k => v.indexOf(k) !== -1);
    if (matchesKeyword) {
      if (isDealerName && rowIndex === keepHeaderRow) {
        // keep
      } else {
        rowsToDelete.push(rowIndex);
      }
    }
  }

  // Удаляем снизу вверх
  for (let i = rowsToDelete.length - 1; i >= 0; i--) {
    try { sheet.deleteRow(rowsToDelete[i]); }
    catch (e) { logBuf.error('ROW_DELETE_FAIL', `Failed to delete row ${rowsToDelete[i]}: ${e}`); }
  }
  let deleted = rowsToDelete.length;

  // Перемещаем единственную «Dealer Name» в строку 1 (если есть)
  const newLastRow = sheet.getLastRow();
  if (keepHeaderRow !== -1 && newLastRow >= 1) {
    const dVals = sheet.getRange(1, 4, newLastRow, 1).getDisplayValues().map(r => normalizeText_(r[0] || '').toLowerCase());
    let currentHeaderRow = -1;
    for (let i = 0; i < dVals.length; i++) {
      if (dVals[i] === 'dealer name') { currentHeaderRow = i + 1; break; }
    }
    if (currentHeaderRow > 1) {
      const lastCol = Math.max(sheet.getLastColumn(), SRC_END_COL);
      const srcRng = sheet.getRange(currentHeaderRow, 1, 1, lastCol);
      const vals = srcRng.getDisplayValues();
      const fmts = srcRng.getNumberFormats();
      const bgs  = srcRng.getBackgrounds();
      const fcs  = srcRng.getFontColors();

      const normVals = normalizeValues2D_(vals);

      sheet.insertRowBefore(1);
      const dstRng = sheet.getRange(1, 1, 1, lastCol);
      dstRng.setValues(normVals);
      dstRng.setNumberFormats(fmts);
      dstRng.setBackgrounds(bgs);
      dstRng.setFontColors(fcs);
      sheet.deleteRow(currentHeaderRow + 1);
    } else if (currentHeaderRow === -1) {
      sheet.insertRowBefore(1);
      sheet.getRange(1, 4).setValue(normalizeText_('Dealer Name'));
    }
  } else {
    sheet.insertRowBefore(1);
    sheet.getRange(1, 4).setValue(normalizeText_('Dealer Name'));
  }

  // Сортировка со 2-й строки по D (A→Z)
  const finalLastRow = sheet.getLastRow();
  if (finalLastRow > 2) {
    try {
      sheet.getRange(2, 1, finalLastRow - 1, sheet.getLastColumn())
           .sort([{column: 4, ascending: true}]);
      logBuf.info('SORT', `"${sheet.getName()}": sorted rows 2..${finalLastRow} by col D (A→Z)`);
    } catch (e) {
      logBuf.error('SORT_FAIL', `Failed sorting rows 2..${finalLastRow} by col D: ${e}`);
    }
  }

  logBuf.info('CLEANUP', `"${sheet.getName()}": deleted ${deleted} rows; kept single "Dealer Name" on top`);
  return deleted;
}

/** === MERGE DUPLICATE DEALERS BY PER-SLOT PRIORITY (Tier1 > Tier2 > Tier3) ===
 * Returns number of merged dealer groups (i.e., dealers that had >=2 rows).
 */
function mergeDuplicateDealerRows_(sheet, logBuf) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;

  const COL_D = 4;
  const SLOT_START_COL = 5; // E
  const SLOT_COUNT = 24;
  const slotLen = Math.min(SLOT_COUNT, sheet.getMaxColumns() - SLOT_START_COL + 1);

  const names = sheet.getRange(2, COL_D, lastRow - 1, 1).getDisplayValues().map(r => normalizeText_(r[0] || ''));
  const slots = sheet.getRange(2, SLOT_START_COL, lastRow - 1, slotLen).getDisplayValues()
                     .map(row => row.map(v => normalizeText_(v)));

  // Сгруппируем по дилеру
  const map = new Map(); // dealer -> { rows: [idxs], slots: [ [slotVals], ... ] }
  for (let i = 0; i < names.length; i++) {
    const dealer = names[i];
    if (!dealer || /^dealer\s*name$/i.test(dealer)) continue;
    if (!map.has(dealer)) map.set(dealer, { rows: [], slots: [] });
    map.get(dealer).rows.push(i + 2);       // фактический номер строки
    map.get(dealer).slots.push(slots[i]);   // массив слотов
  }

  let mergedDealerGroups = 0;
  const dealers = Array.from(map.keys()).sort((a,b)=>a.localeCompare(b));

  // Сформируем новые строки (только D и E:AB); A/B/C позже заполним в annotate
  const newRows = [];
  for (const dealer of dealers) {
    const group = map.get(dealer);
    const rowsCount = group.rows.length;

    if (rowsCount > 1) mergedDealerGroups += 1;

    // Пер-слотная редукция по приоритетам
    const mergedSlots = new Array(slotLen);
    for (let s = 0; s < slotLen; s++) {
      const candidates = [];
      for (let k = 0; k < group.slots.length; k++) {
        candidates.push(group.slots[k][s] || '');
      }
      mergedSlots[s] = chooseByPriority_(candidates);
    }

    // Конечная строка: D=dealer, E:AB=mergedSlots, A..C пустые (будут заполнены позже)
    const row = new Array(SLOT_START_COL + slotLen - 1).fill('');
    row[3] = dealer; // D
    for (let c = 0; c < slotLen; c++) row[SLOT_START_COL - 1 + c] = mergedSlots[c];
    newRows.push(row);
  }

  // Перезапишем таблицу: сохраним шапку (строка 1) и заменим всё с 2-й строки
  // Удалим хвост, если старых строк больше
  const writeRows = Math.max(1, newRows.length);
  const lastColToWrite = Math.max(sheet.getLastColumn(), SRC_END_COL);

  // Очистим блок со 2-й строки
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, lastColToWrite).clearContent();
  }

  if (writeRows > 0) {
    sheet.getRange(2, 1, writeRows, lastColToWrite).clearContent(); // чтобы не потянуть мусор
    sheet.getRange(2, 1, writeRows, newRows[0].length).setValues(newRows);
  }

  // Если остались старые «лишние» строки — удалим их
  const nowLastRow = sheet.getLastRow();
  const expectedLastRow = 1 + writeRows;
  if (nowLastRow > expectedLastRow) {
    try { sheet.deleteRows(expectedLastRow + 1, nowLastRow - expectedLastRow); } catch (_) {}
  }

  logBuf.info('MERGE', `Merged duplicate dealer rows: groups=${mergedDealerGroups}; dealers_total=${dealers.length}`);
  return mergedDealerGroups;
}

/** Выбор значения слота по приоритетам.
 * Tier1: 'late'/'home' → если есть — берём (возвращаем верхним регистром).
 * Tier2: всё остальное (вкл. пустые). Предпочтение НЕНУЛЕВОМУ токену; если все пустые — ''.
 * Tier3: если нет Tier2 вообще (что маловероятно), берём любой из floor-ярлыков.
 */
function chooseByPriority_(candidates) {
  let tier1 = null;
  let tier2NonEmpty = null;
  let tier2Empty = false;
  let tier3 = null;

  for (const raw of candidates) {
    const v = normalizeText_(raw || '');
    const low = v.toLowerCase();
    if (TIER1_NON_WORK_SET.has(low)) {
      tier1 = low; // 'late' or 'home'
      break;       // можно сразу завершить
    }
    if (v === '') {
      tier2Empty = true;
      continue;
    }
    if (TIER3_FLOOR_SET.has(low)) {
      if (!tier3) tier3 = v; // помечаем, но это самый низкий приоритет
      continue;
    }
    // Всё, что не Tier1 и не Tier3, — Tier2 (столы, брейки/коды и т.п.)
    if (!tier2NonEmpty) tier2NonEmpty = v;
  }

  if (tier1) return tier1.toUpperCase();          // LATE/HOME
  if (tier2NonEmpty) return tier2NonEmpty;        // любой ненулевой Tier2
  if (tier2Empty) return '';                      // пустая как активность
  return tier3 || '';                              // если остался только флор — берём его
}

/** === Annotate A/B/C with ACTUAL first/last worked times
 *      and SPLIT multi-segment rows into separate rows
 *      Rule (updated): ONLY LATE/HOME break; empty = activity; touching segments merge naturally. === */
function annotateDateTimeAndComputeHours_(sheet, sheetName, logBuf) {
  const m = sheetName.match(/^(DAY|NIGHT)\s+(\d{1,2})$/i);
  if (!m) return;
  const isDay = m[1].toUpperCase() === 'DAY';
  const dayNum = parseInt(m[2], 10);
  if (isNaN(dayNum)) return;

  let lastRow = sheet.getLastRow();
  if (lastRow < 1) return;

  const COL_D = 4;          // Dealer
  const SLOT_START_COL = 5; // E
  const SLOT_COUNT = 24;    // E..AB
  const slotLen = Math.min(SLOT_COUNT, sheet.getMaxColumns() - SLOT_START_COL + 1);

  // Читаем разом D и E:AB
  let dealers = sheet.getRange(1, COL_D, lastRow, 1).getDisplayValues(); // D
  let slots   = sheet.getRange(1, SLOT_START_COL, lastRow, slotLen).getDisplayValues();

  const dateStr = formatDateDMY_(PERIOD_YEAR, PERIOD_MONTH, dayNum);
  const baseStart = isDay
    ? new Date(PERIOD_YEAR, PERIOD_MONTH - 1, dayNum, 9, 0, 0, 0)   // 09:00
    : new Date(PERIOD_YEAR, PERIOD_MONTH - 1, dayNum, 21, 0, 0, 0); // 21:00

  const NON_WORK = new Set(['home', 'late']);

  let r = 0;               // индекс в массивах dealers/slots (0-based)
  let sheetRow = 1;        // фактический номер строки на листе (1-based)
  let annotated = 0;

  while (sheetRow <= lastRow) {
    const dealer = normalizeText_((dealers[r] && dealers[r][0] || ''));
    const isHeader = /^dealer\s*name$/i.test(dealer);

    // Всегда ставим дату A (через нормализацию)
    sheet.getRange(sheetRow, 1).setValue(normalizeText_(dateStr));

    if (!dealer || isHeader) {
      sheetRow += 1; r += 1;
      continue;
    }

    const rowSlots = slots[r] ? slots[r].map(normalizeText_) : new Array(slotLen).fill('');
    const segments = getWorkingSegments_(rowSlots, NON_WORK);

    if (segments.length === 0) {
      sheet.getRange(sheetRow, 2, 1, 2).clearContent();
      logBuf.debug('HOURS', `"${sheetName}" row ${sheetRow} dealer="${dealer}" no working segments`);
      sheetRow += 1; r += 1;
      continue;
    }

    if (segments.length === 1) {
      const [s0, e0] = segments[0];
      const startDt = addMinutes_(baseStart, s0 * 30);
      const endDt   = addMinutes_(baseStart, (e0 + 1) * 30);
      sheet.getRange(sheetRow, 2).setValue(normalizeText_(formatTimeHMM_(startDt))); // B
      sheet.getRange(sheetRow, 3).setValue(normalizeText_(formatTimeHMM_(endDt)));   // C
      annotated += 1;
      sheetRow += 1; r += 1;
      continue;
    }

    // === Несколько сегментов (разрывы только LATE/HOME) ===
    const baseline = buildBaselineOutsideSegments_(rowSlots, segments);

    // 1) Текущую строку превращаем в первый сегмент
    const [s0, e0] = segments[0];
    const firstStart = addMinutes_(baseStart, s0 * 30);
    const firstEnd   = addMinutes_(baseStart, (e0 + 1) * 30);
    sheet.getRange(sheetRow, 2).setValue(normalizeText_(formatTimeHMM_(firstStart))); // B
    sheet.getRange(sheetRow, 3).setValue(normalizeText_(formatTimeHMM_(firstEnd)));   // C
    const row0 = rowWithSegment_(baseline, rowSlots, s0, e0);
    sheet.getRange(sheetRow, SLOT_START_COL, 1, slotLen).setValues([row0]);
    logBuf.info('SPLIT_ROW', `"${sheetName}" row ${sheetRow} dealer="${dealer}" segments=${segments.length} [1/${segments.length}] ${formatTimeHMM_(firstStart)}–${formatTimeHMM_(firstEnd)}`);

    // 2) Остальные сегменты — вставляем строки
    for (let k = 1; k < segments.length; k++) {
      const [sK, eK] = segments[k];
      const startK = addMinutes_(baseStart, sK * 30);
      const endK   = addMinutes_(baseStart, (eK + 1) * 30);

      sheet.insertRowAfter(sheetRow);
      lastRow += 1;
      sheetRow += 1;

      const rowK = rowWithSegment_(baseline, rowSlots, sK, eK);

      const rowValues = new Array(SLOT_START_COL + slotLen - 1).fill('');
      rowValues[0] = normalizeText_(dateStr);                 // A
      rowValues[1] = normalizeText_(formatTimeHMM_(startK));  // B
      rowValues[2] = normalizeText_(formatTimeHMM_(endK));    // C
      rowValues[3] = dealer;                                  // D
      for (let c = 0; c < slotLen; c++) {
        rowValues[SLOT_START_COL - 1 + c] = rowK[c];
      }
      sheet.getRange(sheetRow, 1, 1, rowValues.length).setValues([rowValues]);
    }

    annotated += 1;
    sheetRow += 1; r += 1;
  }

  logBuf.info('ANNOTATE', `"${sheetName}": annotated ${annotated} dealers with split rows under new interval rules.`);
}

/** Возвращает массив сегментов [startIdx, endIdx] по непрерывным рабочим слотам
 * Новые правила: ПУСТО = активность. Только LATE/HOME = разрыв. */
function getWorkingSegments_(rowVals, nonWorkSet) {
  const segs = [];
  let start = -1;

  for (let i = 0; i < rowVals.length; i++) {
    const raw = normalizeText_(rowVals[i] || '');
    const token = raw.toLowerCase();
    const isWork = !nonWorkSet.has(token); // всё, кроме late/home, считается активностью

    if (isWork) {
      if (start === -1) start = i;
    } else {
      if (start !== -1) { segs.push([start, i - 1]); start = -1; }
    }
  }
  if (start !== -1) segs.push([start, rowVals.length - 1]);

  // Сливаем касающиеся/перекрывающиеся (на всякий случай)
  if (segs.length > 1) {
    const merged = [];
    let [cs, ce] = segs[0];
    for (let k = 1; k < segs.length; k++) {
      const [ns, ne] = segs[k];
      if (ns <= ce + 1) ce = Math.max(ce, ne);
      else { merged.push([cs, ce]); [cs, ce] = [ns, ne]; }
    }
    merged.push([cs, ce]);
    return merged;
  }
  return segs;
}

/** Базовая строка: вне объединения всех интервалов — оригинальное значение; внутри — пусто */
function buildBaselineOutsideSegments_(rowVals, segments) {
  const mask = new Array(rowVals.length).fill(false);
  for (const [s, e] of segments) for (let i = s; i <= e; i++) mask[i] = true;
  const out = new Array(rowVals.length);
  for (let i = 0; i < rowVals.length; i++) out[i] = mask[i] ? '' : normalizeText_(rowVals[i]);
  return out;
}

/** Копия baseline + «вклеенный» свой интервал исходными значениями */
function rowWithSegment_(baseline, rowVals, s, e) {
  const out = baseline.slice();
  for (let i = s; i <= e; i++) out[i] = normalizeText_(rowVals[i]);
  return out;
}

/** === CORE STACKING === */
function ensureFreshSheetFromTemplate_(targetSS, templateSheet, newName, logBuf) {
  const existing = targetSS.getSheetByName(newName);
  if (existing) {
    targetSS.deleteSheet(existing);
    logBuf.info('SHEET_RESET', `Deleted old sheet "${newName}"`);
  }
  const newSheet = templateSheet.copyTo(targetSS);
  newSheet.setName(newName);
  clearRangeDAc_(newSheet);
  logBuf.info('SHEET_CREATE', `Created from template: "${newName}" and cleared D:AC`);
  return newSheet;
}

function clearRangeDAc_(sheet) {
  const lastRow = Math.max(sheet.getMaxRows(), sheet.getLastRow(), 1);
  const numRows = Math.max(1, lastRow);
  const numCols = SRC_END_COL - SRC_START_COL + 1; // 26 колонок: D..AC
  sheet.getRange(1, DST_START_COL, numRows, numCols).clearContent();
}

function stackBlocksForSheetName_NoMissingChecks_(sheetName, targetSheet, logBuf) {
  let writeRow = 1, lastDataRow = 0, copiedBlocks = 0, skippedNoData = 0;

  for (const src of SOURCES) {
    const srcSS = SpreadsheetApp.openById(src.id);
    const srcSheet = srcSS.getSheetByName(sheetName);

    let blockRows = 0;
    let detectionReason = '';
    try {
      if (src.key === 'VIP') {
        blockRows = detectVipCutoff_(srcSheet);
        detectionReason = blockRows > 0 ? 'vip_shift:4' : 'vip_fallback_or_none';
      } else {
        blockRows = detectBlockHeightByShiftLabel_(srcSheet);
        detectionReason = blockRows > 0 ? 'first_shift:' : 'no_shift_found';
      }
    } catch (e) {
      logBuf.error('DETECT_FAIL', `Detection failed for "${src.key}" / "${sheetName}": ${e}`);
      continue;
    }

    if (blockRows === 0) {
      skippedNoData += 1;
      logBuf.warn('SRC_SKIP_NO_DATA', `"${src.key}" / "${sheetName}": no rows to copy (${detectionReason})`);
      continue;
    }

    const numCols = SRC_END_COL - SRC_START_COL + 1;
    try {
      const srcRange = srcSheet.getRange(1, SRC_START_COL, blockRows, numCols);
      let values = srcRange.getDisplayValues();

      values = normalizeValues2D_(values);

      const dstRange = targetSheet.getRange(writeRow, DST_START_COL, blockRows, numCols);
      dstRange.setValues(values);

      logBuf.info('COPY_OK', `"${src.key}" / "${sheetName}": copied D1:AC${blockRows} → target @ row ${writeRow} (${detectionReason})`);
      copiedBlocks += 1;
      lastDataRow = Math.max(lastDataRow, writeRow + blockRows - 1);
      writeRow += blockRows;
      if (INSERT_EMPTY_ROW_BETWEEN_BLOCKS) writeRow += 1;
    } catch (e) {
      logBuf.error('WRITE_FAIL', `"${src.key}" / "${sheetName}": cannot write to target: ${e}`);
      continue;
    }
  }

  return { copiedBlocks, skippedNoData, lastDataRow };
}

function detectBlockHeightByShiftLabel_(sheet) {
  const lastRow = Math.max(sheet.getLastRow(), 1);
  const colD = sheet.getRange(1, 4, lastRow, 1).getDisplayValues();
  const re = /shift:/i;
  for (let i = 0; i < colD.length; i++) {
    const val = normalizeText_(colD[i][0] || '');
    if (re.test(val)) return i + 1;
  }
  return 0;
}

function detectVipCutoff_(sheet) {
  const lastRow = Math.max(sheet.getLastRow(), 1);
  const colD = sheet.getRange(1, 4, lastRow, 1).getDisplayValues();
  const reVip = /shift:\s*4\b/i;
  for (let i = 0; i < colD.length; i++) {
    const val = normalizeText_(colD[i][0] || '');
    if (reVip.test(val)) return i + 1;
  }
  return detectBlockHeightByShiftLabel_(sheet);
}

/** === COLORING UTILITIES === */
function buildNormalizedColorMap_(map) {
  const norm = {};
  for (const k in map) {
    if (!Object.prototype.hasOwnProperty.call(map, k)) continue;
    const kk = normalizeText_(String(k)).toLowerCase();
    norm[kk] = { fg: map[k].fg, bg: map[k].bg };
  }
  return norm;
}
function sanitizeHex_(hex) {
  if (!hex) return hex;
  let h = String(hex).trim();
  if (h.startsWith('##')) h = '#' + h.slice(2);
  return h;
}
function applyColorMapToArea_(sheet, startRow, startCol, numRows, numCols, colorMapNorm, logBuf) {
  if (numRows <= 0 || numCols <= 0) return 0;
  const rng = sheet.getRange(startRow, startCol, numRows, numCols);
  const values = rng.getDisplayValues();
  const bg = rng.getBackgrounds();
  const fg = rng.getFontColors();
  let colored = 0;
  for (let r = 0; r < numRows; r++) {
    for (let c = 0; c < numCols; c++) {
      const raw = normalizeText_(values[r][c]);
      if (!raw) continue;
      const key = raw.toLowerCase();
      const rule = colorMapNorm[key];
      if (!rule) continue;
      const newBg = sanitizeHex_(rule.bg);
      const newFg = sanitizeHex_(rule.fg);
      if (newBg) bg[r][c] = newBg;
      if (newFg) fg[r][c] = newFg;
      colored++;
    }
  }
  try { rng.setBackgrounds(bg); rng.setFontColors(fg); }
  catch (e) { logBuf.error('COLOR_APPLY_FAIL', `Failed to set colors for ${sheet.getName()} D${startRow}:AC${startRow+numRows-1}: ${e}`); }
  return colored;
}

/** === LOGGING UTILITIES === */
class LogBuffer_ {
  constructor(targetSS, logSheetNameOrNull, runId) {
    this.targetSS = targetSS;
    this.sheetName = logSheetNameOrNull;
    this.runId = runId;
    this.buffer = [];
  }
  ensureSheet_() {
    if (!this.sheetName) return null;
    let sh = this.targetSS.getSheetByName(this.sheetName);
    if (!sh) {
      sh = this.targetSS.insertSheet(this.sheetName);
      sh.getRange(1, 1, 1, 6).setValues([[ 'timestamp','run_id','level','stage','message','ms_since_epoch' ]]);
    }
    return sh;
  }
  debug(stage, msg) { this._push_('DEBUG', stage, msg); }
  info(stage, msg)  { this._push_('INFO',  stage, msg); }
  warn(stage, msg)  { this._push_('WARN',  stage, msg); }
  error(stage, msg) { this._push_('ERROR', stage, msg); }
  _push_(level, stage, msg) {
    const now = new Date();
    const row = [this.formatTs_(now), this.runId, level, stage, String(msg), now.getTime()];
    Logger.log(`[${level}] [${stage}] ${msg}`);
    if (this.sheetName) {
      this.buffer.push(row);
      if (this.buffer.length >= LOG_FLUSH_BATCH) this.flush(false);
    }
  }
  flush(force) {
    if (!this.sheetName || this.buffer.length === 0) return;
    const sh = this.ensureSheet_();
    const startRow = sh.getLastRow() + 1;
    sh.getRange(startRow, 1, this.buffer.length, this.buffer[0].length).setValues(this.buffer);
    this.buffer = [];
  }
  formatTs_(d) {
    const date = d.toLocaleDateString(LOG_DATE_FMT_LOCALE);
    const time = d.toLocaleTimeString(LOG_DATE_FMT_LOCALE, { hour12: false });
    return `${date} ${time}`;
  }
}

function makeRunId_() {
  const d = new Date();
  const pad = (n, w=2) => String(n).padStart(w, '0');
  return `${d.getFullYear()}${pad(d.getMonth()+1)}${pad(d.getDate())}-` +
         `${pad(d.getHours())}${pad(d.getMinutes())}${pad(d.getSeconds())}-` +
         `${d.getTime()}`;
}
function formatDuration_(ms) {
  if (ms < 1000) return `${ms} ms`;
  const s = ms / 1000;
  if (s < 60) return `${s.toFixed(2)} s`;
  const m = Math.floor(s / 60);
  const r = (s - m*60).toFixed(2);
  return `${m} m ${r} s`;
}
function ensureLogSheetExists_(ss, logName) {
  let sh = ss.getSheetByName(logName);
  if (!sh) {
    sh = ss.insertSheet(logName);
    sh.getRange(1, 1).setValue('log initialized');
  }
  return sh;
}

/** === TAB POSITIONING HELPERS (NEW) === */
// Целевой индекс для DAY/NIGHT с базой после трёх закреплённых вкладок
function desiredTabIndexFor_(dayNum, prefix) {
  const BASE = 4; // 1-based: 1=ScheduleData, 2=_build_log, 3=template, => 4=DAY 1
  if (INCLUDE_DAY && INCLUDE_NIGHT) {
    const isDay = String(prefix).toUpperCase() === 'DAY';
    return BASE + (dayNum - 1) * 2 + (isDay ? 0 : 1);
  } else {
    return BASE + (dayNum - 1);
  }
}

function moveSheetToIndex_(ss, sheetName, index, logBuf) {
  try {
    const sh = ss.getSheetByName(sheetName);
    if (!sh) return;
    ss.setActiveSheet(sh);
    ss.moveActiveSheet(index);
    logBuf.info('TAB_MOVE', `"${sheetName}" moved to tab index ${index}`);
  } catch (e) {
    logBuf.error('TAB_MOVE_FAIL', `Failed to move "${sheetName}" to index ${index}: ${e}`);
  }
}

function positionFixedTabs_(ss, logBuf) {
  try {
    const sched = ss.getSheetByName('ScheduleData');
    if (sched) { ss.setActiveSheet(sched); ss.moveActiveSheet(1); }

    if (LOG_TO_SHEET) {
      const log = ss.getSheetByName(LOG_SHEET_NAME);
      if (log) { ss.setActiveSheet(log); ss.moveActiveSheet(2); }
    }

    const tmpl = ss.getSheetByName(TEMPLATE_SHEET_NAME);
    if (tmpl) { ss.setActiveSheet(tmpl); ss.moveActiveSheet(3); }

    logBuf.info('TAB_PIN', 'Pinned: ScheduleData→1, _build_log→2, template→3');
  } catch (e) {
    logBuf.error('TAB_PIN_FAIL', `Failed to pin fixed tabs: ${e}`);
  }
}

/** === SMALL HELPERS === */
function addMinutes_(dateObj, minutes) {
  const d = new Date(dateObj.getTime());
  d.setMinutes(d.getMinutes() + minutes);
  return d;
}
function formatTimeHMM_(d) {
  const h = d.getHours();
  const m = String(d.getMinutes()).padStart(2, '0');
  return `${h}:${m}`;
}
function formatDateDMY_(year, month1, day) {
  const d = new Date(year, month1 - 1, day);
  const dd = d.getDate();
  const MM = String(d.getMonth() + 1).padStart(2, '0');
  const yyyy = d.getFullYear();
  return `${MM}/${dd}/${yyyy}`;
}

/** === TEXT NORMALIZATION (NEW) === */
const HOMOGLYPHS_MAP = (function () {
  const map = {};
  Object.assign(map, {
    'А':'A','В':'B','С':'C','Е':'E','Н':'H','К':'K','М':'M','О':'O','Р':'P','Т':'T','Х':'X','У':'Y','І':'I','Ї':'I','Ј':'J',
    'а':'a','с':'c','е':'e','о':'o','р':'p','т':'t','х':'x','у':'y','і':'i','ї':'i','ј':'j'
  });
  Object.assign(map, {
    'Α':'A','Β':'B','Ε':'E','Ζ':'Z','Η':'H','Ι':'I','Κ':'K','Μ':'M','Ν':'N','Ο':'O','Ρ':'P','Τ':'T','Υ':'Y','Χ':'X','Ϊ':'I','Ϋ':'Y'
  });
  Object.assign(map, { 'ο':'o','ρ':'p','χ':'x' });
  return map;
})();

function normalizeText_(val) {
  if (val == null) return '';
  let t = String(val);
  t = t.replace(/[\u00A0\u1680\u180E\u2000-\u200A\u202F\u205F\u3000]/g, ' ');
  t = t.replace(/[\u200B-\u200D\uFEFF]/g, '');
  t = t.replace(/[‐-‒–—−﹘﹣－]/g, '-');

  let out = '';
  for (let i = 0; i < t.length; i++) {
    const ch = t[i];
    const code = ch.charCodeAt(0);
    if ((code >= 0xFF10 && code <= 0xFF19) || (code >= 0xFF21 && code <= 0xFF3A) || (code >= 0xFF41 && code <= 0xFF5A)) {
      out += String.fromCharCode(code - 0xFEE0);
      continue;
    }
    if (HOMOGLYPHS_MAP[ch]) {
      out += HOMOGLYPHS_MAP[ch];
      continue;
    }
    out += ch;
  }
  t = out;
  t = t.replace(/\s+/g, ' ').trim();
  return t;
}

function normalizeValues2D_(values2d) {
  const out = new Array(values2d.length);
  for (let r = 0; r < values2d.length; r++) {
    const row = values2d[r];
    const nr = new Array(row.length);
    for (let c = 0; c < row.length; c++) nr[c] = normalizeText_(row[c]);
    out[r] = nr;
  }
  return out;
}


function syncWithScheduleData_(rotationSheet, schedSheetName, logBuf) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const schedSheet = ss.getSheetByName(schedSheetName);
  if (!schedSheet) {
    logBuf.error('SYNC_FAIL', `No sheet "${schedSheetName}" found`);
    return;
  }

  const schedData = schedSheet.getDataRange().getValues(); // включая заголовки
  const schedHeader = schedData[0];
  const schedRows = schedData.slice(1);

  const schedMap = new Map(); // key = date + "::" + name → [rowIndex1, rowIndex2...]
  for (let i = 0; i < schedRows.length; i++) {
    const row = schedRows[i];
    const key = `${normalizeText_(row[4])}::${normalizeText_(row[10])}`; // date::employee
    if (!schedMap.has(key)) schedMap.set(key, []);
    schedMap.get(key).push(i + 1); // строка в таблице (без заголовка)
  }

  const rotationValues = rotationSheet.getDataRange().getValues();
  const rotationHeader = rotationValues[0];
  const rotationRows = rotationValues.slice(1);

  for (const rotRow of rotationRows) {
    const [date, start, end, name, ...slots] = rotRow;
    let dateStr, nameStr;

    try {
      dateStr = normalizeText_(toMM_dd_yyyy_(date));
    } catch (e) {
      logBuf.error('PARSE_DATE_FAIL', `❌ Failed to parse date "${date}" → ${e}`);
      continue;
    }

    nameStr = normalizeText_(name);
    const key = `${dateStr}::${nameStr}`;

    // 🔍 Поиск по имени в ScheduleData
    const schedMatchesByName = schedRows
      .map((row, idx) => ({
        idx,
        name: normalizeText_(row[10]),
        date: normalizeText_(row[4])
      }))
      .filter(r => r.name === nameStr);

    if (schedMatchesByName.length === 0) {
      logBuf.warn('NOT_FOUND', `🔴 Not found by name: "${nameStr}" → will add new row`);
    } else {
      const dateMatches = schedMatchesByName.filter(r => r.date === dateStr);
      const preview = schedMatchesByName.map(r => `row ${r.idx + 2} → ${r.date}`).join(', ');
      logBuf.info('MATCH_CHECK', `👁️ Found ${schedMatchesByName.length} ScheduleData row(s) for name "${nameStr}": ${preview}`);
      if (dateMatches.length === 0) {
        logBuf.warn('NOT_FOUND_DATE', `🟡 Name matched but no date match for "${dateStr} + ${nameStr}"`);
      } else {
        logBuf.info('MATCH', `🟢 Match: found ${dateMatches.length} row(s) for "${dateStr} + ${nameStr}"`);
      }
    }

    const newStart = normalizeText_(start);
    const newEnd = normalizeText_(end);
    const hasInterval = !!(newStart && newEnd);
    const slotsToInsert = slots.slice(0, 24).map(v => normalizeText_(v));

    const matching = schedMap.get(key) || [];


    // группировка: строки с/без comment
    const withComment = [], noComment = [];
    for (const rIdx of matching) {
      const row = schedRows[rIdx - 1];
      const comment = normalizeText_(row[11]);
      if (comment) withComment.push({ rIdx, row });
      else noComment.push({ rIdx, row });
    }

    // формируем строку для вставки
    const newRow = new Array(schedHeader.length).fill('');
    newRow[5] = dateStr;
    newRow[6] = newStart;
    newRow[7] = newEnd;
    newRow[10] = nameStr;
    for (let i = 0; i < 24; i++) newRow[14 + i] = slotsToInsert[i] || '';

    // интервалы
    if (newStart === '09:00' && newEnd === '21:00') {
      newRow[7] = 12; newRow[8] = ''; newRow[1] = true;
    } else if (newStart === '21:00' && newEnd === '09:00') {
      newRow[7] = ''; newRow[8] = 12; newRow[1] = true;
    } else if (!hasInterval) {
      newRow[5] = ''; newRow[6] = ''; newRow[7] = ''; newRow[8] = ''; newRow[1] = true;
    } else {
      newRow[7] = ''; newRow[8] = ''; newRow[1] = '';
    }

    // 📌 CASE: есть хотя бы одна строка без комментария → обновим первую, удалим остальные
    if (noComment.length > 0) {
      const { rIdx } = noComment[0];
      schedSheet.getRange(rIdx + 1, 1, 1, newRow.length).setValues([newRow]);
      // Было:
      for (let i = 1; i < noComment.length; i++) {
        schedSheet.deleteRow(noComment[i].rIdx + 1 - i);
      }

      // Стало:
      const toDelete = noComment.slice(1).map(o => o.rIdx + 1);
      toDelete.sort((a, b) => b - a);
      for (const delRow of toDelete) {
        try {
          schedSheet.deleteRow(delRow);
        } catch (e) {
          logBuf.error('DELETE_FAIL', `Can't delete row ${delRow}: ${e}`);
        }
      }

    }

    // 📌 CASE: только строки с комментариями
    else if (withComment.length > 0) {
      // ищем точное совпадение → confirm
      let found = false;
      for (const { rIdx, row } of withComment) {
        if (normalizeText_(row[5]) === newStart && normalizeText_(row[6]) === newEnd) {
          newRow[1] = true;
          schedSheet.getRange(rIdx + 1, 1, 1, newRow.length).setValues([newRow]);
          found = true;
          break;
        }
      }
      if (!found) {
        const rIdx = withComment[0].rIdx;
        newRow[1] = ''; // снимем confirm
        schedSheet.getRange(rIdx + 1, 1, 1, newRow.length).setValues([newRow]);
      }
    }

    // 📌 CASE: вообще нет строк
    else {
      const newLast = schedSheet.getLastRow() + 1;
      schedSheet.insertRowAfter(newLast);
      schedSheet.getRange(newLast + 1, 1, 1, newRow.length).setValues([newRow]);
    }
  }

  logBuf.info('SYNC_DONE', `Synchronized ${rotationRows.length} dealers with ScheduleData`);
}

// Date или MM/dd/yyyy → "yyyy-MM-dd" (без UTC-сдвига)
function toISO_yyyy_mm_dd_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, "yyyy-MM-dd");
  const s = String(v).trim();
  const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (!m) throw new Error(`❌ Unparsable date (expected MM/dd/yyyy): ${v}`);
  const mm = parseInt(m[1], 10);
  const dd = parseInt(m[2], 10);
  const yyyy = parseInt(m[3], 10);
  const d = new Date(yyyy, mm - 1, dd);
  return Utilities.formatDate(d, TZ, "yyyy-MM-dd");
}

// Date или MM/dd/yyyy → "MM/dd/yyyy"
function toMM_dd_yyyy_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, "MM/dd/yyyy");
  const s = String(v).trim();
  const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (!m) throw new Error(`❌ Unparsable date (expected MM/dd/yyyy): ${v}`);
  const mm = parseInt(m[1], 10);
  const dd = parseInt(m[2], 10);
  const yyyy = parseInt(m[3], 10);
  const d = new Date(yyyy, mm - 1, dd);
  return Utilities.formatDate(d, TZ, "MM/dd/yyyy");
}

// HH:mm diff (с поддержкой ночи)
function durationMinutes_(start, end) {
  const s = timeToMinutes_(start);
  const e = timeToMinutes_(end);
  if (e >= s) return e - s;
  return (24 * 60 - s) + e;
}

function timeToMinutes_(t) {
  if (t instanceof Date) return t.getHours() * 60 + t.getMinutes();
  const m = String(t).trim().match(/^(\d{1,2})[:.](\d{2})$/);
  if (!m) throw new Error(`❌ Unparsable time: ${t}`);
  const hh = parseInt(m[1], 10);
  const mm = parseInt(m[2], 10);
  return hh * 60 + mm;
}

function fmtTime_(t) {
  if (t instanceof Date) return Utilities.formatDate(t, TZ, "HH:mm");
  const s = String(t).trim();
  const m = s.match(/^(\d{1,2})[:.](\d{2})$/);
  if (!m) return s;
  const hh = String(parseInt(m[1], 10)).padStart(2, '0');
  const mm = String(parseInt(m[2], 10)).padStart(2, '0');
  return `${hh}:${mm}`;
}
