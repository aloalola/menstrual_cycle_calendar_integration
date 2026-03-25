/***************************************
 * CYCLE TRACKER - GOOGLE APPS SCRIPT
 * -------------------------------------
 * USER FLOW
 *
 * 1) "Set up sheets"
 *    - Creates the three required tabs:
 *      a) Cycle_Log
 *      b) Settings
 *      c) Generated_Events
 *    - Seeds default settings
 *    - Shows a popup reminding the user to enter actual cycle start dates
 *
 * 2) "Set up historical cycles"
 *    - Reads all actual cycle starts from Cycle_Log
 *    - Builds completed historical cycles from consecutive actual start dates
 *    - Creates all missing historical phase events in the calendar
 *    - Marks them as "historical" in Generated_Events
 *    - Historical cycles are frozen after creation
 *
 * 3) "Predict cycles"
 *    - This is the normal day-to-day button after setup
 *    - It first checks whether new actual dates now complete another historical cycle
 *    - If so, it creates only the missing historical events
 *    - Then it deletes ONLY mutable events:
 *         - current
 *         - predicted
 *    - Finally it rebuilds:
 *         - the current cycle
 *         - the future predicted cycles
 *
 * IMPORTANT LOGIC
 *
 * Historical cycles:
 * - built once from actual -> actual
 * - never changed again
 *
 * Current cycle:
 * - latest actual -> next predicted
 * - mutable
 *
 * Predicted cycles:
 * - future predicted -> predicted
 * - mutable
 *
 * PHASE LOGIC FOR EVERY CYCLE
 * - Menstruation: fixed N days from cycle start
 * - Follicular: fixed N days
 * - Ovulation: fixed N days
 * - Luteal: whatever remains until next cycle start
 ***************************************/


/* =====================================
   SHEET NAMES
===================================== */

const SHEET_NAMES = {
  CYCLE_LOG: 'Cycle_Log',
  SETTINGS: 'Settings',
  GENERATED_EVENTS: 'Generated_Events',
};


/* =====================================
   DEFAULT SETTINGS
===================================== */

const DEFAULT_SETTINGS = {
  calendar_name: 'Cycle Tracker',
  months_lookback: 6,
  menstruation_days: 5,
  follicular_days: 8,
  ovulation_days: 5,
  predict_cycles_ahead: 6,
  phase_event_prefix: 'Cycle —',
};


/* =====================================
   PHASE DISPLAY / STYLING METADATA
   - Keeps emojis and colors consistent everywhere
===================================== */

const PHASE_META = {
  Menstruation: {
    emoji: '🩸',
    color: CalendarApp.EventColor.PALE_RED,
  },
  Follicular: {
    emoji: '🌱',
    color: CalendarApp.EventColor.PALE_GREEN,
  },
  Ovulation: {
    emoji: '✨',
    color: CalendarApp.EventColor.YELLOW,
  },
  Luteal: {
    emoji: '🌙',
    color: CalendarApp.EventColor.PALE_BLUE,
  },
};


/* =====================================
   CUSTOM MENU
   - This menu appears when the Google Sheet opens
===================================== */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Cycle Tracker')
    .addItem('1) Set up sheets', 'setupSheets')
    .addItem('2) Set up historical cycles', 'setupHistoricalCycles')
    .addItem('3) Predict cycles', 'predictCycles')
    .addSeparator()
    .addItem('Preview prediction', 'previewPrediction')
    .addSeparator()
    .addItem('Install daily trigger', 'installDailyTrigger')
    .addItem('Delete daily triggers', 'deleteDailyTriggers')
    .addToUi();
}


/* =====================================
   STEP 1: SET UP SHEETS
   - Creates missing sheets
   - Creates required headers
   - Seeds default settings
   - Shows user guidance popup
===================================== */

function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Create / prepare Cycle_Log sheet
  let cycleLog = ss.getSheetByName(SHEET_NAMES.CYCLE_LOG);
  if (!cycleLog) cycleLog = ss.insertSheet(SHEET_NAMES.CYCLE_LOG);
  ensureHeader_(cycleLog, ['cycle_start_actual', 'notes']);

  // Create / prepare Settings sheet
  let settings = ss.getSheetByName(SHEET_NAMES.SETTINGS);
  if (!settings) settings = ss.insertSheet(SHEET_NAMES.SETTINGS);
  ensureHeader_(settings, ['key', 'value']);
  seedDefaultSettings_(settings);

  // Create / prepare Generated_Events sheet
  let generated = ss.getSheetByName(SHEET_NAMES.GENERATED_EVENTS);
  if (!generated) generated = ss.insertSheet(SHEET_NAMES.GENERATED_EVENTS);
  ensureHeader_(generated, [
    'anchor_cycle_start',
    'next_cycle_start',
    'phase_name',
    'start_date',
    'end_date',
    'status',
    'calendar_event_id',
    'created_at',
  ]);

  SpreadsheetApp.getUi().alert(
    [
      'Sheets are ready.',
      '',
      'Next step:',
      `Please enter your actual cycle start dates into "${SHEET_NAMES.CYCLE_LOG}"`,
      'under the column "cycle_start_actual".',
      '',
      'After that, click "2) Set up historical cycles".',
    ].join('\n')
  );
}


/* =====================================
   STEP 2: SET UP HISTORICAL CYCLES
   - Creates all missing historical cycles from past actual dates
   - Historical cycles are frozen after creation
   - This can safely be run multiple times; it only creates missing ones
===================================== */

function setupHistoricalCycles() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const settings = getSettings_();
    const calendar = getOrCreateCalendar_(settings.calendar_name);
    const actualStarts = getActualCycleStarts_();

    if (actualStarts.length < 2) {
      throw new Error(
        'You need at least 2 cycle start dates in Cycle_Log to set up historical cycles.'
      );
    }

    // Calculate average cycle length only for metadata / descriptions
    const calc = calculateCyclePrediction_(actualStarts, settings);

    // Build all completed historical cycles from actual -> actual
    const historicalCycles = buildHistoricalCycles_(actualStarts);

    // Read already-existing historical rows so we do not duplicate them
    const existingHistoricalKeys = getExistingGeneratedKeysByStatus_(['historical']);

    // Build only the historical phases that are still missing
    const missingHistoricalPhases = buildMissingHistoricalPhases_(
      historicalCycles,
      settings,
      existingHistoricalKeys
    );

    // Create those missing historical events in the calendar
    const historicalRows = createCalendarEventsForPhases_(
      calendar,
      missingHistoricalPhases,
      calc,
      settings
    );

    // Append them to Generated_Events
    appendGeneratedRows_(historicalRows);

    SpreadsheetApp.getUi().alert(
      [
        'Historical cycle setup complete.',
        `Created ${historicalRows.length} historical phase events.`,
        '',
        'From now on, those historical cycles will remain untouched.',
        'Next, use "3) Predict cycles" for normal ongoing updates.',
      ].join('\n')
    );
  } finally {
    lock.releaseLock();
  }
}


/* =====================================
   STEP 3: PREDICT CYCLES
   - This is the main button used after setup
   - Creates any newly missing historical cycles
   - Deletes only current + predicted events
   - Rebuilds current + predicted cycles
   - Historical cycles remain untouched
===================================== */

function predictCycles() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const settings = getSettings_();
    const calendar = getOrCreateCalendar_(settings.calendar_name);
    const actualStarts = getActualCycleStarts_();

    if (actualStarts.length < 2) {
      throw new Error('You need at least 2 cycle start dates in Cycle_Log.');
    }

    // This uses the latest actual starts to estimate the next start
    const calc = calculateCyclePrediction_(actualStarts, settings);

    // ---------------------------------
    // A) FINALIZE ANY NEW HISTORICAL CYCLES
    // ---------------------------------
    // If the user added a new actual cycle start, then the previously current cycle
    // is now complete and should be converted into a frozen historical cycle.
    // We do this by checking which historical cycles are missing.
    const historicalCycles = buildHistoricalCycles_(actualStarts);
    const existingHistoricalKeys = getExistingGeneratedKeysByStatus_(['historical']);
    const missingHistoricalPhases = buildMissingHistoricalPhases_(
      historicalCycles,
      settings,
      existingHistoricalKeys
    );

    const historicalRows = createCalendarEventsForPhases_(
      calendar,
      missingHistoricalPhases,
      calc,
      settings
    );

    // Historical rows are appended and never touched again later
    appendGeneratedRows_(historicalRows);

    // ---------------------------------
    // B) DELETE ONLY MUTABLE EVENTS
    // ---------------------------------
    // Mutable statuses:
    // - current
    // - predicted
    clearMutableGeneratedEvents_(calendar);

    // ---------------------------------
    // C) REBUILD CURRENT + FUTURE PREDICTED CYCLES
    // ---------------------------------
    const mutableCycles = buildMutableCycles_(calc, settings);
    const mutablePhases = buildAllPhaseWindows_(mutableCycles, settings);

    const mutableRows = createCalendarEventsForPhases_(
      calendar,
      mutablePhases,
      calc,
      settings
    );

    appendGeneratedRows_(mutableRows);

    SpreadsheetApp.getUi().alert(
      [
        'Cycle prediction updated successfully.',
        `Latest actual start: ${formatDateYMD_(calc.latestActualStart)}`,
        `Average cycle length: ${calc.avgCycleLength} days`,
        `Next predicted start: ${formatDateYMD_(calc.nextPredictedStart)}`,
        `Newly finalized historical phase events: ${historicalRows.length}`,
        `Rebuilt mutable phase events: ${mutableRows.length}`,
      ].join('\n')
    );
  } finally {
    lock.releaseLock();
  }
}


/* =====================================
   OPTIONAL PREVIEW
   - Shows what current + predicted cycles would look like
   - Does not write anything
===================================== */

function previewPrediction() {
  const settings = getSettings_();
  const actualStarts = getActualCycleStarts_();

  if (actualStarts.length < 2) {
    SpreadsheetApp.getUi().alert(
      'You need at least 2 cycle start dates in Cycle_Log.'
    );
    return;
  }

  const calc = calculateCyclePrediction_(actualStarts, settings);
  const mutableCycles = buildMutableCycles_(calc, settings);
  const mutablePhases = buildAllPhaseWindows_(mutableCycles, settings);

  const lines = [
    `Latest actual cycle start: ${formatDateYMD_(calc.latestActualStart)}`,
    `Average cycle length (last ${settings.months_lookback} months): ${calc.avgCycleLength}`,
    `Next predicted start: ${formatDateYMD_(calc.nextPredictedStart)}`,
    '',
    'Mutable cycles (current + predicted):',
  ];

  mutableCycles.forEach(cycle => {
    lines.push(
      `${cycle.status.toUpperCase()} | ${formatDateYMD_(cycle.anchor_cycle_start)} → ${formatDateYMD_(cycle.next_cycle_start)}`
    );
  });

  lines.push('', 'Phase windows:');

  mutablePhases.forEach(phase => {
    const visibleEnd = new Date(phase.end_date_exclusive.getTime() - 24 * 60 * 60 * 1000);
    lines.push(
      `${buildPhaseTitle_(phase.phase_name, settings)} | ${phase.status} | ${formatDateYMD_(phase.start_date)} → ${formatDateYMD_(visibleEnd)}`
    );
  });

  SpreadsheetApp.getUi().alert(lines.join('\n'));
}


/* =====================================
   CORE CYCLE BUILDING
===================================== */

/**
 * Reads all actual cycle start dates from Cycle_Log.
 * Returns a sorted array of Date objects.
 */
function getActualCycleStarts_() {
  const sheet = getRequiredSheet_(SHEET_NAMES.CYCLE_LOG);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const values = sheet.getRange(2, 1, lastRow - 1, 2).getValues();

  return values
    .map(row => row[0])
    .filter(v => v instanceof Date && !isNaN(v.getTime()))
    .map(normalizeDate_)
    .sort((a, b) => a - b);
}


/**
 * Calculates rolling average cycle length based on the last X months.
 * This is used to predict the next cycle start after the latest actual start.
 */
function calculateCyclePrediction_(actualStarts, settings) {
  const latestActualStart = actualStarts[actualStarts.length - 1];
  const lookbackStart = addMonths_(latestActualStart, -Number(settings.months_lookback));

  const intervals = [];

  for (let i = 1; i < actualStarts.length; i++) {
    const prev = actualStarts[i - 1];
    const curr = actualStarts[i];

    // Only include intervals where the later cycle start lies inside the lookback window
    if (curr >= lookbackStart) {
      intervals.push(daysBetween_(prev, curr));
    }
  }

  if (intervals.length === 0) {
    throw new Error('No usable cycle intervals found inside the lookback window.');
  }

  const avgRaw = intervals.reduce((sum, x) => sum + x, 0) / intervals.length;
  const avgCycleLength = Math.round(avgRaw);
  const nextPredictedStart = addDays_(latestActualStart, avgCycleLength);

  return {
    latestActualStart,
    lookbackStart,
    intervals,
    avgCycleLength,
    nextPredictedStart,
  };
}


/**
 * Builds all completed historical cycles.
 * Historical means: actual cycle start -> next actual cycle start.
 * These cycles are frozen once created.
 */
function buildHistoricalCycles_(actualStarts) {
  const cycles = [];

  for (let i = 0; i < actualStarts.length - 1; i++) {
    cycles.push({
      anchor_cycle_start: actualStarts[i],
      next_cycle_start: actualStarts[i + 1],
      status: 'historical',
    });
  }

  return cycles;
}


/**
 * Builds the mutable cycles:
 * - current cycle = latest actual -> next predicted
 * - future predicted cycles after that
 *
 * Note:
 * predict_cycles_ahead controls how many future predicted cycles
 * are created AFTER the current cycle.
 */
function buildMutableCycles_(calc, settings) {
  const futurePredictedCycles = Number(settings.predict_cycles_ahead);
  const cycles = [];

  // Current cycle
  cycles.push({
    anchor_cycle_start: calc.latestActualStart,
    next_cycle_start: calc.nextPredictedStart,
    status: 'current',
  });

  // Future predicted cycles
  for (let i = 0; i < futurePredictedCycles; i++) {
    const anchor = addDays_(calc.nextPredictedStart, i * calc.avgCycleLength);
    const nextAnchor = addDays_(anchor, calc.avgCycleLength);

    cycles.push({
      anchor_cycle_start: anchor,
      next_cycle_start: nextAnchor,
      status: 'predicted',
    });
  }

  return cycles;
}


/**
 * Builds all phase windows for all provided cycles.
 * Uses the same phase logic for historical, current, and predicted.
 */
function buildAllPhaseWindows_(cycles, settings) {
  const allPhases = [];

  cycles.forEach(cycle => {
    const phases = buildPhaseWindowsForCycle_(cycle, settings);
    allPhases.push(...phases);
  });

  return allPhases;
}


/**
 * Builds the 4 phases for a single cycle:
 * - Menstruation: fixed N days
 * - Follicular: fixed N days
 * - Ovulation: fixed N days
 * - Luteal: remainder until next cycle start
 */
function buildPhaseWindowsForCycle_(cycle, settings) {
  const menstruationDays = Number(settings.menstruation_days);
  const follicularDays = Number(settings.follicular_days);
  const ovulationDays = Number(settings.ovulation_days);

  const cycleLength = daysBetween_(cycle.anchor_cycle_start, cycle.next_cycle_start);
  const fixedDays = menstruationDays + follicularDays + ovulationDays;

  if (fixedDays >= cycleLength) {
    throw new Error(
      `Cycle ${formatDateYMD_(cycle.anchor_cycle_start)} → ${formatDateYMD_(cycle.next_cycle_start)} is too short (${cycleLength} days) for fixed phase lengths (${fixedDays}).`
    );
  }

  const menstruationStart = cycle.anchor_cycle_start;
  const menstruationEndExclusive = addDays_(menstruationStart, menstruationDays);

  const follicularStart = menstruationEndExclusive;
  const follicularEndExclusive = addDays_(follicularStart, follicularDays);

  const ovulationStart = follicularEndExclusive;
  const ovulationEndExclusive = addDays_(ovulationStart, ovulationDays);

  const lutealStart = ovulationEndExclusive;
  const lutealEndExclusive = cycle.next_cycle_start;

  return [
    {
      anchor_cycle_start: cycle.anchor_cycle_start,
      next_cycle_start: cycle.next_cycle_start,
      phase_name: 'Menstruation',
      start_date: menstruationStart,
      end_date_exclusive: menstruationEndExclusive,
      status: cycle.status,
    },
    {
      anchor_cycle_start: cycle.anchor_cycle_start,
      next_cycle_start: cycle.next_cycle_start,
      phase_name: 'Follicular',
      start_date: follicularStart,
      end_date_exclusive: follicularEndExclusive,
      status: cycle.status,
    },
    {
      anchor_cycle_start: cycle.anchor_cycle_start,
      next_cycle_start: cycle.next_cycle_start,
      phase_name: 'Ovulation',
      start_date: ovulationStart,
      end_date_exclusive: ovulationEndExclusive,
      status: cycle.status,
    },
    {
      anchor_cycle_start: cycle.anchor_cycle_start,
      next_cycle_start: cycle.next_cycle_start,
      phase_name: 'Luteal',
      start_date: lutealStart,
      end_date_exclusive: lutealEndExclusive,
      status: cycle.status,
    },
  ];
}


/**
 * Builds only the historical phases that are still missing.
 * This prevents duplicate creation and keeps historical cycles frozen.
 */
function buildMissingHistoricalPhases_(historicalCycles, settings, existingHistoricalKeys) {
  const allHistoricalPhases = buildAllPhaseWindows_(historicalCycles, settings);

  return allHistoricalPhases.filter(phase => {
    const key = makePhaseKey_(phase);
    return !existingHistoricalKeys.has(key);
  });
}


/* =====================================
   CALENDAR EVENT CREATION
===================================== */

/**
 * Creates calendar events for a list of phase objects and returns
 * rows ready to be written into Generated_Events.
 */
function createCalendarEventsForPhases_(calendar, phases, calc, settings) {
  const rows = [];

  phases.forEach(phase => {
    const title = buildPhaseTitle_(phase.phase_name, settings);

    const description = [
      'Generated by Apps Script',
      `Phase: ${phase.phase_name}`,
      `Cycle type: ${phase.status}`,
      `Cycle start: ${formatDateYMD_(phase.anchor_cycle_start)}`,
      `Next cycle start: ${formatDateYMD_(phase.next_cycle_start)}`,
      `Average cycle length used: ${calc.avgCycleLength}`,
      `Generated at: ${formatDateYMDHMS_(new Date())}`,
    ].join('\n');

    const event = calendar.createAllDayEvent(
      title,
      phase.start_date,
      phase.end_date_exclusive,
      { description }
    );

    applyPhaseColor_(event, phase.phase_name);

    rows.push([
      phase.anchor_cycle_start,
      phase.next_cycle_start,
      phase.phase_name,
      phase.start_date,
      new Date(phase.end_date_exclusive.getTime() - 24 * 60 * 60 * 1000),
      phase.status,
      event.getId(),
      new Date(),
    ]);
  });

  return rows;
}


/* =====================================
   GENERATED EVENTS SHEET MANAGEMENT
===================================== */

/**
 * Appends new rows to Generated_Events.
 */
function appendGeneratedRows_(rows) {
  if (!rows || rows.length === 0) return;

  const sheet = getRequiredSheet_(SHEET_NAMES.GENERATED_EVENTS);
  const startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);
}


/**
 * Deletes only mutable events:
 * - current
 * - predicted
 *
 * Historical rows and events are kept untouched.
 */
function clearMutableGeneratedEvents_(calendar) {
  const sheet = getRequiredSheet_(SHEET_NAMES.GENERATED_EVENTS);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const values = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
  const keepRows = [];

  values.forEach(row => {
    const [
      anchorCycleStart,
      nextCycleStart,
      phaseName,
      startDate,
      endDate,
      status,
      calendarEventId,
      createdAt,
    ] = row;

    const shouldDelete = status === 'current' || status === 'predicted';

    if (shouldDelete && calendarEventId) {
      try {
        const event = calendar.getEventById(calendarEventId);
        if (event) {
          event.deleteEvent();
        }
      } catch (err) {
        Logger.log(`Could not delete event ${calendarEventId}: ${err}`);
      }
    } else {
      keepRows.push([
        anchorCycleStart,
        nextCycleStart,
        phaseName,
        startDate,
        endDate,
        status,
        calendarEventId,
        createdAt,
      ]);
    }
  });

  // Rewrite sheet with only historical rows preserved
  sheet.getRange(2, 1, Math.max(lastRow - 1, 1), 8).clearContent();

  if (keepRows.length > 0) {
    sheet.getRange(2, 1, keepRows.length, 8).setValues(keepRows);
  }
}


/**
 * Reads existing generated rows and returns keys only for the requested statuses.
 * This is used to avoid recreating historical cycles that already exist.
 */
function getExistingGeneratedKeysByStatus_(statuses) {
  const allowed = new Set(statuses);
  const sheet = getRequiredSheet_(SHEET_NAMES.GENERATED_EVENTS);
  const lastRow = sheet.getLastRow();
  const keys = new Set();

  if (lastRow < 2) return keys;

  const values = sheet.getRange(2, 1, lastRow - 1, 8).getValues();

  values.forEach(row => {
    const [
      anchorCycleStart,
      nextCycleStart,
      phaseName,
      startDate,
      endDate,
      status,
      calendarEventId,
      createdAt,
    ] = row;

    if (
      allowed.has(status) &&
      anchorCycleStart instanceof Date &&
      nextCycleStart instanceof Date &&
      startDate instanceof Date &&
      phaseName
    ) {
      const key = [
        formatDateYMD_(anchorCycleStart),
        formatDateYMD_(nextCycleStart),
        phaseName,
        formatDateYMD_(startDate),
        status,
      ].join('|');

      keys.add(key);
    }
  });

  return keys;
}


/**
 * Builds a stable unique key for a phase row.
 * Used for deduplication, especially for frozen historical phases.
 */
function makePhaseKey_(phase) {
  return [
    formatDateYMD_(phase.anchor_cycle_start),
    formatDateYMD_(phase.next_cycle_start),
    phase.phase_name,
    formatDateYMD_(phase.start_date),
    phase.status,
  ].join('|');
}


/* =====================================
   CALENDAR HELPERS
===================================== */

/**
 * Gets the target calendar by name, or creates it if it does not exist.
 */
function getOrCreateCalendar_(calendarName) {
  const existing = CalendarApp.getCalendarsByName(calendarName);
  const calendar = (existing && existing.length > 0)
    ? existing[0]
    : CalendarApp.createCalendar(calendarName);

  calendar.setSelected(true);
  return calendar;
}


/* =====================================
   PHASE TITLE / COLOR HELPERS
===================================== */

/**
 * Builds a consistent phase title including emoji.
 * This keeps naming coherent across all cycle types.
 */
function buildPhaseTitle_(phaseName, settings) {
  const meta = PHASE_META[phaseName];
  const emoji = meta ? meta.emoji : '';
  return `${emoji} ${settings.phase_event_prefix} ${phaseName}`.trim();
}


/**
 * Applies the configured event color for a phase.
 */
function applyPhaseColor_(event, phaseName) {
  const meta = PHASE_META[phaseName];
  if (meta && meta.color) {
    event.setColor(meta.color);
  }
}


/* =====================================
   SETTINGS + SHEET HELPERS
===================================== */

/**
 * Reads settings from the Settings sheet.
 * Missing settings fall back to DEFAULT_SETTINGS.
 */
function getSettings_() {
  const sheet = getRequiredSheet_(SHEET_NAMES.SETTINGS);
  const lastRow = sheet.getLastRow();

  const settings = { ...DEFAULT_SETTINGS };
  if (lastRow < 2) return settings;

  const values = sheet.getRange(2, 1, lastRow - 1, 2).getValues();

  values.forEach(([key, value]) => {
    if (!key) return;
    settings[String(key).trim()] = parseSettingValue_(value);
  });

  return settings;
}


/**
 * Converts settings values from sheet cells into booleans / numbers / strings.
 */
function parseSettingValue_(value) {
  if (typeof value === 'boolean') return value;
  if (typeof value === 'number') return value;
  if (value === null || value === '') return '';

  const str = String(value).trim();

  if (str.toLowerCase() === 'true') return true;
  if (str.toLowerCase() === 'false') return false;
  if (!isNaN(Number(str))) return Number(str);

  return str;
}


/**
 * Ensures the given sheet has exactly the given headers in row 1.
 */
function ensureHeader_(sheet, headers) {
  const existing = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  const matches = headers.every((h, i) => existing[i] === h);

  if (!matches) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
}


/**
 * Seeds default settings if the Settings sheet is still empty.
 */
function seedDefaultSettings_(settingsSheet) {
  const lastRow = settingsSheet.getLastRow();
  if (lastRow >= 2) return;

  const rows = Object.entries(DEFAULT_SETTINGS);
  settingsSheet.getRange(2, 1, rows.length, 2).setValues(rows);
}


/**
 * Returns a required sheet by name or throws an error if it does not exist.
 */
function getRequiredSheet_(name) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sheet) {
    throw new Error(`Missing required sheet: ${name}`);
  }
  return sheet;
}


/* =====================================
   TRIGGERS
===================================== */

/**
 * Installs a daily trigger that runs the normal prediction flow.
 * This is optional. If you use it, it will behave exactly like clicking "Predict cycles".
 */
function installDailyTrigger() {
  deleteDailyTriggers();

  ScriptApp.newTrigger('predictCycles')
    .timeBased()
    .everyDays(1)
    .atHour(6)
    .create();

  SpreadsheetApp.getUi().alert('Daily trigger installed.');
}


/**
 * Deletes all project triggers that run predictCycles.
 */
function deleteDailyTriggers() {
  const triggers = ScriptApp.getProjectTriggers();

  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'predictCycles') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  SpreadsheetApp.getUi().alert('Daily triggers deleted.');
}


/* =====================================
   DATE HELPERS
===================================== */

/**
 * Normalizes a Date to midnight local time.
 * This prevents time-of-day problems when comparing dates.
 */
function normalizeDate_(date) {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}


/**
 * Adds N days to a date and returns a normalized new date.
 */
function addDays_(date, days) {
  const d = normalizeDate_(date);
  d.setDate(d.getDate() + Number(days));
  return d;
}


/**
 * Adds N months to a date and returns a normalized new date.
 */
function addMonths_(date, months) {
  const d = normalizeDate_(date);
  d.setMonth(d.getMonth() + Number(months));
  return d;
}


/**
 * Returns the difference in whole days between two dates.
 */
function daysBetween_(dateA, dateB) {
  const ms = normalizeDate_(dateB).getTime() - normalizeDate_(dateA).getTime();
  return Math.round(ms / (24 * 60 * 60 * 1000));
}


/**
 * Formats a date as YYYY-MM-DD in the script time zone.
 */
function formatDateYMD_(date) {
  return Utilities.formatDate(
    normalizeDate_(date),
    Session.getScriptTimeZone(),
    'yyyy-MM-dd'
  );
}


/**
 * Formats a date-time as YYYY-MM-DD HH:mm:ss in the script time zone.
 */
function formatDateYMDHMS_(date) {
  return Utilities.formatDate(
    date,
    Session.getScriptTimeZone(),
    'yyyy-MM-dd HH:mm:ss'
  );
}