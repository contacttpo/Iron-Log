/**
 * IRON LOG — API BACKEND (Apps Script)
 * This is NOT the mobile UI anymore — App.html/Code.gs's UI job has been
 * replaced by the standalone index.html + app.js pair. This script's only
 * job now is to be a small JSON API in front of your Google Sheet, so the
 * standalone app can read/write workout data from any device/browser.
 *
 * SETUP:
 * 1. Go to https://script.google.com/create (a standalone project —
 *    it does NOT need to be bound to the sheet).
 * 2. Delete the default code, paste this whole file in.
 * 3. Confirm SHEET_ID below matches your sheet (already filled in).
 * 4. Deploy > New deployment > type: Web app
 *      - Execute as: Me
 *      - Who has access: Anyone   (needed so it works from any device
 *        without a Google login prompt; see note in the chat reply)
 * 5. Authorize when prompted (first deploy only).
 * 6. Copy the Web app URL it gives you — you'll paste it into app.js
 *    as API_URL.
 * 7. Re-run "Deploy > Manage deployments > Edit > New version" any time
 *    you change this file, or the live URL won't see your changes.
 */

const SHEET_ID = '1_M0D3dvh_CDNM_M9i5kkeZ__tTkcmHljkQXmFB1n-Fo';
const LOG_SHEET_NAME = 'Log';

const DAY_ROTATION = ['Chest+Triceps', 'Back+Biceps', 'Shoulders+Legs'];

const EXERCISE_LIBRARY = [
  ['Chest', 'Barbell Flat Bench Press', 'Chest+Triceps', 6, 10, 'compound'],
  ['Chest', 'Incline Dumbbell Press', 'Chest+Triceps', 8, 12, 'compound'],
  ['Chest', 'Flat Dumbbell Press', 'Chest+Triceps', 8, 12, 'compound'],
  ['Chest', 'Weighted Dips', 'Chest+Triceps', 8, 12, 'compound'],
  ['Chest', 'Dumbbell Flyes', 'Chest+Triceps', 10, 15, 'isolation'],
  ['Triceps', 'Close-Grip Barbell Bench Press', 'Chest+Triceps', 6, 10, 'compound'],
  ['Triceps', 'Lying Barbell/Dumbbell Skull Crusher', 'Chest+Triceps', 8, 12, 'isolation'],
  ['Triceps', 'Seated Overhead Dumbbell Extension', 'Chest+Triceps', 10, 15, 'isolation'],
  ['Triceps', 'Diamond Push-ups', 'Chest+Triceps', 8, 15, 'compound'],
  ['Triceps', 'Dumbbell Kickback', 'Chest+Triceps', 12, 15, 'isolation'],
  ['Back', 'Weighted Pull-ups / Chin-ups', 'Back+Biceps', 6, 10, 'compound'],
  ['Back', 'Romanian/Conventional Deadlift', 'Back+Biceps', 5, 8, 'compound'],
  ['Back', 'Barbell Bent-over Row', 'Back+Biceps', 6, 10, 'compound'],
  ['Back', 'Single-arm Dumbbell Row', 'Back+Biceps', 8, 12, 'compound'],
  ['Back', 'Barbell/Dumbbell Shrugs', 'Back+Biceps', 10, 15, 'isolation'],
  ['Biceps', 'Standing Barbell Curl', 'Back+Biceps', 6, 10, 'isolation'],
  ['Biceps', 'Incline Dumbbell Curl', 'Back+Biceps', 10, 15, 'isolation'],
  ['Biceps', 'Concentration Curl', 'Back+Biceps', 10, 15, 'isolation'],
  ['Biceps', 'Hammer Curl', 'Back+Biceps', 10, 15, 'isolation'],
  ['Biceps', 'Alternating Dumbbell Curl', 'Back+Biceps', 8, 12, 'isolation'],
  ['Shoulders', 'Standing Barbell/Dumbbell Overhead Press', 'Shoulders+Legs', 6, 10, 'compound'],
  ['Shoulders', 'Arnold Press', 'Shoulders+Legs', 8, 12, 'compound'],
  ['Shoulders', 'Dumbbell Lateral Raise', 'Shoulders+Legs', 12, 15, 'isolation'],
  ['Shoulders', 'Bent-over Dumbbell Reverse Fly', 'Shoulders+Legs', 12, 15, 'isolation'],
  ['Shoulders', 'Dumbbell Front Raise', 'Shoulders+Legs', 10, 15, 'isolation'],
  ['Legs', 'Barbell Back Squat / Goblet Squat', 'Shoulders+Legs', 6, 10, 'compound'],
  ['Legs', 'Romanian Deadlift', 'Shoulders+Legs', 8, 12, 'compound'],
  ['Legs', 'Bulgarian Split Squat', 'Shoulders+Legs', 8, 12, 'compound'],
  ['Legs', 'Dumbbell Walking Lunges', 'Shoulders+Legs', 10, 15, 'compound'],
  ['Legs', 'Standing Dumbbell Calf Raise', 'Shoulders+Legs', 12, 20, 'isolation']
];

// ---------- SHEET ACCESS ----------
function getSheet_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sh = ss.getSheetByName(LOG_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(LOG_SHEET_NAME);
    sh.getRange(1, 1, 1, 9).setValues([[
      'Date', 'Day Type', 'Muscle', 'Exercise', 'Set #', 'Weight (kg)', 'Reps', 'Volume', 'Notes'
    ]]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function readLog_() {
  const sh = getSheet_();
  const values = sh.getDataRange().getValues();
  const rows = [];
  for (let i = 1; i < values.length; i++) {
    const r = values[i];
    if (!r[0]) continue;
    rows.push({ date: r[0], dayType: r[1], muscle: r[2], exercise: r[3], setNum: r[4], weight: r[5], reps: r[6], volume: r[7], notes: r[8] });
  }
  return rows;
}

function appendLogRows_(rows) {
  const sh = getSheet_();
  const startRow = sh.getLastRow() + 1;
  const values = rows.map(r => [r.date, r.dayType, r.muscle, r.exercise, r.setNum, r.weight, r.reps, '', r.notes]);
  sh.getRange(startRow, 1, values.length, 9).setValues(values);
  const formulas = rows.map((r, idx) => {
    const rowNum = startRow + idx;
    return [`=IF(AND(F${rowNum}<>"",G${rowNum}<>""),F${rowNum}*G${rowNum},"")`];
  });
  sh.getRange(startRow, 8, formulas.length, 1).setFormulas(formulas);
}

// ---------- HIDDEN EXERCISES (shared across devices via Script Properties) ----------
function getHiddenSet_() {
  const raw = PropertiesService.getScriptProperties().getProperty('hiddenExercises');
  return new Set(raw ? JSON.parse(raw) : []);
}

function saveHiddenSet_(set) {
  PropertiesService.getScriptProperties().setProperty('hiddenExercises', JSON.stringify([...set]));
}

function setExerciseEnabled(exerciseName, enabled) {
  const hidden = getHiddenSet_();
  hidden.delete(exerciseName);
  if (!enabled) hidden.add(exerciseName);
  saveHiddenSet_(hidden);
  return true;
}

// ---------- NEXT WORKOUT LOGIC ----------
function nextDayFromLog_(data) {
  let lastDate = null, lastDayType = null;
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    if (!row.date) continue;
    const d = new Date(row.date);
    if (!lastDate || d >= lastDate) { lastDate = d; lastDayType = row.dayType; }
  }
  if (!lastDayType) return DAY_ROTATION[0];
  const idx = DAY_ROTATION.indexOf(lastDayType);
  return idx === -1 ? DAY_ROTATION[0] : DAY_ROTATION[(idx + 1) % DAY_ROTATION.length];
}

// ---------- DATA ENDPOINTS ----------
function getTodayData(overrideDay) {
  const data = readLog_();
  const suggested = nextDayFromLog_(data);
  const dayType = (overrideDay && DAY_ROTATION.indexOf(overrideDay) !== -1) ? overrideDay : suggested;
  const hidden = getHiddenSet_();

  const lastMap = {};
  const loggedDates = new Set();
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    if (!row.date) continue;
    loggedDates.add(new Date(row.date).toDateString());
    const name = row.exercise;
    const d = new Date(row.date);
    if (!lastMap[name] || d >= lastMap[name].date) {
      lastMap[name] = { weight: row.weight, reps: row.reps, date: d };
    }
  }

  const exercises = EXERCISE_LIBRARY
    .filter(r => r[2] === dayType && !hidden.has(r[1]))
    .map(r => {
      const last = lastMap[r[1]] || { weight: null, reps: null };
      return { muscle: r[0], exercise: r[1], repLow: r[3], repHigh: r[4], type: r[5], lastWeight: last.weight, lastReps: last.reps };
    });
  const muscles = [...new Set(exercises.map(e => e.muscle))];
  return { dayType, muscles, exercises, suggested, rotation: DAY_ROTATION, streak: loggedDates.size };
}

function getLibraryData() {
  const hidden = getHiddenSet_();
  const grouped = {};
  EXERCISE_LIBRARY.forEach(r => {
    if (!grouped[r[0]]) grouped[r[0]] = [];
    grouped[r[0]].push({ exercise: r[1], dayType: r[2], repLow: r[3], repHigh: r[4], enabled: !hidden.has(r[1]) });
  });
  return grouped;
}

function getProgressData() {
  const data = readLog_();
  const muscles = [...new Set(EXERCISE_LIBRARY.map(r => r[0]))];
  const weekAgo = new Date();
  weekAgo.setDate(weekAgo.getDate() - 7);

  const volumeByMuscle = {};
  muscles.forEach(m => volumeByMuscle[m] = 0);
  const dates = new Set();

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    if (!row.date) continue;
    const d = new Date(row.date);
    dates.add(d.toDateString());
    if (d >= weekAgo && row.volume) {
      volumeByMuscle[row.muscle] = (volumeByMuscle[row.muscle] || 0) + Number(row.volume);
    }
  }

  const maxVolume = Math.max(1, ...Object.values(volumeByMuscle));
  const volumeChart = muscles.map(m => ({
    muscle: m, volume: Math.round(volumeByMuscle[m] || 0),
    pct: Math.round(((volumeByMuscle[m] || 0) / maxVolume) * 100)
  }));

  return { volumeChart, streak: dates.size };
}

function submitSets(entry) {
  const rows = entry.sets.map((s, idx) => {
    const weight = Number(s.weight);
    const reps = Number(s.reps);
    return {
      date: entry.date, dayType: entry.dayType, muscle: entry.muscle, exercise: entry.exercise,
      setNum: idx + 1, weight: weight, reps: reps,
      notes: idx === 0 ? (entry.notes || '') : ''
    };
  });
  appendLogRows_(rows);

  const lib = EXERCISE_LIBRARY.find(r => r[1] === entry.exercise);
  let tip = '';
  if (lib) {
    const topReps = Math.max(...entry.sets.map(s => Number(s.reps)));
    if (topReps >= lib[4]) tip = `Hit ${topReps} reps — top of range. Add weight next session.`;
    else if (topReps < lib[3]) tip = `Below target range (${lib[3]}-${lib[4]}). Same weight next time, chase reps.`;
    else tip = `In range (${lib[3]}-${lib[4]}). Add a rep or two before increasing weight.`;
  }
  return tip;
}

// ---------- WEB API ROUTES ----------
function doGet(e) {
  const action = e.parameter.action;
  let result;
  if (action === 'today') result = getTodayData(e.parameter.day || null);
  else if (action === 'library') result = getLibraryData();
  else if (action === 'progress') result = getProgressData();
  else result = { error: 'Unknown GET action: ' + action };
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}

// Frontend posts with Content-Type "text/plain" on purpose — it avoids a
// CORS preflight (OPTIONS) request, which Apps Script Web Apps don't
// handle. We just parse the JSON body ourselves below.
function doPost(e) {
  const body = JSON.parse(e.postData.contents);
  let result;
  if (body.action === 'submitSets') result = { tip: submitSets(body.entry) };
  else if (body.action === 'setExerciseEnabled') result = { ok: setExerciseEnabled(body.exercise, body.enabled) };
  else result = { error: 'Unknown POST action: ' + body.action };
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}
