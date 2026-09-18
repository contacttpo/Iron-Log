/**
 * IRON LOG — frontend data layer
 * Talks to the Apps Script Web App (Api.gs) instead of localStorage, so
 * data lives in your Google Sheet and is the same across every device.
 *
 * ⚠️ SET THIS after you deploy Api.gs as a Web App:
 */
const API_URL = 'PASTE_YOUR_APPS_SCRIPT_WEB_APP_URL_HERE';

async function apiGet_(action, params) {
  const qs = new URLSearchParams({ action, ...(params || {}) });
  const res = await fetch(`${API_URL}?${qs.toString()}`);
  if (!res.ok) throw new Error(`API GET ${action} failed: ${res.status}`);
  return res.json();
}

async function apiPost_(action, payload) {
  const res = await fetch(API_URL, {
    method: 'POST',
    // text/plain avoids a CORS preflight (Apps Script doesn't answer OPTIONS)
    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify({ action, ...(payload || {}) })
  });
  if (!res.ok) throw new Error(`API POST ${action} failed: ${res.status}`);
  return res.json();
}

async function getTodayData(overrideDay) {
  return apiGet_('today', overrideDay ? { day: overrideDay } : {});
}

async function getLibraryData() {
  return apiGet_('library');
}

async function getProgressData() {
  return apiGet_('progress');
}

async function submitSets(entry) {
  const result = await apiPost_('submitSets', { entry });
  return result.tip;
}

async function setExerciseEnabled(exerciseName, enabled) {
  const result = await apiPost_('setExerciseEnabled', { exercise: exerciseName, enabled });
  return result.ok;
}

window.App = {
  getTodayData,
  getLibraryData,
  getProgressData,
  submitSets,
  setExerciseEnabled
};
