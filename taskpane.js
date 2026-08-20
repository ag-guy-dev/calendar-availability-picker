/* global Office, msal, APP_CONFIG */

let msalInstance;
let graphAccount;

const PX_PER_MIN = 1.5;
const CELL_MINUTES = 30;

let allCells = []; // { start: Date, end: Date, dayLabel, el, selected }

Office.onReady(() => {
  msalInstance = new msal.PublicClientApplication({
    auth: {
      clientId: APP_CONFIG.clientId,
      authority: APP_CONFIG.authority,
      redirectUri: APP_CONFIG.redirectUri,
    },
    cache: { cacheLocation: "sessionStorage" },
  });

  document.getElementById("signin-btn").addEventListener("click", signIn);
  document.getElementById("load-btn").addEventListener("click", loadAvailability);
  document.getElementById("insert-btn").addEventListener("click", insertIntoEmail);
});

function showError(message) {
  const el = document.getElementById("error-msg");
  el.textContent = message;
  el.classList.remove("hidden");
}

function clearError() {
  document.getElementById("error-msg").classList.add("hidden");
}

async function signIn() {
  clearError();
  try {
    const result = await msalInstance.loginPopup({ scopes: APP_CONFIG.graphScopes });
    graphAccount = result.account;
    document.getElementById("signin-status").textContent = `Signed in as ${graphAccount.username}`;
    document.getElementById("signin-btn").classList.add("hidden");
    document.getElementById("controls-section").classList.remove("hidden");
  } catch (err) {
    showError("Sign-in failed: " + err.message);
  }
}

async function getGraphToken() {
  try {
    const result = await msalInstance.acquireTokenSilent({
      scopes: APP_CONFIG.graphScopes,
      account: graphAccount,
    });
    return result.accessToken;
  } catch (err) {
    const result = await msalInstance.acquireTokenPopup({ scopes: APP_CONFIG.graphScopes });
    return result.accessToken;
  }
}

function businessDayRange(lookaheadDays, startHour, endHour) {
  const days = [];
  let cursor = new Date();
  let added = 0;
  let guard = 0;
  while (added < lookaheadDays && guard < lookaheadDays * 3) {
    guard++;
    const dow = cursor.getDay();
    if (dow !== 0 && dow !== 6) {
      const dayStart = new Date(cursor);
      dayStart.setHours(startHour, 0, 0, 0);
      const dayEnd = new Date(cursor);
      dayEnd.setHours(endHour, 0, 0, 0);
      if (dayEnd > new Date()) {
        days.push({ dayStart, dayEnd });
        added++;
      }
    }
    cursor = new Date(cursor.getTime() + 24 * 60 * 60 * 1000);
  }
  return days;
}

async function fetchCalendarIds(token) {
  const resp = await fetch(`${APP_CONFIG.graphBaseUrl}/me/calendars?$select=id,name`, {
    headers: { Authorization: `Bearer ${token}` },
  });

  if (!resp.ok) {
    const text = await resp.text();
    throw new Error(`Graph error ${resp.status}: ${text}`);
  }

  const data = await resp.json();
  return (data.value || []).map((c) => c.id);
}

async function fetchCalendarViewEvents(token, calendarId, rangeStartISO, rangeEndISO) {
  const url =
    `${APP_CONFIG.graphBaseUrl}/me/calendars/${calendarId}/calendarView` +
    `?startDateTime=${encodeURIComponent(rangeStartISO)}` +
    `&endDateTime=${encodeURIComponent(rangeEndISO)}` +
    `&$select=subject,start,end,showAs,isAllDay` +
    `&$top=100`;

  const resp = await fetch(url, {
    headers: {
      Authorization: `Bearer ${token}`,
      Prefer: 'outlook.timezone="' + Intl.DateTimeFormat().resolvedOptions().timeZone + '"',
    },
  });

  if (!resp.ok) {
    const text = await resp.text();
    throw new Error(`Graph error ${resp.status}: ${text}`);
  }

  const data = await resp.json();
  return data.value || [];
}

// Pulls busy time from every calendar the signed-in account has (not just the
// default one), so events synced in from a connected non-Microsoft account
// (e.g. a linked Gmail calendar) are included if Graph exposes them at all.
async function fetchBusyEvents(token, rangeStartISO, rangeEndISO) {
  const calendarIds = await fetchCalendarIds(token);

  const eventsPerCalendar = await Promise.all(
    calendarIds.map((id) => fetchCalendarViewEvents(token, id, rangeStartISO, rangeEndISO))
  );

  return eventsPerCalendar.flat().filter(
    (e) => !e.isAllDay && ["busy", "tentative", "oof"].includes((e.showAs || "").toLowerCase())
  );
}

// Merges overlapping/adjacent busy events (clipped to the day's business hours)
// into contiguous blocks, so back-to-back meetings render as one shape.
function mergeBusyRanges(events, dayStart, dayEnd) {
  const ranges = events
    .map((e) => ({
      start: new Date(e.start.dateTime),
      end: new Date(e.end.dateTime),
      subject: e.subject || "Busy",
    }))
    .filter((r) => r.end > dayStart && r.start < dayEnd)
    .map((r) => ({
      start: r.start < dayStart ? dayStart : r.start,
      end: r.end > dayEnd ? dayEnd : r.end,
      subjects: [r.subject],
    }))
    .sort((a, b) => a.start - b.start);

  const merged = [];
  for (const r of ranges) {
    const last = merged[merged.length - 1];
    if (last && r.start <= last.end) {
      if (r.end > last.end) last.end = r.end;
      if (!last.subjects.includes(r.subjects[0])) last.subjects.push(r.subjects[0]);
    } else {
      merged.push(r);
    }
  }
  return merged;
}

function minutesFrom(dayStart, date) {
  return (date.getTime() - dayStart.getTime()) / 60000;
}

function formatTime(d) {
  return d.toLocaleTimeString([], { hour: "numeric", minute: "2-digit" });
}

function formatDayLabel(d) {
  return d.toLocaleDateString([], { weekday: "long", month: "short", day: "numeric" });
}

function overlaps(aStart, aEnd, bStart, bEnd) {
  return aStart < bEnd && aEnd > bStart;
}

function renderDayTimeline(dayStart, dayEnd, busyEvents) {
  const dayLabel = formatDayLabel(dayStart);
  const totalMinutes = minutesFrom(dayStart, dayEnd);
  const busyRanges = mergeBusyRanges(busyEvents, dayStart, dayEnd);
  const now = new Date();
  const
