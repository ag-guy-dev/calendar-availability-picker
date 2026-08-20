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

async function fetchBusyEvents(token, rangeStartISO, rangeEndISO) {
  const url =
    `${APP_CONFIG.graphBaseUrl}/me/calendarView` +
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
  return (data.value || []).filter(
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
  const pastEnd = now > dayStart ? (now < dayEnd ? now : dayEnd) : dayStart;

  const groupEl = document.createElement("div");
  groupEl.className = "day-group";
  const heading = document.createElement("h3");
  heading.textContent = dayLabel;
  groupEl.appendChild(heading);

  const timelineEl = document.createElement("div");
  timelineEl.className = "timeline";
  timelineEl.style.height = `${totalMinutes * PX_PER_MIN}px`;

  // Hour gridlines + labels.
  for (let m = 0; m <= totalMinutes; m += 60) {
    const hourDate = new Date(dayStart.getTime() + m * 60000);
    const markEl = document.createElement("div");
    markEl.className = "hour-mark";
    markEl.style.top = `${m * PX_PER_MIN}px`;
    const labelEl = document.createElement("span");
    labelEl.className = "hour-label";
    labelEl.textContent = formatTime(hourDate);
    markEl.appendChild(labelEl);
    timelineEl.appendChild(markEl);
  }

  // Past-time shading.
  if (pastEnd > dayStart) {
    const pastEl = document.createElement("div");
    pastEl.className = "past-block";
    pastEl.style.top = "0px";
    pastEl.style.height = `${minutesFrom(dayStart, pastEnd) * PX_PER_MIN}px`;
    timelineEl.appendChild(pastEl);
  }

  // Busy blocks.
  busyRanges.forEach((r) => {
    const busyEl = document.createElement("div");
    busyEl.className = "busy-block";
    busyEl.style.top = `${minutesFrom(dayStart, r.start) * PX_PER_MIN}px`;
    busyEl.style.height = `${minutesFrom(r.start, r.end) * PX_PER_MIN}px`;
    busyEl.title = r.subjects.join(", ");
    if (minutesFrom(r.start, r.end) * PX_PER_MIN >= 16) {
      busyEl.textContent = r.subjects.join(", ");
    }
    timelineEl.appendChild(busyEl);
  });

  // Free, selectable 30-min cells (clock-aligned to the day start).
  let cellStart = new Date(dayStart);
  while (cellStart < dayEnd) {
    const cellEnd = new Date(Math.min(cellStart.getTime() + CELL_MINUTES * 60000, dayEnd.getTime()));
    const isPast = cellEnd <= pastEnd;
    const isBusy = busyRanges.some((r) => overlaps(cellStart, cellEnd, r.start, r.end));

    if (!isPast && !isBusy) {
      const cellEl = document.createElement("div");
      cellEl.className = "free-cell";
      cellEl.style.top = `${minutesFrom(dayStart, cellStart) * PX_PER_MIN}px`;
      cellEl.style.height = `${minutesFrom(cellStart, cellEnd) * PX_PER_MIN}px`;
      cellEl.textContent = `${formatTime(cellStart)} – ${formatTime(cellEnd)}`;

      const cell = {
        start: new Date(cellStart),
        end: new Date(cellEnd),
        dayLabel,
        el: cellEl,
        selected: false,
      };
      cellEl.addEventListener("click", () => {
        cell.selected = !cell.selected;
        cellEl.classList.toggle("selected", cell.selected);
        updateInsertButton();
      });
      allCells.push(cell);
      timelineEl.appendChild(cellEl);
    }

    cellStart = cellEnd;
  }

  groupEl.appendChild(timelineEl);
  return groupEl;
}

async function loadAvailability() {
  clearError();
  const loadBtn = document.getElementById("load-btn");
  loadBtn.disabled = true;
  loadBtn.textContent = "Loading...";

  try {
    const lookaheadDays = parseInt(document.getElementById("range-select").value, 10);
    const [startHour, endHour] = document
      .getElementById("hours-select")
      .value.split("-")
      .map(Number);

    const token = await getGraphToken();
    const days = businessDayRange(lookaheadDays, startHour, endHour);

    if (days.length === 0) {
      showError("No business days found in range — try a longer lookahead.");
      return;
    }

    const rangeStartISO = days[0].dayStart.toISOString();
    const rangeEndISO = days[days.length - 1].dayEnd.toISOString();
    const busyEvents = await fetchBusyEvents(token, rangeStartISO, rangeEndISO);

    allCells = [];
    const listEl = document.getElementById("slots-list");
    listEl.innerHTML = "";

    days.forEach(({ dayStart, dayEnd }) => {
      const dayEvents = busyEvents.filter((e) => {
        const start = new Date(e.start.dateTime);
        const end = new Date(e.end.dateTime);
        return end > dayStart && start < dayEnd;
      });
      listEl.appendChild(renderDayTimeline(dayStart, dayEnd, dayEvents));
    });

    if (allCells.length === 0) {
      showError("No open time found in that range with those business hours.");
    }

    document.getElementById("slots-section").classList.remove("hidden");
    updateInsertButton();
  } catch (err) {
    showError("Couldn't load your calendar: " + err.message);
  } finally {
    loadBtn.disabled = false;
    loadBtn.textContent = "Show my calendar";
  }
}

function updateInsertButton() {
  const anyChecked = allCells.some((c) => c.selected);
  document.getElementById("insert-btn").disabled = !anyChecked;
}

// Merges adjacent selected 30-min cells (same day, back-to-back) into single ranges.
function mergedSelectedRanges() {
  const selected = allCells
    .filter((c) => c.selected)
    .sort((a, b) => a.start - b.start);

  const ranges = [];
  selected.forEach((c) => {
    const last = ranges[ranges.length - 1];
    if (last && last.dayLabel === c.dayLabel && last.end.getTime() === c.start.getTime()) {
      last.end = c.end;
    } else {
      ranges.push({ dayLabel: c.dayLabel, start: c.start, end: c.end });
    }
  });
  return ranges;
}

function insertIntoEmail() {
  clearError();
  const ranges = mergedSelectedRanges();
  if (ranges.length === 0) return;

  const byDay = {};
  ranges.forEach((r) => {
    if (!byDay[r.dayLabel]) byDay[r.dayLabel] = [];
    byDay[r.dayLabel].push(`${formatTime(r.start)} – ${formatTime(r.end)}`);
  });

  let text = "Here are some times that work on my end:\n\n";
  Object.keys(byDay).forEach((day) => {
    text += `${day}\n`;
    byDay[day].forEach((label) => { text += `  • ${label}\n`; });
    text += "\n";
  });
  text += "Let me know what fits best for you.\n";

  Office.context.mailbox.item.body.setSelectedDataAsync(
    text,
    { coercionType: Office.CoercionType.Text },
    (result) => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        showError("Couldn't insert text into the email: " + result.error.message);
      }
    }
  );
}
