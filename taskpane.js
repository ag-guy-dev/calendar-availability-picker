/* global Office, msal, APP_CONFIG */

let msalInstance;
let graphAccount;
let computedSlots = []; // { id, dayLabel, label, checked }

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

function computeFreeSlots(dayStart, dayEnd, busyEvents, durationMinutes) {
  const busyRanges = busyEvents
    .map((e) => ({ start: new Date(e.start.dateTime), end: new Date(e.end.dateTime) }))
    .filter((r) => r.end > dayStart && r.start < dayEnd)
    .sort((a, b) => a.start - b.start);

  const free = [];
  let cursor = new Date(Math.max(dayStart.getTime(), new Date().getTime()));

  for (const busy of busyRanges) {
    if (busy.start > cursor) addSlotsInGap(free, cursor, busy.start, durationMinutes);
    if (busy.end > cursor) cursor = busy.end;
  }
  if (cursor < dayEnd) addSlotsInGap(free, cursor, dayEnd, durationMinutes);
  return free;
}

function addSlotsInGap(free, gapStart, gapEnd, durationMinutes) {
  const durationMs = durationMinutes * 60 * 1000;
  let slotStart = new Date(gapStart);
  while (slotStart.getTime() + durationMs <= gapEnd.getTime()) {
    const slotEnd = new Date(slotStart.getTime() + durationMs);
    free.push({ start: new Date(slotStart), end: slotEnd });
    slotStart = slotEnd;
  }
}

function formatTime(d) {
  return d.toLocaleTimeString([], { hour: "numeric", minute: "2-digit" });
}

function formatDayLabel(d) {
  return d.toLocaleDateString([], { weekday: "long", month: "short", day: "numeric" });
}

async function loadAvailability() {
  clearError();
  const loadBtn = document.getElementById("load-btn");
  loadBtn.disabled = true;
  loadBtn.textContent = "Loading...";

  try {
    const lookaheadDays = parseInt(document.getElementById("range-select").value, 10);
    const durationMinutes = parseInt(document.getElementById("duration-select").value, 10);
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

    computedSlots = [];
    const listEl = document.getElementById("slots-list");
    listEl.innerHTML = "";

    days.forEach(({ dayStart, dayEnd }) => {
      const freeSlots = computeFreeSlots(dayStart, dayEnd, busyEvents, durationMinutes);
      if (freeSlots.length === 0) return;

      const groupEl = document.createElement("div");
      groupEl.className = "day-group";
      const heading = document.createElement("h3");
      heading.textContent = formatDayLabel(dayStart);
      groupEl.appendChild(heading);

      freeSlots.forEach((slot) => {
        const id = `slot-${computedSlots.length}`;
        const label = `${formatTime(slot.start)} – ${formatTime(slot.end)}`;
        computedSlots.push({ id, dayLabel: formatDayLabel(dayStart), label, checked: false });

        const row = document.createElement("div");
        row.className = "slot-item";
        row.innerHTML = `<input type="checkbox" id="${id}" /><label for="${id}">${label}</label>`;
        row.querySelector("input").addEventListener("change", (e) => {
          const s = computedSlots.find((s) => s.id === id);
          s.checked = e.target.checked;
          updateInsertButton();
        });
        groupEl.appendChild(row);
      });

      listEl.appendChild(groupEl);
    });

    if (computedSlots.length === 0) {
      showError("No open slots found in that range with those business hours.");
    }

    document.getElementById("slots-section").classList.remove("hidden");
    updateInsertButton();
  } catch (err) {
    showError("Couldn't load your calendar: " + err.message);
  } finally {
    loadBtn.disabled = false;
    loadBtn.textContent = "Find open times";
  }
}

function updateInsertButton() {
  const anyChecked = computedSlots.some((s) => s.checked);
  document.getElementById("insert-btn").disabled = !anyChecked;
}

function insertIntoEmail() {
  clearError();
  const chosen = computedSlots.filter((s) => s.checked);
  if (chosen.length === 0) return;

  const byDay = {};
  chosen.forEach((s) => {
    if (!byDay[s.dayLabel]) byDay[s.dayLabel] = [];
    byDay[s.dayLabel].push(s.label);
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
