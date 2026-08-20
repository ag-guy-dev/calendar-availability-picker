// ---------------------------------------------------------------------------
// FILL THESE IN after registering the Azure AD app (see README.md, Step 1).
// This version is set up for a PERSONAL Microsoft account (outlook.com,
// hotmail.com, live.com) — not a work/school tenant, and no admin needed.
// ---------------------------------------------------------------------------
const APP_CONFIG = {
  // Azure AD "Application (client) ID" from the app registration.
  clientId: "d6c03c94-2795-4c52-8593-74690446ebed",

  // "consumers" = personal Microsoft accounts only. Use "common" instead
  // if you might ever want to sign in with a work/school account too.
  authority: "https://login.microsoftonline.com/consumers",

  // Must exactly match a Redirect URI (type: Single-page application)
  // configured on the Azure AD app registration.
  redirectUri: "https://ag-guy-dev.github.io/calendar-availability-picker/taskpane.html",

  // Delegated Graph permission needed to read free/busy + events.
  graphScopes: ["Calendars.Read"],

  graphBaseUrl: "https://graph.microsoft.com/v1.0",
};
