const msalConfig = {
  auth: {
    clientId: "0486fae2-afeb-4044-ab8d-0c060910b0a8",
    authority: "https://login.microsoftonline.com/c06fea01-72bf-415d-ac1d-ac0382f8d39f",
    redirectUri: "https://sambugopan1998.github.io/teams-app/hello.html"
  }
};

const loginScopes = ["User.Read", "Directory.Read.All"];
const state = {
  msalInstance: new msal.PublicClientApplication(msalConfig),
  accessToken: ""
};

// Check if running in Teams
function isRunningInTeams() {
  return typeof microsoftTeams !== "undefined";
}

// Log to screen
function logToPage(msg, isError = false) {
  const div = document.getElementById("log-output");
  const p = document.createElement("p");
  p.textContent = msg;
  p.style.color = isError ? "red" : "green";
  div.appendChild(p);
}

// Wait for Teams SDK to be ready
async function waitForTeamsInit() {
  try {
    await microsoftTeams.app.initialize();
    logToPage("✅ Teams SDK initialized");
  } catch (e) {
    logToPage("❌ Teams SDK init failed: " + e.message, true);
  }
}

// MSAL + Teams login
async function authenticate() {
  await state.msalInstance.initialize();
  const accounts = state.msalInstance.getAllAccounts();

  if (accounts.length === 0) {
    if (isRunningInTeams()) {
      await waitForTeamsInit();
      logToPage("🔁 Logging in with redirect (Teams)");
      await state.msalInstance.loginRedirect({ scopes: loginScopes });
    } else {
      try {
        logToPage("🔁 Logging in with popup (Browser)");
        const loginResp = await state.msalInstance.loginPopup({ scopes: loginScopes });
        state.msalInstance.setActiveAccount(loginResp.account);
      } catch (e) {
        logToPage("❌ Login popup failed: " + e.message, true);
      }
    }
  } else {
    state.msalInstance.setActiveAccount(accounts[0]);
    logToPage("✅ User already signed in");
  }

  try {
    const tokenResp = await state.msalInstance.acquireTokenSilent({
      scopes: loginScopes,
      account: state.msalInstance.getActiveAccount()
    });
    state.accessToken = tokenResp.accessToken;
    logToPage("✅ Token acquired silently");
    return tokenResp.accessToken;
  } catch (e) {
    logToPage("⚠️ Silent token failed: " + e.message, true);

    if (!isRunningInTeams()) {
      try {
        const popupResp = await state.msalInstance.acquireTokenPopup({ scopes: loginScopes });
        state.accessToken = popupResp.accessToken;
        logToPage("✅ Token acquired via popup");
        return popupResp.accessToken;
      } catch (popupErr) {
        logToPage("❌ Token popup failed: " + popupErr.message, true);
        return null;
      }
    } else {
      logToPage("🔐 Teams requires redirect for token", true);
      await state.msalInstance.acquireTokenRedirect({ scopes: loginScopes });
    }
  }
}

// Fetch MS Graph data
async function fetchGraphData(token) {
  const headers = { Authorization: `Bearer ${token}` };

  try {
    const profileRes = await fetch("https://graph.microsoft.com/v1.0/me", { headers });
    const profile = await profileRes.json();
    let html = "<h3>👤 Profile</h3><ul>";
    for (const [k, v] of Object.entries(profile)) {
      html += `<li><b>${k}</b>: ${v}</li>`;
    }
    html += "</ul>";
    document.getElementById("user-info").innerHTML = html;
  } catch (err) {
    logToPage("❌ Profile fetch failed: " + err.message, true);
  }
}

// Entry point
state.msalInstance.handleRedirectPromise().then(async (response) => {
  if (response && response.account) {
    state.msalInstance.setActiveAccount(response.account);
    logToPage("✅ Redirect login complete");
  }

  const token = await authenticate();
  if (token) {
    document.getElementById("access-token").textContent = token;
    await fetchGraphData(token);
  }
}).catch((err) => {
  logToPage("❌ Redirect error: " + err.message, true);
});
