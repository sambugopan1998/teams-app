// Include MSAL in your HTML:
// <script src="https://alcdn.msauth.net/browser/2.37.0/js/msal-browser.min.js"></script>
// <script src="auth.js" type="module"></script>

const msalConfig = {
  auth: {
    clientId: "0486fae2-afeb-4044-ab8d-0c060910b0a8",
    authority: "https://login.microsoftonline.com/c06fea01-72bf-415d-ac1d-ac0382f8d39f",
    redirectUri: "https://sambugopan1998.github.io/teams-app/hello.html",
  },
};

const loginScopes = ["User.Read", "Directory.Read.All"];

const state = {
  msalInstance: new msal.PublicClientApplication(msalConfig),
  accessToken: "",
};

// Utility: log messages to a div on the page
function logToPage(message, isError = false) {
  const el = document.getElementById("log-output");
  const p = document.createElement("p");
  p.textContent = message;
  p.style.color = isError ? "red" : "green";
  el.appendChild(p);
}

// Check if inside Microsoft Teams app
function isRunningInTeams() {
  return typeof microsoftTeams !== "undefined";
}

// Wait for Teams SDK init
async function waitForTeamsInit() {
  return new Promise((resolve) => {
    microsoftTeams.app.initialize().then(() => {
      logToPage("✅ Microsoft Teams SDK initialized");
      resolve(true);
    }).catch((e) => {
      logToPage("❌ Teams SDK init failed: " + e.message, true);
      resolve(false);
    });
  });
}

// MSAL login and token
async function initializeMsalAndLogin() {
  await state.msalInstance.initialize();

  const accounts = state.msalInstance.getAllAccounts();

  if (accounts.length === 0) {
    if (isRunningInTeams()) {
      await waitForTeamsInit();
      logToPage("🔁 Logging in using redirect (Teams)...");
      state.msalInstance.loginRedirect({ scopes: loginScopes });
      return null; // Let redirect happen
    } else {
      logToPage("🔁 Logging in using popup (browser)...");
      const loginResponse = await state.msalInstance.loginPopup({ scopes: loginScopes });
      state.msalInstance.setActiveAccount(loginResponse.account);
    }
  } else {
    state.msalInstance.setActiveAccount(accounts[0]);
    logToPage("✅ User already signed in.");
  }

  try {
    const tokenResponse = await state.msalInstance.acquireTokenSilent({
      scopes: loginScopes,
      account: state.msalInstance.getActiveAccount(),
    });
    state.accessToken = tokenResponse.accessToken;
    logToPage("✅ Access token acquired silently.");
    return tokenResponse.accessToken;
  } catch (error) {
    logToPage("⚠️ Silent token failed: " + error.message, true);

    if (!isRunningInTeams()) {
      try {
        const popupToken = await state.msalInstance.acquireTokenPopup({ scopes: loginScopes });
        state.accessToken = popupToken.accessToken;
        logToPage("✅ Token acquired via popup.");
        return popupToken.accessToken;
      } catch (popupError) {
        logToPage("❌ Token popup failed: " + popupError.message, true);
        return null;
      }
    } else {
      logToPage("🔒 Token acquisition in Teams requires redirect. Try reloading.", true);
      return null;
    }
  }
}

// Call Microsoft Graph API
async function fetchGraphData(token) {
  const headers = { Authorization: `Bearer ${token}` };

  // 1. Get profile
  try {
    const profileRes = await fetch("https://graph.microsoft.com/v1.0/me", { headers });
    const profile = await profileRes.json();

    let html = "<h3>👤 Profile Info</h3><ul>";
    for (const [key, value] of Object.entries(profile)) {
      html += `<li><strong>${key}</strong>: ${value ?? "N/A"}</li>`;
    }
    html += "</ul>";
    document.getElementById("user-info").innerHTML = html;
  } catch (err) {
    logToPage("❌ Failed to load profile: " + err.message, true);
  }

  // 2. Photo
  try {
    const photoRes = await fetch("https://graph.microsoft.com/v1.0/me/photo/$value", { headers });
    if (photoRes.ok) {
      const blob = await photoRes.blob();
      const imgURL = URL.createObjectURL(blob);
      const imgTag = `<h3>🖼️ Photo</h3><img src="${imgURL}" style="height:100px;border-radius:50%">`;
      document.getElementById("user-info").insertAdjacentHTML("afterbegin", imgTag);
    } else {
      logToPage("⚠️ Photo not available.", true);
    }
  } catch (err) {
    logToPage("❌ Failed to fetch photo: " + err.message, true);
  }

  // 3. Roles / Groups
  try {
    const rolesRes = await fetch("https://graph.microsoft.com/v1.0/me/memberOf", { headers });
    if (rolesRes.ok) {
      const roles = await rolesRes.json();
      let rolesHtml = "<h3>🔐 Roles / Groups</h3><ul>";
      roles.value.forEach((entry) => {
        rolesHtml += `<li><strong>${entry["@odata.type"]}</strong>: ${entry.displayName}</li>`;
      });
      rolesHtml += "</ul>";
      document.getElementById("user-info").insertAdjacentHTML("beforeend", rolesHtml);
    } else {
      logToPage("⚠️ Unable to fetch roles/groups.", true);
    }
  } catch (err) {
    logToPage("❌ Roles fetch error: " + err.message, true);
  }
}

// Entry point
state.msalInstance.handleRedirectPromise().then(async (response) => {
  if (response && response.account) {
    state.msalInstance.setActiveAccount(response.account);
    logToPage("✅ Redirect login successful.");
  }

  const token = await initializeMsalAndLogin();
  if (token) {
    document.getElementById("access-token").textContent = token;
    await fetchGraphData(token);
  }
}).catch((err) => {
  logToPage("❌ Auth error: " + err.message, true);
});
