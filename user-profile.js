// auth.js (type="module")
const msalConfig = {
  auth: {
    clientId: "0486fae2-afeb-4044-ab8d-0c060910b0a8",
    authority: "https://login.microsoftonline.com/c06fea01-72bf-415d-ac1d-ac0382f8d39f",
    redirectUri: "https://sambugopan1998.github.io/teams-app/hello.html", // Must be Teams app registered URL
  },
};
const loginScopes = ["User.Read", "Directory.Read.All"];

const state = {
  msalInstance: new msal.PublicClientApplication(msalConfig),
  accessToken: "",
};

function log(msg, isError = false) {
  const el = document.getElementById("log-output");
  const p = document.createElement("p");
  p.textContent = msg;
  p.style.color = isError ? "red" : "green";
  el.appendChild(p);
}

function isInTeams() {
  return typeof microsoftTeams !== "undefined";
}

async function waitForTeamsInit() {
  return new Promise((resolve) => {
    microsoftTeams.app.initialize().then(() => {
      log("✅ Teams SDK initialized");
      resolve(true);
    }).catch((err) => {
      log("❌ Teams SDK init failed: " + err.message, true);
      resolve(false);
    });
  });
}

async function loginFlow() {
  await state.msalInstance.initialize();
  const accounts = state.msalInstance.getAllAccounts();

  if (accounts.length === 0) {
    if (isInTeams()) {
      await waitForTeamsInit();
      log("🔁 Logging in with redirect (Teams)");
      state.msalInstance.loginRedirect({ scopes: loginScopes });
    } else {
      log("🔁 Logging in with popup (browser)");
      const loginResponse = await state.msalInstance.loginPopup({ scopes: loginScopes });
      state.msalInstance.setActiveAccount(loginResponse.account);
    }
  } else {
    state.msalInstance.setActiveAccount(accounts[0]);
    log("✅ Already signed in");
  }

  try {
    const tokenResponse = await state.msalInstance.acquireTokenSilent({
      scopes: loginScopes,
      account: state.msalInstance.getActiveAccount(),
    });
    state.accessToken = tokenResponse.accessToken;
    document.getElementById("access-token").textContent = tokenResponse.accessToken;
    log("✅ Token acquired");
    await fetchGraphData(tokenResponse.accessToken);
  } catch (err) {
    log("❌ Silent token error: " + err.message, true);
  }
}

async function fetchGraphData(token) {
  const headers = { Authorization: `Bearer ${token}` };

  // Profile
  const res = await fetch("https://graph.microsoft.com/v1.0/me", { headers });
  const profile = await res.json();
  const info = document.getElementById("user-info");
  info.innerHTML = `<h3>👤 Profile</h3><pre>${JSON.stringify(profile, null, 2)}</pre>`;
}

// Handle redirect
state.msalInstance.handleRedirectPromise().then(async (response) => {
  if (response && response.account) {
    state.msalInstance.setActiveAccount(response.account);
    log("✅ Redirect login complete");
  }
  await loginFlow();
}).catch((err) => {
  log("❌ Redirect error: " + err.message, true);
});
