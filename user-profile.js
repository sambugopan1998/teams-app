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

function logToPage(message, isError = false) {
  const el = document.getElementById("log-output");
  const p = document.createElement("p");
  p.textContent = message;
  p.style.color = isError ? "red" : "green";
  el.appendChild(p);
}

function isRunningInTeams() {
  return typeof microsoftTeams !== "undefined";
}

async function waitForTeamsInit() {
  return new Promise((resolve) => {
    microsoftTeams.app.initialize().then(() => {
      logToPage("✅ Teams SDK initialized");
      resolve(true);
    }).catch((e) => {
      logToPage("❌ Teams SDK init failed: " + e.message, true);
      resolve(false);
    });
  });
}

async function authenticate() {
  await state.msalInstance.initialize();
  const accounts = state.msalInstance.getAllAccounts();

  if (accounts.length === 0) {
    if (isRunningInTeams()) {
      await waitForTeamsInit();
      logToPage("🔁 Login redirect (Teams)");
      return state.msalInstance.loginRedirect({ scopes: loginScopes });
    } else {
      try {
        logToPage("🔁 Login popup (Browser)");
        const loginResp = await state.msalInstance.loginPopup({ scopes: loginScopes });
        state.msalInstance.setActiveAccount(loginResp.account);
      } catch (err) {
        logToPage("❌ Login popup failed: " + err.message, true);
        return null;
      }
    }
  } else {
    state.msalInstance.setActiveAccount(accounts[0]);
    logToPage("✅ Existing session");
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

    if (isRunningInTeams()) {
      logToPage("🔁 Token redirect (Teams)");
      return state.msalInstance.acquireTokenRedirect({ scopes: loginScopes });
    } else {
      try {
        const popupResp = await state.msalInstance.acquireTokenPopup({ scopes: loginScopes });
        state.accessToken = popupResp.accessToken;
        logToPage("✅ Token acquired via popup");
        return popupResp.accessToken;
      } catch (popupErr) {
        logToPage("❌ Token popup failed: " + popupErr.message, true);
        return null;
      }
    }
  }
}

async function fetchGraphData(token) {
  const headers = { Authorization: `Bearer ${token}` };

  // 1. Profile
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
    logToPage("❌ Profile fetch failed: " + err.message, true);
  }

  // 2. Photo
  try {
    const photoRes = await fetch("https://graph.microsoft.com/v1.0/me/photo/$value", { headers });
    if (photoRes.ok) {
      const blob = await photoRes.blob();
      const imgURL = URL.createObjectURL(blob);
      document.getElementById("user-info").insertAdjacentHTML("afterbegin",
        `<h3>🖼️ Photo</h3><img src="${imgURL}" style="height:100px;border-radius:50%">`
      );
    } else {
      logToPage("⚠️ No photo found");
    }
  } catch (err) {
    logToPage("❌ Photo fetch failed: " + err.message, true);
  }

  // 3. Groups / Roles
  try {
    const groupRes = await fetch("https://graph.microsoft.com/v1.0/me/memberOf", { headers });
    if (groupRes.ok) {
      const groups = await groupRes.json();
      let rolesHTML = "<h3>🔐 Roles / Groups</h3><ul>";
      groups.value.forEach(g => {
        rolesHTML += `<li><strong>${g['@odata.type']}</strong>: ${g.displayName || "Unnamed"}</li>`;
      });
      rolesHTML += "</ul>";
      document.getElementById("user-info").insertAdjacentHTML("beforeend", rolesHTML);
    } else {
      logToPage("⚠️ Group fetch failed");
    }
  } catch (err) {
    logToPage("❌ Roles fetch failed: " + err.message, true);
  }
}

state.msalInstance.handleRedirectPromise().then(async (response) => {
  if (response && response.account) {
    state.msalInstance.setActiveAccount(response.account);
    logToPage("✅ Redirect login success");
  }

  const token = await authenticate();
  if (token) {
    document.getElementById("access-token").textContent = token;
    await fetchGraphData(token);
  }
}).catch((err) => {
  document.getElementById("user-info").innerHTML = 
    `<p style="color:red;">❌ Auth Error: ${err.name} — ${err.message}</p>`;
});
