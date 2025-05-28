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

function isRunningInTeams() {
  return typeof microsoftTeams !== "undefined";
}

async function waitForTeamsInit() {
  return new Promise((resolve) => {
    microsoftTeams.app.initialize().then(() => {
      console.log("✅ Microsoft Teams SDK initialized");
      resolve(true);
    });
  });
}

async function initializeMsalAndLogin() {
  await state.msalInstance.initialize();

  const accounts = state.msalInstance.getAllAccounts();

  if (accounts.length === 0) {
    if (isRunningInTeams()) {
      await waitForTeamsInit();
      state.msalInstance.loginRedirect({ scopes: loginScopes });
      return null; // Let redirect happen
    } else {
      const loginResponse = await state.msalInstance.loginPopup({ scopes: loginScopes });
      state.msalInstance.setActiveAccount(loginResponse.account);
    }
  } else {
    state.msalInstance.setActiveAccount(accounts[0]);
  }

  try {
    const tokenResponse = await state.msalInstance.acquireTokenSilent({
      scopes: loginScopes,
      account: state.msalInstance.getActiveAccount(),
    });
    state.accessToken = tokenResponse.accessToken;
    return tokenResponse.accessToken;
  } catch (error) {
    if (!isRunningInTeams()) {
      // Fallback for browser only
      const popupToken = await state.msalInstance.acquireTokenPopup({ scopes: loginScopes });
      state.accessToken = popupToken.accessToken;
      return popupToken.accessToken;
    } else {
      console.error("Silent token failed in Teams. Re-login required.");
      return null;
    }
  }
}

async function fetchGraphData(token) {
  const headers = { Authorization: `Bearer ${token}` };

  // 1. Get profile
  const profileRes = await fetch("https://graph.microsoft.com/v1.0/me", { headers });
  const profile = await profileRes.json();
  let html = "<h3>👤 Profile Info</h3><ul>";
  for (const [key, value] of Object.entries(profile)) {
    html += `<li><strong>${key}</strong>: ${value ?? "N/A"}</li>`;
  }
  html += "</ul>";
  document.getElementById("user-info").innerHTML = html;

  // 2. Photo
  try {
    const photoRes = await fetch("https://graph.microsoft.com/v1.0/me/photo/$value", { headers });
    if (photoRes.ok) {
      const blob = await photoRes.blob();
      const imgURL = URL.createObjectURL(blob);
      const imgTag = `<h3>🖼️ Photo</h3><img src="${imgURL}" style="height:100px;border-radius:50%">`;
      document.getElementById("user-info").insertAdjacentHTML("afterbegin", imgTag);
    }
  } catch (err) {
    console.warn("No photo found.");
  }

  // 3. Roles / Groups
  const rolesRes = await fetch("https://graph.microsoft.com/v1.0/me/memberOf", { headers });
  if (rolesRes.ok) {
    const roles = await rolesRes.json();
    let rolesHtml = "<h3>🔐 Roles / Groups</h3><ul>";
    roles.value.forEach((entry) => {
      rolesHtml += `<li><strong>${entry["@odata.type"]}</strong>: ${entry.displayName}</li>`;
    });
    rolesHtml += "</ul>";
    document.getElementById("user-info").insertAdjacentHTML("beforeend", rolesHtml);
  }
}

// Entry point
state.msalInstance.handleRedirectPromise().then(async (response) => {
  if (response && response.account) {
    state.msalInstance.setActiveAccount(response.account);
  }

  const token = await initializeMsalAndLogin();
  if (token) {
    document.getElementById("access-token").textContent = token;
    await fetchGraphData(token);
  }
}).catch((err) => {
  console.error("Auth error:", err);
  document.getElementById("user-info").textContent = `❌ ${err.name}: ${err.message}`;
});