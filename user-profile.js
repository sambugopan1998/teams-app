// Include this script in an HTML page that loads MSAL first:
// <script src="https://alcdn.msauth.net/browser/2.37.0/js/msal-browser.min.js"></script>
// <script src="auth.js" type="module"></script>

const msalConfig = {
  auth: {
    clientId: "0486fae2-afeb-4044-ab8d-0c060910b0a8", // Your Azure AD App ID
    authority: "https://login.microsoftonline.com/c06fea01-72bf-415d-ac1d-ac0382f8d39f", // Your tenant ID
    redirectUri: "https://sambugopan1998.github.io/teams-app/hello.html", // Your redirect URI
  },
};

const loginScopes = ["User.Read", "Directory.Read.All"];

const state = {
  msalInstance: null,
  accessToken: "",
};

function isRunningInTeams() {
  return window.self !== window.top && window.navigator.userAgent.includes("Teams");
}

async function initializeMsalInstance() {
  if (!state.msalInstance) {
    const instance = new msal.PublicClientApplication(msalConfig);
    state.msalInstance = instance;
    try {
      await instance.initialize();
      console.log("✅ MSAL instance initialized.");
    } catch (err) {
      console.error("❌ MSAL initialization failed:", err);
    }
  }
}

async function signInAndGetToken() {
  await initializeMsalInstance();
  try {
    const accounts = state.msalInstance.getAllAccounts();
    if (accounts.length === 0) {
      if (isRunningInTeams()) {
        state.msalInstance.loginRedirect({ scopes: loginScopes });
        return null; // redirect in progress
      } else {
        const loginResponse = await state.msalInstance.loginPopup({ scopes: loginScopes });
        state.msalInstance.setActiveAccount(loginResponse.account);
      }
    } else {
      state.msalInstance.setActiveAccount(accounts[0]);
    }

    const tokenResponse = await state.msalInstance.acquireTokenSilent({
      scopes: loginScopes,
      account: state.msalInstance.getActiveAccount(),
    });

    state.accessToken = tokenResponse.accessToken;
    return tokenResponse.accessToken;
  } catch (error) {
    console.warn("Silent token failed, trying popup:", error);
    try {
      const tokenResponse = await state.msalInstance.acquireTokenPopup({ scopes: loginScopes });
      state.accessToken = tokenResponse.accessToken;
      return tokenResponse.accessToken;
    } catch (popupError) {
      console.error("Popup failed:", popupError);
      return null;
    }
  }
}

async function fetchGraphData(token) {
  const headers = { Authorization: `Bearer ${token}` };

  // 1. Profile Info
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
    console.warn("No photo available.");
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

// Handle redirect if applicable and then run logic
state.msalInstance = new msal.PublicClientApplication(msalConfig);
state.msalInstance.handleRedirectPromise().then(async (response) => {
  if (response) {
    state.msalInstance.setActiveAccount(response.account);
  }

  const token = await signInAndGetToken();
  if (token) {
    document.getElementById("access-token").textContent = token;
    await fetchGraphData(token);
  }
});
