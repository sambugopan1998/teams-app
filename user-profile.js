// app.js
const msalConfig = {
  auth: {
    clientId: "0486fae2-afeb-4044-ab8d-0c060910b0a8",
    authority: "https://login.microsoftonline.com/c06fea01-72bf-415d-ac1d-ac0382f8d39f",
    redirectUri: "https://sambugopan1998.github.io/teams-app/hello.html"
  }
};

const msalInstance = new msal.PublicClientApplication(msalConfig);
const loginRequest = { scopes: ["User.Read", "Directory.Read.All"] };

// Utility to check if running inside Teams iframe
function isRunningInTeams() {
  return window.self !== window.top || navigator.userAgent.includes("Teams");
}

async function acquireAccessToken() {
  const accounts = msalInstance.getAllAccounts();
  if (accounts.length === 0) {
    if (isRunningInTeams()) {
      msalInstance.loginRedirect(loginRequest); // Teams: use redirect
      return null; // stop execution
    } else {
      const loginResponse = await msalInstance.loginPopup(loginRequest); // Browser: use popup
      msalInstance.setActiveAccount(loginResponse.account);
    }
  } else {
    msalInstance.setActiveAccount(accounts[0]);
  }

  try {
    const tokenResponse = await msalInstance.acquireTokenSilent({
      ...loginRequest,
      account: msalInstance.getActiveAccount(),
    });
    return tokenResponse.accessToken;
  } catch (error) {
    // fallback to popup if needed (e.g., consent not yet given)
    if (!isRunningInTeams()) {
      const tokenResponse = await msalInstance.acquireTokenPopup(loginRequest);
      return tokenResponse.accessToken;
    } else {
      document.getElementById("user-info").innerHTML =
        `❌ Teams Tab Error: ${error.name} — ${error.message}`;
      throw error;
    }
  }
}

// MSAL redirect flow + data load
msalInstance.handleRedirectPromise().then(async (response) => {
  if (response && response.account) {
    msalInstance.setActiveAccount(response.account);
  }

  const token = await acquireAccessToken();
  if (!token) return;

  document.getElementById("access-token").textContent = token;
  const headers = { Authorization: `Bearer ${token}` };

  // 1. Get profile
  const profileRes = await fetch("https://graph.microsoft.com/v1.0/me", { headers });
  const user = await profileRes.json();

  let html = "<h3>👤 Profile Info</h3><ul>";
  for (const [key, value] of Object.entries(user)) {
    html += `<li><strong>${key}</strong>: ${value ?? "N/A"}</li>`;
  }
  html += "</ul>";
  document.getElementById("user-info").innerHTML = html;

  // 2. Get photo
  try {
    const photoRes = await fetch("https://graph.microsoft.com/v1.0/me/photo/$value", { headers });
    if (photoRes.ok) {
      const blob = await photoRes.blob();
      const url = URL.createObjectURL(blob);
      document.getElementById("user-info").insertAdjacentHTML("afterbegin",
        `<h3>🖼️ Photo</h3><img src="${url}" style="height:100px;border-radius:50%;">`);
    }
  } catch (err) {
    console.warn("No profile photo.");
  }

  // 3. Get roles / groups
  const groupRes = await fetch("https://graph.microsoft.com/v1.0/me/memberOf", { headers });
  if (groupRes.ok) {
    const groups = await groupRes.json();
    let rolesHTML = "<h3>🔐 Roles / Groups</h3><ul>";
    groups.value.forEach(g => {
      rolesHTML += `<li><strong>${g['@odata.type']}</strong>: ${g.displayName || "Unnamed"}</li>`;
    });
    rolesHTML += "</ul>";
    document.getElementById("user-info").insertAdjacentHTML("beforeend", rolesHTML);
  }
}).catch(err => {
  console.error("Auth flow failed", err);
  document.getElementById("user-info").innerHTML =
    `❌ Auth Error: ${err.name} — ${err.message}`;
});
