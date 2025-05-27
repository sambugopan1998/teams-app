const msalConfig = {
  auth: {
    clientId: "0486fae2-afeb-4044-ab8d-0c060910b0a8",
    authority: "https://login.microsoftonline.com/c06fea01-72bf-415d-ac1d-ac0382f8d39f",
    redirectUri: "https://sambugopan1998.github.io/teams-app/hello.html"
  }
};

const msalInstance = new msal.PublicClientApplication(msalConfig);
const loginRequest = { scopes: ["User.Read", "Directory.Read.All"] };

// Utility to check if app is running inside Microsoft Teams iframe
function isRunningInTeams() {
  return window.self !== window.top && window.navigator.userAgent.includes("Teams");
}

// Acquire token based on environment
async function acquireAccessToken() {
  const accounts = msalInstance.getAllAccounts();

  if (accounts.length === 0) {
    if (isRunningInTeams()) {
      msalInstance.loginRedirect(loginRequest); // ✅ for Teams iframe
      return null;
    } else {
      const loginResponse = await msalInstance.loginPopup(loginRequest); // ✅ for browser
      msalInstance.setActiveAccount(loginResponse.account);
    }
  } else {
    msalInstance.setActiveAccount(accounts[0]);
  }

  const tokenResponse = await msalInstance.acquireTokenSilent({
    ...loginRequest,
    account: msalInstance.getActiveAccount()
  });

  return tokenResponse.accessToken;
}

// Main flow: Handle redirect, then fetch data
msalInstance.handleRedirectPromise().then(async (response) => {
  if (response) {
    msalInstance.setActiveAccount(response.account);
  }

  const token = await acquireAccessToken();
  if (!token) return; // redirect started, wait for return

  document.getElementById("access-token").textContent = token;

  const headers = { Authorization: `Bearer ${token}` };

  // 1. Profile
  const profileRes = await fetch("https://graph.microsoft.com/v1.0/me", { headers });
  const user = await profileRes.json();

  let html = "<h3>👤 Profile Info</h3><ul>";
  for (const [key, value] of Object.entries(user)) {
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
    console.warn("Photo not available.");
  }

  // 3. Roles
  const roleRes = await fetch("https://graph.microsoft.com/v1.0/me/memberOf", { headers });
  if (roleRes.ok) {
    const roles = await roleRes.json();
    let html = "<h3>🔐 Roles / Groups</h3><ul>";
    roles.value.forEach(entry => {
      html += `<li><strong>${entry["@odata.type"]}</strong>: ${entry.displayName || "Unnamed"}</li>`;
    });
    html += "</ul>";
    document.getElementById("user-info").insertAdjacentHTML("beforeend", html);
  }
}).catch(err => {
  console.error("Auth error:", err);
  document.getElementById("user-info").textContent =
    `❌ ${err.name} — ${err.message}`;
});
