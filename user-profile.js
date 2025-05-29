import * as msal from "https://cdn.jsdelivr.net/npm/@azure/msal-browser@3.11.0/+esm";

// Replace with your actual values
const msalConfig = {
  auth: {
    clientId: "0486fae2-afeb-4044-ab8d-0c060910b0a8",
    authority: "https://login.microsoftonline.com/c06fea01-72bf-415d-ac1d-ac0382f8d39f",
    redirectUri: 'https://sambugopan1998.github.io/teams-app/hello.html',
  }
};

// Scopes
const apiScopes = ["api://40306a83-ec51-4935-b95e-485d3804873c/read"]; // App B: custom API
const graphScopes = ["User.Read"]; // Microsoft Graph

const msalInstance = new msal.PublicClientApplication(msalConfig);
const app = document.querySelector(".app");

function render(content) {
  app.innerHTML = content;
}

function renderError(error) {
  const errMsg = error?.message || JSON.stringify(error) || "Unknown error";
  app.innerHTML += `
    <div style="color: red;">
      <h4>Error:</h4>
      <pre>${errMsg}</pre>
    </div>`;
}

function signInButton() {
  render(`<button id="signin">🔐 Sign in with Microsoft</button>`);
  document.getElementById("signin").onclick = signIn;
}

async function signIn() {
  try {
    const loginResponse = await msalInstance.loginPopup({ scopes: apiScopes });
    msalInstance.setActiveAccount(loginResponse.account);

    // Token for App B
    const apiTokenResponse = await msalInstance.acquireTokenSilent({
      scopes: apiScopes,
      account: loginResponse.account
    });
    console.log("✅ Token for App B:", apiTokenResponse.accessToken);

    // Token for Microsoft Graph
    const graphTokenResponse = await msalInstance.acquireTokenSilent({
      scopes: graphScopes,
      account: loginResponse.account
    });
    console.log("✅ Microsoft Graph Token:", graphTokenResponse.accessToken);

    // Call Microsoft Graph
    await fetchGraphData(graphTokenResponse.accessToken);

  } catch (err) {
    renderError(err);
  }
}

async function fetchGraphData(token) {
  try {
    const headers = { Authorization: `Bearer ${token}` };
    const [profileRes, groupsRes] = await Promise.all([
      fetch("https://graph.microsoft.com/v1.0/me", { headers }),
      fetch("https://graph.microsoft.com/v1.0/me/memberOf", { headers }),
    ]);

    const profile = await profileRes.json();
    const groups = await groupsRes.json();

    let userDetails = "<h3>👤 User Profile</h3><ul>";
    for (const [key, value] of Object.entries(profile)) {
      userDetails += `<li><strong>${key}</strong>: ${value || "N/A"}</li>`;
    }
    userDetails += "</ul>";

    let groupHTML = "<h3>🔐 Groups</h3><ul>";
    (groups.value || []).forEach(g => {
      groupHTML += `<li>${g.displayName || g.id}</li>`;
    });
    groupHTML += "</ul>";

    render(`${userDetails}${groupHTML}<h3>🎯 Done</h3>`);
  } catch (err) {
    renderError(err);
  }
}

msalInstance.handleRedirectPromise().then(async (response) => {
  if (response && response.account) {
    msalInstance.setActiveAccount(response.account);
  }
  const accounts = msalInstance.getAllAccounts();
  if (accounts.length > 0) {
    msalInstance.setActiveAccount(accounts[0]);
    signIn(); // auto continue if already logged in
  } else {
    signInButton();
  }
}).catch(err => renderError(err));
