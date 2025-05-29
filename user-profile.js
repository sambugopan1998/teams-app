const msalConfig = {
  auth: {
    clientId: "0486fae2-afeb-4044-ab8d-0c060910b0a8",
    authority: "https://login.microsoftonline.com/c06fea01-72bf-415d-ac1d-ac0382f8d39f",
    redirectUri:'https://sambugopan1998.github.io/teams-app/hello.html',
  }
};

const scopes = ["User.Read", "Directory.Read.All"];
const msalInstance = new msal.PublicClientApplication(msalConfig);
const app = document.querySelector(".app");

function render(content) {
  app.innerHTML = content;
}

function renderError(error) {
  const errMsg = error?.message || JSON.stringify(error) || "Unknown error";
  app.innerHTML += `
    <div style="color: red;">
      <h4>❌ Error:</h4>
      <pre>${errMsg}</pre>
    </div>`;
}

function renderUser(user, groupsHTML, token) {
  let userDetails = "<h3>👤 User Profile</h3><ul>";
  for (const [key, value] of Object.entries(user)) {
    userDetails += `<li><strong>${key}</strong>: ${value || "N/A"}</li>`;
  }
  userDetails += "</ul>";

  app.innerHTML = `
    ${userDetails}
    <h3>🔐 Roles / Groups</h3>
    ${groupsHTML}
    <h3>🔑 Access Token</h3>
    <textarea readonly>${token}</textarea>
  `;
}

function signInButton() {
  render(`<button id="signin">🔐 Sign in with Microsoft</button>`);
  document.getElementById("signin").onclick = signIn;
}

async function signIn() {
  try {
    await microsoftTeams.app.initialize();

    const isInIframe = window.parent !== window;
    const loginMethod = isInIframe ? msalInstance.loginPopup : msalInstance.loginRedirect;

    const loginResponse = await loginMethod.call(msalInstance, { scopes });
    msalInstance.setActiveAccount(loginResponse.account);
    await handleAuth();
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

    let groupHTML = "<ul>";
    (groups.value || []).forEach(g => {
      groupHTML += `<li>${g.displayName || g.id}</li>`;
    });
    groupHTML += "</ul>";

    renderUser(profile, groupHTML, token);
  } catch (err) {
    renderError(err);
  }
}

async function handleAuth() {
  try {
    const accounts = msalInstance.getAllAccounts();
    if (accounts.length === 0) {
      signInButton();
      return;
    }

    msalInstance.setActiveAccount(accounts[0]);

    const tokenResp = await msalInstance.acquireTokenSilent({
      scopes,
      account: accounts[0]
    });

    await fetchGraphData(tokenResp.accessToken);
  } catch (e) {
    renderError(e);
    signInButton();
  }
}

msalInstance.handleRedirectPromise().then(async (response) => {
  if (response && response.account) {
    msalInstance.setActiveAccount(response.account);
  }
  await handleAuth();
}).catch(err => renderError(err));