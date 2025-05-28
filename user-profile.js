const msalConfig = {
  auth: {
    clientId: "0486fae2-afeb-4044-ab8d-0c060910b0a8",
    authority: "https://login.microsoftonline.com/c06fea01-72bf-415d-ac1d-ac0382f8d39f",
    redirectUri: window.location.href,
  }
};

const scopes = ["User.Read", "Directory.Read.All"];
const msalInstance = new msal.PublicClientApplication(msalConfig);
const app = document.querySelector(".app");

function render(content) {
  app.innerHTML = content;
}

function renderError(error) {
  const errMsg = `
    <div style="color: red; margin-top: 1em;">
      <h4>❌ Error:</h4>
      <pre>${JSON.stringify(error, null, 2)}</pre>
    </div>`;
  app.innerHTML += errMsg;
}

function renderUser(user, groupsHTML, token) {
  let userDetails = "<h3>👤 User Profile</h3><ul>";
  for (const [key, value] of Object.entries(user)) {
    userDetails += `<li><strong>${key}</strong>: ${value || "N/A"}</li>`;
  }
  userDetails += "</ul>";

  app.innerHTML = `
    ${userDetails}
    <h3>🔐 Groups</h3>${groupsHTML}
    <h4>🪪 Access Token</h4><textarea style="width:100%;height:100px">${token}</textarea>
  `;
}

function signInUI() {
  render(`<button id="signin">🔐 Sign in with Microsoft</button>`);
  document.getElementById("signin").addEventListener("click", signIn);
}

async function signIn() {
  try {
    await microsoftTeams.app.initialize();

    microsoftTeams.authentication.authenticate({
      url: window.location.href,
      width: 600,
      height: 535,
      successCallback: () => attemptSilentSignIn(),
      failureCallback: (err) => renderError(err),
    });
  } catch (e) {
    renderError(e);
  }
}

async function attemptSilentSignIn() {
  const account = msalInstance.getAllAccounts()[0];
  if (!account) return signInUI();

  try {
    const tokenResponse = await msalInstance.acquireTokenSilent({
      scopes,
      account
    });

    const accessToken = tokenResponse.accessToken;
    await fetchGraphData(accessToken);
  } catch (e) {
    renderError(e);
    signInUI();
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
  } catch (e) {
    renderError(e);
  }
}

// Initial auth flow
msalInstance.handleRedirectPromise().then(async (response) => {
  if (response && response.account) {
    msalInstance.setActiveAccount(response.account);
  }
  await attemptSilentSignIn();
}).catch(err => {
  renderError(err);
});
