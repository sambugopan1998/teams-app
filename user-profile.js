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

function renderError(err) {
  app.innerHTML += `<p style="color:red">❌ ${err.name || "Error"}: ${err.message || err}</p>`;
}

function renderUser(user, groupsHTML, token) {
  app.innerHTML = `
    <h2>Hello, ${user.displayName}</h2>
    <p><strong>Access Token:</strong><br><small>${token}</small></p>
    ${groupsHTML}
  `;
}

function signInUI() {
  render(`<button id="signin">Sign in</button>`);
  document.getElementById("signin").addEventListener("click", () => signIn());
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
    signInUI(); // fallback if silent auth fails
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

    let groupHTML = "<h3>🔐 Roles / Groups</h3><ul>";
    (groups.value || []).forEach(g => {
      groupHTML += `<li>${g.displayName || g.id}</li>`;
    });
    groupHTML += "</ul>";

    renderUser(profile, groupHTML, token);
  } catch (e) {
    renderError(e);
  }
}

// Entry point
msalInstance.handleRedirectPromise().then(async (response) => {
  if (response && response.account) {
    msalInstance.setActiveAccount(response.account);
  }
  await attemptSilentSignIn();
}).catch(err => {
  renderError(err);
});