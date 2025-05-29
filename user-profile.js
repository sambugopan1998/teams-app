import * as msal from "https://cdn.jsdelivr.net/npm/@azure/msal-browser@3.11.0/+esm";

const msalConfig = {
  auth: {
    clientId: "0486fae2-afeb-4044-ab8d-0c060910b0a8",
    authority: "https://login.microsoftonline.com/c06fea01-72bf-415d-ac1d-ac0382f8d39f",
    redirectUri: "https://sambugopan1998.github.io/teams-app/hello.html"
  }
};

const graphScopes = ["User.Read", "Directory.Read.All"];
const msalInstance = new msal.PublicClientApplication(msalConfig);
const app = document.querySelector(".app");

function render(content) {
  app.innerHTML = content;
}

function renderError(error) {
  const errMsg = error?.message || JSON.stringify(error) || "Unknown error";
  app.innerHTML = `
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

(async () => {
  try {
    await msalInstance.initialize();
    await microsoftTeams.app.initialize();

    microsoftTeams.authentication.getAuthToken({
      successCallback: async (ssoToken) => {
        console.log("✅ Teams SSO Token:", ssoToken);

        const accounts = msalInstance.getAllAccounts();
        if (accounts.length > 0) {
          msalInstance.setActiveAccount(accounts[0]);

          try {
            const graphToken = await msalInstance.acquireTokenSilent({
              scopes: graphScopes,
              account: accounts[0]
            });

            console.log("✅ Graph token acquired silently");
            await fetchGraphData(graphToken.accessToken);
          } catch (silentErr) {
            renderError("⚠️ Silent token failed. User must consent via popup or backend exchange.");
          }
        } else {
          renderError("🛑 No MSAL account found. You need to sign in via backend or handle consent.");
        }
      },
      failureCallback: (error) => {
        renderError("🛑 Teams SSO failed: " + error);
      }
    });

  } catch (err) {
    renderError(err);
  }
})();