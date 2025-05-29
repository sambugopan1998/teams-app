import * as msal from "https://cdn.jsdelivr.net/npm/@azure/msal-browser@3.11.0/+esm";

const msalConfig = {
  auth: {
    clientId: "0486fae2-afeb-4044-ab8d-0c060910b0a8",
    authority: "https://login.microsoftonline.com/c06fea01-72bf-415d-ac1d-ac0382f8d39f",
    redirectUri: "https://sambugopan1998.github.io/teams-app/hello.html"
  }
};

const graphScopes = ["User.Read"];
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

function renderUser(user, token) {
  let html = "<h3>👤 User Profile</h3><ul>";
  for (const [key, value] of Object.entries(user)) {
    html += `<li><strong>${key}</strong>: ${value || "N/A"}</li>`;
  }
  html += `</ul><h3>🔑 Access Token</h3><textarea readonly>${token}</textarea>`;
  render(html);
}

async function fetchGraphData(token) {
  const headers = { Authorization: `Bearer ${token}` };
  const profileRes = await fetch("https://graph.microsoft.com/v1.0/me", { headers });
  const profile = await profileRes.json();
  renderUser(profile, token);
}

(async () => {
  try {
    await msalInstance.initialize();
    await microsoftTeams.app.initialize();

    // Get Teams SSO token (identity)
    microsoftTeams.authentication.getAuthToken({
      successCallback: async () => {
        console.log("✅ Teams SSO Token acquired");

        // Use MSAL to get Graph token
        const accounts = msalInstance.getAllAccounts();
        if (accounts.length === 0) {
          renderError("No MSAL account. Please sign in via Azure AD once.");
          return;
        }

        const graphTokenResponse = await msalInstance.acquireTokenSilent({
          scopes: graphScopes,
          account: accounts[0]
        });

        console.log("✅ Microsoft Graph Token:", graphTokenResponse.accessToken);
        await fetchGraphData(graphTokenResponse.accessToken);
      },
      failureCallback: (error) => {
        renderError("Teams SSO failed: " + error);
      }
    });
  } catch (err) {
    renderError(err);
  }
})();
