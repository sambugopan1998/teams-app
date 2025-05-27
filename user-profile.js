const msalConfig = {
      auth: {
        clientId: '0486fae2-afeb-4044-ab8d-0c060910b0a8', // App with Graph permissions
        authority: "https://login.microsoftonline.com/common",
        redirectUri: "https://sambugopan1998.github.io/teams-app"
      }
    };

const msalInstance = new msal.PublicClientApplication(msalConfig);

async function acquireAccessToken() {
  const accounts = msalInstance.getAllAccounts();

  if (accounts.length === 0) {
    try {
      const loginResponse = await msalInstance.loginPopup({
        scopes: ["User.Read"]
      });
      console.log("Login response:", loginResponse);
    } catch (err) {
      console.error("Login error:", err);
      return;
    }
  }

  try {
    const response = await msalInstance.acquireTokenSilent({
      scopes: ["User.Read"],
      account: msalInstance.getAllAccounts()[0]
    });
    return response.accessToken;
  } catch (err) {
    console.warn("Silent token failed, trying popup...", err);

    // Only call acquireTokenPopup if no interaction is already happening
    if (err instanceof msal.InteractionRequiredAuthError) {
      return msalInstance.acquireTokenPopup({
        scopes: ["User.Read"]
      }).then(response => response.accessToken);
    } else {
      throw err;
    }
  }
}

microsoftTeams.app.initialize().then(async () => {
  try {
    const token = await acquireAccessToken();

    const res = await fetch("https://graph.microsoft.com/v1.0/me", {
      headers: {
        Authorization: `Bearer ${token}`
      }
    });

    const user = await res.json();
    document.getElementById("output").textContent = `
      👤 ${user.displayName}
      📧 ${user.mail || user.userPrincipalName}
    `;
  } catch (err) {
    console.error("Final error:", err);
    document.getElementById("output").textContent = "❌ Error: " + err;
  }
});


