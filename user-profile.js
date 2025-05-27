const msalConfig = {
  auth: {
    clientId: "0486fae2-afeb-4044-ab8d-0c060910b0a8",
    authority: "https://login.microsoftonline.com/c06fea01-72bf-415d-ac1d-ac0382f8d39f",
    redirectUri: "https://sambugopan1998.github.io/teams-app/hello.html"
  }
};

const msalInstance = new msal.PublicClientApplication(msalConfig);
const loginRequest = { scopes: ["User.Read"] };

async function acquireAccessToken() {
  try {
    const accounts = msalInstance.getAllAccounts();

    if (accounts.length === 0) {
      const loginResponse = await msalInstance.loginPopup(loginRequest);
      msalInstance.setActiveAccount(loginResponse.account);
    } else {
      msalInstance.setActiveAccount(accounts[0]);
    }

    const tokenResponse = await msalInstance.acquireTokenSilent(loginRequest);
    return tokenResponse.accessToken;

  } catch (error) {
    console.warn("Silent token failed, using popup", error);

    try {
      const tokenResponse = await msalInstance.acquireTokenPopup(loginRequest);
      return tokenResponse.accessToken;
    } catch (popupError) {
      // 👇 Show popup error in UI
      document.getElementById("user-info").textContent =
        `❌ Popup error: ${popupError.name} — ${popupError.message}`;
      throw popupError;
    }
  }
}

(async () => {
  try {
    const token = await acquireAccessToken();
    document.getElementById("access-token").textContent = token;

    const res = await fetch("https://graph.microsoft.com/v1.0/me", {
      headers: { Authorization: `Bearer ${token}` }
    });

    if (!res.ok) {
      throw new Error("Graph API failed: " + res.status);
    }

    const user = await res.json();
    document.getElementById("user-info").innerHTML = `
      ✅ Name: ${user.displayName}<br>
      📧 Email: ${user.mail || user.userPrincipalName}
    `;
  } catch (err) {
    // 👇 Catch-all fallback
    document.getElementById("user-info").textContent =
      `❌ Auth or Graph error: ${err.name} — ${err.message}`;
    console.error("Final error:", err);
  }
})();
