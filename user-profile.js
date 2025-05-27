const msalConfig = {
      auth: {
        clientId: "0486fae2-afeb-4044-ab8d-0c060910b0a8",
        authority: "https://login.microsoftonline.com/c06fea01-72bf-415d-ac1d-ac0382f8d39f",
        redirectUri: "https://sambugopan1998.github.io/teams-app/hello.html"
      }
    };

    const msalInstance = new msal.PublicClientApplication(msalConfig);
    const loginRequest = { scopes: ["User.Read"] };

    microsoftTeams.app.initialize().then(async () => {
      try {
        const response = await msalInstance.handleRedirectPromise();
        if (response) {
          msalInstance.setActiveAccount(response.account);
        }

        const accounts = msalInstance.getAllAccounts();
        if (accounts.length === 0) {
          // 🔁 Redirect login if not logged in
          msalInstance.loginRedirect(loginRequest);
          return;
        } else {
          msalInstance.setActiveAccount(accounts[0]);
        }

        // Now try silent token acquisition
        const tokenResponse = await msalInstance.acquireTokenSilent({
          scopes: ["User.Read"],
          account: msalInstance.getActiveAccount()
        });

        const token = tokenResponse.accessToken;
        document.getElementById("access-token").textContent = token;

        const res = await fetch("https://graph.microsoft.com/v1.0/me", {
          headers: { Authorization: `Bearer ${token}` }
        });

        const user = await res.json();
        document.getElementById("user-info").innerHTML = `
          👤 ${user.displayName}<br>
          📧 ${user.mail || user.userPrincipalName}
        `;
      } catch (err) {
        console.error("❌ Auth error:", err);
        document.getElementById("user-info").textContent = "❌ " + err.message;
      }
    });
