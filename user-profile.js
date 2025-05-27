const msalConfig = {
  auth: {
    clientId: '0486fae2-afeb-4044-ab8d-0c060910b0a8',
    authority: 'https://login.microsoftonline.com/c06fea01-72bf-415d-ac1d-ac0382f8d39f',
    redirectUri: "https://sambugopan1998.github.io/teams-app/hello.html"
  }
};

const msalInstance = new msal.PublicClientApplication(msalConfig);
const loginRequest = { scopes: ["User.Read"] };

// Check redirect result first (to complete the flow)
msalInstance.handleRedirectPromise().then(async (response) => {
  if (response) {
    msalInstance.setActiveAccount(response.account);
  } else {
    const currentAccounts = msalInstance.getAllAccounts();
    if (currentAccounts.length === 0) {
      // ✅ SAFE inside Teams iframe
      msalInstance.loginRedirect(loginRequest);
    } else {
      msalInstance.setActiveAccount(currentAccounts[0]);
    }
  }

  // Try to acquire token silently
  const tokenResponse = await msalInstance.acquireTokenSilent(loginRequest);
  const token = tokenResponse.accessToken;

  document.getElementById("access-token").textContent = token;

  const res = await fetch("https://graph.microsoft.com/v1.0/me", {
    headers: { Authorization: `Bearer ${token}` }
  });

  const user = await res.json();
  document.getElementById("user-info").innerHTML = `
    ✅ Name: ${user.displayName}<br>
    📧 Email: ${user.mail || user.userPrincipalName}
  `;
}).catch(err => {
  console.error("Auth error:", err);
  document.getElementById("user-info").textContent = "❌ Auth failed: " + err.message;
});
