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
  const token = await acquireAccessToken();

  document.getElementById("access-token").textContent = token;

  // 1. Get basic user profile
  const profileRes = await fetch("https://graph.microsoft.com/v1.0/me", {
    headers: { Authorization: `Bearer ${token}` }
  });
  const user = await profileRes.json();

  let html = "<h3>👤 Profile Info</h3><ul>";
  for (const [key, value] of Object.entries(user)) {
    html += `<li><strong>${key}</strong>: ${value ?? "N/A"}</li>`;
  }
  html += "</ul>";
  document.getElementById("user-info").innerHTML = html;

  // 2. Get user photo
  const photoRes = await fetch("https://graph.microsoft.com/v1.0/me/photo/$value", {
    headers: { Authorization: `Bearer ${token}` }
  });

  if (photoRes.ok) {
    const blob = await photoRes.blob();
    const imgURL = URL.createObjectURL(blob);
    const imgTag = `<h3>🖼️ Profile Photo</h3><img src="${imgURL}" alt="Profile Photo" style="height:100px;border-radius:50%;">`;
    document.getElementById("user-info").insertAdjacentHTML("afterbegin", imgTag);
  }

  // 3. Get directory roles or group memberships
  const roleRes = await fetch("https://graph.microsoft.com/v1.0/me/memberOf", {
    headers: { Authorization: `Bearer ${token}` }
  });

  if (roleRes.ok) {
    const roleData = await roleRes.json();
    let rolesHTML = "<h3>🔐 Roles / Groups</h3><ul>";

    roleData.value.forEach(entry => {
      rolesHTML += `<li><strong>${entry["@odata.type"]}</strong>: ${entry.displayName}</li>`;
    });

    rolesHTML += "</ul>";
    document.getElementById("user-info").insertAdjacentHTML("beforeend", rolesHTML);
  } else {
    console.warn("Group/role info could not be fetched. Missing permissions?");
  }
})();
