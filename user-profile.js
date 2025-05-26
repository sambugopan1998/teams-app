microsoftTeams.app.initialize().then(() => {
  microsoftTeams.authentication.getAuthToken({
    successCallback: async (token) => {
      document.getElementById("access-token").textContent = token;
      const res = await fetch("https://graph.microsoft.com/v1.0/me", {
        headers: { Authorization: `Bearer ${token}` }
      });
      const user = await res.json();
      document.getElementById("user-info").innerHTML = `
        👤 Name: ${user.displayName}<br>
        📧 Email: ${user.mail || user.userPrincipalName}<br>
        🏢 Location: ${user.officeLocation || "Not set"}
      `;
    },
    failureCallback: (error) => {
      document.getElementById("access-token").textContent = `❌ Token failed: ${error}`;
      console.error("Token error", error);
    }
  });
});
