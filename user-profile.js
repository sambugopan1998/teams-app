if (window.microsoftTeams) {
  microsoftTeams.app.initialize().then(() => {
    microsoftTeams.authentication.getAuthToken({
      successCallback: async (token) => {
        document.getElementById("access-token").textContent = token;
        const res = await fetch("https://graph.microsoft.com/v1.0/me", {
          headers: { Authorization: `Bearer ${token}` }
        });
        const user = await res.json();

        document.getElementById("user-info").innerHTML = `
          ✅ <strong>Name:</strong> ${user.displayName}<br>
          👤 <strong>Given Name:</strong> ${user.givenName || "-"}<br>
          📧 <strong>Email:</strong> ${user.mail || user.userPrincipalName}<br>
        `;
      },
      failureCallback: (err) => {
        console.error("❌ Token failed", err);
      }
    });
  });
} else {
  console.error("❌ Microsoft Teams SDK not loaded.");
}
