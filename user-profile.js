microsoftTeams.app.initialize().then(() => {
  // Use Teams SSO to get the token
  microsoftTeams.authentication.getAuthToken({
    successCallback: async (token) => {
      console.log("✅ Token received from Teams SSO");
      document.getElementById("access-token").textContent = token;

      try {
        // Fetch user details from Microsoft Graph
        const res = await fetch("https://graph.microsoft.com/v1.0/me", {
          headers: {
            Authorization: `Bearer ${token}`
          }
        });

        const user = await res.json();

        // Display user details
        document.getElementById("user-info").innerHTML = `
          ✅ <strong>Display Name:</strong> ${user.displayName}<br>
          👤 <strong>Given Name:</strong> ${user.givenName || "-"}<br>
          📧 <strong>Email:</strong> ${user.mail || user.userPrincipalName}<br>
          🧑‍💼 <strong>Job Title:</strong> ${user.jobTitle || "-"}<br>
          🏢 <strong>Office Location:</strong> ${user.officeLocation || "-"}<br>
          🆔 <strong>ID:</strong> ${user.id}
        `;
      } catch (error) {
        console.error("❌ Microsoft Graph call failed:", error);
        document.getElementById("user-info").textContent = "Failed to load user profile.";
      }
    },
    failureCallback: (error) => {
      console.error("❌ Failed to get token via Teams SSO:", error);
      document.getElementById("access-token").textContent = "Token fetch failed.";
      document.getElementById("user-info").textContent = "Authentication failed.";
    }
  });
});

