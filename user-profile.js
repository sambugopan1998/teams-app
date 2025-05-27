microsoftTeams.app.initialize().then(() => {
  microsoftTeams.authentication.getAuthToken({
    successCallback: async (token) => {
      document.getElementById("access-token").textContent = token;
      try {
        const res = await fetch("https://graph.microsoft.com/v1.0/me", {
          headers: { Authorization: `Bearer ${token}` }
        });
        if (!res.ok) throw new Error("Graph call failed");
        const user = await res.json();
        document.getElementById("user-info").innerHTML = `
          👤 Name: ${user.displayName}<br>
          📧 Email: ${user.mail || user.userPrincipalName}<br>
          🏢 Location: ${user.officeLocation || "Not set"}
        `;
      } catch (err) {
        console.error("Graph error:", err);
        document.getElementById("user-info").textContent = "❌ Error fetching profile.";
      }
    },
    failureCallback: (error) => {
      document.getElementById("access-token").textContent = `❌ Token failed: ${error}`;
      console.error("Token error", error);
    },
      resources: ["api://0486fae2-afeb-4044-ab8d-0c060910b0a8"]
  });
});
