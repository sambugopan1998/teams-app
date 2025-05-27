microsoftTeams.app.initialize().then(() => {
  microsoftTeams.authentication.getAuthToken({
    resources: ["api://0486fae2-afeb-4044-ab8d-0c060910b0a8"], // ✅ custom app
    successCallback: async (token) => {
      console.log("Token received:", token);
      const res = await fetch("https://graph.microsoft.com/v1.0/me", {
        headers: { Authorization: `Bearer ${token}` }
      });
      const user = await res.json();
      document.getElementById("user-info").innerHTML = `
        👤 Name: ${user.displayName}<br>
        📧 Email: ${user.mail || user.userPrincipalName}
      `;
    },
    failureCallback: (err) => {
      console.error("Token error:", err);
      document.getElementById("user-info").textContent = "❌ Token failed";
    }
  });
});

