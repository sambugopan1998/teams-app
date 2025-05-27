const msalConfig = {
      auth: {
        clientId: '0486fae2-afeb-4044-ab8d-0c060910b0a8', // App with Graph permissions
        authority: "https://login.microsoftonline.com/common",
        redirectUri: "https://sambugopan1998.github.io/teams-app"
      }
    };

    const msalInstance = new msal.PublicClientApplication(msalConfig);

    microsoftTeams.app.initialize().then(() => {
      msalInstance.loginPopup({
        scopes: ["User.Read"]
      }).then((response) => {
        const accessToken = response.accessToken;
        console.log("accessToken:",accessToken);
        fetch("https://graph.microsoft.com/v1.0/me", {
          headers: {
            Authorization: `Bearer ${accessToken}`
          }
        }).then(res => res.json())
          .then(data => {
            document.getElementById("output").textContent = `
              👤 ${data.displayName}
              📧 ${data.mail || data.userPrincipalName}
            `;
          });
      }).catch(err => {
        document.getElementById("output").textContent = "Login failed: " + err;
      });
    });


