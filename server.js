const express = require("express");
const session = require("express-session");
const { ConfidentialClientApplication } = require("@azure/msal-node");
const { Client } = require("@microsoft/microsoft-graph-client");
require("isomorphic-fetch");
require("dotenv").config();

const app = express();
const PORT = 3000;

// MSAL configuration
const msalConfig = {
  auth: {
    clientId: process.env.CLIENT_ID,
    authority: `https://login.microsoftonline.com/${process.env.TENANT_ID}`,
    clientSecret: process.env.CLIENT_SECRET, // Confidential client requires clientSecret
  },
};

const msalClient = new ConfidentialClientApplication(msalConfig);

// Session setup
app.use(
  session({
    secret: process.env.SESSION_SECRET,
    resave: false,
    saveUninitialized: true,
  })
);

// Redirect to Microsoft login
app.get("/auth", (req, res) => {
  const authUrlParams = {
    scopes: ["user.read", "mail.read"],
    redirectUri: process.env.REDIRECT_URI,
  };

  msalClient
    .getAuthCodeUrl(authUrlParams)
    .then((url) => res.redirect(url))
    .catch((err) => {
      console.error("Error generating auth URL", err);
      res.status(500).send("Error generating authentication URL");
    });
});

// Handle redirect and acquire token
app.get("/auth/callback", async (req, res) => {
  const tokenRequest = {
    code: req.query.code,
    scopes: ["user.read", "mail.read"],
    redirectUri: process.env.REDIRECT_URI,
  };

  try {
    const authResponse = await msalClient.acquireTokenByCode(tokenRequest);
    req.session.token = authResponse.accessToken;
    res.redirect("/emails");
  } catch (error) {
    console.error("Error acquiring token", error);
    res.status(500).send("Error during authentication");
  }
});

// Fetch emails using Microsoft Graph API
app.get("/emails", async (req, res) => {
  const accessToken = req.session.token;

  if (!accessToken) {
    return res.redirect("/auth");
  }

  const client = Client.init({
    authProvider: (done) => done(null, accessToken),
  });

  try {
    const messages = await client.api("/me/messages").get();
    res.json(messages.value);
  } catch (error) {
    console.error("Error fetching emails", error);
    res.status(500).send("Error fetching emails");
  }
});

// Start the server
app.listen(PORT, () => {
  console.log(`Server is running on http://localhost:${PORT}`);
});
