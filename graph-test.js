import { PublicClientApplication } from "@azure/msal-browser";
import { verifyGraphMe } from "./src/graph.js";

const signInButton = document.getElementById("sign-in");
const graphButton = document.getElementById("test-graph");
const status = document.getElementById("status");

const clientId = (import.meta.env.VITE_ENTRA_CLIENT_ID ?? "").trim();
const tenantId = (import.meta.env.VITE_ENTRA_TENANT_ID ?? "").trim();
const isUuid = (value) => /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(value);

let msal = null;
let signedInAccount = null;

function setStatus(message) {
  status.textContent = message;
}

async function initialise() {
  if (!isUuid(clientId) || !isUuid(tenantId)) {
    setStatus("Configure VITE_ENTRA_CLIENT_ID and VITE_ENTRA_TENANT_ID in an untracked .env.local file first.");
    return;
  }

  const redirectUri = new URL("/redirect.html", window.location.origin).href;
  msal = new PublicClientApplication({
    auth: {
      clientId,
      authority: `https://login.microsoftonline.com/${tenantId}`,
      redirectUri,
      postLogoutRedirectUri: redirectUri,
    },
    cache: { cacheLocation: "sessionStorage" },
  });

  try {
    await msal.initialize();
    signInButton.disabled = false;
    setStatus("Ready. Sign in using the development tenant account.");
  } catch {
    setStatus("Microsoft authentication could not initialise. Check the local configuration.");
  }
}

signInButton.addEventListener("click", async () => {
  if (!msal) return;
  signInButton.disabled = true;
  graphButton.disabled = true;
  signedInAccount = null;
  setStatus("Opening Microsoft sign-in...");
  try {
    const authentication = await msal.loginPopup({ scopes: ["User.Read"], prompt: "select_account" });
    if (!authentication.account) throw new Error("Missing signed-in account");
    signedInAccount = authentication.account;
    msal.setActiveAccount(signedInAccount);
    graphButton.disabled = false;
    setStatus("Signed in. Click 'Test Graph' to make one read-only request.");
  } catch {
    setStatus("Sign-in did not complete. Check pop-up blocking, tenant access and registered SPA redirect URI.");
  } finally {
    signInButton.disabled = false;
  }
});

graphButton.addEventListener("click", async () => {
  if (!msal || !signedInAccount) return;
  graphButton.disabled = true;
  setStatus("Requesting a delegated token and making GET /me...");
  try {
    const authentication = await msal.acquireTokenSilent({
      scopes: ["User.Read"],
      account: signedInAccount,
    });
    const result = await verifyGraphMe(authentication.accessToken);
    setStatus(`PASS: Live Microsoft Graph /me returned HTTP ${result.httpStatus}. No profile data was retained.`);
  } catch {
    setStatus("HOLD: Live Graph test did not complete. Check sign-in, User.Read consent and connectivity. Do not paste tokens into reports.");
  } finally {
    graphButton.disabled = false;
  }
});

initialise();
