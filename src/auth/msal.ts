import { PublicClientApplication, type Configuration } from "@azure/msal-browser";

const msalConfig: Configuration = {
  auth: {
    /* clientId: "4fa634d0-d0eb-4e93-ba60-11202871091d",  // from Azure app registration
    // authority must point to Azure AD (tenant or common), not to localhost
    authority: "https://login.microsoftonline.com/614a3022-85c5-41b0-8544-44baff944166",
    // redirectUri should be the add-in taskpane URL registered in Azure
    redirectUri: "https://localhost:3000/taskpane.html" */
    clientId: "1869ede3-2b24-47ea-a5c1-0eee6327df9a",  // from Azure app registration
    authority: "https://login.microsoftonline.com/77b82cb6-96cf-4c9f-81b5-fe2a70e55d58",
    // redirectUri should be the add-in taskpane URL registered in Azure
    redirectUri: "https://localhost:3000/taskpane.html"
  },
  cache: {
    cacheLocation: "localStorage",
    storeAuthStateInCookie: false
  }
};

export const pca = new PublicClientApplication(msalConfig);

// MSAL requires initialize() to be called before other API calls in some environments.
export async function initializeMsal(): Promise<void> {
  // initialize() is idempotent; calling it multiple times is safe.
  await pca.initialize();
}

const loginRequest = {
  scopes: ["User.Read", "Mail.Read"] // add Mail.Read so /me/messages is allowed
};

export async function signIn() {
  // interactive login (popup or redirect)
  const loginResponse = await pca.loginPopup(loginRequest);
  // set the active account so acquireTokenSilent can find it later
  if (loginResponse && loginResponse.account) {
    pca.setActiveAccount(loginResponse.account);
  }
  return loginResponse;
}

// To avoid multiple interactive prompts, this function first tries to get a token silently.
// If that fails (e.g. no cached token, or expired), it falls back to an interactive method.
export async function getAccessToken(): Promise<string> {
  let account = pca.getActiveAccount();
  if (!account) {
    const loginResponse = await signIn();
    account = loginResponse.account!;
  }

  // Attempt silent acquisition first
  try {
    const silentResult = await pca.acquireTokenSilent({
      ...loginRequest,
      account
    });
    return silentResult.accessToken;
  } catch (err) {
    console.warn("Silent token failed, falling back to interactive", err);
    const interactiveResult = await pca.acquireTokenPopup(loginRequest);
    if (interactiveResult && interactiveResult.account) {
      pca.setActiveAccount(interactiveResult.account);
    }
    return interactiveResult.accessToken;
  }
}

export async function getAccessTokenSilent(): Promise<string> {
  const account = pca.getActiveAccount();
  if (!account) {
    throw new Error("No active account - cannot acquire token silently");
  }
  const silentResult = await pca.acquireTokenSilent({
    ...loginRequest,
    account
  });
  return silentResult.accessToken;
}

export async function callGraphMe() {
  const token = await getAccessToken();
  const res = await fetch("https://graph.microsoft.com/v1.0/me", {
    headers: { Authorization: `Bearer ${token}` }
  });
  return await res.json();
}
