import { PublicClientApplication, type Configuration } from "@azure/msal-browser";

const msalConfig: Configuration = {
  auth: {
    clientId: "<CLIENT_ID>",  // from Azure app registration
    // authority must point to Azure AD (tenant or common), not to localhost
    authority: "https://login.microsoftonline.com/<TENANT_ID>",
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
  scopes: ["User.Read"] // scopes you need for Graph
};

export async function signIn() {
  // interactive login (popup or redirect)
  const loginResponse = await pca.loginPopup(loginRequest);
  return loginResponse;
}

export async function getAccessToken(): Promise<string> {
  let account = pca.getActiveAccount();
  if (!account) {
    const loginResponse = await signIn();
    account = loginResponse.account!;
  }

  try {
    const silentResult = await pca.acquireTokenSilent({
      ...loginRequest,
      account
    });
    return silentResult.accessToken;
  } catch (err) {
    console.warn("Silent token failed, using popup", err);
    const interactiveResult = await pca.acquireTokenPopup(loginRequest);
    return interactiveResult.accessToken;
  }
}

export async function callGraphMe() {
  const token = await getAccessToken();
  const res = await fetch("https://graph.microsoft.com/v1.0/me", {
    headers: { Authorization: `Bearer ${token}` }
  });
  return await res.json();
}