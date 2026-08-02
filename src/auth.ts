import { PublicClientApplication, type AccountInfo, InteractionRequiredAuthError } from "@azure/msal-browser";

// Diese Werte stammen aus der Azure-App-Registrierung "P03 Packliste App".
const CLIENT_ID = "ddfae93a-0623-4d37-b6dd-f25c385d06d1";
const TENANT_AUTHORITY = "https://login.microsoftonline.com/common"; // "common" erlaubt private Microsoft-Konten und Organisationskonten

// BASE_URL wird von Vite anhand von vite.config.ts (base: "/p03-packliste/") gesetzt.
// Dadurch passt die Redirect-URI automatisch zu GitHub Pages und zu "npm run dev".
const REDIRECT_URI = `${window.location.origin}${import.meta.env.BASE_URL}`;

export const msalInstance = new PublicClientApplication({
  auth: {
    clientId: CLIENT_ID,
    authority: TENANT_AUTHORITY,
    redirectUri: REDIRECT_URI,
    postLogoutRedirectUri: REDIRECT_URI,
  },
  cache: {
    cacheLocation: "localStorage", // übersteht Browser-Neustarts; nur das Zugriffstoken, keine Packlistendaten
    storeAuthStateInCookie: false,
  },
});

let initialized = false;

export async function initMsal(): Promise<void> {
  if (initialized) return;
  await msalInstance.initialize();
  const response = await msalInstance.handleRedirectPromise();
  if (response?.account) {
    msalInstance.setActiveAccount(response.account);
  } else {
    const accounts = msalInstance.getAllAccounts();
    if (accounts.length > 0) msalInstance.setActiveAccount(accounts[0]);
  }
  initialized = true;
}

export function getAccount(): AccountInfo | null {
  return msalInstance.getActiveAccount();
}

const SCOPES = ["Files.ReadWrite", "User.Read"];

export async function login(): Promise<AccountInfo> {
  const result = await msalInstance.loginPopup({ scopes: SCOPES });
  msalInstance.setActiveAccount(result.account);
  return result.account;
}

export async function logout(): Promise<void> {
  const account = getAccount();
  await msalInstance.logoutPopup({ account: account ?? undefined });
}

export async function getAccessToken(): Promise<string> {
  const account = getAccount();
  if (!account) throw new Error("Nicht angemeldet.");
  try {
    const result = await msalInstance.acquireTokenSilent({ scopes: SCOPES, account });
    return result.accessToken;
  } catch (error) {
    if (error instanceof InteractionRequiredAuthError) {
      const result = await msalInstance.acquireTokenPopup({ scopes: SCOPES, account });
      return result.accessToken;
    }
    throw error;
  }
}
