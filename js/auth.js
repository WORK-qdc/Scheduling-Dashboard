// js/auth.js
import { PublicClientApplication } from 'https://alcdn.msauth.net/browser/2.27.0/js/msal-browser.min.js';
import { msalConfig }            from './config.js';

const msalInstance = new PublicClientApplication(msalConfig);

export async function signIn() {
  const accounts = msalInstance.getAllAccounts();
  if (!accounts.length) {
    await msalInstance.loginPopup({
      scopes: ['User.Read','Sites.Read.All','Sites.ReadWrite.All']
    });
  }
}

export async function getToken(scopes) {
  try {
    const resp = await msalInstance.acquireTokenSilent({ scopes });
    return resp.accessToken;
  } catch {
    const resp = await msalInstance.acquireTokenPopup({ scopes });
    return resp.accessToken;
  }
}
