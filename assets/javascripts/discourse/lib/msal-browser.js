import "../../vendor/msal-browser.min";

const msalGlobal = window.msal;

if (!msalGlobal?.PublicClientApplication) {
  throw new Error("MSAL global not available after loading bundle");
}

export const InteractionRequiredAuthError =
  msalGlobal.InteractionRequiredAuthError;
export const PublicClientApplication = msalGlobal.PublicClientApplication;
export const BrowserCacheLocation = msalGlobal.BrowserCacheLocation;

export default msalGlobal;
