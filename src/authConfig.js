export const msalConfig = {
  auth: {
    clientId: import.meta.env.VITE_MSAL_CLIENT_ID,
    authority: `https://login.microsoftonline.com/${import.meta.env.VITE_MSAL_TENANT_ID}`,
    redirectUri: window.location.origin,
  },
  cache: {
    cacheLocation: "sessionStorage",
    storeAuthStateInCookie: false,
  },
};

// Permissions we need to read/write SharePoint files
export const loginRequest = {
  scopes: ["Files.ReadWrite", "Sites.ReadWrite.All", "User.Read"],
};

// SharePoint / Teams folder config
export const sharePointConfig = {
  hostname: import.meta.env.VITE_SP_HOSTNAME,
  siteName: import.meta.env.VITE_SP_SITE_NAME,
  folderPath: import.meta.env.VITE_SP_FOLDER_PATH,
};