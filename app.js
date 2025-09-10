document.addEventListener('DOMContentLoaded', async () => {
  const statoElement = document.getElementById("stato");
  const elencoFogli = document.getElementById("elencoFogli");

  const msalConfig = {
    auth: {
      clientId: "c3893db8-ca5a-4193-8cfd-08feb16832b1",
      authority: "https://login.microsoftonline.com/common",
      redirectUri: "https://stevr82.github.io/Gestione-Liste-Giorno/"
    }
  };

  const msalInstance = new msal.PublicClientApplication(msalConfig);
  const loginRequest = {
    scopes: ["User.Read", "Sites.Read.All", "Files.ReadWrite.All"]
  };

  const siteId = "nordestholding.sharepoint.com,8c3c9c2e-xxxx-xxxx-xxxx-xxxxxxxxxxxx,xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"; // ← Inserisci il tuo siteId
  const fileId = "A3856CCE-D8CC-4C35-92E3-02EAB1E3B368";

  const handleRedirect = async () => {
    try {
      const response = await msalInstance.handleRedirectPromise();
      if (response && response.account) {
        msalInstance.setActiveAccount(response.account);
        localStorage.setItem("msalAccount", response.account.homeAccountId);
      }
    } catch (error) {
      console.error("Errore nel redirect:", error);
    }
  };

  await handleRedirect();

  const savedAccountId = localStorage.getItem("msalAccount");
  const accounts = msalInstance.getAllAccounts();
  const account = accounts.find(acc => acc.homeAccountId === savedAccountId);

  if (account) {
    msalInstance.setActiveAccount(account);
  } else {
    msalInstance.loginRedirect(loginRequest);
    return;
  }

  async function getAccessToken() {
    try {
      const response = await msalInstance.acquireTokenSilent({
        ...loginRequest,
        account: msalInstance.getActiveAccount()
     
