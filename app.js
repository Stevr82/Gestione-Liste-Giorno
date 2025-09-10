document.addEventListener('DOMContentLoaded', async () => {
  const statoElement = document.getElementById("stato");

  const msalConfig = {
    auth: {
      clientId: "c3893db8-ca5a-4193-8cfd-08feb16832b1",
      authority: "https://login.microsoftonline.com/common",
      redirectUri: "https://stevr82.github.io/Gestione-Liste-Giorno/"
    }
  };

  const msalInstance = new msal.PublicClientApplication(msalConfig);
  const loginRequest = {
    scopes: ["User.Read", "Files.ReadWrite"]
  };

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
      });
      return response.accessToken;
    } catch (error) {
      console.error("Errore nel token:", error);
      statoElement.innerText = "Autenticazione fallita.";
      return null;
    }
  }

  const menuPrincipale = document.querySelector('.pulsanti-container');
  const formContainer = document.getElementById('formInserisciNominativo');
  const btnInserisciNominativo = document.getElementById('btnInserisci');
  const btnIndietro = document.getElementById('btnIndietro');
  const nominativo
