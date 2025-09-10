// Aspetta che il DOM sia completamente caricato
document.addEventListener('DOMContentLoaded', async () => {
    console.log("Script caricato correttamente!");

    // Configurazione MSAL
    const msalConfig = {
        auth: {
            clientId: "c3893db8-ca5a-4193-8cfd-08feb16832b1",
            authority: "https://login.microsoftonline.com/common",
            redirectUri: "http://localhost:8080"
        }
    };

    const msalInstance = new msal.PublicClientApplication(msalConfig);
    const loginRequest = {
        scopes: ["User.Read", "Files.ReadWrite"]
    };

    // Imposta account attivo se già presente
    const accounts = msalInstance.getAllAccounts();
    if (accounts.length > 0) {
        msalInstance.setActiveAccount(accounts[0]);
        console.log("Account attivo impostato automaticamente:", accounts[0]);
    }

    // Funzione per effettuare login se necessario
    async function ensureLogin() {
        let account = msalInstance.getActiveAccount();
        if (!account) {
            try {
                const loginResponse = await msalInstance.loginPopup(loginRequest);
                msalInstance.setActiveAccount(loginResponse.account);
                account = loginResponse.account;
                console.log("Login effettuato:", account);
            } catch (error) {
                console.error("Errore nel login:", error);
                statoElement.innerText = "Errore durante il login. Riprova.";
                return null;
            }
        }
        return account;
    }

    // Funzione per ottenere il token
    async function getAccessToken() {
        const account = await ensureLogin();
        if (!account) return null;

        try {
            const response = await msalInstance.acquireTokenSilent({
                ...loginRequest,
                account: account
            });
            return response.accessToken;
        } catch (error) {
            console.warn("Token silenzioso fallito, provo con popup...");
            try {
                const popupResponse = await msalInstance.acquireTokenPopup(loginRequest);
                return popupResponse.accessToken;
            } catch (popupError) {
                console.error("Errore nel token tramite popup:", popupError);
                statoElement.innerText = "Autenticazione fallita.";
                return null;
            }
        }
    }

    // Elementi DOM
    const menuPrincipale = document.querySelector('.pulsanti-container');
    const formContainer = document.getElementById('formInserisciNominativo');
    const btnInserisciNominativo = document.getElementById('btnInserisci');
    const btnIndietro = document.getElementById('btnIndietro');
    const nominativoForm = document.getElementById('nominativoForm');
    const statoElement = document.getElementById("stato");

    // Popola dropdown
    function popolaDropdown() {
        const giornoDropdown = document.getElementById('giorno');
        const meseDropdown = document.getElementById('mese');
        const orarioDropdown = document.getElementById('orario');

        giornoDropdown.innerHTML = '';
        for (let i = 1; i <= 31; i++) {
            giornoDropdown.innerHTML += `<option value="${i}">${i}</option>`;
        }

        meseDropdown.innerHTML = '';
        for (let i = 1; i <= 12; i++) {
            meseDropdown.innerHTML += `<option value="${i}">${i}</option>`;
        }

        const orari = ['9:00', '9:30', '10:00', '10:30', '11:00', '11:30', '12:00', '12:30', '13:00', '13:30', '14:00', '14:30', '15:00', '15:30', '16:00', '16:30', '17:00', '17:30', '18:00', '18:30', '19:00', '19:30'];
        orarioDropdown.innerHTML = '';
        orari.forEach(orario => {
            orarioDropdown.innerHTML += `<option value="${orario}">${orario}</option>`;
        });
    }

    // Mostra form
    btnInserisciNominativo.addEventListener('click', async () => {
        const account = await ensureLogin();
        if (!account) return;

        menuPrincipale.classList.add('hidden');
        formContainer.classList.remove('hidden');
        popolaDropdown();
    });

    // Torna al menu
    btnIndietro.addEventListener('click', () => {
        formContainer.classList.add('hidden');
        menuPrincipale.classList.remove('hidden');
    });

    // Invia dati
    nominativoForm.addEventListener('submit', async (e) => {
        e.preventDefault();

        const dati = {
            cognome: document.getElementById('cognome').value,
            nome: document.getElementById('nome').value,
            ambiente: document.getElementById('ambiente').value,
            gruppo: document.getElementById('gruppo').value,
            consulente: document.getElementById('consulente').value,
            arredatore: document.getElementById('arredatore').value,
            giorno: document.getElementById('giorno').value,
            mese: document.getElementById('mese').value,
            orario: document.getElementById('orario').value
        };

        statoElement.innerText = "Invio in corso...";

        const accessToken = await getAccessToken();
        if (!accessToken) return;

        const fileId = "A3856CCE-D8CC-4C35-92E3-02EAB1E3B368";
        const nomiMesi = ['gen', 'feb', 'mar', 'apr', 'mag', 'giu', 'lug', 'ago', 'set', 'ott', 'nov', 'dic'];
        const mese = nomiMesi[parseInt(dati.mese) - 1];
        const worksheetName = `${dati.giorno}-${mese}`;
        const rangeAddress = "A4:T4";
        const apiUrl = `https://graph.microsoft.com/v1.0/me/drive/items/${fileId}/workbook/worksheets('${worksheetName}')/range(address='${rangeAddress}')`;

        const valoriRiga = [
            [dati.orario, '', `${dati.cognome} ${dati.nome}`, dati.ambiente, ...Array(13).fill(''), dati.gruppo, dati.consulente, dati.arredatore]
        ];

        try {
            await fetch(apiUrl, {
                method: 'PATCH',
                headers: {
                    'Authorization': `Bearer ${accessToken}`,
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify({ values: valoriRiga })
            });

            statoElement.innerText = `Dati inseriti nel foglio "${worksheetName}"!`;
            nominativoForm.reset();
            setTimeout(() => {
                formContainer.classList.add('hidden');
                menuPrincipale.classList.remove('hidden');
            }, 2000);
        } catch (error) {
            console.error("Errore durante l'invio:", error);
            statoElement.innerText = `Errore: ${error.message}`;
        }
    });

    // Funzioni placeholder per gli altri pulsanti
    document.getElementById("btnRicerca").addEventListener("click", () => {
        statoElement.innerText = "Funzionalità 'RICERCA NOMINATIVO' ancora da implementare.";
    });

    document.getElementById("btnVisualizza").addEventListener("click", () => {
        statoElement.innerText = "Funzionalità 'VISUALIZZA LISTA GIORNO' ancora da implementare.";
    });

    document.getElementById("btnCompila").addEventListener("click", () => {
        statoElement.innerText = "Funzionalità 'COMPILA LISTA GIORNO' ancora da implementare.";
    });
});
