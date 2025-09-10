document.addEventListener('DOMContentLoaded', async () => {
    const statoElement = document.getElementById("stato");
    const foglioDropdown = document.getElementById("foglio");

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

    async function elencaFogliNelFile() {
        const accessToken = await getAccessToken();
        if (!accessToken) return;

        const fileId = "A3856CCE-D8CC-4C35-92E3-02EAB1E3B368";
        const url = `https://graph.microsoft.com/v1.0/me/drive/items/${fileId}/workbook/worksheets`;

        try {
            const response = await fetch(url, {
                headers: { Authorization: `Bearer ${accessToken}` }
            });

            const data = await response.json();
            foglioDropdown.innerHTML = "";

            if (data.value && data.value.length > 0) {
                data.value.forEach(ws => {
                    const option = document.createElement("option");
                    option.value = ws.name;
                    option.textContent = ws.name;
                    foglioDropdown.appendChild(option);
                });
                statoElement.innerText = `✅ Fogli caricati: ${data.value.length}`;
            } else {
                statoElement.innerText = "⚠️ Nessun foglio trovato nel file.";
            }
        } catch (error) {
            console.error("Errore nel recupero dei fogli:", error);
            statoElement.innerText = `❌ Errore: ${error.message}`;
        }
    }

    const menuPrincipale = document.querySelector('.pulsanti-container');
    const formContainer = document.getElementById('formInserisciNominativo');
    const btnInserisciNominativo = document.getElementById('btnInserisci');
    const btnIndietro = document.getElementById('btnIndietro');
    const nominativoForm = document.getElementById('nominativoForm');

    function popolaDropdown() {
        const orario = document.getElementById('orario');
        const orari = ['9:00', '9:30', '10:00', '10:30', '11:00', '11:30', '12:00', '12:30', '13:00', '13:30', '14:00', '14:30', '15:00', '15:30', '16:00', '16:30', '17:00', '17:30', '18:00', '18:30', '19:00', '19:30'];
        orario.innerHTML = '';
        orari.forEach(o => {
            orario.innerHTML += `<option value="${o}">${o}</option>`;
        });
    }

    btnInserisciNominativo.addEventListener('click', async () => {
        menuPrincipale.classList.add('hidden');
        formContainer.classList.remove('hidden');
        popolaDropdown();
        await elencaFogliNelFile();
    });

    btnIndietro.addEventListener('click', () => {
        formContainer.classList.add('hidden');
        menuPrincipale.classList.remove('hidden');
    });

    nominativoForm.addEventListener('submit', async (e) => {
        e.preventDefault();

        const dati = {
            cognome: document.getElementById('cognome').value,
            nome: document.getElementById('nome').value,
            ambiente: document.getElementById('ambiente').value,
            gruppo: document.getElementById('gruppo').value,
            consulente: document.getElementById('consulente').value,
            arredatore: document.getElementById('arredatore').value,
            orario: document.getElementById('orario').value,
            foglio: document.getElementById('foglio').value
        };

        statoElement.innerText = "Invio in corso...";

        const accessToken = await getAccessToken();
        if (!accessToken) return;

        const fileId = "A3856CCE-D8CC-4C35-92E3-02EAB1E3B368";

        const sessionResponse = await fetch(`https://graph.microsoft.com/v1.0/me/drive/items/${fileId}/workbook/createSession`, {
            method: 'POST',
            headers: {
                'Authorization': `Bearer ${accessToken}`,
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({ persistChanges: true })
        });

        const sessionId = (await sessionResponse.json()).id;

        const rangeAddress = "A4:T4";
        const apiUrl = `https://graph.microsoft.com/v1.0/me/drive/items/${fileId}/workbook/worksheets('${dati.foglio}')/range(address='${rangeAddress}')`;

        const valoriRiga = [[
            dati.orario,                     // A4
            '',                              // B4
            `${dati.cognome} ${dati.nome}`, // C4
            dati.ambiente,                  // D4
            ...Array(12).fill(''),          // E4 → Q4
            dati.gruppo,                    // R4
            dati.consulente,                // S4
            dati.arredatore                 // T4
        ]];

        try {
            await fetch(apiUrl, {
                method: 'PATCH',
                headers: {
                    'Authorization': `Bearer ${accessToken}`,
                    'Content-Type': 'application/json',
                    'workbook-session-id': sessionId
                },
                body: JSON.stringify({ values: valoriRiga })
            });

            statoElement.innerText = `✅ Dati inseriti nel foglio "${dati.foglio}"!`;
            nominativoForm.reset();
            setTimeout(() => {
                formContainer.classList.add('hidden');
                menuPrincipale.classList.remove('hidden');
            }, 2000);
        } catch (error) {
            statoElement.innerText = `❌ Errore: ${error.message}`;
            console.error(error);
        }
    });

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
