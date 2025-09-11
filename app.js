document.addEventListener('DOMContentLoaded', async () => {

    // Riferimenti agli elementi HTML

    const statoElement = document.getElementById("stato");

    const elencoFogli = document.getElementById("elencoFogli");

    const menuPrincipale = document.querySelector('.pulsanti-container');

    const formContainer = document.getElementById('formInserisciNominativo');

    const btnInserisciNominativo = document.getElementById('btnInserisci');

    const btnIndietro = document.getElementById('btnIndietro');

    const nominativoForm = document.getElementById('nominativoForm');



    // Configurazione e variabili globali

    const msalConfig = {

        auth: {

            clientId: "c3893db8-ca5a-4193-8cfd-08feb16832b1",

            authority: "https://login.microsoftonline.com/common",

            redirectUri: "https://stevr82.github.io/Gestione-Liste-Giorno/"

        }

    };

    const loginRequest = {

        scopes: ["User.Read", "Files.ReadWrite.All"]

    };

    

    // Il tuo account

    const userEmail = "stefano.bresolin.vr@gmail.com";

    

    // Il nome del file

    const fileName = "LISTE GIORNO VERONA 2024.xlsx";



    const msalInstance = new msal.PublicClientApplication(msalConfig);



    // Gestione reindirizzamento e autenticazione

    try {

        const response = await msalInstance.handleRedirectPromise();

        if (response && response.account) {

            msalInstance.setActiveAccount(response.account);

            localStorage.setItem("msalAccount", response.account.homeAccountId);

        }

    } catch (error) {

        console.error("Errore nel redirect:", error);

    }



    const savedAccountId = localStorage.getItem("msalAccount");

    const accounts = msalInstance.getAllAccounts();

    const account = accounts.find(acc => acc.homeAccountId === savedAccountId);



    if (account) {

        msalInstance.setActiveAccount(account);

        await mostraFogliDisponibili();

    } else {

        await msalInstance.loginRedirect(loginRequest);

    }



    // Funzione per ottenere il token di accesso

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

    

    async function findSharedFileId() {

        const accessToken = await getAccessToken();

        if (!accessToken) return null;



        try {

            const url = `https://graph.microsoft.com/v1.0/me/drive/sharedWithMe`;

            const response = await fetch(url, {

                headers: { Authorization: `Bearer ${accessToken}` }

            });



            if (!response.ok) {

                const errorData = await response.json();

                throw new Error(errorData.error.message || `Errore HTTP: ${response.status}`);

            }



            const data = await response.json();

            const file = data.value.find(item => item.name === fileName);



            if (file && file.remoteItem && file.remoteItem.id) {

                return file.remoteItem.id;

            } else {

                throw new Error(`File '${fileName}' non trovato nella cartella 'Condivisi con me'.`);

            }

        } catch (error) {

            console.error("Errore nel recupero dell'ID del file:", error);

            statoElement.innerText = `❌ Errore: ${error.message}`;

            return null;

        }

    }



    // Funzione per mostrare i fogli di lavoro disponibili

    async function mostraFogliDisponibili() {

        const accessToken = await getAccessToken();

        if (!accessToken) return;

        

        const fileId = await findSharedFileId();

        if (!fileId) return;



        const url = `https://graph.microsoft.com/v1.0/me/drive/items/${fileId}/workbook/worksheets`;



        try {

            const response = await fetch(url, {

                headers: { Authorization: `Bearer ${accessToken}` }

            });



            if (!response.ok) {

                const errorData = await response.json();

                throw new Error(errorData.error.message || `Errore HTTP: ${response.status}`);

            }



            const data = await response.json();

            if (data.value && data.value.length > 0) {

                const nomi = data.value.map(ws => ws.name);

                elencoFogli.innerHTML = `📄 Fogli disponibili nel file:<br>${nomi.join(" • ")}`;

            } else {

                elencoFogli.innerText = "⚠️ Nessun foglio trovato nel file.";

            }

        } catch (error) {

            console.error("Errore nel recupero dei fogli:", error);

            elencoFogli.innerText = `❌ Errore: ${error.message}`;

        }

    }



    // Funzione per popolare i dropdown del form

    function popolaDropdown() {

        const giorno = document.getElementById('giorno');

        const mese = document.getElementById('mese');

        const orario = document.getElementById('orario');



        giorno.innerHTML = '';

        for (let i = 1; i <= 31; i++) {

            giorno.innerHTML += `<option value="${i}">${i}</option>`;

        }



        mese.innerHTML = '';

        for (let i = 1; i <= 12; i++) {

            mese.innerHTML += `<option value="${i}">${i}</option>`;

        }



        const orari = ['9:00', '9:30', '10:00', '10:30', '11:00', '11:30', '12:00', '12:30', '13:00', '13:30', '14:00', '14:30', '15:00', '15:30', '16:00', '16:30', '17:00', '17:30', '18:00', '18:30', '19:00', '19:30'];

        orario.innerHTML = '';

        orari.forEach(o => {

            orario.innerHTML += `<option value="${o}">${o}</option>`;

        });

    }



    // Gestione degli eventi sui pulsanti

    btnInserisciNominativo.addEventListener('click', () => {

        menuPrincipale.classList.add('hidden');

        formContainer.classList.remove('hidden');

        popolaDropdown();

    });



    btnIndietro.addEventListener('click', () => {

        formContainer.classList.add('hidden');

        menuPrincipale.classList.remove('hidden');

    });



    // Gestione invio del form

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

        

        const fileId = await findSharedFileId();

        if (!fileId) return;



        const nomiMesi = ['gen', 'feb', 'mar', 'apr', 'mag', 'giu', 'lug', 'ago', 'set', 'ott', 'nov', 'dic'];

        const mese = nomiMesi[parseInt(dati.mese) - 1];

        const worksheetName = `${dati.giorno}-${mese}`;

        const rangeAddress = "A4:T4";

        

        const valoriRiga = [

            [

                dati.orario, '', `${dati.cognome} ${dati.nome}`, dati.ambiente,

                ...Array(12).fill(''),

                dati.gruppo, dati.consulente, dati.arredatore

            ]

        ];



        let sessionId = null;

        try {

            // URL AGGIORNATO per creare la sessione

            const sessionResponse = await fetch(`https://graph.microsoft.com/v1.0/me/drive/items/${fileId}/workbook/createSession`, {

                method: 'POST',

                headers: {

                    'Authorization': `Bearer ${accessToken}`,

                    'Content-Type': 'application/json'

                },

                body: JSON.stringify({ persistChanges: true })

            });



            if (!sessionResponse.ok) {

                const errorData = await sessionResponse.json();

                throw new Error(errorData.error.message || `Errore nella creazione della sessione: ${sessionResponse.status}`);

            }



            sessionId = (await sessionResponse.json()).id;



            // URL AGGIORNATO per inviare i dati

            const apiUrl = `https://graph.microsoft.com/v1.0/me/drive/items/${fileId}/workbook/worksheets('${worksheetName}')/range(address='${rangeAddress}')`;

            const writeResponse = await fetch(apiUrl, {

                method: 'PATCH',

                headers: {

                    'Authorization': `Bearer ${accessToken}`,

                    'Content-Type': 'application/json',

                    'workbook-session-id': sessionId

                },

                body: JSON.stringify({ values: valoriRiga })

            });



            if (!writeResponse.ok) {

                const errorData = await writeResponse.json();

                throw new Error(errorData.error.message || `Errore nell'invio dei dati: ${writeResponse.status}`);

            }



            statoElement.innerText = `✅ Dati inseriti nel foglio "${worksheetName}"!`;

            nominativoForm.reset();

            

            setTimeout(() => {

                formContainer.classList.add('hidden');

                menuPrincipale.classList.remove('hidden');

            }, 2000);



        } catch (error) {

            statoElement.innerText = `❌ Errore: ${error.message}`;

            console.error("Errore completo:", error);

        } finally {

            // Chiusura della sessione di lavoro (importantissimo per evitare blocchi)

            if (sessionId) {

                await fetch(`https://graph.microsoft.com/v1.0/me/drive/items/${fileId}/workbook/closeSession`, {

                    method: 'POST',

                    headers: {

                        'Authorization': `Bearer ${accessToken}`,

                        'workbook-session-id': sessionId

                    }

                }).catch(err => console.error("Errore nella chiusura della sessione:", err));

            }

        }

    });



    // Gestione degli altri pulsanti

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
