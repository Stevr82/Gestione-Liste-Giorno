// Aspetta che il DOM sia completamente caricato prima di eseguire lo script
document.addEventListener('DOMContentLoaded', () => {

    console.log("Il tuo script è stato caricato!"); // Messaggio di debug

    // Configurazione di MSAL (con il tuo ID applicazione)
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

    // Seleziona gli elementi del DOM
    const menuPrincipale = document.querySelector('.pulsanti-container');
    const formContainer = document.getElementById('formInserisciNominativo');
    const btnInserisciNominativo = document.getElementById('btnInserisci');
    const btnIndietro = document.getElementById('btnIndietro');
    const nominativoForm = document.getElementById('nominativoForm');
    const statoElement = document.getElementById("stato");

    // Funzione per riempire le caselle a tendina
    function popolaDropdown() {
        const giornoDropdown = document.getElementById('giorno');
        const meseDropdown = document.getElementById('mese');
        const orarioDropdown = document.getElementById('orario');

        // Popola i giorni (da 1 a 31)
        giornoDropdown.innerHTML = '';
        for (let i = 1; i <= 31; i++) {
            const option = document.createElement('option');
            option.value = i;
            option.textContent = i;
            giornoDropdown.appendChild(option);
        }
        
        // Popola i mesi (da 1 a 12)
        meseDropdown.innerHTML = '';
        for (let i = 1; i <= 12; i++) {
            const option = document.createElement('option');
            option.value = i;
            option.textContent = i;
            meseDropdown.appendChild(option);
        }
        
        // Popola gli orari
        orarioDropdown.innerHTML = '';
        const orari = ['9:00', '9:30', '10:00', '10:30', '11:00', '11:30', '12:00', '12:30', '13:00', '13:30', '14:00', '14:30', '15:00', '15:30', '16:00', '16:30', '17:00', '17:30', '18:00', '18:30', '19:00', '19:30'];
        orari.forEach(orario => {
            const option = document.createElement('option');
            option.value = orario;
            option.textContent = orario;
            orarioDropdown.appendChild(option);
        });
    }

    // Event listener per mostrare il modulo di inserimento
    btnInserisciNominativo.addEventListener('click', () => {
        console.log("Pulsante 'Inserisci Nominativo' cliccato!"); // Messaggio di debug
        formContainer.style.display = 'flex';
        menuPrincipale.style.display = 'none';
        popolaDropdown(); // Popola le caselle a tendina
    });

    // Event listener per tornare al menu principale
    btnIndietro.addEventListener('click', () => {
        formContainer.style.display = 'none';
        menuPrincipale.style.display = 'grid';
    });

    // Gestisce l'invio del form per inserire i dati in Excel
    nominativoForm.addEventListener('submit', async (e) => {
        e.preventDefault(); 
        
        // Raccogli i dati dal form
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

        statoElement.innerText = "Sto inviando i dati al file Excel...";

        try {
            const response = await msalInstance.acquireTokenSilent(loginRequest);
            const accessToken = response.accessToken;

            // ID del file
            const fileId = "A3856CCE-D8CC-4C35-92E3-02EAB1E3B368"; 

            // Mappa per i nomi dei mesi
            const nomiMesi = [
                'gen', 'feb', 'mar', 'apr', 'mag', 'giu',
                'lug', 'ago', 'set', 'ott', 'nov', 'dic'
            ];

            // Ottieni il nome del foglio di lavoro basato sulla data
            const mese = nomiMesi[parseInt(dati.mese) - 1];
            const worksheetName = `${dati.giorno}-${mese}`;
            
            // Per scrivere in un intervallo di celle, devi specificare la posizione esatta.
            // Qui usiamo un intervallo che inizia dalla riga 4 e copre le colonne necessarie.
            // Questo sovrascriverà i dati esistenti in quelle celle.
            const rangeAddress = "A4:T4";

            // Costruisci l'URL dell'API con il nome del foglio dinamico e l'indirizzo del range
            const apiUrl = `https://graph.microsoft.com/v1.0/me/drive/items/${fileId}/workbook/worksheets('${worksheetName}')/range(address='${rangeAddress}')`;

            // L'ordine dei valori deve corrispondere all'ordine delle colonne nell'intervallo
            // (A, B, C, D, ..., R, S, T)
            const valoriRiga = [
                [dati.orario, '', `${dati.cognome} ${dati.nome}`, dati.ambiente, ...Array(13).fill(''), dati.gruppo, dati.consulente, dati.arredatore]
            ];
            
            const body = {
                values: valoriRiga
            };

            // Effettua la richiesta POST per aggiornare le celle
            await fetch(apiUrl, {
                method: 'PATCH', // Usiamo PATCH per aggiornare i dati
                headers: {
                    'Authorization': `Bearer ${accessToken}`,
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify(body)
            });

            statoElement.innerText = `Dati inseriti con successo nel foglio "${worksheetName}"!`;

            // Resetta il form e nascondi il modulo dopo un breve ritardo
            nominativoForm.reset();
            setTimeout(() => {
                formContainer.style.display = 'none';
                menuPrincipale.style.display = 'grid';
            }, 2000); 

        } catch (error) {
            console.error("Errore durante l'invio dei dati:", error);
            statoElement.innerText = `Errore: ${error.message}`;
            if (error instanceof msal.InteractionRequiredAuthError) {
                 alert("L'accesso è scaduto. Riprova.");
                 msalInstance.acquireTokenPopup(loginRequest);
            }
        }
    });

    // Aggiungi un ascoltatore per gli altri pulsanti, con messaggi di funzionalità non implementata
    document.getElementById("btnRicerca").addEventListener("click", () => {
        statoElement.innerText = "Funzionalità 'RICERCA NOMINATIVO' ancora da implementare.";
    });
    document.getElementById("btnVisualizza").addEventListener("click", () => {
        statoElement.innerText = "Funzionalità 'VISUALIZZA LISTA GIORNO' ancora da implementare.";
    });
    document.getElementById("btnCompila").addEventListener("click", async () => {
        statoElement.innerText = "Funzionalità 'COMPILA LISTA GIORNO' ancora da implementare.";
    });
});
