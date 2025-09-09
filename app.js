// Passo 1: Configurazione di MSAL.js
// Sostituisci "IL_TUO_ID_CLIENT" con l'ID applicazione che hai ottenuto da Azure.
const msalConfig = {
    auth: {
        clientId: "c3893db8-ca5a-4193-8cfd-08feb16832b1", // <-- SOSTITUISCI QUI!
        authority: "https://login.microsoftonline.com/common",
        // Sostituisci l'URI di reindirizzamento se necessario.
        redirectUri: "http://localhost:8080" 
    }
};

const msalInstance = new msal.PublicClientApplication(msalConfig);
const loginRequest = {
    scopes: ["User.Read", "Files.ReadWrite"]
};

// Passo 2: Funzione per aggiornare la cella Excel
// Questa funzione gestisce l'autenticazione e la chiamata API
async function updateExcelCell(valore) {
    try {
        // Tentativo di acquisire un token in modo silenzioso
        const response = await msalInstance.acquireTokenSilent(loginRequest);
        const accessToken = response.accessToken;

        // ID del tuo file Excel (estratto dal link)
        // Questo è l'ID univoco che hai trovato prima
        const fileId = "A3856CCE-D8CC-4C35-92E3-02EAB1E3B368"; 
        
        // Dati da modificare: nome del foglio e cella
        // ASSICURATI DI CAMBIARE QUESTE VARIABILI PERCHÉ SIANO CORRETTE PER IL TUO FILE!
        const worksheetName = "Foglio1"; // Es. 'Dati' o 'Sheet1'
        const cellAddress = "A1";        // Es. 'C5' o 'B2'

        // URL dell'API di Microsoft Graph per modificare la cella
        const apiUrl = `https://graph.microsoft.com/v1.0/me/drive/items/${fileId}/workbook/worksheets('${worksheetName}')/range(address='${cellAddress}')`;

        // Dati della richiesta in formato JSON
        const body = {
            values: [[valore]]
        };

        // Effettua la richiesta PATCH per aggiornare la cella
        await fetch(apiUrl, {
            method: 'PATCH',
            headers: {
                'Authorization': `Bearer ${accessToken}`,
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(body)
        });

    } catch (error) {
        // Se l'acquisizione silenziosa del token fallisce, prova con un popup
        if (error instanceof msal.InteractionRequiredAuthError) {
            await msalInstance.acquireTokenPopup(loginRequest);
            // Dopo il popup, la funzione si può richiamare da sola per riprovare
            return updateExcelCell(valore);
        } else {
            console.error("Errore durante l'aggiornamento del file Excel:", error);
            document.getElementById("stato").innerText = `Errore: ${error.message}`;
            throw error;
        }
    }
}

// Passo 3: Gestione dell'invio del form
document.getElementById("excelForm").addEventListener("submit", async (e) => {
    e.preventDefault();
    const valore = document.getElementById("valore").value;
    document.getElementById("stato").innerText = "Sto aggiornando il file...";
    
    try {
        // Chiede all'utente di loggarsi (se non lo è già) e poi chiama la funzione
        await msalInstance.loginPopup(loginRequest);
        await updateExcelCell(valore);
        document.getElementById("stato").innerText = "File aggiornato con successo!";
    } catch (error) {
        console.error("Errore durante l'accesso o l'aggiornamento:", error);
        document.getElementById("stato").innerText = `Errore: ${error.message}`;
    }
});
