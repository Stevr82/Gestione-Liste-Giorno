// Configurazione di MSAL (usa il tuo ID applicazione!)
const msalConfig = {
    auth: {
        clientId: "IL_TUO_ID_CLIENT", // <-- SOSTITUISCI QUI!
        authority: "https://login.microsoftonline.com/common",
        redirectUri: "http://localhost:8080" 
    }
};

const msalInstance = new msal.PublicClientApplication(msalConfig);
const loginRequest = {
    scopes: ["User.Read", "Files.ReadWrite"]
};

// Funzione per aggiornare la cella Excel
// Questa funzione è la stessa che avevamo, ma ora è chiamata da un pulsante
async function updateExcelCell(valore) {
    try {
        const response = await msalInstance.acquireTokenSilent(loginRequest);
        const accessToken = response.accessToken;

        const fileId = "A3856CCE-D8CC-4C35-92E3-02EAB1E3B368"; 
        const worksheetName = "Foglio1"; // Assicurati che il nome sia corretto
        const cellAddress = "A1";        // Assicurati che la cella sia corretta

        const apiUrl = `https://graph.microsoft.com/v1.0/me/drive/items/${fileId}/workbook/worksheets('${worksheetName}')/range(address='${cellAddress}')`;

        const body = {
            values: [[valore]]
        };

        await fetch(apiUrl, {
            method: 'PATCH',
            headers: {
                'Authorization': `Bearer ${accessToken}`,
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(body)
        });

    } catch (error) {
        if (error instanceof msal.InteractionRequiredAuthError) {
            await msalInstance.acquireTokenPopup(loginRequest);
            return updateExcelCell(valore);
        } else {
            console.error("Errore durante l'aggiornamento del file Excel:", error);
            alert(`Errore: ${error.message}`);
            throw error;
        }
    }
}

// Gestione dei click sui pulsanti
document.getElementById("btnInserisci").addEventListener("click", () => {
    alert("Funzionalità 'INSERISCI NOMINATIVO' ancora da implementare.");
});

document.getElementById("btnRicerca").addEventListener("click", () => {
    alert("Funzionalità 'RICERCA NOMINATIVO' ancora da implementare.");
});

document.getElementById("btnVisualizza").addEventListener("click", () => {
    alert("Funzionalità 'VISUALIZZA LISTA GIORNO' ancora da implementare.");
});

// Questo è il pulsante che usa la nostra logica esistente
document.getElementById("btnCompila").addEventListener("click", async () => {
    const stato = document.getElementById("stato");
    
    // Chiede all'utente di inserire il valore
    const valore = prompt("Inserisci il valore da scrivere:");

    if (valore) {
        stato.innerText = "Sto aggiornando il file...";
        try {
            await msalInstance.loginPopup(loginRequest);
            await updateExcelCell(valore);
            stato.innerText = "File aggiornato con successo!";
        } catch (error) {
            console.error("Errore durante l'accesso o l'aggiornamento:", error);
            stato.innerText = `Errore: ${error.message}`;
        }
    } else {
        stato.innerText = "Operazione annullata.";
    }
});
