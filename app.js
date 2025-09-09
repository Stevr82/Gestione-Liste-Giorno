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
    for (let i = 1; i <= 31; i++) {
        const option = document.createElement('option');
        option.value = i;
        option.textContent = i;
        giornoDropdown.appendChild(option);
    }
    
    // Popola i mesi (da 1 a 12)
    for (let i = 1; i <= 12; i++) {
        const option = document.createElement('option');
        option.value = i;
        option.textContent = i;
        meseDropdown.appendChild(option);
    }
    
    // Popola gli orari
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

        const fileId = "A3856CCE-D8CC-4C35-92E3-02EAB1E3B368"; 
        const worksheetName = "Foglio1"; // Assicurati che il nome sia corretto
        // In questo caso, presupponiamo che i tuoi dati siano in una tabella chiamata 'Tabella1'
        const tableName = "Tabella1"; 

        const apiUrl = `https://graph.microsoft.com/v1.0/me/drive/items/${fileId}/workbook/worksheets('${worksheetName}')/tables('${tableName}')/rows`;

        // L'ordine dei valori deve corrispondere all'ordine delle colonne della tabella
        const valoriRiga = [
            dati.cognome,
            dati.nome,
            dati.ambiente,
            dati.gruppo,
            dati.consulente,
            dati.arredatore,
            dati.giorno,
            dati.mese,
            dati.orario
        ];

        const body = {
            values: [valoriRiga]
        };

        // Effettua la richiesta POST per aggiungere la riga
        await fetch(apiUrl, {
            method: 'POST',
            headers: {
                'Authorization': `Bearer ${accessToken}`,
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(body)
        });

        statoElement.innerText = "Dati inseriti con successo!";

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
    const giornoDropdown = document.getElementById('giorno');
    const meseDropdown = document.getElementById('mese');
    const orarioDropdown = document.getElementById('orario');

    // Popola i giorni (da 1 a 31)
    for (let i = 1; i <= 31; i++) {
        const option = document.createElement('option');
        option.value = i;
        option.textContent = i;
        giornoDropdown.appendChild(option);
    }
    
    // Popola i mesi (da 1 a 12)
    for (let i = 1; i <= 12; i++) {
        const option = document.createElement('option');
        option.value = i;
        option.textContent = i;
        meseDropdown.appendChild(option);
    }
    
    // Popola gli orari
    const orari = ['9:00', '9:30', '10:00', '10:30', '11:00', '11:30', '12:00', '12:30', '13:00', '13:30', '14:00', '14:30', '15:00', '15:30', '16:00', '16:30', '17:00', '17:30', '18:00', '18:30', '19:00', '19:30'];
    orari.forEach(orario => {
        const option = document.createElement('option');
        option.value = orario;
        option.textContent = orario;
        orarioDropdown.appendChild(option);
    });
}

// Aggiungi un ascoltatore al pulsante "Inserisci Nominativo"
btnInserisciNominativo.addEventListener('click', () => {
    // Mostra il modulo di inserimento
    formContainer.style.display = 'flex';
    menuPrincipale.style.display = 'none';
    popolaDropdown(); // Popola le caselle a tendina
});

// Aggiungi un ascoltatore al pulsante "Indietro"
btnIndietro.addEventListener('click', () => {
    // Nasconde il modulo di inserimento
    formContainer.style.display = 'none';
    menuPrincipale.style.display = 'grid';
});

// Gestisci il click sul pulsante "INSERISCI"
nominativoForm.addEventListener('submit', async (e) => {
    e.preventDefault(); // Previene il ricaricamento della pagina
    
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

    alert("Dati raccolti. Ora dovremmo inviarli al file Excel: " + JSON.stringify(dati));
    // La logica per inviare questi dati al file Excel va qui
    // e userà la Microsoft Graph API per aggiungere una nuova riga

    // Riporta l'interfaccia al menu principale dopo l'invio
    // formContainer.style.display = 'none';
    // menuPrincipale.style.display = 'grid';
});

// Aggiungi un ascoltatore anche per gli altri pulsanti (come prima)
document.getElementById("btnRicerca").addEventListener("click", () => {
    alert("Funzionalità 'RICERCA NOMINATIVO' ancora da implementare.");
});
document.getElementById("btnVisualizza").addEventListener("click", () => {
    alert("Funzionalità 'VISUALIZZA LISTA GIORNO' ancora da implementare.");
});
document.getElementById("btnCompila").addEventListener("click", async () => {
    alert("Funzionalità 'COMPILA LISTA GIORNO' ancora da implementare. Ora si usa 'Inserisci Nominativo'.");
});

