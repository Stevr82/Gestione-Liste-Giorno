<!DOCTYPE html>
<html>
<head>
    <title>Gestione Lista</title>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <style>
        /* Stili per l'interfaccia principale */
        body {
            font-family: sans-serif;
            display: flex;
            flex-direction: column;
            justify-content: center;
            align-items: center;
            height: 100vh;
            margin: 0;
            background-color: #f0f0f0;
        }
        .pulsanti-container {
            display: grid;
            grid-template-columns: 1fr 1fr;
            grid-gap: 20px;
            padding: 20px;
        }
        .pulsante {
            padding: 30px;
            font-size: 1.2em;
            font-weight: bold;
            color: white;
            background-color: #0078d4;
            border: none;
            border-radius: 8px;
            cursor: pointer;
            text-align: center;
            transition: background-color 0.3s ease;
        }
        .pulsante:hover {
            background-color: #005a9e;
        }
        h1 {
            color: #333;
            text-align: center;
            margin-bottom: 40px;
        }

        /* Stili per il modulo di inserimento (modal) */
        #formInserisciNominativo {
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            background-color: rgba(0, 0, 0, 0.7);
            justify-content: center;
            align-items: center;
        }
        .modulo-contenuto {
            background-color: white;
            padding: 30px;
            border-radius: 10px;
            width: 90%;
            max-width: 500px;
            box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1);
            display: flex;
            flex-direction: column;
            gap: 15px;
        }
        .modulo-contenuto input, .modulo-contenuto select {
            width: 100%;
            padding: 10px;
            box-sizing: border-box;
            border: 1px solid #ccc;
            border-radius: 5px;
        }
        .modulo-contenuto .pulsanti-form {
            display: flex;
            justify-content: space-between;
            margin-top: 20px;
        }
        .pulsanti-form button {
            width: 48%;
            padding: 12px;
            border-radius: 5px;
            font-weight: bold;
        }
        #btnInserisciDati {
            background-color: #28a745;
            color: white;
            border: none;
        }
        #btnIndietro {
            background-color: #dc3545;
            color: white;
            border: none;
        }
        #stato {
            text-align: center;
            margin-top: 20px;
            font-weight: bold;
            color: #333;
        }

        /* Classe per nascondere gli elementi in modo definitivo */
        .hidden {
            display: none !important;
        }
    </style>
</head>
<body>
    <h1>Menu Principale</h1>
    <div class="pulsanti-container">
        <button class="pulsante" id="btnInserisci">INSERISCI NOMINATIVO</button>
        <button class="pulsante" id="btnRicerca">RICERCA NOMINATIVO</button>
        <button class="pulsante" id="btnVisualizza">VISUALIZZA LISTA GIORNO</button>
        <button class="pulsante" id="btnCompila">COMPILA LISTA GIORNO</button>
    </div>

    <!-- Messaggio di stato -->
    <p id="stato"></p>

    <!-- Modulo di Inserimento Nominativo (inizialmente nascosto) -->
    <div id="formInserisciNominativo" class="hidden">
        <div class="modulo-contenuto">
            <h2>Inserisci Nominativo</h2>
            <form id="nominativoForm">
                <label for="cognome">Cognome:</label>
                <input type="text" id="cognome" name="cognome" required>

                <label for="nome">Nome:</label>
                <input type="text" id="nome" name="nome" required>

                <label for="ambiente">Ambiente:</label>
                <input type="text" id="ambiente" name="ambiente">

                <label for="gruppo">Gruppo:</label>
                <input type="text" id="gruppo" name="gruppo">

                <label for="consulente">Consulente:</label>
                <input type="text" id="consulente" name="consulente">

                <label for="arredatore">Arredatore:</label>
                <input type="text" id="arredatore" name="arredatore">

                <label for="giorno">Giorno:</label>
                <select id="giorno" name="giorno"></select>

                <label for="mese">Mese:</label>
                <select id="mese" name="mese"></select>

                <label for="orario">Orario:</label>
                <select id="orario" name="orario"></select>

                <div class="pulsanti-form">
                    <button type="submit" id="btnInserisciDati">INSERISCI</button>
                    <button type="button" id="btnIndietro">INDIETRO</button>
                </div>
            </form>
        </div>
    </div>

    <!-- Script per l'autenticazione MSAL e la nostra logica -->
    <script src="https://alcdn.msauth.net/browser/2.16.0/js/msal-browser.js"></script>
    <script src="app.js"></script>
</body>
</html>
