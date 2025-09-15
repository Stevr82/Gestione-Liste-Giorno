const fileIDs = {
  gennaio: 'INSERISCI_ID_DI_GENNAIO',
  febbraio: 'INSERISCI_ID_DI_FEBBRAIO',
  marzo: 'INSERISCI_ID_DI_MARZO',
  aprile: 'INSERISCI_ID_DI_APRILE',
  maggio: 'INSERISCI_ID_DI_MAGGIO',
  giugno: 'INSERISCI_ID_DI_GIUGNO',
  luglio: 'INSERISCI_ID_DI_LUGLIO',
  agosto: 'INSERISCI_ID_DI_AGOSTO',
  settembre: 'INSERISCI_ID_DI_SETTEMBRE',
  ottobre: 'INSERISCI_ID_DI_OTTOBRE',
  novembre: 'INSERISCI_ID_DI_NOVEMBRE',
  dicembre: 'INSERISCI_ID_DI_DICEMBRE'
};

function caricaExcel() {
  const mese = document.getElementById('mese').value;
  const giorno = document.getElementById('giorno').value;
  const fileID = fileIDs[mese];

  if (!fileID || !giorno) {
    document.getElementById('contenuto').innerHTML = "Seleziona mese e giorno validi.";
    return;
  }

  const url = `https://corsproxy.io/?https://drive.google.com/uc?export=download&id=${fileID}`;

  fetch(url)
    .then(res => res.arrayBuffer())
    .then(data => {
      const workbook = XLSX.read(data, { type: 'array' });
      const sheetName = workbook.SheetNames[giorno - 1];
      const sheet = workbook.Sheets[sheetName];

      if (!sheet) {
        document.getElementById('contenuto').innerHTML = "Foglio non trovato per questo giorno.";
        return;
      }

      const html = XLSX.utils.sheet_to_html(sheet);
      document.getElementById('contenuto').innerHTML = html;
    })
    .catch(err => {
      document.getElementById('contenuto').innerHTML = "Errore nel caricamento del file.";
      console.error(err);
    });
}
