const fileIDs = {
  gennaio: '1kEsCsEKl6C9ZGkdOpzLxE2bSMWnkH_r3F34Yvk2zDFU',
  febbraio: '1wh9JZy0kTksXXtu2rvDTRXceh7RQ-NXcdulw6in4AjI',
  marzo: '1SNHIGxX6jtKB2gK03e2m0YO9Oj-wd0KTTHAmyiJG3_E',
  aprile: '1yyH1k3dBSfQn6XiTLrmPJ1KeUORKCnLvYBvanjogVEk',
  maggio: '17ejoWO2hVNekJpwCwbwgcVGmmqwuFqeC6SzoLhMxC6Q',
  giugno: '1ACQDVLh1Whdpm0vmmr1c4ZvNBv-4dFmS_oCDKa1pY3g',
  luglio: '1c_K7PCXRPtn3v5IsYr2qaPoFAgWUiHkc05j7roOD8_E',
  agosto: '1TWk5RWnUn--cIkkFBZFFzBEvdziBoX73GpyPngir1SE',
  settembre: '1Pbwo64rcv0RGzJiDa84ev8hRwUzKydrlYNRGgI2sjaA',
  ottobre: '1M37_o_PvujXozP_461GZgyeGB50X9Mc5RoNi-I5rBbI',
  novembre: '1Xih-zeZPD3Wgo5wSY7IzP9GAhkXSaScU0cJJMFm_xGM',
  dicembre: '1Sd5lssnfPF0okKLMD_mcC7I08lwfc99YdoQL5I7i-9o'
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
