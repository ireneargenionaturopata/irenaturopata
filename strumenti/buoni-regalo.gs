/**
 * BUONI REGALO – Irene Argenio Naturopata
 *
 * Cosa fa:
 * 1. Riceve dal sito ogni richiesta di buono e la scrive nel foglio "Buoni regalo – Irene Argenio",
 *    salvando nella cartella "Buoni regalo" di Drive l'immagine del buono già compilato e il suo PDF.
 * 2. Quando nel foglio spunti la casella "Pagato":
 *    - se il cliente ha scelto l'EMAIL, gli manda subito il buono in PDF (e l'immagine);
 *    - se ha scelto WHATSAPP, trovi nella colonna "Buono (immagine)" il link da aprire e inoltrare;
 *    - in ogni caso scrive la scadenza (12 mesi da oggi).
 *
 * Installazione (una volta sola): vedi la guida che ti ha mandato Claude.
 * 1) Esegui la funzione "installa" e dai i permessi.
 * 2) Distribuisci come "App web" (Esegui come: Me · Chi ha accesso: Chiunque) e manda il link a Claude.
 */

const CONFIG = {
  NOME: 'Irene Argenio',
  FIRMA: 'Irene Argenio · Naturopata',
  EMAIL: 'argenioirene@gmail.com',
  WHATSAPP: '392 893 8682',
  SITO: 'https://ireneargenio.it/',
  MESI_VALIDITA: 12,
  NOME_FOGLIO: 'Buoni regalo – Irene Argenio',
  NOME_CARTELLA: 'Buoni regalo'
};

const COLONNE = [
  'Codice', 'Richiesta del', 'Stato', 'Pagato', 'Per', 'Da parte di', 'Regalo', 'Valore €',
  'Studio', 'Occasione', 'Dedica', 'Consegna', 'Email cliente', 'Pagamento', 'Scadenza',
  'Buono (PDF)', 'Buono (immagine)', 'Inviato il', 'Usato il', 'Note'
];
const C = {}; COLONNE.forEach((n, i) => C[n] = i + 1);

/* ---------- 1. INSTALLAZIONE ---------- */
function installa() {
  const props = PropertiesService.getScriptProperties();
  let ss;
  if (props.getProperty('FOGLIO_ID')) {
    ss = SpreadsheetApp.openById(props.getProperty('FOGLIO_ID'));
  } else {
    ss = SpreadsheetApp.create(CONFIG.NOME_FOGLIO);
    props.setProperty('FOGLIO_ID', ss.getId());
  }
  const sh = ss.getSheets()[0];
  sh.setName('Buoni');
  sh.getRange(1, 1, 1, COLONNE.length).setValues([COLONNE])
    .setFontWeight('bold').setBackground('#E3EDDD').setFontColor('#2E3B33');
  sh.setFrozenRows(1);
  sh.setFrozenColumns(1);
  sh.setColumnWidth(C['Codice'], 140);
  sh.setColumnWidth(C['Stato'], 130);
  sh.setColumnWidth(C['Regalo'], 240);
  sh.setColumnWidth(C['Dedica'], 260);
  sh.getRange(2, C['Pagato'], 500, 1).insertCheckboxes();
  sh.getRange(2, C['Richiesta del'], 500, 1).setNumberFormat('dd/mm/yyyy HH:mm');
  sh.getRange(2, C['Scadenza'], 500, 1).setNumberFormat('dd/mm/yyyy');
  sh.getRange(2, C['Inviato il'], 500, 1).setNumberFormat('dd/mm/yyyy HH:mm');
  sh.getRange(2, C['Usato il'], 500, 1).setNumberFormat('dd/mm/yyyy');

  // colori dello stato
  const stato = sh.getRange(2, C['Stato'], 500, 1);
  const r = (testo, sfondo, colore) => SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo(testo).setBackground(sfondo).setFontColor(colore).setRanges([stato]).build();
  sh.setConditionalFormatRules([
    r('Da pagare', '#FFF4DD', '#8A5A00'),
    r('Pagato', '#E3EDDD', '#2E5A34'),
    r('Inviato', '#DDEBF7', '#1F4E79'),
    r('Usato', '#EEEEEE', '#666666'),
    r('Scaduto', '#FBE1E1', '#9C1C1C')
  ]);

  // cartella per immagini e PDF
  if (!props.getProperty('CARTELLA_ID')) {
    const folder = DriveApp.createFolder(CONFIG.NOME_CARTELLA);
    props.setProperty('CARTELLA_ID', folder.getId());
  }

  // la spunta "Pagato" fa partire l'invio
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  PropertiesService.getScriptProperties().deleteProperty('TRIGGER_GIORNALIERO');
  ScriptApp.newTrigger('quandoModifichi').forSpreadsheet(ss).onEdit().create();

  // prova d'invio email a te stessa, così Google chiede subito tutti i permessi
  MailApp.sendEmail(CONFIG.EMAIL, 'Buoni regalo: installazione riuscita ✅',
    'Il foglio dei buoni regalo è pronto:\n' + ss.getUrl() +
    '\n\nOra distribuisci lo script come App web e manda il link a Claude.');

  Logger.log('FOGLIO: ' + ss.getUrl());
}

/* ---------- 2. RICEVE LE RICHIESTE DAL SITO ---------- */
function doPost(e) {
  const p = (e && e.parameter) || {};
  if (p.website) return risposta('ok'); // trappola anti-spam: i robot compilano anche i campi nascosti
  const codice = String(p.codice || '').replace(/[^A-Z0-9-]/gi, '').slice(0, 20);
  if (!codice) return risposta('manca il codice');

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = foglio();
    const riga = trovaRiga(sh, codice) || primaRigaLibera(sh);

    // immagine e PDF del buono (il sito li manda già compilati).
    // Se qualcosa va storto con i file, la riga viene scritta lo stesso e l'errore finisce nelle Note.
    let linkPdf = '', linkImg = '', nota = '';
    if (p.immagine && p.immagine.indexOf('base64,') > 0) {
      try {
        const folder = DriveApp.getFolderById(PropertiesService.getScriptProperties().getProperty('CARTELLA_ID'));
        ['.jpg', '.pdf'].forEach(ext => { const it = folder.getFilesByName(codice + ext); while (it.hasNext()) it.next().setTrashed(true); });
        const dati = Utilities.base64Decode(p.immagine.split('base64,')[1]);
        const img = folder.createFile(Utilities.newBlob(dati, 'image/jpeg', codice + '.jpg'));
        linkImg = img.getUrl();
        // PDF: pagina orizzontale con il buono a tutta larghezza
        const html = '<html><head><style>@page{size:A4 landscape;margin:12mm}body{margin:0}img{width:100%}</style></head>' +
                     '<body><img src="' + p.immagine + '"></body></html>';
        const pdf = folder.createFile(Utilities.newBlob(html, 'text/html', codice + '.html').getAs('application/pdf').setName(codice + '.pdf'));
        linkPdf = pdf.getUrl();
      } catch (err) {
        nota = 'Problema con i file del buono: ' + err.message;
      }
    }

    const t = (s, n) => String(s || '').slice(0, n || 200);
    const valori = {
      'Codice': codice,
      'Richiesta del': new Date(),
      'Per': t(p.per, 80),
      'Da parte di': t(p.da, 80),
      'Regalo': t(p.regalo, 120),
      'Valore €': Number(String(p.valore || '').replace(/[^\d]/g, '')) || '',
      'Studio': t(p.studio, 60),
      'Occasione': t(p.occasione, 60),
      'Dedica': t(p.dedica, 120),
      'Consegna': t(p.consegna, 40),
      'Email cliente': t(p.email, 120),
      'Pagamento': t(p.pagamento, 40),
      'Buono (PDF)': linkPdf,
      'Buono (immagine)': linkImg,
      'Note': nota
    };
    Object.keys(valori).forEach(k => sh.getRange(riga, C[k]).setValue(valori[k]));
    sh.getRange(riga, C['Pagato']).insertCheckboxes();
    impostaStato(sh, riga);
  } finally {
    lock.releaseLock();
  }
  return risposta('ok');
}

/* ---------- 3. QUANDO SPUNTI "PAGATO" ---------- */
function quandoModifichi(e) {
  const sh = e.range.getSheet();
  if (sh.getName() !== 'Buoni') return;
  const col = e.range.getColumn(), riga = e.range.getRow();
  if (riga < 2) return;
  if (col === C['Usato il']) { impostaStato(sh, riga); return; }
  if (col !== C['Pagato']) return;
  if (sh.getRange(riga, C['Pagato']).getValue() !== true) { impostaStato(sh, riga); return; }

  if (!sh.getRange(riga, C['Scadenza']).getValue()) {
    const scad = new Date();
    scad.setMonth(scad.getMonth() + CONFIG.MESI_VALIDITA);
    sh.getRange(riga, C['Scadenza']).setValue(scad);
  }

  const consegna = String(sh.getRange(riga, C['Consegna']).getValue()).toLowerCase();
  const giaInviato = sh.getRange(riga, C['Inviato il']).getValue();

  if (consegna.indexOf('email') >= 0 && !giaInviato) {
    try {
      inviaEmail(sh, riga);
      e.source.toast('Buono inviato per email ✅', 'Buoni regalo', 6);
    } catch (err) {
      sh.getRange(riga, C['Note']).setValue('Email NON inviata: ' + err.message);
      e.source.toast('Email non inviata: guarda la colonna Note', 'Buoni regalo', 10);
    }
  } else if (consegna.indexOf('whatsapp') >= 0) {
    e.source.toast('Apri il link in "Buono (immagine)" e mandalo su WhatsApp 📲', 'Buoni regalo', 10);
  } else if (consegna.indexOf('studio') >= 0) {
    e.source.toast('Ritiro in studio: prepara la confezione 🎁', 'Buoni regalo', 8);
  }
  impostaStato(sh, riga);
}

function inviaEmail(sh, riga) {
  const get = n => sh.getRange(riga, C[n]).getValue();
  const email = String(get('Email cliente')).trim();
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) throw new Error('indirizzo email mancante o sbagliato');
  const codice = get('Codice'), per = get('Per'), da = get('Da parte di'), regalo = get('Regalo');
  const scad = Utilities.formatDate(new Date(get('Scadenza')), 'Europe/Rome', 'dd/MM/yyyy');

  const folder = DriveApp.getFolderById(PropertiesService.getScriptProperties().getProperty('CARTELLA_ID'));
  const allegati = [];
  const f1 = folder.getFilesByName(codice + '.pdf'); if (f1.hasNext()) allegati.push(f1.next().getBlob());
  const f2 = folder.getFilesByName(codice + '.jpg'); if (f2.hasNext()) allegati.push(f2.next().getBlob());

  const oggetto = 'Il tuo buono regalo 🎁' + (per ? ' per ' + per : '');
  const testo =
    'Ciao,\n\ngrazie per aver scelto un mio buono regalo!\n\n' +
    'In allegato trovi il buono' + (per ? ' per ' + per : '') + (da ? ', da parte di ' + da : '') +
    ', pronto da stampare o da inoltrare.\n\n' +
    'Regalo: ' + regalo + '\nCodice: ' + codice + '\nValido fino al: ' + scad + '\n\n' +
    'Per prenotare basta scrivermi su WhatsApp al ' + CONFIG.WHATSAPP + ' o prenotare dal sito:\n' + CONFIG.SITO + '\n\n' +
    'Un abbraccio,\n' + CONFIG.FIRMA;
  const html =
    '<div style="font-family:Arial,sans-serif;color:#2E3B33;font-size:15px;line-height:1.6;max-width:560px">' +
    '<p>Ciao,</p><p>grazie per aver scelto un mio buono regalo! 🌿</p>' +
    '<p>In allegato trovi il buono' + (per ? ' per <b>' + esc(per) + '</b>' : '') + (da ? ', da parte di <b>' + esc(da) + '</b>' : '') +
    ', pronto da stampare o da inoltrare.</p>' +
    '<p style="background:#F4F8F1;border-left:4px solid #A9C5A0;padding:10px 14px">' +
    '<b>Regalo:</b> ' + esc(regalo) + '<br><b>Codice:</b> ' + esc(codice) + '<br><b>Valido fino al:</b> ' + scad + '</p>' +
    '<p>Per prenotare basta scrivermi su WhatsApp al <b>' + CONFIG.WHATSAPP + '</b> o prenotare dal <a href="' + CONFIG.SITO + '">sito</a>.</p>' +
    '<p>Un abbraccio,<br>' + CONFIG.FIRMA + '</p></div>';

  MailApp.sendEmail({ to: email, subject: oggetto, body: testo, htmlBody: html, name: CONFIG.NOME, replyTo: CONFIG.EMAIL, attachments: allegati });
  sh.getRange(riga, C['Inviato il']).setValue(new Date());
}

/* ---------- utilità ---------- */
function foglio() {
  return SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('FOGLIO_ID')).getSheetByName('Buoni');
}
// prima riga con la colonna "Codice" vuota (le caselle "Pagato" già pronte non contano)
function primaRigaLibera(sh) {
  const n = Math.max(sh.getLastRow(), 2);
  const codici = sh.getRange(2, 1, n - 1, 1).getValues();
  let ultima = 1;
  codici.forEach((r, i) => { if (String(r[0]).trim() !== '') ultima = i + 2; });
  return ultima + 1;
}
function trovaRiga(sh, codice) {
  const n = sh.getLastRow();
  if (n < 2) return 0;
  const codici = sh.getRange(2, 1, n - 1, 1).getValues();
  for (let i = 0; i < codici.length; i++) if (codici[i][0] === codice) return i + 2;
  return 0;
}
function impostaStato(sh, riga) {
  // stato scritto come testo (niente formule: così non dà errori con le impostazioni italiane)
  const v = n => sh.getRange(riga, C[n]).getValue();
  let stato = 'Da pagare';
  const scad = v('Scadenza'), oggi = new Date(); oggi.setHours(0, 0, 0, 0);
  if (v('Usato il')) stato = 'Usato';
  else if (scad && new Date(scad) < oggi) stato = 'Scaduto';
  else if (v('Inviato il')) stato = 'Inviato';
  else if (v('Pagato') === true) stato = 'Pagato';
  sh.getRange(riga, C['Stato']).setValue(stato);
  controllaScadenzeOgniGiorno();
}
// una volta al giorno ricontrolla gli stati (per segnare i buoni scaduti)
function controllaScadenzeOgniGiorno() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('TRIGGER_GIORNALIERO')) return;
  ScriptApp.newTrigger('aggiornaTuttiGliStati').timeBased().everyDays(1).atHour(6).create();
  props.setProperty('TRIGGER_GIORNALIERO', 'si');
}
function aggiornaTuttiGliStati() {
  const sh = foglio(), n = sh.getLastRow();
  for (let r = 2; r <= n; r++) if (String(sh.getRange(r, C['Codice']).getValue()).trim()) impostaStato(sh, r);
}
function esc(s) { return String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c])); }
function risposta(t) { return ContentService.createTextOutput(t); }
