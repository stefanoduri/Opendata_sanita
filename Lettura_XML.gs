// esempio file http://trasparenza.asl1abruzzo.it/archiviofile/asl1abruzzo/ANAC/avcpLegge190.xml

var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheetLotti = ss.getSheetByName('Lotti');
var sheetPartecipazioni = ss.getSheetByName('Partecipazioni');
var sheetMetadati = ss.getSheetByName('Metadati');
var sheetAggiudicazioni = ss.getSheetByName('Aggiudicazioni');
var headersLotti = ['id', 'cig', 'oggetto', 'codiceFiscaleProp', 'denominazione', 'sceltaContraente', 'importoAggiudicazione', 'importoSommeLiquidate', 'dataInizio', 'dataUltimazione', 'bagOfWords' ];
var headersPartecipazioni = ['lottoId', 'id', 'cf', 'ragioneSociale', 'gruppoId', 'ruolo', 'aggiudicatario' ];
var headersMetadata = ['titolo', 'abstract', 'dataPubbicazioneDataset', 'entePubblicatore', 'dataUltimoAggiornamentoDataset', 
    'annoRiferimento', 'urlFile', 'licenza'];
// colonna importi nel foglio lotti
var colImporti = 9;
var ssStopwords = ss.getSheetByName('Stopwords');
var stopwords = ssStopwords.getRange(1,1,ssStopwords.getLastRow()).getValues();

var rsUniche = {};

function onOpen() {
  // Add a custom menu to the spreadsheet.
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
    .createMenu('Lettura XML')
    .addItem('Start', 'setup')
    .addToUi();
}

function setup() {
  // http://trasparenza.asl1abruzzo.it/archiviofile/asl1abruzzo/ANAC/avcpLegge190.xml
  var ui = SpreadsheetApp.getUi();
  var url = ui.prompt('URL da analizzare').getResponseText();
  
  if(!url) url='http://trasparenza.asl1abruzzo.it/archiviofile/asl1abruzzo/ANAC/avcpLegge190.xml';
  
  sheetLotti.clear().setFrozenRows(1);
  sheetLotti.getRange(1, 1, 1,headersLotti.length).setValues([headersLotti]);
  sheetLotti.getRange('A:F').setNumberFormat('@STRING@');
  sheetLotti.getRange('G:H').setNumberFormat('0.00');

  sheetPartecipazioni.clear().setFrozenRows(1);
  sheetPartecipazioni.getRange(1, 1, 1,headersPartecipazioni.length).setValues([headersPartecipazioni]);
  sheetPartecipazioni.getRange('A:E').setNumberFormat('@STRING@');
  
  sheetMetadati.clear().setFrozenColumns(1);
    
  ss.toast('Leggo il file xml');
  parseXml(url)
}


function parseXml(url) {
  var xml = UrlFetchApp.fetch(url).getContentText();
  if (! xml) {
    ss.toast('File illeggibile');
    return;
  }
  
  var document = XmlService.parse(xml);
  var metadata = document.getRootElement().getChild('metadata');
  for (var i=0, max_i=headersMetadata.length; i<max_i; i++) {
    var fld = headersMetadata[i];
    sheetMetadati.appendRow([fld,metadata.getChild(fld).getText()]);
  }
    
  var data = document.getRootElement().getChild('data');
  var items = data.getChildren('lotto');
  var lotti = [];
  var partecipazioni = [];
  var numCol = sheetLotti.getLastColumn();
  for (var i=0, max_i=items.length; i<max_i; i++) {
    var item = items[i];
    var lotto = [];
    lotto.push(item.getChild('cig').getText());
    var oggetto = item.getChild('oggetto').getText();
    lotto.push(oggetto);
    lotto.push(item.getChild('strutturaProponente').getChild('codiceFiscaleProp').getText());
    lotto.push(item.getChild('strutturaProponente').getChild('denominazione').getText());
    lotto.push(item.getChild('sceltaContraente').getText());
    lotto.push(parseFloat(item.getChild('importoAggiudicazione').getText()));
    lotto.push(parseFloat(item.getChild('importoSommeLiquidate').getText()));
    lotto.push(item.getChild('tempiCompletamento').getChild('dataInizio').getText());
    lotto.push(item.getChild('tempiCompletamento').getChild('dataUltimazione').getText());
    var lottoId = md5(lotto.join(""));
    // bag of words, esclusa da calcolo id
    lotto.push(bow(oggetto));

    lotto.unshift(lottoId);
    lotti.push(lotto);
    if(lotti.length>=250) {
      sheetLotti.getRange(sheetLotti.getLastRow()+1, 1, lotti.length, lotto.length).setValues(lotti);
      lotti = [];
      ss.toast('Letti '+(sheetLotti.getLastRow()-1)+' lotti');
    }
    
    // partecipazioni di singoli
    var partecipanti = item.getChild('partecipanti').getChildren('partecipante');
    var aggiudicatari = item.getChild('aggiudicatari').getChildren('aggiudicatario');
    var gruppiPartecipanti = item.getChild('partecipanti').getChildren('raggruppamento');
    var gruppiAggiudicatari = item.getChild('aggiudicatari').getChildren('aggiudicatarioRaggruppamento');
    var savecf = [];
    var savers = [];
    var savegruppi = [];
    var partId, gruppoId;
    for (a=0, max_a=aggiudicatari.length; a<max_a; a++) {
      var agg = aggiudicatari[a];
      if (! agg) continue;
      var cf = agg.getChild('codiceFiscale')? agg.getChild('codiceFiscale').getText(): '';
      var rs = agg.getChild('ragioneSociale').getText();
      if (cf>'') {
        savecf.push(cf);
        rsUniche[cf] ? rs = rsUniche[cf]: rsUniche[cf] = rs;
      }
      savers.push(rs);
      partId = creaPartId(cf,rs);
      partecipazioni.push( [lottoId, partId, cf, rs, '', '', 1] );
    }
    for (p=0, max_p=partecipanti.length; p<max_p; p++) {
      var part = partecipanti[p];
      if (! part) continue;
      var cf = part.getChild('codiceFiscale')? part.getChild('codiceFiscale').getText(): '';
      var rs = part.getChild('ragioneSociale').getText();
      if (cf>'') {
        rsUniche[cf] ? rs = rsUniche[cf]: rsUniche[cf] = rs;
      }
      if (! (savecf.indexOf(cf)>=0||savers.indexOf(rs)>=0)) {
        partId = creaPartId(cf,rs);
        partecipazioni.push( [lottoId, partId, cf, part.getChild('ragioneSociale').getText(), '', '', 0] );
      }
    }
    
    // gruppi
    for (a=0, max_a=gruppiAggiudicatari.length; a<max_a; a++) {
      var agg = gruppiAggiudicatari[a];
      if (! agg) continue;
      membri = agg.getChildren('membro');
      gruppoId = md5(membri.join(""));
      for (m=0, max_m=membri.length; m<max_m; m++) {
        var membro = membri[m];
        var cf = membro.getChild('codiceFiscale')? membro.getChild('codiceFiscale').getText(): '';
        var rs = membro.getChild('ragioneSociale').getText();
        if (cf>'') {
          savecf.push(cf);
          rsUniche[cf] ? rs = rsUniche[cf]: rsUniche[cf] = rs;
        }
        savers.push(rs);
        partId = creaPartId(cf,rs);
        partecipazioni.push( [lottoId, partId, cf, rs, gruppoId, membro.getChild('ruolo').getText(), 1] );
      }
    }
    for (p=0, max_p=gruppiPartecipanti.length; p<max_p; p++) {
      var part = gruppiPartecipanti[p];
      if (! part) continue;
      membri = part.getChildren('membro');
      gruppoId = md5(membri.join(""));
      for (m=0, max_m=membri.length; m<max_m; m++) {
        var membro = membri[m];
        var cf = membro.getChild('codiceFiscale')? membro.getChild('codiceFiscale').getText(): '';
        var rs = membro.getChild('ragioneSociale').getText();
        if (cf>'') {
          rsUniche[cf] ? rs = rsUniche[cf]: rsUniche[cf] = rs;
        }
        if (! (savecf.indexOf(cf)>=0||savers.indexOf(rs)>=0)) {
          partId = creaPartId(cf,rs);
          partecipazioni.push( [lottoId, partId, cf, rs, gruppoId, membro.getChild('ruolo').getText(), 0] );
        }
      }
    }
    if(partecipazioni.length>=250) {
      sheetPartecipazioni.getRange(sheetPartecipazioni.getLastRow()+1, 1, partecipazioni.length, 7).setValues(partecipazioni);
      partecipazioni = [];
    }
  }
  
  if(lotti.length) sheetLotti.getRange(sheetLotti.getLastRow()+1, 1, lotti.length, lotto.length).setValues(lotti);
  if(partecipazioni.length) sheetPartecipazioni.getRange(sheetPartecipazioni.getLastRow()+1, 1, partecipazioni.length, 7).setValues(partecipazioni);
  // intervalli denominati
  ss.setNamedRange("Dati_lotti", sheetLotti.getRange(2,1,sheetLotti.getLastRow()-1, sheetLotti.getLastColumn()));
  ss.setNamedRange("Dati_partecipazioni", sheetPartecipazioni.getRange(2,1,sheetPartecipazioni.getLastRow()-1, sheetPartecipazioni.getLastColumn()));
  ss.setNamedRange("Aggiudicatari", sheetAggiudicazioni.getRange(2,2,sheetAggiudicazioni.getLastRow()-1, 3));
  ss.setNamedRange("Importi", sheetLotti.getRange(2,colImporti,sheetLotti.getLastRow()-1, 1));
  
  ss.toast('Letti '+(sheetLotti.getLastRow()-1)+' lotti e '+(sheetPartecipazioni.getLastRow()-1)+' partecipazioni');
}

function creaPartId(cf,rs) {
  rs = rs.trim().toLowerCase().replace(/ +/, " ").replace(/[^a-z0-9 ]/g,"");
  return md5(cf+rs);
}

function md5(str) {
  // Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, str) restituisce un array di 16 int.
  // Quelli negativi vanno sommati a 256 per ottenere il codice carattere corretto (il segno - 'interpreta' il bit di ord sup) 
  return Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, str).reduce(function(str,chr){
    chr = (chr < 0 ? chr + 256 : chr).toString(16);
    return str + (chr.length==1?'0':'') + chr;
  },'');
  Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, lotto.join(""))
}

// bag of words
function bow(text) {
//      Logger.log('testo: '+text);
  text = text.replace(/[^a-z0-9 ]/gi, " ").replace(/\b\d+\b/g," ");
  for (var i=0, max_i = stopwords.length; i<max_i; i++) {
    var regexp = new RegExp("\\b"+stopwords[i][0]+"\\b", 'ig');
    text = text.replace(regexp, " ");
//    Logger.log(regexp+' -> '+text);
  }
  var arr = text.replace(/\b[^ ]{1,2}\b/g," ").trim().toLowerCase().split(/\s+/);
  var dedupedArr = [];
  arr.forEach( function(element, index, array) {
    if(dedupedArr.indexOf(element) == -1) dedupedArr.push(element);
  });
  return dedupedArr.join(',');
}