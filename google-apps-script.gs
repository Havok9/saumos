// Saumos · synchronisation du suivi + tableau de bord (Google Sheets)
// 1) Colle tout ce fichier dans Extensions > Apps Script (remplace tout).
// 2) Choisis la fonction "setup" en haut, clique Exécuter : ça crée les onglets, le tableau et les graphiques.
//    (Sinon ils se créent tout seuls à la première utilisation du site.)
// 3) Déployer > Gérer les déploiements > Modifier > Nouvelle version.

const SHEET_ID = '1fkXPdDXtpnLmlKneHFBOoc0EO0YKoXCWhs6jn9MR9P0';
const LOG = 'Suivi', CONF = 'Réglages', DASH = 'Tableau de bord';
const TASKS = [
  ['eau', 'Vérifier niveau eau et propreté', 'd'],
  ['croq', 'Vérifier niveau croquettes et propreté', 'd'],
  ['litiere', 'Retirer déjection litière', 'd'],
  ['litiere_tout', 'Changer intégralité litière', 's'],
  ['fontaine', 'Nettoyage cuve et fontaine', 's'],
  ['menage', 'Passer le balai bureau et serpillière', 's'],
  ['friandise', 'Friandise : moitié sachet', 's'],
  ['visuel', 'Vérification visuelle du chat', 's']
];
const NAVY = '#1E2A33', BLUE = '#3F86B5', LIGHT = '#E3EFF6', RED = '#E0533D', PINK = '#FBE3DE';

function ss_() { return SpreadsheetApp.openById(SHEET_ID); }
function tz_() { return ss_().getSpreadsheetTimeZone(); }
function day_(v) { return v instanceof Date ? Utilities.formatDate(v, tz_(), 'yyyy-MM-dd') : String(v); }
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function label_(k) { const t = TASKS.find(x => x[0] === k); return t ? t[1] : k; }

function logSheet_() {
  const ss = ss_();
  let s = ss.getSheetByName(LOG);
  if (s && String(s.getRange('A1').getValue()).toLowerCase() === 'type' && s.getRange('C1').getValue() !== 'Tâche') {
    s.setName(LOG + ' (ancien ' + Utilities.formatDate(new Date(), tz_(), 'dd-MM HH:mm') + ')');
    s = null;
  }
  if (!s) {
    s = ss.insertSheet(LOG, 0);
    s.getRange(1, 1, 1, 7).setValues([['Type', 'Date', 'Tâche', 'Fait par', 'Heure', 'Commentaire', 'clé']]);
    styleLog_(s);
  }
  return s;
}
function styleLog_(s) {
  s.setFrozenRows(1);
  s.getRange('A1:G1').setBackground(NAVY).setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(11).setVerticalAlignment('middle');
  s.setRowHeight(1, 36);
  [110, 110, 300, 130, 150, 360, 80].forEach((w, i) => s.setColumnWidth(i + 1, w));
  s.getRange('B:B').setNumberFormat('dd/mm/yyyy');
  s.getRange('E:E').setNumberFormat('dd/mm HH:mm');
  s.getRange('F:F').setWrap(true);
  s.hideColumns(7);
  s.getBandings().forEach(b => b.remove());
  s.getRange('A1:G1000').applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);
  s.getRange('A1:G1').setBackground(NAVY);
  const typeRange = s.getRange('A2:A1000');
  s.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Tâche').setBackground(LIGHT).setFontColor('#1F5A80').setBold(true).setRanges([typeRange]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Commentaire').setBackground(PINK).setFontColor('#9A2E20').setBold(true).setRanges([typeRange]).build()
  ]);
}
function confSheet_() {
  const ss = ss_();
  let s = ss.getSheetByName(CONF);
  if (!s) { s = ss.insertSheet(CONF); s.getRange('A1').setValue('{}'); s.hideSheet(); }
  return s;
}

function ensure_() {
  if (!ss_().getSheetByName(DASH)) setup();
}

function doGet() {
  ensure_();
  const rows = logSheet_().getDataRange().getValues().slice(1);
  const out = { entries: [], comments: [], config: null };
  rows.forEach(r => {
    if (!r[0]) return;
    const at = r[4] instanceof Date ? r[4].toISOString() : String(r[4]);
    if (r[0] === 'Commentaire') out.comments.push({ date: day_(r[1]), by: r[3], at, text: r[5] });
    else out.entries.push({ date: day_(r[1]), task: r[6], by: r[3], at });
  });
  try { out.config = JSON.parse(confSheet_().getRange('A1').getValue() || 'null'); } catch (x) {}
  return json_(out);
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    ensure_();
    const b = JSON.parse(e.postData.contents);
    if (b.type === 'config') { confSheet_().getRange('A1').setValue(b.text || '{}'); return json_({ ok: true }); }
    const s = logSheet_();
    const date = Utilities.parseDate(b.date, tz_(), 'yyyy-MM-dd');
    const at = b.at ? new Date(b.at) : new Date();
    if (b.type === 'entry') {
      const done = s.getDataRange().getValues().some(r => r[0] === 'Tâche' && day_(r[1]) === b.date && r[6] === b.task);
      if (done) return json_({ ok: false, reason: 'already' });
      s.appendRow(['Tâche', date, label_(b.task), b.by || '', at, '', b.task]);
    } else {
      s.appendRow(['Commentaire', date, '', b.by || '', at, b.text || '', '']);
    }
    return json_({ ok: true });
  } finally {
    lock.releaseLock();
  }
}

function setup() {
  const ss = ss_();
  styleLog_(logSheet_());
  confSheet_();
  let d = ss.getSheetByName(DASH);
  if (d) ss.deleteSheet(d);
  d = ss.insertSheet(DASH, 0);
  d.setHiddenGridlines(true);
  d.getRange('A1').setValue('🐾 Suivi de Saumos').setFontSize(20).setFontWeight('bold').setFontColor(NAVY);
  d.getRange('A2').setValue('Semaine du').setFontColor('#6B7A86');
  d.getRange('B2').setFormula('=TODAY()-WEEKDAY(TODAY();3)').setNumberFormat('dd/mm/yyyy').setFontWeight('bold');

  // grille de la semaine
  d.getRange('A3').setValue('Tâche');
  for (let i = 0; i < 7; i++) d.getRange(3, 2 + i).setFormula('=$B$2+' + i);
  d.getRange('B3:H3').setNumberFormat('ddd dd/mm');
  d.getRange('A3:H3').setBackground(NAVY).setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center');
  d.getRange('A3').setHorizontalAlignment('left');
  TASKS.forEach((t, i) => {
    const r = 4 + i;
    d.getRange(r, 1).setValue(t[1]);
    d.getRange(r, 9).setValue(t[0]);
    d.getRange(r, 10).setValue(t[2]);
    for (let c = 0; c < 7; c++) {
      const col = String.fromCharCode(66 + c);
      d.getRange(r, 2 + c).setFormula(`=IF(AND($J${r}="s";WEEKDAY(${col}$3)<>1);"";IF(COUNTIFS('${LOG}'!$G:$G;$I${r};'${LOG}'!$B:$B;${col}$3)>0;"✓";IF(${col}$3<TODAY();"✕";"·")))`);
    }
  });
  const grid = d.getRange('B4:H11');
  grid.setHorizontalAlignment('center').setFontWeight('bold').setFontSize(13);
  d.getRange('A4:A11').setFontSize(11);
  d.setColumnWidth(1, 300);
  for (let c = 2; c <= 8; c++) d.setColumnWidth(c, 82);
  d.hideColumns(9, 2);
  d.getRange('A3:H11').setBorder(true, true, true, true, true, true, '#DCE4EA', SpreadsheetApp.BorderStyle.SOLID);
  d.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('✓').setBackground(BLUE).setFontColor('#FFFFFF').setRanges([grid]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('✕').setBackground(RED).setFontColor('#FFFFFF').setRanges([grid]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('·').setBackground('#F2F5F8').setFontColor('#9CC3D5').setRanges([grid]).build()
  ]);

  // résumé du jour
  d.getRange('A13').setValue("Aujourd'hui").setFontWeight('bold');
  d.getRange('B13').setFormula(`=COUNTIFS('${LOG}'!A:A;"Tâche";'${LOG}'!B:B;TODAY())&" / "&(3+IF(WEEKDAY(TODAY())=1;5;0))&" tâches faites"`);
  d.getRange('A14').setValue('Commentaires cette semaine').setFontWeight('bold');
  d.getRange('B14').setFormula(`=COUNTIFS('${LOG}'!A:A;"Commentaire";'${LOG}'!B:B;">="&$B$2)`);

  // données des graphiques
  d.getRange('L3:M3').setValues([['Jour', 'Tâches faites']]).setFontWeight('bold');
  for (let i = 0; i < 14; i++) {
    d.getRange(4 + i, 12).setFormula('=TODAY()-' + (13 - i)).setNumberFormat('dd/mm');
    d.getRange(4 + i, 13).setFormula(`=COUNTIFS('${LOG}'!$A:$A;"Tâche";'${LOG}'!$B:$B;L${4 + i})`);
  }
  d.getRange('O3').setFormula(`=IFERROR(QUERY('${LOG}'!A2:D;"select D, count(A) where A = 'Tâche' group by D label D 'Personne', count(A) 'Tâches'";0);{"Personne"\\"Tâches"})`);
  d.getRange('L3:P20').setFontColor('#6B7A86');

  d.getCharts().forEach(c => d.removeChart(c));
  d.insertChart(d.newChart().asColumnChart()
    .addRange(d.getRange('L3:M17')).setNumHeaders(1)
    .setOption('title', 'Tâches faites · 14 derniers jours')
    .setOption('colors', [BLUE]).setOption('legend', { position: 'none' })
    .setPosition(16, 1, 0, 10).setOption('width', 620).setOption('height', 300).build());
  d.insertChart(d.newChart().asPieChart()
    .addRange(d.getRange('O3:P15')).setNumHeaders(1)
    .setOption('title', 'Qui a fait quoi')
    .setOption('pieHole', 0.45).setOption('colors', [BLUE, '#9CC3D5', NAVY, '#6B7A86', RED, '#CFE3EF'])
    .setPosition(16, 6, 40, 10).setOption('width', 420).setOption('height', 300).build());

  ss.setActiveSheet(d);
}
