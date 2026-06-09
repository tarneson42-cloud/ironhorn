// IRONHORN Data Engine + Dashboard — Google Apps Script v3
// Burnt Orange Theme · Terry L. Arneson

// ── COLORS ──────────────────────────────────
var C = {
  // Burnt orange palette
  ORANGE_DARK:   '#7B3300',  // Dark burnt orange
  ORANGE:        '#BF5700',  // Primary burnt orange  
  ORANGE_MED:    '#D4621A',  // Medium orange
  ORANGE_LIGHT:  '#F08030',  // Light orange
  GOLD:          '#F0B429',  // Gold accent
  // Backgrounds
  BG_HEADER:     '#1C1410',  // Near black
  BG_DARK:       '#2A1A0E',  // Dark brown
  BG_MED:        '#3D2410',  // Medium brown
  BG_LIGHT:      '#4A2E14',  // Light brown
  BG_ROW_ALT:    '#251608',  // Alternate row
  // Status colors
  GREEN:         '#2E7D32',
  GREEN_LIGHT:   '#4CAF50',
  YELLOW:        '#F57F17',
  RED:           '#C62828',
  RED_LIGHT:     '#EF5350',
  // Text
  TEXT_WHITE:    '#F0EDE8',
  TEXT_GOLD:     '#F0B429',
  TEXT_LIGHT:    '#C4A882',
  TEXT_MUTED:    '#7A5F40',
};

// ── WEB API ─────────────────────────────────
function doGet(e) {
  var p = (e && e.parameter) ? e.parameter : {};
  var action = p.action || 'ping';
  if(action === 'ping')    return respond({ok:true, msg:'IRONHORN Engine Online'});
  if(action === 'log')     return respond(logDaily(p));
  if(action === 'week')    return respond(getReport('week'));
  if(action === 'month')   return respond(getReport('month'));
  if(action === 'format')  { formatDashboard(); return respond({ok:true, msg:'Dashboard formatted'}); }
  return respond({ok:false, msg:'Unknown action'});
}

function doPost(e) {
  try{
    var p = (e && e.postData) ? JSON.parse(e.postData.contents) : {};
    var action = p.action || 'log';
    if(action === 'log')   return respond(logDaily(p));
    if(action === 'week')  return respond(getReport('week'));
    if(action === 'month') return respond(getReport('month'));
    return respond({ok:false, msg:'Unknown action'});
  }catch(err){
    return respond({ok:false, error:err.toString()});
  }
}

function respond(data){
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}

function today(){
  return Utilities.formatDate(new Date(), 'America/Chicago', 'yyyy-MM-dd');
}

// ── SHEET SETUP ─────────────────────────────
var HEADERS = [
  'Date','Weight','BP Sys','BP Dia','Heart Rate','Glucose','SpO2','Temp',
  'Oura Sleep','Calories','Protein','Carbs','Fat','Water Oz',
  'Cal Burned','Ex Minutes','Stand Hrs','Meds Done','Meds Total',
  'myAir Score','myAir AHI','Notes'
];

var COL_GROUPS = {
  'Date':        {color: C.ORANGE,       group: 'DATE'},
  'Weight':      {color: C.ORANGE_MED,   group: 'VITALS'},
  'BP Sys':      {color: '#1565C0',      group: 'VITALS'},
  'BP Dia':      {color: '#1976D2',      group: 'VITALS'},
  'Heart Rate':  {color: '#C62828',      group: 'VITALS'},
  'Glucose':     {color: '#6A1B9A',      group: 'VITALS'},
  'SpO2':        {color: '#00838F',      group: 'VITALS'},
  'Temp':        {color: '#558B2F',      group: 'VITALS'},
  'Oura Sleep':  {color: '#4527A0',      group: 'VITALS'},
  'Calories':    {color: C.ORANGE_LIGHT, group: 'NUTRITION'},
  'Protein':     {color: '#2E7D32',      group: 'NUTRITION'},
  'Carbs':       {color: '#1565C0',      group: 'NUTRITION'},
  'Fat':         {color: C.GOLD,         group: 'NUTRITION'},
  'Water Oz':    {color: '#00838F',      group: 'NUTRITION'},
  'Cal Burned':  {color: C.ORANGE,       group: 'ACTIVITY'},
  'Ex Minutes':  {color: '#2E7D32',      group: 'ACTIVITY'},
  'Stand Hrs':   {color: '#1565C0',      group: 'ACTIVITY'},
  'Meds Done':   {color: '#558B2F',      group: 'MEDS'},
  'Meds Total':  {color: '#558B2F',      group: 'MEDS'},
  'myAir Score': {color: '#4527A0',      group: 'SLEEP'},
  'myAir AHI':   {color: '#6A1B9A',      group: 'SLEEP'},
  'Notes':       {color: C.TEXT_MUTED,   group: 'NOTES'},
};

function getSheet(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('IRONHORN');
  if(!sh){
    sh = ss.insertSheet('IRONHORN');
    sh.getRange(1,1,1,HEADERS.length).setValues([HEADERS]);
    sh.setFrozenRows(1);
  }
  return sh;
}

// ── DATA LOGGING ────────────────────────────
function normalizeDate(val){
  if(!val) return '';
  try{
    var d = new Date(val);
    if(!isNaN(d.getTime())) return Utilities.formatDate(d, 'America/Chicago', 'yyyy-MM-dd');
  }catch(e){}
  return String(val).trim();
}

function findRow(sh, date){
  var lastRow = sh.getLastRow();
  if(lastRow < 2) return -1;
  var dates = sh.getRange(2, 1, lastRow-1, 1).getValues();
  for(var i=0; i<dates.length; i++){
    if(normalizeDate(dates[i][0]) === date) return i + 2;
  }
  return -1;
}

function logDaily(p){
  var sh   = getSheet();
  var date = p.date ? normalizeDate(p.date) : today();

  var newRow = [
    date,
    p.weight||'', p.bp_sys||'', p.bp_dia||'', p.heart_rate||'',
    p.glucose||'', p.spo2||'', p.temp||'', p.oura_sleep||'',
    p.calories||'', p.protein||'', p.carbs||'', p.fat||'',
    p.water_oz||'', p.cal_burned||'', p.ex_minutes||'', p.stand_hours||'',
    p.meds_done||'', p.meds_total||'',
    p.myair_score||'', p.myair_ahi||'',
    p.notes||''
  ];

  var rowIdx = findRow(sh, date);
  if(rowIdx > 0){
    var existing = sh.getRange(rowIdx, 1, 1, newRow.length).getValues()[0];
    for(var j=0; j<newRow.length; j++){
      if(newRow[j]===''||newRow[j]===null) newRow[j] = existing[j];
    }
    sh.getRange(rowIdx, 1, 1, newRow.length).setValues([newRow]);
    formatRow(sh, rowIdx);
    return {ok:true, date:date, action:'updated', row:rowIdx};
  } else {
    sh.appendRow(newRow);
    var newRow2 = sh.getLastRow();
    formatRow(sh, newRow2);
    return {ok:true, date:date, action:'added', row:newRow2};
  }
}

// ── DASHBOARD FORMATTING ────────────────────
function formatDashboard(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = getSheet();

  // ── Spreadsheet title ──
  ss.rename('IRONHORN Health Dashboard');

  // ── Sheet setup ──
  sh.setTabColor(C.ORANGE);
  sh.setName('IRONHORN');

  // ── Column widths ──
  var widths = [100,70,65,65,80,75,60,60,85,80,70,65,60,75,85,85,75,80,80,90,75,200];
  for(var i=0; i<widths.length; i++){
    sh.setColumnWidth(i+1, widths[i]);
  }

  // ── Row height ──
  sh.setRowHeight(1, 48);

  // ── Header row ──
  var headerRange = sh.getRange(1, 1, 1, HEADERS.length);
  headerRange.setValues([HEADERS]);
  headerRange.setBackground(C.BG_HEADER);
  headerRange.setFontColor(C.TEXT_GOLD);
  headerRange.setFontWeight('bold');
  headerRange.setFontSize(11);
  headerRange.setHorizontalAlignment('center');
  headerRange.setVerticalAlignment('middle');
  headerRange.setBorder(false, false, true, false, false, false, C.ORANGE, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  // Color each header cell individually
  for(var h=0; h<HEADERS.length; h++){
    var headerCell = sh.getRange(1, h+1);
    var colInfo = COL_GROUPS[HEADERS[h]];
    if(colInfo){
      headerCell.setFontColor(colInfo.color);
    }
  }

  // ── Format existing data rows ──
  var lastRow = sh.getLastRow();
  if(lastRow >= 2){
    for(var r=2; r<=lastRow; r++){
      formatRow(sh, r);
    }
  }

  // ── Freeze header ──
  sh.setFrozenRows(1);
  sh.setFrozenColumns(1);

  // ── Add conditional formatting ──
  addConditionalFormatting(sh);

  // ── Add summary dashboard tab ──
  buildSummaryTab(ss, sh);

  SpreadsheetApp.flush();
}

function formatRow(sh, rowNum){
  var isAlt = (rowNum % 2 === 0);
  var bg    = isAlt ? C.BG_ROW_ALT : C.BG_DARK;
  var rowRange = sh.getRange(rowNum, 1, 1, HEADERS.length);

  rowRange.setBackground(bg);
  rowRange.setFontColor(C.TEXT_WHITE);
  rowRange.setFontSize(10);
  rowRange.setVerticalAlignment('middle');
  rowRange.setBorder(false, false, true, false, false, false, C.BG_MED, SpreadsheetApp.BorderStyle.SOLID);

  // Date column
  sh.getRange(rowNum, 1).setFontColor(C.GOLD).setFontWeight('bold').setNumberFormat('mmm d, yyyy');

  // Weight column - orange
  sh.getRange(rowNum, 2).setFontColor(C.ORANGE_LIGHT).setHorizontalAlignment('center').setFontWeight('bold').setNumberFormat('0.0');

  // BP columns - blue
  sh.getRange(rowNum, 3).setFontColor('#64B5F6').setHorizontalAlignment('center');
  sh.getRange(rowNum, 4).setFontColor('#64B5F6').setHorizontalAlignment('center');

  // Heart rate - red
  sh.getRange(rowNum, 5).setFontColor('#EF9A9A').setHorizontalAlignment('center');

  // Glucose - purple
  sh.getRange(rowNum, 6).setFontColor('#CE93D8').setHorizontalAlignment('center');

  // SpO2 - teal
  sh.getRange(rowNum, 7).setFontColor('#80DEEA').setHorizontalAlignment('center');

  // Temp
  sh.getRange(rowNum, 8).setFontColor('#A5D6A7').setHorizontalAlignment('center');

  // Oura Sleep - indigo
  sh.getRange(rowNum, 9).setFontColor('#9FA8DA').setHorizontalAlignment('center');

  // Calories - orange
  sh.getRange(rowNum, 10).setFontColor(C.ORANGE_LIGHT).setHorizontalAlignment('center').setFontWeight('bold');

  // Protein - green
  sh.getRange(rowNum, 11).setFontColor('#A5D6A7').setHorizontalAlignment('center').setFontWeight('bold');

  // Carbs - blue
  sh.getRange(rowNum, 12).setFontColor('#90CAF9').setHorizontalAlignment('center');

  // Fat - gold
  sh.getRange(rowNum, 13).setFontColor(C.GOLD).setHorizontalAlignment('center');

  // Water - teal
  sh.getRange(rowNum, 14).setFontColor('#80DEEA').setHorizontalAlignment('center');

  // Cal Burned
  sh.getRange(rowNum, 15).setFontColor(C.ORANGE_LIGHT).setHorizontalAlignment('center');

  // Ex Minutes - green
  sh.getRange(rowNum, 16).setFontColor('#A5D6A7').setHorizontalAlignment('center').setFontWeight('bold');

  // Stand Hrs
  sh.getRange(rowNum, 17).setHorizontalAlignment('center');

  // Meds
  sh.getRange(rowNum, 18).setFontColor('#A5D6A7').setHorizontalAlignment('center');
  sh.getRange(rowNum, 19).setHorizontalAlignment('center');

  // myAir Score - indigo bold
  sh.getRange(rowNum, 20).setFontColor('#B39DDB').setHorizontalAlignment('center').setFontWeight('bold');

  // myAir AHI
  sh.getRange(rowNum, 21).setHorizontalAlignment('center');

  // Notes - muted
  sh.getRange(rowNum, 22).setFontColor(C.TEXT_LIGHT).setFontStyle('italic');

  sh.setRowHeight(rowNum, 32);
}

function addConditionalFormatting(sh){
  var lastRow = Math.max(sh.getLastRow(), 100);
  var rules   = [];

  // Weight — highlight loss in green, gain in red vs previous
  // Glucose > 140 = red, 100-140 = yellow, <100 = green
  var glucoseRange = sh.getRange(2, 6, lastRow-1, 1);
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(140)
    .setBackground('#4E1414').setFontColor('#EF9A9A')
    .setRanges([glucoseRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberBetween(100, 140)
    .setBackground('#4E3D00').setFontColor('#FFE082')
    .setRanges([glucoseRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(100)
    .setBackground('#1B3A1F').setFontColor('#A5D6A7')
    .setRanges([glucoseRange]).build());

  // BP Sys > 140 = red
  var bpRange = sh.getRange(2, 3, lastRow-1, 1);
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(140)
    .setBackground('#4E1414').setFontColor('#EF9A9A')
    .setRanges([bpRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThanOrEqualTo(130)
    .setBackground('#1B3A1F').setFontColor('#A5D6A7')
    .setRanges([bpRange]).build());

  // myAir >= 90 = green, 70-89 = yellow, <70 = red
  var myairRange = sh.getRange(2, 20, lastRow-1, 1);
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThanOrEqualTo(90)
    .setBackground('#1B3A1F').setFontColor('#A5D6A7')
    .setRanges([myairRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberBetween(70, 89)
    .setBackground('#4E3D00').setFontColor('#FFE082')
    .setRanges([myairRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(70)
    .setBackground('#4E1414').setFontColor('#EF9A9A')
    .setRanges([myairRange]).build());

  // Protein >= 160 = green
  var protRange = sh.getRange(2, 11, lastRow-1, 1);
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThanOrEqualTo(160)
    .setBackground('#1B3A1F').setFontColor('#A5D6A7')
    .setRanges([protRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberBetween(100, 159)
    .setBackground('#4E3D00').setFontColor('#FFE082')
    .setRanges([protRange]).build());

  sh.setConditionalFormatRules(rules);
}

function buildSummaryTab(ss, dataSheet){
  // Delete existing summary if exists
  var existing = ss.getSheetByName('SUMMARY');
  if(existing) ss.deleteSheet(existing);

  var sum = ss.insertSheet('SUMMARY', 0);
  sum.setTabColor(C.GOLD);

  // Get data
  var lastRow = dataSheet.getLastRow();
  var numRows = Math.max(lastRow - 1, 0);

  // ── Title ──
  sum.setColumnWidth(1, 180);
  for(var c=2; c<=8; c++) sum.setColumnWidth(c, 120);
  sum.setRowHeight(1, 60);
  sum.setRowHeight(2, 20);

  var titleCell = sum.getRange('A1:H1');
  titleCell.merge();
  titleCell.setValue('⚔️  IRONHORN HEALTH COMMAND  ⚔️');
  titleCell.setBackground(C.BG_HEADER);
  titleCell.setFontColor(C.GOLD);
  titleCell.setFontSize(20);
  titleCell.setFontWeight('bold');
  titleCell.setHorizontalAlignment('center');
  titleCell.setVerticalAlignment('middle');

  var subCell = sum.getRange('A2:H2');
  subCell.merge();
  subCell.setValue('Terry L. Arneson  ·  For God. For Vic. For family.');
  subCell.setBackground(C.BG_HEADER);
  subCell.setFontColor(C.TEXT_MUTED);
  subCell.setFontSize(10);
  subCell.setHorizontalAlignment('center');
  subCell.setFontStyle('italic');

  // ── Latest Vitals Section ──
  sum.setRowHeight(3, 10);
  sum.setRowHeight(4, 36);

  var vitHdr = sum.getRange('A4:H4');
  vitHdr.merge();
  vitHdr.setValue('LATEST VITALS');
  vitHdr.setBackground(C.ORANGE_DARK);
  vitHdr.setFontColor(C.GOLD);
  vitHdr.setFontWeight('bold');
  vitHdr.setFontSize(12);
  vitHdr.setHorizontalAlignment('center');
  vitHdr.setVerticalAlignment('middle');
  vitHdr.setBorder(false,false,true,false,false,false,C.ORANGE,SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  var vitLabels = ['Weight','Blood Pressure','Heart Rate','Glucose','SpO2','Oura Sleep','myAir Score'];
  var vitCols   = [2,       '3&4',            5,           6,         7,      9,            20];
  var vitColors = ['#F08030','#64B5F6',        '#EF9A9A',   '#CE93D8', '#80DEEA','#9FA8DA', '#B39DDB'];

  sum.setRowHeight(5, 42);
  sum.setRowHeight(6, 42);

  for(var v=0; v<vitLabels.length; v++){
    sum.getRange(5, v+1).setValue(vitLabels[v])
      .setBackground(C.BG_MED).setFontColor(vitColors[v])
      .setFontWeight('bold').setFontSize(10)
      .setHorizontalAlignment('center').setVerticalAlignment('middle');

    var formula = '';
    if(v === 1){
      // BP = Sys/Dia
      formula = '=IFERROR(LOOKUP(2,1/(IRONHORN!C2:C100<>""),IRONHORN!C2:C100)&"/"&LOOKUP(2,1/(IRONHORN!D2:D100<>""),IRONHORN!D2:D100),"—")';
    } else if(v === 0){
      formula = '=IFERROR(LOOKUP(2,1/(IRONHORN!B2:B100<>""),IRONHORN!B2:B100)&" lbs","—")';
    } else {
      var colMap = ['','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W']; var col = colMap[vitCols[v]];
      if(col) formula = '=IFERROR(LOOKUP(2,1/(IRONHORN!'+col+'2:'+col+'100<>""),IRONHORN!'+col+'2:'+col+'100),"—")';
    }

    sum.getRange(6, v+1).setFormula(formula)
      .setBackground(C.BG_DARK).setFontColor(vitColors[v])
      .setFontWeight('bold').setFontSize(16)
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
  }

  // ── 50-Day Mission Progress ──
  sum.setRowHeight(7, 10);
  sum.setRowHeight(8, 36);
  var missionHdr = sum.getRange('A8:H8');
  missionHdr.merge();
  missionHdr.setValue('50-DAY MISSION — 211 lbs → 186 lbs');
  missionHdr.setBackground(C.ORANGE_DARK);
  missionHdr.setFontColor(C.GOLD);
  missionHdr.setFontWeight('bold').setFontSize(12).setHorizontalAlignment('center').setVerticalAlignment('middle');
  missionHdr.setBorder(false,false,true,false,false,false,C.ORANGE,SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  sum.setRowHeight(9, 42);
  sum.setRowHeight(10, 42);

  var mLabels = ['Start','Current','Lost','To Goal','% Complete','Pace/Week','Days Logged'];
  for(var m=0; m<mLabels.length; m++){
    sum.getRange(9, m+1).setValue(mLabels[m])
      .setBackground(C.BG_MED).setFontColor(C.GOLD)
      .setFontWeight('bold').setFontSize(10)
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
  }

  var mFormulas = [
    '"211 lbs"',
    '=IFERROR(LOOKUP(2,1/(IRONHORN!B2:B100<>""),IRONHORN!B2:B100)&" lbs","—")',
    '=IFERROR((211-LOOKUP(2,1/(IRONHORN!B2:B100<>""),IRONHORN!B2:B100))&" lbs","—")',
    '=IFERROR((LOOKUP(2,1/(IRONHORN!B2:B100<>""),IRONHORN!B2:B100)-186)&" lbs","—")',
    '=IFERROR(TEXT((211-LOOKUP(2,1/(IRONHORN!B2:B100<>""),IRONHORN!B2:B100))/25,"0%"),"—")',
    '=IFERROR(TEXT((211-LOOKUP(2,1/(IRONHORN!B2:B100<>""),IRONHORN!B2:B100))/(COUNTA(IRONHORN!B2:B100)/7),"0.0")&" lbs/wk","—")',
    '=IFERROR(COUNTA(IRONHORN!A2:A100)&" days","0 days")',
  ];

  for(var mf=0; mf<mFormulas.length; mf++){
    sum.getRange(10, mf+1).setFormula('='+mFormulas[mf])
      .setBackground(C.BG_DARK).setFontColor(C.ORANGE_LIGHT)
      .setFontWeight('bold').setFontSize(14)
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
  }

  // ── Weekly Averages ──
  sum.setRowHeight(11, 10);
  sum.setRowHeight(12, 36);
  var weekHdr = sum.getRange('A12:H12');
  weekHdr.merge();
  weekHdr.setValue('7-DAY AVERAGES');
  weekHdr.setBackground(C.ORANGE_DARK).setFontColor(C.GOLD)
    .setFontWeight('bold').setFontSize(12).setHorizontalAlignment('center').setVerticalAlignment('middle');
  weekHdr.setBorder(false,false,true,false,false,false,C.ORANGE,SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  sum.setRowHeight(13, 42);
  sum.setRowHeight(14, 42);

  var wLabels = ['Avg Weight','Avg BP','Avg Glucose','Avg Calories','Avg Protein','Avg Water','Avg myAir'];
  for(var w=0; w<wLabels.length; w++){
    sum.getRange(13, w+1).setValue(wLabels[w])
      .setBackground(C.BG_MED).setFontColor(C.TEXT_GOLD)
      .setFontWeight('bold').setFontSize(10)
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
  }

  var wFormulas = [
    '=IFERROR(ROUND(AVERAGEIF(IRONHORN!A2:A,">="&TODAY()-7,IRONHORN!B2:B),1)&" lbs","—")',
    '=IFERROR(ROUND(AVERAGEIF(IRONHORN!A2:A,">="&TODAY()-7,IRONHORN!C2:C),0)&"/"&ROUND(AVERAGEIF(IRONHORN!A2:A,">="&TODAY()-7,IRONHORN!D2:D),0),"—")',
    '=IFERROR(ROUND(AVERAGEIF(IRONHORN!A2:A,">="&TODAY()-7,IRONHORN!F2:F),0)&" mg/dL","—")',
    '=IFERROR(ROUND(AVERAGEIF(IRONHORN!A2:A,">="&TODAY()-7,IRONHORN!J2:J),0)&" cal","—")',
    '=IFERROR(ROUND(AVERAGEIF(IRONHORN!A2:A,">="&TODAY()-7,IRONHORN!K2:K),1)&"g","—")',
    '=IFERROR(ROUND(AVERAGEIF(IRONHORN!A2:A,">="&TODAY()-7,IRONHORN!N2:N),0)&" oz","—")',
    '=IFERROR(ROUND(AVERAGEIF(IRONHORN!A2:A,">="&TODAY()-7,IRONHORN!T2:T),0)&"/100","—")',
  ];

  for(var wf=0; wf<wFormulas.length; wf++){
    sum.getRange(14, wf+1).setFormula(wFormulas[wf])
      .setBackground(C.BG_DARK).setFontColor(C.ORANGE_LIGHT)
      .setFontWeight('bold').setFontSize(14)
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
  }

  // ── Footer ──
  sum.setRowHeight(15, 10);
  sum.setRowHeight(16, 30);
  var footer = sum.getRange('A16:H16');
  footer.merge();
  footer.setValue('Your dad Alvin left at 55. You will not.  ⚔️');
  footer.setBackground(C.BG_HEADER).setFontColor(C.TEXT_MUTED)
    .setFontStyle('italic').setFontSize(11).setHorizontalAlignment('center').setVerticalAlignment('middle');

  // Set summary tab background
  sum.setTabColor(C.GOLD);
  SpreadsheetApp.flush();
}

// ── REPORTS ────────────────────────────────
function getReport(period){
  var sh = getSheet();
  var lastRow = sh.getLastRow();
  if(lastRow < 2) return {ok:false, msg:'No data yet'};

  var days = period === 'month' ? 30 : 7;
  var cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - days);

  var allData = sh.getRange(2, 1, lastRow-1, 22).getValues();
  var rows = [];
  for(var i=0; i<allData.length; i++){
    try{
      var d = new Date(allData[i][0]);
      if(!isNaN(d.getTime()) && d >= cutoff) rows.push(allData[i]);
    }catch(e){}
  }

  if(!rows.length) return {ok:false, msg:'No data in this period yet.'};

  var wts=[],sys=[],dia=[],gluc=[],cal=[],prot=[],water=[],myair=[],exmin=[];
  rows.forEach(function(r){
    if(r[1]>0)  wts.push(parseFloat(r[1]));
    if(r[2]>0)  sys.push(parseFloat(r[2]));
    if(r[3]>0)  dia.push(parseFloat(r[3]));
    if(r[5]>0)  gluc.push(parseFloat(r[5]));
    if(r[9]>0)  cal.push(parseFloat(r[9]));
    if(r[10]>0) prot.push(parseFloat(r[10]));
    if(r[13]>0) water.push(parseFloat(r[13]));
    if(r[19]>0) myair.push(parseFloat(r[19]));
    if(r[15]>0) exmin.push(parseFloat(r[15]));
  });

  function avg(arr){ return arr.length ? Math.round(arr.reduce(function(a,b){return a+b},0)/arr.length*10)/10 : null }

  var label = period==='month'?'30-Day':'7-Day';
  var report = label+' Report — '+rows.length+' days. ';
  if(wts.length>=2) report += 'Weight: '+wts[0]+' → '+wts[wts.length-1]+' lbs ('+(wts[0]-wts[wts.length-1]>0?'-':'+')+Math.abs(wts[0]-wts[wts.length-1]).toFixed(1)+' lbs). ';
  if(sys.length)   report += 'Avg BP: '+avg(sys)+'/'+avg(dia)+'. ';
  if(gluc.length)  report += 'Avg Glucose: '+avg(gluc)+'. ';
  if(cal.length)   report += 'Avg Cal: '+avg(cal)+'. ';
  if(prot.length)  report += 'Avg Protein: '+avg(prot)+'g. ';
  if(water.length) report += 'Avg Water: '+avg(water)+' oz. ';
  if(exmin.length) report += 'Avg Exercise: '+avg(exmin)+' min. ';
  if(myair.length) report += 'Avg myAir: '+avg(myair)+'/100. ';

  return {ok:true, period:period, days:rows.length, report:report};
}

// ── RUN FORMAT (call this manually once) ──
function runFormat(){
  formatDashboard();
}
