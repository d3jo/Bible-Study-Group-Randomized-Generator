/******** CONFIG (edit if you rename tabs) ********/
const RESP_SHEET_NAME = 'Form Responses 1'; // form responses tab
const BOARD_SHEET     = 'Board';            // final visual page
const CHART_DATA_SHEET= '_ChartData';       // hidden helper for chart

// Group labeling + layout
const LABEL_PREFIX    = 'Group ';           // "Group 1", "Group 2", ...
const MAX_COL_TEAMS   = 12;                 // max groups shown side-by-side
const COL_GAP         = 1;                  // spacing between group columns (A,C,E,...)
const TEAM_COL_WIDTH  = 220;                // px; fixed width for readable columns

// Team assignment mode
const USE_TEAM_SIZE   = false;              // true: fixed size; false: team count
const TEAM_SIZE       = 4;                  // when USE_TEAM_SIZE = true

// Dynamic team-count rule (your mapping)
// 0–5 → 2, 6–11 → 3, 12–16 → 4, 17–21 → 5, 22–26 → 6, >26 → 6
function teamCountFor(n) {
  if (n <= 5)  return 2;
  if (n <= 11) return 3;
  if (n <= 16) return 4;
  if (n <= 21) return 5;
  if (n <= 26) return 6;
  return 6;
}

/******** MENUS & TRIGGERS ********/
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Team Tools')
    .addItem('Rebuild Board now', 'rebuildBoardFromResponses')
    .addItem('Freeze Board to new tab', 'freezeBoard')
    .addToUi();
}

function onFormSubmit(e) {
  rebuildBoardFromResponses();
}

// Optional: refresh when you manually edit responses (e.g., fix a name/email)
function onResponsesEdit(e) {
  if (!e || !e.range) return;
  const sh = e.range.getSheet();
  if (sh.getName() !== RESP_SHEET_NAME) return;
  if (e.range.getRow() === 1) return; // ignore header
  Utilities.sleep(150);
  rebuildBoardFromResponses();
}

/******** CORE ********/
function rebuildBoardFromResponses() {
  const ss = SpreadsheetApp.getActive();
  const rs = ss.getSheetByName(RESP_SHEET_NAME);
  if (!rs) throw new Error('Sheet "' + RESP_SHEET_NAME + '" not found.');

  // Read all responses
  const lastRow = rs.getLastRow(), lastCol = rs.getLastColumn();
  if (lastRow < 2) { clearBoard_(); return; }

  const header = rs.getRange(1,1,1,lastCol).getValues()[0];
  const rows   = rs.getRange(2,1,lastRow-1,lastCol).getValues();

  // Find Email + Name columns (supports "이름")
  var emailIdx = -1, nameIdx = -1;
  for (var i=0;i<header.length;i++){
    var h = String(header[i]);
    if (emailIdx === -1 && /email/i.test(h)) emailIdx = i;
    if (nameIdx  === -1 && /(name|이름)/i.test(h)) nameIdx  = i;
  }
  if (nameIdx === -1) throw new Error('No "Name" column found.');

  // Dedupe by email (if present)
  var data = rows.slice();

  // Shuffle in-place
  data.sort(function(){ return Math.random() - 0.5; });

  // Assign groups
  var assignments = []; // [groupNumber, name]
  if (USE_TEAM_SIZE) {
    for (var j=0;j<data.length;j++){
      assignments.push([ Math.floor(j/TEAM_SIZE)+1, String(data[j][nameIdx]) ]);
    }
  } else {
    var tc = teamCountFor(data.length);
    for (var k=0;k<data.length;k++){
      assignments.push([ (k % tc) + 1, String(data[k][nameIdx]) ]);
    }
  }

  // Group map: {1:[names], 2:[names], ...}
  var map = {}, maxLen = 0, teams = [];
  for (var a=0;a<assignments.length;a++){
    var n = assignments[a][0], nm = assignments[a][1];
    if (!map[n]) { map[n] = []; teams.push(n); }
    map[n].push(nm);
    if (map[n].length > maxLen) maxLen = map[n].length;
  }
  teams.sort(function(a,b){ return a-b; });

  // Build Board
  var board = ss.getSheetByName(BOARD_SHEET) || ss.insertSheet(BOARD_SHEET);
  board.clear();

  var colors = ['#fde2e2','#e2f0fe','#e6f4ea','#fff3cd','#efe1ff','#d7f3f7',
                '#ffdfe5','#e8eaf6','#f1f8e9','#fff8e1','#ede7f6','#e0f2f1'];

  var colsUsed = 0;
  teams.forEach(function(num, idx){
    var col = 1 + idx * COL_GAP; colsUsed = col;
    var members = map[num].map(function(n){ return [n]; });
    var label = LABEL_PREFIX + num;

    board.getRange(1, col).setValue(label).setFontWeight('bold').setHorizontalAlignment('center');
    if (members.length) board.getRange(2, col, members.length, 1).setValues(members);

    board.getRange(1, col, Math.max(2, maxLen+1), 1)
         .setBackground(colors[idx % colors.length])
         .setBorder(true,true,true,true,true,true);
    board.setColumnWidth(col, TEAM_COL_WIDTH);
    board.getRange(2, col, Math.max(1, maxLen), 1).setWrap(true);
  });
  board.setFrozenRows(1);

  // ---- Chart with group labels, data stored on hidden sheet ----
  var counts = teams.map(function(num){ return [LABEL_PREFIX + num, map[num].length]; });
  var dataSheet = ss.getSheetByName(CHART_DATA_SHEET) || ss.insertSheet(CHART_DATA_SHEET);
  dataSheet.clear(); dataSheet.hideSheet();
  dataSheet.getRange(1,1,1,2).setValues([['Group','# Members']]);
  if (counts.length) dataSheet.getRange(2,1,counts.length,2).setValues(counts);

  // Place pie/donut on the right (or below if narrow)
  var placeRight = colsUsed <= 16;
  var chartRow = placeRight ? 1 : (maxLen + 3);
  var chartCol = placeRight ? (colsUsed + 2) : 1;

  board.getCharts().forEach(function(c){ board.removeChart(c); });
}

function freezeBoard() {
  const ss = SpreadsheetApp.getActive();
  const b  = ss.getSheetByName(BOARD_SHEET);
  if (!b) { SpreadsheetApp.getUi().alert('No board yet.'); return; }
  const name = 'Board ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  const copy = b.copyTo(ss).setName(name);
  ss.setActiveSheet(copy);
}

function clearBoard_() {
  var ss = SpreadsheetApp.getActive();
  var b  = ss.getSheetByName(BOARD_SHEET) || ss.insertSheet(BOARD_SHEET);
  b.clear();
  b.getRange(1,1).setValue('No responses yet.');
}

