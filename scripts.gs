/** @OnlyCurrentDoc */
function compteurProprietaire(coproprietaire = "dupont") {
  
  var sheetName = 'Releve ' + coproprietaire;

  var spreadsheet = SpreadsheetApp.getActive();
  try { 
    var sheet = spreadsheet.getSheetByName(sheetName);
    if(sheet != null) {
      spreadsheet.deleteSheet(sheet);
    }
  } finally {    
  }


  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Compteurs eau'),true);
  var sourceData = spreadsheet.getRange('A1:Y162');
  spreadsheet.insertSheet(spreadsheet.getActiveSheet().getIndex() + 1).activate();
  spreadsheet.getActiveSheet().setHiddenGridlines(true);
  
  spreadsheet.getActiveSheet().setName(sheetName);

  var pivotTable = spreadsheet.getRange('A1').createPivotTable(sourceData);
  var pivotGroup = pivotTable.addRowGroup(2);
  pivotTable = spreadsheet.getRange('A1').createPivotTable(sourceData);
  pivotGroup = pivotTable.addRowGroup(2);
  pivotGroup.showTotals(false);
  pivotGroup = pivotTable.addRowGroup(12);
  pivotGroup.showTotals(false);
  pivotGroup = pivotTable.addRowGroup(13);
  pivotGroup.showTotals(false);
  pivotGroup = pivotTable.addRowGroup(14);
  pivotGroup.showTotals(false);
  pivotGroup = pivotTable.addRowGroup(16);
  pivotGroup.showTotals(false);

  var thisYear = Utilities.formatDate(new Date(), "GMT+1", "yyyy").toString();
  
  var pivotValue = pivotTable.addPivotValue(17, SpreadsheetApp.PivotTableSummarizeFunction.SUM);
  pivotValue.setDisplayName(thisYear-3);
  pivotValue = pivotTable.addPivotValue(18, SpreadsheetApp.PivotTableSummarizeFunction.SUM);
  pivotValue.setDisplayName(thisYear-2);
  pivotValue = pivotTable.addPivotValue(19, SpreadsheetApp.PivotTableSummarizeFunction.SUM);
  pivotValue.setDisplayName(thisYear-1);
  var criteria = SpreadsheetApp.newFilterCriteria()
  .setVisibleValues([coproprietaire])
  .build();
  pivotTable.addFilter(2, criteria);
  spreadsheet.getRange('I1').activate();
  spreadsheet.getCurrentCell().setValue(thisYear);

  spreadsheet.getRange('1:15').activate();
  spreadsheet.getActiveSheet().insertRowsBefore(spreadsheet.getActiveRange().getRow(), 15);
  spreadsheet.getRange('B1').activate();
  spreadsheet.getCurrentCell().setValue('Blabla');
  spreadsheet.getActiveRangeList().setFontWeight('bold');

  spreadsheet.getRange('I16:I22').activate();
  spreadsheet.getActiveRangeList().setBackground('#00ff00');
  spreadsheet.getRange('A1').activate();


  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Relevé'), true);
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName(sheetName), true);
  spreadsheet.getRange('A1').activate();
  spreadsheet.getRange('\'Relevé\'!1:14').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Relevé'), true);
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName(sheetName), true);
  spreadsheet.getRange('I4').activate();

  spreadsheet.getRange('15:23').activate();
  spreadsheet.getActiveRangeList().setFontSize(22);
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Relevé'), true);
  spreadsheet.getRange('24:24').activate();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName(sheetName), true);
  spreadsheet.getRange('24:24').activate();
  spreadsheet.getRange('\'Relevé\'!24:24').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);

  spreadsheet.getRange('R:R').activate();
  spreadsheet.getActiveSheet().hideColumns(spreadsheet.getActiveRange().getColumn(), spreadsheet.getActiveRange().getNumColumns());

  spreadsheet.getRange('N14:N24').activate();
  spreadsheet.getActiveRangeList().setBorder(null, null, null, true, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_THICK)
  .setBorder(null, null, null, true, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_THICK);
  spreadsheet.getRange('A14:A24').activate();
  spreadsheet.getActiveRangeList().setBorder(null, true, null, null, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_THICK)
  .setBorder(null, true, null, null, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_THICK);

  spreadsheet.getRange('A:N').activate();
  spreadsheet.setCurrentCell(spreadsheet.getRange('A5'));
  spreadsheet.getActiveSheet().setColumnWidths(1, 14, 152);

  spreadsheet.getRange('A1').activate();
  
  insertGraphique();
};



function tousLesCompteurs() {

  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Compteurs eau'), true);

  var range = spreadsheet.getActiveRangeList
  var range = spreadsheet.getRange('B2:B');
  var values = range.getValues();
  var dejavue = [];
  
  // Print values from a 3x3 box.
  for (var row in values) {
      var val = values[row][0];
      if (dejavue.indexOf(val) == -1){
        var newLength = dejavue.push(val);
        Logger.log(dejavue);
        compteurProprietaire(val);
      } 
  }
};



function insertGraphique() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('S15').activate();
  spreadsheet.getCurrentCell().setValue('Consommations');
  spreadsheet.getRange('S15:U15').activate().mergeAcross();
  spreadsheet.getActiveRangeList().setHorizontalAlignment('right');
  spreadsheet.getRange('S16').activate();
  spreadsheet.getCurrentCell().setValue('Cpt');
  spreadsheet.getRange('T16').activate();
  spreadsheet.getCurrentCell().setValue('2018');
  spreadsheet.getRange('U16').activate();
  spreadsheet.getCurrentCell().setValue('2019');
  spreadsheet.getRange('S17').activate();
  spreadsheet.getCurrentCell().setFormula('=e17');
  spreadsheet.getRange('S18').activate();
  spreadsheet.getRange('S17').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getRange('S19').activate();
  spreadsheet.getRange('S17').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getRange('S20').activate();
  spreadsheet.getRange('S17').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getRange('S21').activate();
  spreadsheet.getRange('S17').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getRange('S17').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getRange('S22').activate();
  spreadsheet.getRange('S17').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getRange('S23').activate();
  spreadsheet.getRange('S17').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getRange('T17').activate();
  spreadsheet.getCurrentCell().setFormula('=g17-f17');
  spreadsheet.getRange('U17').activate();
  spreadsheet.getCurrentCell().setFormula('=h17-g17');
  spreadsheet.getRange('T18').activate();
  spreadsheet.getRange('T17:U17').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getRange('T19').activate();
  spreadsheet.getRange('T17:U17').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getRange('T20').activate();
  spreadsheet.getRange('T17:U17').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getRange('T21').activate();
  spreadsheet.getRange('T17:U17').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getRange('T22').activate();
  spreadsheet.getRange('T17:U17').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getRange('T23').activate();
  spreadsheet.getRange('T17:U17').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getRange('S16:U23').activate();
  var sheet = spreadsheet.getActiveSheet();
  var chart = sheet.newChart()
  .asColumnChart()
  .addRange(spreadsheet.getRange('S16:U23'))
  .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
  .setTransposeRowsAndColumns(false)
  .setNumHeaders(-1)
  .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH)
  .setPosition(7, 10, 145, 18)
  .build();
  sheet.insertChart(chart);
  var charts = sheet.getCharts();
  chart = charts[charts.length - 1];
  sheet.removeChart(chart);
  chart = sheet.newChart()
  .asColumnChart()
  .addRange(spreadsheet.getRange('S16:U23'))
  .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
  .setTransposeRowsAndColumns(false)
  .setNumHeaders(-1)
  .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH)
  .setPosition(12, 10, 150, 20)
  .build();
  sheet.insertChart(chart);
  charts = sheet.getCharts();
  chart = charts[charts.length - 1];
  sheet.removeChart(chart);
  chart = sheet.newChart()
  .asColumnChart()
  .addRange(spreadsheet.getRange('S16:U23'))
  .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
  .setTransposeRowsAndColumns(false)
  .setNumHeaders(-1)
  .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH)
  .setOption('height', 294)
  .setOption('width', 476)
  .setPosition(15, 11, 122, 15)
  .build();
  sheet.insertChart(chart);
  charts = sheet.getCharts();
  chart = charts[charts.length - 1];
  sheet.removeChart(chart);
  chart = sheet.newChart()
  .asColumnChart()
  .addRange(spreadsheet.getRange('S16:U23'))
  .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
  .setTransposeRowsAndColumns(false)
  .setNumHeaders(1)
  .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH)
  .setOption('bubble.stroke', '#000000')
  .setOption('height', 294)
  .setOption('width', 476)
  .setPosition(15, 11, 122, 15)
  .build();
  sheet.insertChart(chart);
  charts = sheet.getCharts();
  chart = charts[charts.length - 1];
  sheet.removeChart(chart);
  chart = sheet.newChart()
  .asColumnChart()
  .addRange(spreadsheet.getRange('S16:U23'))
  .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
  .setTransposeRowsAndColumns(false)
  .setNumHeaders(1)
  .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH)
  .setOption('bubble.stroke', '#000000')
  .setOption('useFirstColumnAsDomain', true)
  .setOption('height', 294)
  .setOption('width', 476)
  .setPosition(15, 11, 122, 15)
  .build();
  sheet.insertChart(chart);
};


