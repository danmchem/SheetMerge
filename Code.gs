function onInstall(e) {
  onOpen(e);
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Merge Tool')
    .addItem('Open Merge Dialog', 'showMergeDialog')
    .addToUi();
}

function showMergeDialog() {
  var html = HtmlService.createHtmlOutputFromFile('Dialog')
    .setWidth(500)
    .setHeight(600)
    .setTitle('Merge Sheets');
  SpreadsheetApp.getUi().showModalDialog(html, 'Merge Sheets');
}

function getSheetNames() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  return sheets.map(function(sheet) {
    return sheet.getName();
  });
}

function getColumnHeaders(sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  return headers.map(function(header, index) {
    return {name: header, index: index};
  });
}

function performMerge(params) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var primarySheet = ss.getSheetByName(params.primarySheet);
    var mergingSheet = ss.getSheetByName(params.mergingSheet);
    
    if (!primarySheet || !mergingSheet) {
      throw new Error('Could not find one or both sheets');
    }
    
    // Backup if requested
    if (params.createBackup) {
      var backupName = params.primarySheet + '_backup_' + new Date().getTime();
      primarySheet.copyTo(ss).setName(backupName);
    }
    
    // Get all data
    var primaryData = primarySheet.getDataRange().getValues();
    var mergingData = mergingSheet.getDataRange().getValues();
    
    if (primaryData.length < 2 || mergingData.length < 2) {
      throw new Error('Sheets must have headers and at least one data row');
    }
    
    // Get column indices
    var primaryHeaders = primaryData[0];
    var mergingHeaders = mergingData[0];
    
    var primaryCompareCol = primaryHeaders.indexOf(params.primaryCompareColumn);
    var mergingCompareCol = mergingHeaders.indexOf(params.mergingCompareColumn);
    
    if (primaryCompareCol === -1 || mergingCompareCol === -1) {
      throw new Error('Comparison columns not found');
    }
    
    // Build column mappings
    var columnMappings = [];
    for (var i = 0; i < params.columnsToTransfer.length; i++) {
      var transferCol = params.columnsToTransfer[i];
      var mergingColIndex = mergingHeaders.indexOf(transferCol.mergingColumn);
      
      if (mergingColIndex === -1) continue;
      
      var primaryColIndex;
      if (transferCol.replaceColumn) {
        primaryColIndex = primaryHeaders.indexOf(transferCol.replaceColumn);
        if (primaryColIndex === -1) {
          // Column doesn't exist, will add new column
          primaryColIndex = primaryHeaders.length;
          primaryHeaders.push(transferCol.replaceColumn);
        }
      } else {
        // Add as new column
        primaryColIndex = primaryHeaders.length;
        primaryHeaders.push(transferCol.mergingColumn);
      }
      
      columnMappings.push({
        mergingIndex: mergingColIndex,
        primaryIndex: primaryColIndex
      });
    }
    
    // Extend primary data rows if needed
    for (var i = 0; i < primaryData.length; i++) {
      while (primaryData[i].length < primaryHeaders.length) {
        primaryData[i].push('');
      }
    }
    
    // Build lookup for primary sheet rows
    var primaryLookup = {};
    for (var i = 1; i < primaryData.length; i++) {
      var key = primaryData[i][primaryCompareCol];
      if (params.skipEmptyComparison && (key === '' || key === null || key === undefined)) {
        continue;
      }
      if (!params.matchCase && typeof key === 'string') {
        key = key.toLowerCase();
      }
      
      if (!primaryLookup[key]) {
        primaryLookup[key] = [];
      }
      primaryLookup[key].push(i);
    }
    
    // Track matched rows in merging sheet
    var matchedMergingRows = {};
    var newRows = [];
    
    // Process merging data
    for (var i = 1; i < mergingData.length; i++) {
      var mergingKey = mergingData[i][mergingCompareCol];
      
      if (params.skipEmptyComparison && (mergingKey === '' || mergingKey === null || mergingKey === undefined)) {
        continue;
      }
      
      var lookupKey = mergingKey;
      if (!params.matchCase && typeof lookupKey === 'string') {
        lookupKey = lookupKey.toLowerCase();
      }
      
      var matchingPrimaryRows = primaryLookup[lookupKey];
      
      if (matchingPrimaryRows && matchingPrimaryRows.length > 0) {
        // Match found - update all matching rows (Cartesian product)
        matchedMergingRows[i] = true;
        
        for (var j = 0; j < matchingPrimaryRows.length; j++) {
          var primaryRowIndex = matchingPrimaryRows[j];
          
          // Copy data from merging sheet to primary sheet
          for (var k = 0; k < columnMappings.length; k++) {
            var mapping = columnMappings[k];
            primaryData[primaryRowIndex][mapping.primaryIndex] = mergingData[i][mapping.mergingIndex];
          }
        }
      } else if (params.appendUnmatched) {
        // No match found - prepare to append
        var newRow = new Array(primaryHeaders.length).fill('');
        
        // Set comparison column value
        newRow[primaryCompareCol] = mergingKey;
        
        // Set transferred column values
        for (var k = 0; k < columnMappings.length; k++) {
          var mapping = columnMappings[k];
          newRow[mapping.primaryIndex] = mergingData[i][mapping.mergingIndex];
        }
        
        newRows.push(newRow);
      }
    }
    
    // Write back to primary sheet
    primarySheet.clear();
    primarySheet.getRange(1, 1, primaryData.length, primaryHeaders.length).setValues(primaryData);
    
    // Append new rows if any
    if (newRows.length > 0) {
      primarySheet.getRange(primaryData.length + 1, 1, newRows.length, primaryHeaders.length).setValues(newRows);
    }
    
    return {
      success: true,
      message: 'Merge completed successfully. Updated/added ' + (Object.keys(matchedMergingRows).length + newRows.length) + ' rows.'
    };
    
  } catch (error) {
    return {
      success: false,
      message: 'Error: ' + error.message
    };
  }
}
