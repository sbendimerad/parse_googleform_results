function calculateScores() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var mainSheet = spreadsheet.getSheetByName('Reponse'); // Replace 'my_sheet' with your sheet name
  var correspondenceSheet = spreadsheet.getSheetByName('Correspondence'); // Correspondence sheet
  var lastRow = mainSheet.getLastRow();
  var lastColumn = mainSheet.getLastColumn();
  
  // Get the correspondence data as an object for efficient lookup
  var correspondenceData = getCorrespondenceData(correspondenceSheet);
  
  // Add a new column for total scores
  mainSheet.insertColumnAfter(lastColumn);
  mainSheet.getRange(1, lastColumn + 1).setValue('Total Score'); // Set column header
  
  // Loop through each row (excluding the header row)
  for (var row = 2; row <= lastRow; row++) {
    var rowData = mainSheet.getRange(row, 1, 1, lastColumn).getValues()[0];
    var totalScore = 0;
    
    // Loop through each column (starting from the third column)
    for (var col = 3; col <= lastColumn; col++) {
      var cellValue = rowData[col - 1];
      
      // Check if the cell contains a string before splitting
      if (typeof cellValue === 'string') {
        var listAnswers = cellValue.split(", "); // Split responses by comma
        
        // Calculate the score for this cell based on the correspondence data
        var cellScore = calculateCellScore(listAnswers, correspondenceData);
        
        // Add cell score to the total score for the row
        totalScore += cellScore;
      }
    }
    
    // Set the total score in the new column
    mainSheet.getRange(row, lastColumn + 1).setValue(totalScore);
  }
}

function getCorrespondenceData(correspondenceSheet) {
  var data = correspondenceSheet.getRange(2, 1, correspondenceSheet.getLastRow() - 1, 2).getValues();
  var correspondenceData = {};
  
  // Create an object for efficient lookup
  for (var i = 0; i < data.length; i++) {
    var responseOption = data[i][0].trim(); // Trim spaces from the response option
    var pointValue = data[i][1];
    correspondenceData[responseOption] = pointValue;
  }
  
  return correspondenceData;
}

function calculateCellScore(listAnswers, correspondenceData) {
  var cellScore = 0;
  
  for (var i = 0; i < listAnswers.length; i++) {
    var responseOption = listAnswers[i];
    
    // Check if the response option exists in the correspondence data
    if (correspondenceData.hasOwnProperty(responseOption)) {
      cellScore += correspondenceData[responseOption];
    }
  }
  
  return cellScore;
}
