function onEdit(e) {
  var sheet = e.source.getActiveSheet();
  var range = e.range;
  
  // Check if the edited cell is in column C (3) and not empty
  if (range.getColumn() === 3 && range.getValue() !== "") {
    // Get the active cell row
    var row = range.getRow();
    
    // Get the current date and time
    var date = new Date();
    
    // Write the date in column A and the time in column B
    sheet.getRange(row, 1).setValue(date);
    sheet.getRange(row, 2).setValue(date.toLocaleTimeString());
    
    // Automatically increment the value in the adjacent tab (column D)
    var currentValue = sheet.getRange(row, 4).getValue() || 0;
    sheet.getRange(row, 4).setValue(currentValue + 1);
  }
  
  // Calculate the sum at the end of the sheet
  if (range.getColumn() === 1 || range.getColumn() === 2 || range.getColumn() === 4) {
    var lastRow = sheet.getLastRow();
    var sumRange = sheet.getRange(2, 4, lastRow - 1, 1);
    var sum = sumRange.getValues().reduce(function(acc, val) {
      return acc + (val[0] || 0);
    }, 0);
    sheet.getRange(lastRow, 4).setValue(sum);
  }
}
