

function createSavingDictionary() {
  Logger.log("Starting createSavingDictionary");

  var userProperties = PropertiesService.getUserProperties();

  var empTemplateHeaderCellsJSON = userProperties.getProperty('empTemplateHeaderCells');
  var empTemplateHeaderCells = JSON.parse(empTemplateHeaderCellsJSON);

  var selectedEmployeeJSON = userProperties.getProperty(selectedEmployeeKey);
  var selectedEmployee = JSON.parse(selectedEmployeeJSON);

  if (empTemplateHeaderCells) {
    var savingDictionary = {};

    var editedCellsExist = false; 

    for (var namedRange in empTemplateHeaderCells) {
      savingDictionary[namedRange] = {};

      for (var key in empTemplateHeaderCells[namedRange]) {
        var cellInfo = empTemplateHeaderCells[namedRange][key];
        var row = cellInfo.row;
        var column = cellInfo.column;

        if (empSheet.getRange(row, column).getBackground() === editedColour) {
          var value = empSheet.getRange(row, column).getValue();

          savingDictionary[namedRange][key] = value;
          editedCellsExist = true; 
        }
      }
    }

    if (editedCellsExist) {
      var message = "WARNING: You are about to save and overwrite the following person's details:\n\n";
      message += selectedEmployee["First Name"] + " " + selectedEmployee["Surname"] + " - " + selectedEmployee["Department"] + "\n\n";
      message += "Details that are going to be saved:\n\n";

      for (var namedRange in savingDictionary) {
        if (Object.keys(savingDictionary[namedRange]).length > 0) { 
          message += "--- " + namedRange + " ---\n";

          for (var header in savingDictionary[namedRange]) {
            var value = savingDictionary[namedRange][header];
            message += header + ": " + value + "\n";
          }

          message += "\n"; 
        }
      }

      SpreadsheetApp.getUi().alert(message);
    } else {

      SpreadsheetApp.getUi().alert("No details to be saved.");
    }
  }

  Logger.log("Finished createSavingDictionary");
}


