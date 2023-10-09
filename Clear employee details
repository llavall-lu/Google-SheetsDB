//This clears all the saved Employee Details from the Employee Details sheet.
function clearEmployeeDetails(empTemplateHeaderCells, selectedEmployeeKey) {
  Logger.log("Starting clearEmployeeDetails")
  //Logger.log("Starting clearEmployeeDetails: empTemplateHeadersCell: " + JSON.stringify(empTemplateHeadersCell) + ', selectedEmployeeKey: '+ JSON.stringify(selectedEmployeeKey));

  var userProperties = PropertiesService.getUserProperties();
  
  // Clear the cells if '<Create New Employee>' is selected
  for (var namedRange in empTemplateHeaderCells) {
    for (var key in empTemplateHeaderCells[namedRange]) {
      var cellLocation = empTemplateHeaderCells[namedRange][key];
      var cell = empSheet.getRange(cellLocation.row, cellLocation.column);
      cell.setValue("");
    }
  }

  // Retrieve the existing selectedEmployee dictionary and clear its values
  var currentSelectedEmployee = JSON.parse(userProperties.getProperty(selectedEmployeeKey));
  for (var key in currentSelectedEmployee) {
    currentSelectedEmployee[key] = "";
  }

  // Store the updated selectedEmployee dictionary in user properties
  userProperties.setProperty(selectedEmployeeKey, JSON.stringify(currentSelectedEmployee));
  
  var dropdownCellPos = empSheet.getRange(dropdownCellValue);
  var dropdownValue = dropdownCellPos.getValue();

  //Logger.log ('dropdownValue = ' + JSON.stringify(dropdownValue));

  if (dropdownValue != ('<Create New Employee>')) {
    // Call the backgroundColour function on the edited cells
    for (var namedRange in empTemplateHeaderCells) {
      for (var key in empTemplateHeaderCells[namedRange]) {
        var cellLocation = empTemplateHeaderCells[namedRange][key];
        var cell = empSheet.getRange(cellLocation.row, cellLocation.column);
        backgroundColour(cell, empTemplateHeaderCells, currentSelectedEmployee);
      }
    }
  } else {
    Logger.log('clearEmployeeDetails: No need to set Background Cell Colours');
  }

  Logger.log("Finishing clearEmployeeDetails")
}

/** **************************************************************************************************** */
