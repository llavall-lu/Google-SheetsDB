

function showSavedEmployeeDetails() {
  Logger.log("Starting showSavedEmployeeDetails");

  var userProperties = PropertiesService.getUserProperties();

  var selectedEmployeeJSON = userProperties.getProperty(selectedEmployeeKey);
  var selectedEmployee = JSON.parse(selectedEmployeeJSON);

  var empTemplateHeaderCellsJSON = userProperties.getProperty('empTemplateHeaderCells'); 
  var empTemplateHeaderCells = JSON.parse(empTemplateHeaderCellsJSON);

  var message;
  if (selectedEmployee !== '') {
    message = "The following Employee Details have been loaded:\n\n";
  } else {
    message = "Either no Employee Details are loaded or you are creating a new Employee manually:\n\n";
  }

  for (var namedRange in empTemplateHeaderCells) {
    message += "---   " + namedRange + "   ---\n";

    for (var key in empTemplateHeaderCells[namedRange]) {
      if (key.toLowerCase().includes('date') && selectedEmployee[key] !== '') {

        var formattedDate = Utilities.formatDate(new Date(selectedEmployee[key]), "GMT", "dd/MM/yyyy");
        message += key + ":       " + formattedDate + "\n";
      } else {
        message += key + ":       " + (selectedEmployee[key] || '') + "\n";
      }
    }

    message += "\n"; 
  }

  SpreadsheetApp.getUi().alert(message);

  Logger.log("Finished showSavedEmployeeDetails");
}
function populateEmployeeDetails(response, empNameMatch, storeToSelectedEmployee = true) {

  Logger.log('Starting populateEmployeeDetails')

  Logger.log('empNameMatch = ' + empNameMatch)

  var userProperties = PropertiesService.getUserProperties();
  var empTemplateHeaderCells = JSON.parse(userProperties.getProperty('empTemplateHeaderCells'));

  var unhiddenNamedRanges = JSON.parse(userProperties.getProperty('unhiddenNamedRanges'));

  if (empNameMatch) {
    var selectedEmployee = {}
    unhiddenNamedRanges.push("mainDetails", "hoursSection") 
    Logger.log('The following namedRanges are unhidden: ' + unhiddenNamedRanges.join(", "))
  }

  if (response) {
    for (var namedRange in empTemplateHeaderCells) { 

      if (!unhiddenNamedRanges.includes(namedRange)) {
        continue;
      }

      Logger.log('Checking namedRange: ' + namedRange);

      for (var key in response.response) { 

        if (empTemplateHeaderCells[namedRange].hasOwnProperty(key)) {

          var cellDetails = empTemplateHeaderCells[namedRange][key];
          Logger.log('Checking key: ' + key);

          var cell = empSheet.getRange(cellDetails.row, cellDetails.column);

          if (key.toLowerCase().includes('date') && response.response[key] !== '') {

            var date = new Date(response.response[key]);
            var formattedDate = Utilities.formatDate(date, "GMT", "dd/MM/yyyy");
            cell.setValue(formattedDate);
          } else {

            cell.setValue(response.response[key]);
          }

          if (storeToSelectedEmployee) {
            selectedEmployee[key] = response.response[key];
          }
        }
      }
    }
  }

  for (var namedRange in empTemplateHeaderCells) {
    for (var key in selectedEmployee) {
      if (empTemplateHeaderCells[namedRange].hasOwnProperty(key)) {
        var cellDetails = empTemplateHeaderCells[namedRange][key];
        var cell = empSheet.getRange(cellDetails.row, cellDetails.column);
        backgroundColour(cell, empTemplateHeaderCells, selectedEmployee);
      }
    }
  }

  userProperties.setProperty('selectedEmployee', JSON.stringify(selectedEmployee));

  Logger.log('selectedEmployee = ' + JSON.stringify(selectedEmployee));

  Logger.log('Finishing populateEmployeeDetails')
}
function populateEmployeeDetailsOLD(value, row, column, focusSection = null) {
  Logger.log("Starting populateEmployeeDetails: " + value);

  Logger.log('value = ' + value);

  var userProperties = PropertiesService.getUserProperties();
  var objectNameDictionary = JSON.parse(userProperties.getProperty('nameDictionary'));
  var filmDetails = JSON.parse(userProperties.getProperty('filmDetails'));
  var empTemplateHeaderCells = JSON.parse(userProperties.getProperty('empTemplateHeaderCells'));
  var empTemplateNamedRanges = JSON.parse(userProperties.getProperty('empTemplateNamedRanges'));
  var empTemplateDropdownCells = JSON.parse(userProperties.getProperty('empTemplateDropdownCells'));
  var storedMatchingResponse = JSON.parse(userProperties.getProperty('matchingResponse'));

  var matchingEntry = objectNameDictionary.find(entry => entry.objectName === value);

  Logger.log('main Dropdown matchingEntry = ' + JSON.stringify(matchingEntry))

  if (!matchingEntry) {
    Logger.log("No matching entry in the main ObjectName Dropdown found for: " + value);

    var valueNoSpace = value.replace(/\s+/g, '')

    Logger.log("valueNoSpace: " + valueNoSpace);

    var namedRangeMatch = empTemplateNamedRanges[valueNoSpace];

    if (namedRangeMatch) {
      Logger.log("YES!, namedRangeMatch matches entry found in NamedRange for: " + JSON.stringify(valueNoSpace));

      var dropdownHeader = getHeaderNameByCell(row, column, value);
      Logger.log('dropdownHeader = ' + JSON.stringify(dropdownHeader))

      if (!dropdownHeader) {
        Logger.log("Unable to determine the dropdownHeader for the given row/column. Finishing populateEmployeeDetails");

        return;
      }

      if (storedMatchingResponse.response && storedMatchingResponse.response[dropdownHeader] !== value) {

        Logger.log(value + "(value) is different from " + storedMatchingResponse.response[dropdownHeader] + "(the stored matching response!)");

        loadingSection(value, row, column, dropdownHeader)

      } else {

        Logger.log('values are the same. Finishing populateEmployeeDetails')

        return;
      }

    } else {
      Logger.log('No namedRange matches this row or column. Please refresh the page and start again.') 
      Logger.log("Finishing populateEmployeeDetails");

      return;
    }

  } else {

    var formResponseRow = formSheet.getRange(matchingEntry.rowNumber, 1, 1, formSheet.getLastColumn()).getValues();
    var formHeaders = formSheet.getRange(1, 1, 1, formSheet.getLastColumn()).getValues()[0];

    var matchingResponse = {
      objectName: value,
      response: {}
    };
    for (var i = 0; i < formHeaders.length; i++) {
      matchingResponse.response[formHeaders[i]] = formResponseRow[0][i];
    }

    userProperties.setProperty('matchingResponse', JSON.stringify(matchingResponse));
    Logger.log('matchingResponse = ' + JSON.stringify(matchingResponse))

    Logger.log('empTemplateDropdownCells = ' + JSON.stringify(empTemplateDropdownCells))

    for (var dropdownHeader in empTemplateDropdownCells) {

      if (matchingResponse.response.hasOwnProperty(dropdownHeader)) {

        var dropdownValue = matchingResponse.response[dropdownHeader].replace(/\s+/g, '');

        Logger.log('dropdownValue = ' + dropdownValue)

        if (empTemplateNamedRanges.hasOwnProperty(dropdownValue)) {

          var namedRangeInfo = empTemplateNamedRanges[dropdownValue];
          empSheet.showRows(namedRangeInfo.rowStart, namedRangeInfo.rowEnd - namedRangeInfo.rowStart + 1);
        }
      }
    }

    var visibleNamedRanges = {};

    for (var namedRange in empTemplateNamedRanges) {
      var firstRow = empTemplateNamedRanges[namedRange].rowStart;

      var isRowHidden = empSheet.isRowHiddenByUser(firstRow);

      if (!isRowHidden) {
        visibleNamedRanges[namedRange] = true;
      }
    }

    Logger.log("Visible named ranges: " + JSON.stringify(visibleNamedRanges));
}
function clearEmployeeDetails(empTemplateHeaderCells, selectedEmployeeKey) {
  Logger.log("Starting clearEmployeeDetails")

  var userProperties = PropertiesService.getUserProperties();

  for (var namedRange in empTemplateHeaderCells) {
    for (var key in empTemplateHeaderCells[namedRange]) {
      var cellLocation = empTemplateHeaderCells[namedRange][key];
      var cell = empSheet.getRange(cellLocation.row, cellLocation.column);
      cell.setValue("");
    }
  }

  var currentSelectedEmployee = JSON.parse(userProperties.getProperty(selectedEmployeeKey));
  for (var key in currentSelectedEmployee) {
    currentSelectedEmployee[key] = "";
  }

  userProperties.setProperty(selectedEmployeeKey, JSON.stringify(currentSelectedEmployee));

  var dropdownCellPos = empSheet.getRange(dropdownCellValue);
  var dropdownValue = dropdownCellPos.getValue();

  if (dropdownValue != ('<Create New Employee>')) {

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
}
