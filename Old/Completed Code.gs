function resetLoadingIndicator(){
  var loadingCellPos = empSheet.getRange(loadingCellValue);
  loadingCellPos.setValue('');
  loadingCellPos.setBackground(whiteColour);
  loadingCellPos.setFontColor(blackColour);

}

function loadingIndicator(){
  var loadingCellPos = empSheet.getRange(loadingCellValue);
  var loadingCell = loadingCellPos.getValue();

  if (loadingCell !== 'Loading...')  {
    loadingCellPos.setValue('Loading...');
    //Logger.log('loadingCell = ' + JSON.stringify(loadingCell))
    loadingCellPos.setBackground(loadingColour);
    loadingCellPos.setFontColor(whiteColour);
    // Force the changes to apply immediately
    SpreadsheetApp.flush();
  } else {
    // Force the changes to apply immediately
    SpreadsheetApp.flush();
    // Clear the 'Loading...' indicator
    //Logger.log('loadingCell = ' + JSON.stringify(loadingCell))
    loadingCellPos.setValue('');
    loadingCellPos.setBackground(whiteColour);
    loadingCellPos.setFontColor(blackColour);
  }
}

/** **************************************************************************************************** */

//Process the OBJECT NAMES and populate the dropdown menu in Employee Details
function processFormResponses() {
  Logger.log("Starting processFormResponses");

  var formDataRange = formSheet.getDataRange();
  var formData = formDataRange.getValues();
  var headers = formData[0];
  var nameDictionary = [];

  var firstNameIndex = headers.indexOf('First Name');
  var lastNameIndex = headers.indexOf('Surname');
  var deptIndex = headers.indexOf('Department');

  for (var i = 1; i < formData.length; i++) {
    var row = formData[i];

    var firstName = row[firstNameIndex];
    var lastName = row[lastNameIndex];
    var deptName = row[deptIndex];
    var objectName = firstName + ' ' + lastName + ' ' + deptName;
    
    var entry = {
      objectName: objectName,
      rowNumber: i + 1 // since rows in Spreadsheet start from 1 and not 0
    };

    nameDictionary.push(entry);
  }

  // Sort nameDictionary alphabetically based on objectName
  nameDictionary.sort(function(a, b) {
    return a.objectName.localeCompare(b.objectName);
  });

  // Convert the structured dictionary into a list of object names for the dropdown
  var dropdownNames = nameDictionary.map(function(entry) {
    return entry.objectName;
  });
  dropdownNames.unshift('<Create New Employee>');

  // Store the nameDictionary array in user properties
  var userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty('nameDictionary', JSON.stringify(nameDictionary));

  // Populate dropdown menu in Employee Details
  var dropdownRange = empSheet.getRange(dropdownRangeValue);
  dropdownRange.clearDataValidations();

  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(dropdownNames)
    .build();

  dropdownRange.setDataValidation(rule);

  Logger.log('processFormResponses: nameDictionary = ' + JSON.stringify(nameDictionary))

  Logger.log("Finished processFormResponses");
}

/** **************************************************************************************************** */

function processFilmDetails() {

  Logger.log("Starting processFilmDetails");

  var filmDetailsRange = filmSheet.getDataRange(); // Get all data from Film Details
  var filmDetailsValues = filmDetailsRange.getValues(); // Get the values of the data

  var filmDetailsDictionary = {}; // Create a dictionary to store film details

  // Loop through each row in Film Details sheet
  for (var i = 0; i < filmDetailsValues.length; i++) {
    var header = filmDetailsValues[i][0]; // header is in column A
    var value = filmDetailsValues[i][1]; // value is in column B

    // Ignore rows where header is empty
    if (header !== "") {
      filmDetailsDictionary[header] = value; // Add the header-value pair to the dictionary
    }
  }

  // Create the filmDetails object
  var filmDetails = [{"objectName": "Film Details", "response": filmDetailsDictionary}];

  // Save the filmDetails to the user properties
  var userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty('filmDetails', JSON.stringify(filmDetails));

  //Logger.log('filmDetails = ' + JSON.stringify(filmDetails));

  Logger.log("Finished processFilmDetails");
}

/** **************************************************************************************************** 
 * This finds the cells based of the namedRanges from the 'Employee Template' sheet. (The namedRanges are embedded within the google sheet)
*/

function processEmployeeTemplateDetailsCells() {
  Logger.log("Starting processEmployeeTemplateDetailsCells");

  // Assuming formSheet and empSheet are defined elsewhere in your code
  var formDataRange = formSheet.getDataRange();
  var formData = formDataRange.getValues();
  var headers = formData[0];
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var namedRanges = ss.getNamedRanges();

  // filter namedRanges to remove 'employeeDetails'
  var filteredNamedRanges = namedRanges.filter(function(namedRange) {
    return namedRange.getName() !== 'employeeDetails';
  });

  var empTemplateHeaderCells = {};
  var empTemplateNamedRanges = {};
  var empTemplateDropdownCells = {};

  filteredNamedRanges.forEach(function(namedRange) {
    var rangeName = namedRange.getName();
    var range = namedRange.getRange();
    var rangeRowStart = range.getRow();
    var rangeColStart = range.getColumn();
    var rangeData = range.getValues();

    // Add to empTemplateHeaderCells
    empTemplateHeaderCells[rangeName] = {};

    // Add to empTemplateNamedRanges
    empTemplateNamedRanges[rangeName] = {};
    var rangeRowEnd = rangeRowStart + range.getNumRows() - 1;
    empTemplateNamedRanges[rangeName]['rowStart'] = rangeRowStart;
    empTemplateNamedRanges[rangeName]['rowEnd'] = rangeRowEnd;

    for (var i = 0; i < rangeData.length; i++) {
      for (var j = 0; j < rangeData[i].length; j++) {
        var cellValue = rangeData[i][j];
        var row = rangeRowStart + i;
        var column = rangeColStart + j + 1; // Add 1 to get editable cell

        if (headers.includes(cellValue)) {
          empTemplateHeaderCells[rangeName][cellValue] = {
            row: row,
            column: column
          };

          var dropdownValues = checkIfCellIsDropdown(cellValue, row, column);
          // Check if the cell is a dropdown and if so, store it
          if (dropdownValues) {
            empTemplateDropdownCells[cellValue] = {
              dropdown:true,
              values: dropdownValues  // Store the dropdown values
            };
          }
        }
      }
    }
  });

  // Store dictionaries in user properties
  var userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty('empTemplateHeaderCells', JSON.stringify(empTemplateHeaderCells));
  userProperties.setProperty('empTemplateNamedRanges', JSON.stringify(empTemplateNamedRanges));
  userProperties.setProperty('empTemplateDropdownCells', JSON.stringify(empTemplateDropdownCells));

  Logger.log('empTemplateHeaderCells = ' + JSON.stringify(empTemplateHeaderCells, null, 2));
  Logger.log('empTemplateNamedRanges = ' + JSON.stringify(empTemplateNamedRanges, null, 2));
  Logger.log('empTemplateDropdownCells = ' + JSON.stringify(empTemplateDropdownCells, null, 2));

  Logger.log("Finished processEmployeeTemplateDetailsCells");
}

/** **************************************************************************************************** */

// Removes all rows below A4 from the 'Employee Details' sheet and then loads the ranges from 'Employee Templates' sheet.
function resetEmployeeDetailsLayout() {

  Logger.log("Starting resetEmployeeDetailsLayout");
  
  // If there are rows below A5, delete them
  var totalRows = empSheet.getMaxRows();
  if (totalRows > 3) {
    //Logger.log('totalRows =' + totalRows);
    empSheet.deleteRows(4, totalRows - 3);
    //Logger.log('totalRows being deleted =' + totalRows);
  }
  
  // Ensure extra row exists
  empSheet.insertRowAfter(3);

  // Copy the 'mainSection' named range from the 'Employee Templates' sheet to the 'Employee Details' sheet
  var employeeDetails = ss.getRangeByName('employeeDetails');

  if (employeeDetails) {
    // Insert the exact number of rows required for the employeeDetails
    empSheet.insertRows(4, employeeDetails.getNumRows() - 1);  // Subtract 1 since we already added a row earlier
    employeeDetails.copyTo(empSheet.getRange(4, 1));
  } else {
    Logger.log("'employeeDetails' named range not found.");
  }

  // Fetch all named ranges in the spreadsheet
  var namedRanges = ss.getNamedRanges();

  // Iterate through each named range
  namedRanges.forEach(function(namedRange) {
    var rangeName = namedRange.getName();
    var range = namedRange.getRange();
    // If range name insn't employeeDetails or mainDetails, hide the rows;
    if (rangeName !== 'employeeDetails' && rangeName !== 'mainDetails' && rangeName !== 'hoursSection') {
      empSheet.hideRows(range.getRow(), range.getNumRows());
    }
  });
  
  Logger.log("Finished resetEmployeeDetailsLayout");

}

/** **************************************************************************************************** */

function autoDetailsSwitch() {

  Logger.log("Starting autoDetailsSwitch");
  
  loadingIndicator()

  var userProperties = PropertiesService.getUserProperties();
  var autoDetails = userProperties.getProperty('autoDetails') || 'ON';
  var cell = empSheet.getRange(autoDetailsCell);

  if (autoDetails === 'ON') {
    autoDetails = 'OFF';
    cell.setValue('Auto Details: OFF');
  } else {
    autoDetails = 'ON';
    cell.setValue('Auto Details: ON');
  }

  userProperties.setProperty('autoDetails', autoDetails); // Save the updated state

  loadingIndicator()

  Logger.log("Finished autoDetailsSwitch");
}

/** **************************************************************************************************** */


function getHeaderNameByCell(row, column, value) {
    Logger.log('Starting getHeaderNameByCell');
    Logger.log(value + '= row: ' + row +', column: ' + column);

    var userProperties = PropertiesService.getUserProperties();
    var empTemplateHeaderCells = JSON.parse(userProperties.getProperty('empTemplateHeaderCells'));

    for (var namedRange in empTemplateHeaderCells) {
        Logger.log('namedRange = ' + namedRange)
        for (var header in empTemplateHeaderCells[namedRange]) {
          //Logger.log('header2 = ' + header + ' row ' + empTemplateHeaderCells[namedRange][header].row + ' column ' + empTemplateHeaderCells[namedRange][header].column)
            if (empTemplateHeaderCells[namedRange][header].row == row && empTemplateHeaderCells[namedRange][header].column == column) {
                Logger.log('Finishing getHeaderNameByCell and found header ' + header)
                return header;
            }
        }
    }
    Logger.log('Failed getHeaderNameByCell and returning "null"')
    return null;
}

/** **************************************************************************************************** */

function backgroundColour(cell, empTemplateHeaderCells, selectedEmployee) {
  //Logger.log("Starting backgroundColour: cell: " + JSON.stringify(cell) + ', empTemplateHeaderCells: ' + JSON.stringify(empTemplateHeaderCells) + ', selectedEmployee: ' + JSON.stringify(selectedEmployee));

  // Obtain the row and column of the cell, and its current value
  var cellRow = cell.getRow();
  var cellCol = cell.getColumn();
  var cellValue = cell.getValue();

  // Get the namedRange and key in empTemplateHeaderCells where the value matches the cell's row and column
  var foundNamedRangeAndKey;
  for (var namedRange in empTemplateHeaderCells) {
    var key = Object.keys(empTemplateHeaderCells[namedRange]).find(
      key => JSON.stringify(empTemplateHeaderCells[namedRange][key]) === JSON.stringify({row: cellRow, column: cellCol})
    );
    if (key) {
      foundNamedRangeAndKey = { namedRange: namedRange, key: key };
      break;
    }
  }

  // If the cell is not in empTemplateHeaderCells, then exit the function
  if (!foundNamedRangeAndKey) {
    return;
  }

  var cellKey = foundNamedRangeAndKey.key;

  // If the key contains 'date', then handle it as a date
  if (cellKey && cellKey.toLowerCase().includes('date')) {
    if (selectedEmployee[cellKey] !== '' && selectedEmployee[cellKey] !== null) {
      selectedEmployee[cellKey] = Utilities.formatDate(new Date(selectedEmployee[cellKey]), "GMT", "dd/MM/yyyy");
    }
  }

  // Handle the case where the cell is empty
  if (cellValue === '' || cellValue === null) {
    if (selectedEmployee[cellKey] === '' || selectedEmployee[cellKey] === null) {
      cell.setBackground(emptyColour); // Empty cell color
    } else {
      cell.setBackground(editedColour); // Edited cell color
    }
  } 
  // Handle the case where the cell is not empty
  else {
    var cellValueLowerCase = String(cellValue).toLowerCase();
    var selectedValueLowerCase = String(selectedEmployee[cellKey]).toLowerCase();

    if (cellKey && selectedValueLowerCase === cellValueLowerCase) {
      cell.setBackground(loadedColour); // Loaded cell color
    } else {
      cell.setBackground(editedColour); // Edited cell color
    }
  }
  //Logger.log("Finished backgroundColour: cell: " + JSON.stringify(cell) + ', empTemplateHeaderCells: ' + JSON.stringify(empTemplateHeaderCells) + ', selectedEmployee: ' + JSON.stringify(selectedEmployee));
}

/** **************************************************************************************************** */

function createSavingDictionary() {
  Logger.log("Starting createSavingDictionary");

  var userProperties = PropertiesService.getUserProperties();

  // Retrieve and parse the empTemplateHeaderCells dictionary from user properties
  var empTemplateHeaderCellsJSON = userProperties.getProperty('empTemplateHeaderCells');
  var empTemplateHeaderCells = JSON.parse(empTemplateHeaderCellsJSON);

  // Retrieve and parse the selectedEmployee dictionary from user properties
  var selectedEmployeeJSON = userProperties.getProperty(selectedEmployeeKey);
  var selectedEmployee = JSON.parse(selectedEmployeeJSON);

  if (empTemplateHeaderCells) {
    var savingDictionary = {};

    var editedCellsExist = false; // Flag to track if any cells were edited

    // Loop through the named ranges in empTemplateHeaderCells
    for (var namedRange in empTemplateHeaderCells) {
      savingDictionary[namedRange] = {};
      
      for (var key in empTemplateHeaderCells[namedRange]) {
        var cellInfo = empTemplateHeaderCells[namedRange][key];
        var row = cellInfo.row;
        var column = cellInfo.column;

        // Check if the cell has been edited (global edited color value)
        if (empSheet.getRange(row, column).getBackground() === editedColour) {
          var value = empSheet.getRange(row, column).getValue();

          // Add the key-value pair to the savingDictionary
          savingDictionary[namedRange][key] = value;
          editedCellsExist = true; // Set the flag to true since at least one cell was edited
        }
      }
    }

    // Create the message
    if (editedCellsExist) {
      var message = "WARNING: You are about to save and overwrite the following person's details:\n\n";
      message += selectedEmployee["First Name"] + " " + selectedEmployee["Surname"] + " - " + selectedEmployee["Department"] + "\n\n";
      message += "Details that are going to be saved:\n\n";

      // Loop through the savingDictionary and add the named range groupings and header-value pairs to the message
      for (var namedRange in savingDictionary) {
        if (Object.keys(savingDictionary[namedRange]).length > 0) { // Check if this namedRange has edited values
          message += "--- " + namedRange + " ---\n";

          for (var header in savingDictionary[namedRange]) {
            var value = savingDictionary[namedRange][header];
            message += header + ": " + value + "\n";
          }

          message += "\n"; // Add an extra newline between different named ranges
        }
      }

      // Display the warning message
      SpreadsheetApp.getUi().alert(message);
    } else {
      // Display an alert if no details to be saved
      SpreadsheetApp.getUi().alert("No details to be saved.");
    }
  }

  Logger.log("Finished createSavingDictionary");
}

/** **************************************************************************************************** */

function checkIfCellIsDropdown(cellValue, row,column)  {

  //Logger.log ('Starting to check if this cell has a dropdown menu')

  var cell = empSheet.getRange(row,column)
  var dataValidation = cell.getDataValidation()

  if (dataValidation && dataValidation.getCriteriaType() === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
    var dropdownValues = dataValidation.getCriteriaValues();
    Logger.log('cellValue: ' + cellValue + ' is a dropdown menu with values: ' + dropdownValues.join(", "));
    return dropdownValues;
  } else {
    return false;
  }

}

/** **************************************************************************************************** */

function employeeCalc(e, value, row, column){

  Logger.log('Starting employeeCalc: value = ' + value +' Row = ' +row + ' Column = ' + column)

  var userProperties = PropertiesService.getUserProperties();
  var autoDetails = userProperties.getProperty('autoDetails');

  if (autoDetails === 'ON'){
    Logger.log('AutoDetails is ON')
    // Get all the named ranges in the spreadsheet
    var namedRanges = ss.getNamedRanges();
    //filter namedRanges to take out the whole page range aka 'employeeDetails'
    var filteredNamedRanges = namedRanges.filter(function(namedRange) {
    return namedRange.getName() !== 'employeeDetails';
    });

    for (var i = 0; i <filteredNamedRanges.length; i++){
      var namedRange  = filteredNamedRanges[i];
      var range = namedRange.getRange()
      if (range.getRow() <= row && row <= range.getLastRow() && range.getColumn() <= column && column <= range.getLastColumn()) {
        // Check if the named range is 'WeeklyDailyFee' or 'hoursSection' and run the appropriate function
        if (namedRange.getName() === 'WeeklyDailyFee') {
          // Run the function for 'WeeklyDailyFee'
          Logger.log('FEECALCULATIONS GO!')
          feeCalculations(e)
          
        } else if (namedRange.getName() === 'hoursSection') {
          // Run the function for 'hoursSection'
          Logger.log('WORKINGDATES GO!')
          workingDates(e)

        }else{
         Logger.log ('NO CALC MATCHES') 
        }
      }
    }

  }else if (autoDetails === 'OFF'){
    Logger.log('AutoDetails is OFF')
    return
    
  }

  Logger.log('Finishing employeeCalc: value = ' + value +' Row = ' +row + ' Column = ' + column)

}

/** **************************************************************************************************** */

function feeCalculations(e) {

  Logger.log("Starting feeCalculations");

  var userProperties = PropertiesService.getUserProperties();

  var autoDetails = userProperties.getProperty('autoDetails');
  if (autoDetails === 'OFF') {
    return;
  }

  var empTemplateHeaderCells = JSON.parse(userProperties.getProperty('empTemplateHeaderCells'));

  var selectedEmployeeJSON = userProperties.getProperty(selectedEmployeeKey);
  var selectedEmployee = JSON.parse(selectedEmployeeJSON);

  var activeCell = e.range;
  var activeCellRow = activeCell.getRow();
  var activeCellCol = activeCell.getColumn();

  var activeKey = null;

  // Identify the active key
  for (var key in empTemplateHeaderCells) {
    var cellInfo = empTemplateHeaderCells[key];
    if (cellInfo.row === activeCellRow && cellInfo.column === activeCellCol) {
      activeKey = key;
      break;
    }
  }

  // If active key is not one of the keys of interest, return early
  if (!['Daily Rate inc Hol Pay', 'Weekly Rate inc Hol Pay', 'Holiday Pay Percentage'].includes(activeKey)) {
    return;
  }

  // Retrieve cells
  var holidayPayPercentageInfo = empTemplateHeaderCells['Holiday Pay Percentage'];
  var holidayPayPercentageCell = empSheet.getRange(holidayPayPercentageInfo.row, holidayPayPercentageInfo.column);
  var holidayPayPercentageCellValue = 1 + Number(holidayPayPercentageCell.getValue());
  var dailyInclHolPayCellInfo = empTemplateHeaderCells['Daily Rate inc Hol Pay'];
  var dailyInclHolPayCell = empSheet.getRange(dailyInclHolPayCellInfo.row, dailyInclHolPayCellInfo.column);
  var weeklyInclHolPayCellInfo = empTemplateHeaderCells['Weekly Rate inc Hol Pay'];
  var weeklyInclHolPayCell = empSheet.getRange(weeklyInclHolPayCellInfo.row, weeklyInclHolPayCellInfo.column);
  var weeklyExclHolPayCellInfo = empTemplateHeaderCells['Weekly Rate exc Hol Pay'];
  var weeklyExclHolPayCell = empSheet.getRange(weeklyExclHolPayCellInfo.row, weeklyExclHolPayCellInfo.column);
  var weeklyHolPayValueCellInfo = empTemplateHeaderCells['Weekly Hol Pay Value'];
  var weeklyHolPayValueCell = empSheet.getRange(weeklyHolPayValueCellInfo.row, weeklyHolPayValueCellInfo.column);
  var dailyExcHolPayCellInfo = empTemplateHeaderCells['Daily Rate exc Hol Pay'];
  var dailyExclHolPayCell = empSheet.getRange(dailyExcHolPayCellInfo.row, dailyExcHolPayCellInfo.column);
  var dailyHolPayValueCellInfo = empTemplateHeaderCells['Daily Hol Pay Value'];
  var dailyHolPayValueCell = empSheet.getRange(dailyHolPayValueCellInfo.row, dailyHolPayValueCellInfo.column);

  // Calculate based on active key
  switch (activeKey) {
    case 'Daily Rate inc Hol Pay':
      var weeklyHolPayFormula = '=round((' + dailyInclHolPayCell.getA1Notation() + '*5), 2)';
      weeklyInclHolPayCell.setFormula(weeklyHolPayFormula);
      backgroundColour(weeklyInclHolPayCell, empTemplateHeaderCells, selectedEmployee);
      break;
    case 'Weekly Rate inc Hol Pay':
      var dailyHolPayFormula = '=round((' + weeklyInclHolPayCell.getA1Notation() + '/5), 2)';
      dailyInclHolPayCell.setFormula(dailyHolPayFormula);
      backgroundColour(dailyInclHolPayCell, empTemplateHeaderCells, selectedEmployee);
      break;
  }

  // Calculate weekly and daily rates excluding holiday pay
  var weeklyExclHolPayFormula = '=round((' + weeklyInclHolPayCell.getA1Notation() + '/' + holidayPayPercentageCellValue + '), 2)';
  var dailyExclHolPayFormula = '=round((' + dailyInclHolPayCell.getA1Notation() + '/' + holidayPayPercentageCellValue + '), 2)';
  weeklyExclHolPayCell.setFormula(weeklyExclHolPayFormula);
  dailyExclHolPayCell.setFormula(dailyExclHolPayFormula);

  backgroundColour(weeklyExclHolPayCell, empTemplateHeaderCells, selectedEmployee);
  backgroundColour(dailyExclHolPayCell, empTemplateHeaderCells, selectedEmployee);

  // Calculate weekly and daily holiday pay value
  var weeklyHolPayValueFormula = '=round((' + weeklyInclHolPayCell.getA1Notation() + '-' + weeklyExclHolPayCell.getA1Notation() +'),2)';
  var dailyHolPayValueFormula = '=round((' + dailyInclHolPayCell.getA1Notation() + '-' + dailyExclHolPayCell.getA1Notation() +'),2)';
  weeklyHolPayValueCell.setFormula(weeklyHolPayValueFormula);
  dailyHolPayValueCell.setFormula(dailyHolPayValueFormula);

  backgroundColour(weeklyHolPayValueCell, empTemplateHeaderCells, selectedEmployee);
  backgroundColour(dailyHolPayValueCell, empTemplateHeaderCells, selectedEmployee);

  Logger.log("Finished feeCalculations");
}

/** **************************************************************************************************** */

function workingDates(e) {

  Logger.log("Starting workingDates");

  var userProperties = PropertiesService.getUserProperties();
  var autoDetails = userProperties.getProperty('autoDetails');

  if (autoDetails === 'OFF') {
    return;
  }

  var empTemplateHeaderCells = JSON.parse(userProperties.getProperty('empTemplateHeaderCells'));

  var selectedEmployeeJSON = userProperties.getProperty(selectedEmployeeKey);
  var selectedEmployee = JSON.parse(selectedEmployeeJSON);

  var activeCell = e.range;
  var activeCellRow = activeCell.getRow();
  var activeCellCol = activeCell.getColumn();

  var activeKey = null;
  var activeCategory = null;

  for (var category in empTemplateHeaderCells) {
  for (var key in empTemplateHeaderCells[category]) {
    var cellInfo = empTemplateHeaderCells[category][key];
    if (cellInfo.row === activeCellRow && cellInfo.column === activeCellCol) {
      activeKey = key;
      activeCategory = category;
      Logger.log('activeCategory = ' + activeCategory + ' and activeKey = ' + activeKey)
      break;
    }
  }
  if (activeKey) break;  // exit outer loop if activeKey is found
  }

  /*
  // Identify the active key
  for (var key in empTemplateHeaderCells) {
    var cellInfo = empTemplateHeaderCells[key];
    if (cellInfo.row === activeCellRow && cellInfo.column === activeCellCol) {
      activeKey = key;
      break;
    }
  }
  */

  // If active key is not one of the keys of interest, return early
  if (!['Start Date', 'End Date', 'Weeks'].includes(activeKey)) {
    Logger.log('No working Dates selected. Finished workingDates' )
    return;
  }

  // Retrieve cells
  var startDateInfo = empTemplateHeaderCells[activeCategory]['Start Date'];
  var startDateCell = empSheet.getRange(startDateInfo.row, startDateInfo.column);
  var startDateValue = new Date(startDateCell.getValue());
  var endDateInfo = empTemplateHeaderCells[activeCategory]['End Date'];
  var endDateCell = empSheet.getRange(endDateInfo.row, endDateInfo.column);
  var endDateValue = new Date(endDateCell.getValue());
  var weeksInfo = empTemplateHeaderCells[activeCategory]['Weeks'];
  var weeksCell = empSheet.getRange(weeksInfo.row, weeksInfo.column);
  var weeksValue = weeksCell.getValue();

  // Calculate based on active key
  switch (activeKey) {
    case 'Start Date':
      if (startDateValue && endDateValue) {
        var timeDiff = endDateValue - startDateValue + (1000 * 3600 * 24); // Add one day
        var diffWeeks = Math.floor(timeDiff / (1000 * 3600 * 24 * 7));
        var diffDays = Math.ceil((timeDiff % (1000 * 3600 * 24 * 7)) / (1000 * 3600 * 24));
        var weekString = diffWeeks + ' weeks ' + diffDays + ' days';
        weeksCell.setValue(weekString);
        backgroundColour(weeksCell, empTemplateHeaderCells, selectedEmployee);
      } else {
        // If the start date is removed, clear the end date and weeks values
        endDateCell.setValue('');
        weeksCell.setValue('');
      }
    break;
    case 'Weeks':
      if (weeksValue === '' || isNaN(weeksValue)) {
        // If Weeks is cleared, clear End Date too
        endDateCell.clearContent();
        backgroundColour(endDateCell, empTemplateHeaderCells, selectedEmployee);
      } else {
        // Calculate new end date
        var endDate = new Date(startDateValue);
        endDate.setHours(0, 0, 0, 0); // Set time to midnight
        endDate.setDate(endDate.getDate() + weeksValue * 7 - 1); // Subtract one day since start date is inclusive

        var totalDays = weeksValue * 7;
        var exactWeeks = Math.floor(totalDays / 7);
        var remainingDays = totalDays % 7;

        // Set the formatted weeks and days value
        weeksCell.setValue(exactWeeks + ' weeks ' + remainingDays + ' days');

        endDateCell.setValue(endDate);
        backgroundColour(endDateCell, empTemplateHeaderCells, selectedEmployee);
      }
    break;
    case 'End Date':
      if (endDateValue) {
        var timeDiff = endDateValue - startDateValue + (1000 * 3600 * 24); // Add one day
        var diffWeeks = Math.floor(timeDiff / (1000 * 3600 * 24 * 7));
        var diffDays = Math.ceil((timeDiff % (1000 * 3600 * 24 * 7)) / (1000 * 3600 * 24));
        var weekString = diffWeeks + ' weeks ' + diffDays + ' days';
        weeksCell.setValue(weekString);
        backgroundColour(weeksCell, empTemplateHeaderCells, selectedEmployee);
      } else {
        // If the end date is removed, clear the weeks value
        weeksCell.setValue('');
      }
    break;
  }
  Logger.log("Finished workingDates");
}

/** **************************************************************************************************** */

function showSavedEmployeeDetails() {
  Logger.log("Starting showSavedEmployeeDetails");

  // Retrieve from user properties
  var userProperties = PropertiesService.getUserProperties();

  var selectedEmployeeJSON = userProperties.getProperty(selectedEmployeeKey);
  var selectedEmployee = JSON.parse(selectedEmployeeJSON);

  var empTemplateHeaderCellsJSON = userProperties.getProperty('empTemplateHeaderCells'); // Assuming this is the key you're using to save it in user properties
  var empTemplateHeaderCells = JSON.parse(empTemplateHeaderCellsJSON);

  var message;
  if (selectedEmployee !== '') {
    message = "The following Employee Details have been loaded:\n\n";
  } else {
    message = "Either no Employee Details are loaded or you are creating a new Employee manually:\n\n";
  }

  // Iterate over the named ranges
  for (var namedRange in empTemplateHeaderCells) {
    message += "---   " + namedRange + "   ---\n";
    
    for (var key in empTemplateHeaderCells[namedRange]) {
      if (key.toLowerCase().includes('date') && selectedEmployee[key] !== '') {
        // Format the date value
        var formattedDate = Utilities.formatDate(new Date(selectedEmployee[key]), "GMT", "dd/MM/yyyy");
        message += key + ":       " + formattedDate + "\n";
      } else {
        message += key + ":       " + (selectedEmployee[key] || '') + "\n";
      }
    }
    
    message += "\n"; // Add an extra newline between different named ranges
  }

  SpreadsheetApp.getUi().alert(message);

  Logger.log("Finished showSavedEmployeeDetails");
}



/** **************************************************************************************************** */
