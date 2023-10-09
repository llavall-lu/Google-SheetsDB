/** ****************************************************************************************************

/** this is original process of emploee details cells.  This takes one cell at at time, its row/column and puts it in a dictionary.
 
function processEmployeeDetailsCells() { 

  Logger.log("Starting processEmployeeDetailsCells");

  var empDataRange = empSheet.getDataRange()
  var empData = empDataRange.getValues()
  var formDataRange = formSheet.getDataRange();
  var formData = formDataRange.getValues();
  var headers = formData[0]
  var lastRow = empSheet.getLastRow()
  var lastColumn = empSheet.getMaxColumns()
  var empHeaderCells = {}

  for (var row = 0; row < lastRow; row++){
    for (var col = 0; col < lastColumn; col++) {
      var cellValue = empData[row][col];
      if (headers.includes(cellValue)) {
        if(col+1 < lastColumn) { // Check if there's a cell to the right
          // Store the row and column index of the cell one to the right (adding 1 because rows and columns are 0-based)
          empHeaderCells[cellValue] = {row: row + 1, column: col + 2};
        }
      }
    }
  }

  // Store the empHeaderCells dictionary in user properties
  var empHeaderCellsJSON = JSON.stringify(empHeaderCells);
  var userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty('empHeaderCells', empHeaderCellsJSON);
  
  Logger.log('EmpHeaderCells = ' + JSON.stringify(empHeaderCells, null, 2))

  Logger.log("Finished processEmployeeDetailsCells");

}

/** ****************************************************************************************************
 * 
 * This is the original populateEmployeeDetails.  This takes one cell at at time

function populateEmployeeDetails(value) {

  Logger.log("Starting populateEmployeeDetails: " + value);

  Logger.log('value = ' +value)

  // Retrieve and parse the formResponses and filmDetails dictionaries from user properties
  var userProperties = PropertiesService.getUserProperties();
  var formResponses = JSON.parse(userProperties.getProperty('formResponses'));
  var filmDetails = JSON.parse(userProperties.getProperty('filmDetails'));

  // Retrieve and parse the empHeaderCells dictionary from user properties
  var empHeaderCells = JSON.parse(userProperties.getProperty('empHeaderCells'));

  // Find the formResponse and filmDetails objects where the objectName matches the selected value
  var matchingResponse = formResponses.find(response => response.objectName === value);
  var filmMatchingResponse = filmDetails.find(response => response.objectName === "Film Details");

  // Initialize selectedEmployee object
  var selectedEmployee = {};

  // Function to handle processing of response
  function processResponse(response, storeToSelectedEmployee = true) {
    if (response) {
      // Go through each key-value pair in the response
      for (var key in response.response) {
        // Check if the key exists in the empHeaderCells dictionary
        if (empHeaderCells.hasOwnProperty(key)) {
          // Get the cell details
          var cellDetails = empHeaderCells[key];
          // Get the cell
          var cell = empSheet.getRange(cellDetails.row, cellDetails.column);

          // Check if the key includes 'date'
          if (key.toLowerCase().includes('date') && response.response[key] !== '') {
            // Parse the date from the response, format it, and set the cell value to the formatted date
            var date = new Date(response.response[key]);
            var formattedDate = Utilities.formatDate(date, "GMT", "dd/MM/yyyy");
            cell.setValue(formattedDate);
          } else {
            // Set the cell value to the value from the response
            cell.setValue(response.response[key]);
          }
          // Add value to selectedEmployee if storeToSelectedEmployee is true
          if (storeToSelectedEmployee) {
            selectedEmployee[key] = response.response[key];
          }
        }
      }
    }
  }

  processResponse(matchingResponse);
  processResponse(filmMatchingResponse, false);

  /*
  //If any sections need to be loaded.
  if (selectedEmployee['Fee Code'] || selectedEmployee['Employment Status']){

      // Check if the 'Fee Code' is populated in the selectedEmployee dictionary
    if (selectedEmployee['Fee Code'] && selectedEmployee['Fee Code'] !== '') {
        feeSection(selectedEmployee['Fee Code']);  ///BUILD FEESECTION - HIDING/DISPLAYING ETC
    }

    // Check if the 'Employment Status' is populated in the selectedEmployee dictionary
    if (selectedEmployee['Employment Status'] && selectedEmployee['Employment Status'] !== '') {
        employmentSection(selectedEmployee['Employment Status']); ///BUILD EMPLOYMENTSECTION - HIDING/DISPLAYING ETC
    }

  }


  // Now that all data has been loaded and stored, set the cell background colors
  for (var key in selectedEmployee) {
    if (empHeaderCells.hasOwnProperty(key)) {
      var cellDetails = empHeaderCells[key];
      var cell = empSheet.getRange(cellDetails.row, cellDetails.column);
      backgroundColour(cell, empHeaderCells, selectedEmployee);
    }
  }

  // Store the selectedEmployee dictionary in user properties
  userProperties.setProperty('selectedEmployee', JSON.stringify(selectedEmployee));

  Logger.log("Finished populateEmployeeDetails: " + value);
}

/** ****************************************************************************************************
 * 
 * This is the original backgroundCoours tath uses empHeaderCells
// Function to update the background color of a cell based on its content and the content of the selectedEmployee dictionary
function backgroundColour(cell, empHeaderCells, selectedEmployee) {

  //Logger.log("Starting backgroundColour: cell: " + JSON.stringify(cell) + ', empHeaderCells: ' + JSON.stringify(empHeaderCells) + ', selectedEmployee: ' + JSON.stringify(selectedEmployee));

  // Obtain the row and column of the cell, and its current value
  var cellRow = cell.getRow();
  var cellCol = cell.getColumn();
  var cellValue = cell.getValue();

  // Get the key in empHeaderCells where the value matches the cell's row and column
  var cellKey = Object.keys(empHeaderCells).find(
    key => JSON.stringify(empHeaderCells[key]) === JSON.stringify({row: cellRow, column: cellCol})
  );

    // If the cell is not in empHeaderCells, then exit the function
  if (!cellKey) {
    return;
  }

  // If the key contains 'date', then handle it as a date
  if (cellKey && cellKey.toLowerCase().includes('date')) {
    // If the selectedEmployee value isn't blank, convert it to the same date format as cellValue
    if (selectedEmployee[cellKey] !== '' && selectedEmployee[cellKey] !== null) {
      selectedEmployee[cellKey] = Utilities.formatDate(new Date(selectedEmployee[cellKey]), "GMT", "dd/MM/yyyy");
      // Log the values
      //Logger.log("Cell value: " + cellValue);
      //Logger.log("Cell location: Row " + cellRow + ", Column " + cellCol);
      //Logger.log("Selected Employee value: " + selectedEmployee[cellKey]);
    }
  }

  // Handle the case where the cell is empty
  if (cellValue === '' || cellValue === null) {
    // If the corresponding value in selectedEmployee is also empty, color the cell as an empty cell
    if (selectedEmployee[cellKey] === '' || selectedEmployee[cellKey] === null) {
      cell.setBackground(emptyColour); // Empty cell color
    } 
    // If the corresponding value in selectedEmployee is not empty, color the cell as an edited cell
    else {
      cell.setBackground(editedColour); // Edited cell color
    }
  } 
  // Handle the case where the cell is not empty
  else {

    //covert both selected and saved to lowercase so it doesn't recognise 'abc' and 'ABC' as different.
    var cellValueLowerCase = String(cellValue).toLowerCase();
    var selectedValueLowerCase = String(selectedEmployee[cellKey]).toLowerCase();

     // If the cell's value matches the corresponding value in selectedEmployee, color the cell as a loaded cell
    if (cellKey && selectedValueLowerCase === cellValueLowerCase) {
      cell.setBackground(loadedColour); // Loaded cell color
    } 
    // If the cell's value does not match the corresponding value in selectedEmployee, color the cell as an edited cell
    else {
      cell.setBackground(editedColour); // Edited cell color
    }
  }
  //Logger.log("Finished backgroundColour: cell: " + JSON.stringify(cell) + ', empHeaderCells: ' + JSON.stringify(empHeaderCells) + ', selectedEmployee: ' + JSON.stringify(selectedEmployee));
}

/** ****************************************************************************************************

/** original processEmployeeTemplateDetailsCells
function processEmployeeTemplateDetailsCells() {
  Logger.log("Starting processEmployeeTemplateDetailsCells");

  var formDataRange = formSheet.getDataRange();
  var formData = formDataRange.getValues();
  var headers = formData[0];

  var namedRanges = ss.getNamedRanges(); // Gets all named ranges
  var empTemplateHeadersCells = {};

  namedRanges.forEach(function(namedRange) {
    var rangeName = namedRange.getName();
    var range = namedRange.getRange();
    var rangeData = range.getValues();

    empTemplateHeadersCells[rangeName] = {}; // Initializes dictionary for the named range

    for (var i = 0; i < rangeData.length; i++) {
      for (var j = 0; j < rangeData[i].length; j++) {
        var cellValue = rangeData[i][j];
        if (headers.includes(cellValue)) {
          empTemplateHeadersCells[rangeName][cellValue] = {row: row + 1, column: col + 2};
        }
      }
    }
  });

  // Store the empTemplateHeadersCells dictionary in user properties
  var empTemplateHeadersCellsJSON = JSON.stringify(empTemplateHeadersCells);
  var userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty('empTemplateHeadersCells', empTemplateHeadersCellsJSON);
  
  Logger.log('empTemplateHeadersCells = ' + JSON.stringify(empTemplateHeadersCells, null, 2));

  Logger.log("Finished processEmployeeTemplateDetailsCells");
}

/** ****************************************************************************************************

/** original clearEmployeeDetails
function clearEmployeeDetails(empTemplateHeaderCells, selectedEmployeeKey) {

  //Logger.log("Starting clearEmployeeDetails: empTemplateHeaderCells: " + JSON.stringify(empTemplateHeaderCells) + ', selectedEmployeeKey: '+ JSON.stringify(selectedEmployeeKey));

  var userProperties = PropertiesService.getUserProperties();
  
  // Clear the cells if '<Create New Employee>' is selected
  for (var key in empTemplateHeaderCells) {
    var cellLocation = empTemplateHeaderCells[key];
    var cell = empSheet.getRange(cellLocation.row, cellLocation.column);
    cell.setValue("");
  }

  // Retrieve the existing selectedEmployee dictionary and clear its values
  var currentSelectedEmployee = JSON.parse(userProperties.getProperty(selectedEmployeeKey));
  for (var key in currentSelectedEmployee) {
    currentSelectedEmployee[key] = "";
  }

  // Store the updated selectedEmployee dictionary in user properties
  userProperties.setProperty(selectedEmployeeKey, JSON.stringify(currentSelectedEmployee));

  
  var dropdownCellPos = empSheet.getRange(dropdownCellValue)
  var dropdownValue = dropdownCellPos.getValue()

  //Logger.log ('dropdownValue = ' + JSON.stringify(dropdownValue))

  if (dropdownValue != ('<Create New Employee>')) {

    // Call the backgroundColour function on the edited cells
    for (var key in empTemplateHeaderCells) {
      var cellLocation = empTemplateHeaderCells[key];
      var cell = empSheet.getRange(cellLocation.row, cellLocation.column);
      backgroundColour(cell, empTemplateHeaderCells, currentSelectedEmployee);
    }
  }else{
    Logger.log ('No need to set Background')
  }

  //Logger.log("Finished clearEmployeeDetails: empTemplateHeaderCells: " + JSON.stringify(empTemplateHeaderCells) + ', selectedEmployeeKey: '+ selectedEmployeeKey); 
}


/** ****************************************************************************************************

function showSavedEmployeeDetails(){

  Logger.log("Starting showSavedEmployeeDetails");

  // Retrieve and parse the selectedEmployee dictionary from user properties
  var userProperties = PropertiesService.getUserProperties();
  var selectedEmployeeJSON = userProperties.getProperty(selectedEmployeeKey);
  var selectedEmployee = JSON.parse(selectedEmployeeJSON);
  if (selectedEmployee !== ''){
    var message = "The following Employee Details have been loaded:\n\n";
  }else{
    var message = "Either no Employee Details are loaded or you are creating a new Employee manually:\n\n";
  }
  
    for (var key in selectedEmployee) {
      if (key.toLowerCase().includes('date')) {
        if(selectedEmployee[key] !== ''){
          // Format the date value
          var formattedDate = Utilities.formatDate(new Date(selectedEmployee[key]), "GMT", "dd/MM/yyyy");
          message += key + ": " + formattedDate + "\n";
        } else {
          message += key + ":  \n";
        }
      } else {
        message += key + ": " + selectedEmployee[key] + "\n";
      }
    }
    SpreadsheetApp.getUi().alert(message);

  Logger.log("Finished showSavedEmployeeDetails");
}
/** original createSavingDictionary
 * 
 * 
  function createSavingDictionary() {

  Logger.log("Starting createSavingDictionary");

  var userProperties = PropertiesService.getUserProperties();
  // Retrieve and parse the empHeaderCells dictionary from user properties
  var empHeaderCellsJSON = userProperties.getProperty('empHeaderCells');
  var empHeaderCells = JSON.parse(empHeaderCellsJSON);

  // Retrieve and parse the selectedEmployee dictionary from user properties
  var selectedEmployeeJSON = userProperties.getProperty(selectedEmployeeKey);
  var selectedEmployee = JSON.parse(selectedEmployeeJSON);

  if (empHeaderCells) {
    // Clear the previous content of the savingDictionary
    var savingDictionary = {};

    var editedCellsExist = false; // Flag to track if any cells were edited

    // Loop through the empHeaderCells dictionary
    for (var key in empHeaderCells) {
      if (empHeaderCells.hasOwnProperty(key)) {
        var cellInfo = empHeaderCells[key];
        var row = cellInfo.row;
        var column = cellInfo.column;

        // Check if the cell has been edited (global editec color value)
        if (empSheet.getRange(row, column).getBackground() === editedColour) {
          var value = empSheet.getRange(row, column).getValue();

          // Add the key-value pair to the savingDictionary
          savingDictionary[key] = value;

          editedCellsExist = true; // Set the flag to true since at least one cell was edited
        }
      }
    }

    // Check if any cells were edited
    if (editedCellsExist) {
      var message = "WARNING: You are about to save and overwrite the following person's details:\n\n";
      message += selectedEmployee["First Name"] + " " + selectedEmployee["Surname"] + " - " + selectedEmployee["Department"] + "\n\n";
      message += "Details that are going to be saved:\n\n";

      // Loop through the savingDictionary and add the header-value pairs to the message
      for (var header in savingDictionary) {
        if (savingDictionary.hasOwnProperty(header)) {
          var value = savingDictionary[header];
          message += header + ": " + value + "\n";
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


function processEmployeeTemplateDetailsCells() {
  Logger.log("Starting processEmployeeTemplateDetailsCells");

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
        if (headers.includes(cellValue)) {
          empTemplateHeaderCells[rangeName][cellValue] = {
            row: rangeRowStart + i,
            column: rangeColStart + j + 1 // Add 1 to get editable cell
          };
        }
      }
    }
  });

  // Store dictionaries in user properties
  var userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty('empTemplateHeaderCells', JSON.stringify(empTemplateHeaderCells));
  userProperties.setProperty('empTemplateNamedRanges', JSON.stringify(empTemplateNamedRanges));

  Logger.log('empTemplateHeaderCells = ' + JSON.stringify(empTemplateHeaderCells, null, 2));
  Logger.log('empTemplateNamedRanges = ' + JSON.stringify(empTemplateNamedRanges, null, 2));

  Logger.log("Finished processEmployeeTemplateDetailsCells");
}

function processEmployeeTemplateDetailsCells() {
  Logger.log("Starting processEmployeeTemplateDetailsCells");

  var formDataRange = formSheet.getDataRange();
  var formData = formDataRange.getValues();
  var headers = formData[0];

  var namedRanges = ss.getNamedRanges(); // Gets all named ranges
  //filter namedRanges to take out the whole page range aka 'employeeDetails'
  var filteredNamedRanges = namedRanges.filter(function(namedRange) {
  return namedRange.getName() !== 'employeeDetails';
  });
  var empTemplateHeaderCells = {};

  filteredNamedRanges.forEach(function(namedRange) {
    var rangeName = namedRange.getName();
    var range = namedRange.getRange();
    var rangeRowStart = range.getRow(); // Get the starting row of the range
    var rangeColStart = range.getColumn(); // Get the starting column of the range
    var rangeData = range.getValues();

    empTemplateHeaderCells[rangeName] = {}; // Initializes dictionary for the named range

    for (var i = 0; i < rangeData.length; i++) {
      for (var j = 0; j < rangeData[i].length; j++) {
        var cellValue = rangeData[i][j];
        if (headers.includes(cellValue)) {
          empTemplateHeaderCells[rangeName][cellValue] = {
            row: rangeRowStart + i, // Adjusting for 0-based loop index
            column: (rangeColStart + j) +1 // Adjusting for 0-based loop index and +1 to get the editable cell, not just the header.
          };
        }
      }
    }
  });

  // Store the empTemplateHeaderCells dictionary in user properties
  var empTemplateHeaderCellsJSON = JSON.stringify(empTemplateHeaderCells);
  var userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty('empTemplateHeaderCells', empTemplateHeaderCellsJSON);
  
  Logger.log('empTemplateHeaderCells = ' + JSON.stringify(empTemplateHeaderCells, null, 2));

  Logger.log("Finished processEmployeeTemplateDetailsCells");
}

function populateEmployeeDetails(value, row, column) {
  Logger.log("Starting populateEmployeeDetails: " + value);
  
  Logger.log('value = ' + value);

  // Retrieve and parse the user properties
  var userProperties = PropertiesService.getUserProperties();
  var objectNameDictionary = JSON.parse(userProperties.getProperty('nameDictionary'));
  var filmDetails = JSON.parse(userProperties.getProperty('filmDetails'));
  var empTemplateHeaderCells = JSON.parse(userProperties.getProperty('empTemplateHeaderCells'));

  //Logger.log ('objectNameDictionary = ' +JSON.stringify(objectNameDictionary))

  // Find the entry in objectNameDictionary where the objectName matches the selected value
  var matchingEntry = objectNameDictionary.find(entry => entry.objectName === value);
  //Logger.log ('matchingEntry = ' +JSON.stringify(matchingEntry))


  //Initialze the variable dictionaries for the 
  var selectedEmployeeDropdown = {}
  var selectedEmployeeDropdownRow = {}
  var selectedEmployeeDropdownColumn = {}

  // If there's no matching entry, exit the function early
  if (!matchingEntry) {
    Logger.log("No matching entry found for: " + value);
    return;
  }

  // Get the relevant row from formSheet using the rowNumber from matchingEntry
  var formResponseRow = formSheet.getRange(matchingEntry.rowNumber, 1, 1, formSheet.getLastColumn()).getValues();
  var formHeaders = formSheet.getRange(1, 1, 1, formSheet.getLastColumn()).getValues()[0];

  // Convert the row data and headers into an object for easier processing
  var matchingResponse = {
    objectName: value,
    response: {}
  };
  for (var i = 0; i < formHeaders.length; i++) {
    matchingResponse.response[formHeaders[i]] = formResponseRow[0][i];
  }

  // Initialize selectedEmployee object
  var selectedEmployee = {}

  // Function to handle processing of response
  function processResponse(response, storeToSelectedEmployee = true) {
    if (response) {
      for (var namedRange in empTemplateHeaderCells) {  // Iterate over named ranges
        for (var key in response.response) {  // Iterate over response keys
          // Check if the key exists in the empTemplateHeaderCells dictionary under the namedRange
          if (empTemplateHeaderCells[namedRange].hasOwnProperty(key)) {
            // Get the cell details
            var cellDetails = empTemplateHeaderCells[namedRange][key];
            
            // Get the cell
            var cell = empSheet.getRange(cellDetails.row, cellDetails.column);

            // Check if the key includes 'date'
            if (key.toLowerCase().includes('date') && response.response[key] !== '') {
              // Parse the date from the response, format it, and set the cell value to the formatted date
              var date = new Date(response.response[key]);
              var formattedDate = Utilities.formatDate(date, "GMT", "dd/MM/yyyy");
              cell.setValue(formattedDate);
            } else {
              // Set the cell value to the value from the response
              cell.setValue(response.response[key]);
            }
            // Add value to selectedEmployee if storeToSelectedEmployee is true
            if (storeToSelectedEmployee) {
              selectedEmployee[key] = response.response[key];
            }
          }
        }
      }
    }
  }

  processResponse(matchingResponse);
  processResponse(filmDetails.find(response => response.objectName === "Film Details"), false);

  // Set cell background colors
  for (var namedRange in empTemplateHeaderCells) {  
    for (var key in selectedEmployee) {  
      if (empTemplateHeaderCells[namedRange].hasOwnProperty(key)) {
        var cellDetails = empTemplateHeaderCells[namedRange][key];
        var cell = empSheet.getRange(cellDetails.row, cellDetails.column);
        backgroundColour(cell, empTemplateHeaderCells, selectedEmployee);
      }
    }
  }

  // Store the selectedEmployee dictionary in user properties
  userProperties.setProperty('selectedEmployee', JSON.stringify(selectedEmployee));

  Logger.log("Finished populateEmployeeDetails: " + value);
}

 LAST VERSION
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
    nameDictionary.push(objectName);
  }

  // Sort nameDictionary alphabetically
  nameDictionary.sort();

  // Add '<new employee>' option to the nameDictionary
  nameDictionary.unshift('<Create New Employee>');

  // Store the nameDictionary array in user properties
  var userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty('nameDictionary', JSON.stringify(nameDictionary));

  // Populate dropdown menu in Employee Details
  var dropdownRange = empSheet.getRange(dropdownRangeValue);
  dropdownRange.clearDataValidations();

  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(nameDictionary)
    .build();

  dropdownRange.setDataValidation(rule);

  //Logger.log('nameDictionary = ' + nameDictionary)

  Logger.log("Finished processFormResponses");
}

/**ORIGINAL PROCESSFORMSRESPONSES.
function processFormResponses(){

  Logger.log("Starting processFormResponses");

  var formDataRange = formSheet.getDataRange()
  var formData = formDataRange.getValues()
  var headers = formData[0];
  var formResponses = []
  var nameDictionary = []

  var firstNameIndex = headers.indexOf('First Name');
  var lastNameIndex = headers.indexOf('Surname');
  var deptIndex = headers.indexOf('Department');

  for (var i = 1; i < formData.length; i++) {
    var row = formData[i];
    var formRepsonse = {};


    // Combine First Name, Last Name, and Department to create the nameDictionary that populates the Dropdown Menu
    var firstName = row[firstNameIndex];
    var lastName = row[lastNameIndex];
    var deptName = row[deptIndex];
    var objectName = firstName + ' ' + lastName + ' ' + deptName;
    nameDictionary.push(objectName);

    for (var j = 0; j < headers.length; j++) {
      var header = headers[j];
      var value = row[j];
      formRepsonse[header] = value;
    }

    formResponses.push({ objectName: objectName, response: formRepsonse});
  }

  dropDownNames = nameDictionary

  // Store the formRepsonses array in user properties
  var userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty('formResponses', JSON.stringify(formResponses));

  

   // Add '<new employee>' option to the nameDictionary
  nameDictionary.unshift('<Create New Employee>');

  //Populate dropdown menu in Employee Details
  // Get the range where you want to create the dropdown menu
  var dropdownRange = empSheet.getRange(dropdownRangeValue);

  // Clear the existing data validation in the range
  dropdownRange.clearDataValidations();

  // Create a new data validation rule
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(nameDictionary)
    .build();

  // Apply the data validation rule to the range
  dropdownRange.setDataValidation(rule);

  Logger.log('Form Responses = ' + JSON.stringify(formResponses));
  
  Logger.log("Finished processFormResponses");
}

function populateEmployeeDetails(value, row, column) {
  Logger.log("Starting populateEmployeeDetails: " + value);
  
  Logger.log('value = ' + value);

  // Retrieve and parse the formResponses and filmDetails dictionaries from user properties
  var userProperties = PropertiesService.getUserProperties();
  var formResponses = JSON.parse(userProperties.getProperty('formResponses'));
  Logger.log('formResponses = ' + JSON.stringify(formResponses))
  var filmDetails = JSON.parse(userProperties.getProperty('filmDetails'));

  // Retrieve and parse the empTemplateHeaderCells dictionary from user properties
  var empTemplateHeaderCells = JSON.parse(userProperties.getProperty('empTemplateHeaderCells'));

  var objectNameDictionary = JSON.parse(userProperties.getProperty('nameDictionary'))

  // Find the formResponse and filmDetails objects where the objectName matches the selected value
  var matchingResponse = formResponses.find(response => response.objectName === value);
  var filmMatchingResponse = filmDetails.find(response => response.objectName === "Film Details");

  // Initialize selectedEmployee object
  var selectedEmployee = {};
  var selectedEmployeeDropdown = {}
  var selectedEmployeeDropdownRow = {}
  var selectedEmployeeDropdownColumn = {}

  // Function to handle processing of response
  function processResponse(response, storeToSelectedEmployee = true) {
    if (response) {
      for (var namedRange in empTemplateHeaderCells) {  // Iterate over named ranges
        for (var key in response.response) {  // Iterate over response keys
          // Check if the key exists in the empTemplateHeaderCells dictionary under the namedRange
          if (empTemplateHeaderCells[namedRange].hasOwnProperty(key)) {
            // Get the cell details
            var cellDetails = empTemplateHeaderCells[namedRange][key];
            
            // Get the cell
            var cell = empSheet.getRange(cellDetails.row, cellDetails.column);

            // Check if the key includes 'date'
            if (key.toLowerCase().includes('date') && response.response[key] !== '') {
              // Parse the date from the response, format it, and set the cell value to the formatted date
              var date = new Date(response.response[key]);
              var formattedDate = Utilities.formatDate(date, "GMT", "dd/MM/yyyy");
              cell.setValue(formattedDate);
            } else {
              // Set the cell value to the value from the response
              cell.setValue(response.response[key]);
            }
            // Add value to selectedEmployee if storeToSelectedEmployee is true
            if (storeToSelectedEmployee) {
              selectedEmployee[key] = response.response[key];
            }
          }
        }
      }
    }
  }

  processResponse(matchingResponse);
  processResponse(filmMatchingResponse, false);

  // Now that all data has been loaded and stored, set the cell background colors
  for (var namedRange in empTemplateHeaderCells) {  // Iterate over named ranges
    for (var key in selectedEmployee) {  // Iterate over selectedEmployee keys
      if (empTemplateHeaderCells[namedRange].hasOwnProperty(key)) {
        var cellDetails = empTemplateHeaderCells[namedRange][key];
        var cell = empSheet.getRange(cellDetails.row, cellDetails.column);
        backgroundColour(cell, empTemplateHeaderCells, selectedEmployee);
      }
    }
  }

  // Store the selectedEmployee dictionary in user properties
  userProperties.setProperty('selectedEmployee', JSON.stringify(selectedEmployee));

  Logger.log("Finished populateEmployeeDetails: " + value);
}

    WE MIGHT NEED THIS BELOW BUT IN A DIFFERENT FORMAT
    
    if (feeSectionRange) { // ensure the named range exists
    // Check if the edited cell's row is within the FeeSection range and the column is not 1
    if (range.getRow() >= feeSectionRange.getRow() && range.getRow() <= feeSectionRange.getLastRow() && range.getColumn() !== 1) {
      feeCalculations(e);
    }
  }

  if (hoursSectionRange) { // ensure the named range exists
    // Check if the edited cell's row is within the FeeSection range and the column is not 1
    if (range.getRow() >= hoursSectionRange.getRow() && range.getRow() <= hoursSectionRange.getLastRow() && range.getColumn() !== 1) {
      hoursCalculations(e);
    }  
  }
  
  WE MIGHT NEED THIS ABOVE BUT IN A DIFFERENT FORMAT

function populateEmployeeDetails(value, row, column, focusSection = null) {
  Logger.log("Starting populateEmployeeDetails: " + value);
  
  Logger.log('value = ' + value);

  // Retrieve and parse the user properties
  var userProperties = PropertiesService.getUserProperties();
  var objectNameDictionary = JSON.parse(userProperties.getProperty('nameDictionary'));
  var filmDetails = JSON.parse(userProperties.getProperty('filmDetails'));
  var empTemplateHeaderCells = JSON.parse(userProperties.getProperty('empTemplateHeaderCells'));
  var empTemplateNamedRanges = JSON.parse(userProperties.getProperty('empTemplateNamedRanges'));
  var empTemplateDropdownCells = JSON.parse(userProperties.getProperty('empTemplateDropdownCells'));
  var storedMatchingResponse = JSON.parse(userProperties.getProperty('matchingResponse'));

  //Logger.log ('objectNameDictionary = ' +JSON.stringify(objectNameDictionary))

  /** MATCHING RESPONSE START

  // Find the entry in objectNameDictionary where the objectName matches the selected value
  var matchingEntry = objectNameDictionary.find(entry => entry.objectName === value);
  
  Logger.log ('main Dropdown matchingEntry = ' +JSON.stringify(matchingEntry))

  // If there's no matching entry, exit the function early
  if (!matchingEntry) {
    Logger.log("No matching entry in the main ObjectName Dropdown found for: " + value);
    // Attempt to find the named range that matches the value
    
    //convert value to NamedRange format with no spaces - ONLY use this to find value in namedRange
    var valueNoSpace = value.replace(/\s+/g, '')

    Logger.log("valueNoSpace: " + valueNoSpace);

    var namedRangeMatch = empTemplateNamedRanges[valueNoSpace];
    
    // If a match is found
    if (namedRangeMatch) {
        Logger.log("YES!, namedRangeMatch matches entry found in NamedRange for: " + JSON.stringify(valueNoSpace));


        // Get the header name using the provided row and column
        var dropdownHeader = getHeaderNameByCell(row, column, value);
        Logger.log('dropdownHeader = ' + JSON.stringify(dropdownHeader))
        
        // If dropdownHeader is not identified, log a message and proceed further.
        if (!dropdownHeader) {
        Logger.log("Unable to determine the dropdownHeader for the given row/column. Finishing populateEmployeeDetails");
        
        return;
    }

      // Check if the value is different to the storedMatchingResponse value
      if (storedMatchingResponse.response && storedMatchingResponse.response[dropdownHeader] !== value) {
        // This is where you'd handle the case where the values are different
        // For now, just logging it
        Logger.log(value + "(value) is different from "+ storedMatchingResponse.response[dropdownHeader] + "(the stored matching response!)");

        loadingSection(value, row, column, dropdownHeader)

      }else{
        
        Logger.log('values are the same. Finishing populateEmployeeDetails')
        
        return;
      }

    } else {
      Logger.log('No namedRange matches this row or column. Please refresh the page and start again.')
      Logger.log("Finishing populateEmployeeDetails");

      return;
    }

  } else {

    // Get the relevant row from formSheet using the rowNumber from matchingEntry
    var formResponseRow = formSheet.getRange(matchingEntry.rowNumber, 1, 1, formSheet.getLastColumn()).getValues();
    var formHeaders = formSheet.getRange(1, 1, 1, formSheet.getLastColumn()).getValues()[0];

    // Convert the row data and headers into an object for easier processing
    var matchingResponse = {
      objectName: value,
      response: {}
    };
    for (var i = 0; i < formHeaders.length; i++) {
      matchingResponse.response[formHeaders[i]] = formResponseRow[0][i];
    }

    /** Store matchingResponse to the userProperties
    userProperties.setProperty('matchingResponse', JSON.stringify(matchingResponse));
    Logger.log ('matchingResponse = ' + JSON.stringify(matchingResponse))

    Logger.log ('empTemplateDropdownCells = ' + JSON.stringify(empTemplateDropdownCells))

    /** MATCHING RESPONSE END

    /** UNHIDE SECTION START

    // Iterate over the empTemplateDropdownCells headers
      for (var dropdownHeader in empTemplateDropdownCells) {
        // Check if the header exists in matchingResponse.response
        if (matchingResponse.response.hasOwnProperty(dropdownHeader)) {
          // Get the corresponding value from matchingResponse
          var dropdownValue = matchingResponse.response[dropdownHeader].replace(/\s+/g, '');

          Logger.log('dropdownValue = ' + dropdownValue)

          // Check if the value matches a namedRange in empTemplateNamedRanges
          if (empTemplateNamedRanges.hasOwnProperty(dropdownValue)) {
            // Unhide the namedRange using the rowStart and rowEnd
            var namedRangeInfo = empTemplateNamedRanges[dropdownValue];
            empSheet.showRows(namedRangeInfo.rowStart, namedRangeInfo.rowEnd - namedRangeInfo.rowStart + 1);
        }
      }
    }

    /** UNHIDE SECTION END

    /** VISIBLE NAMED RANGES START

    // Initialize an empty dictionary to store the visibility status of each named range
    var visibleNamedRanges = {};

    // Iterate through each named range to check if its first row is hidden
    for (var namedRange in empTemplateNamedRanges) {
      var firstRow = empTemplateNamedRanges[namedRange].rowStart;
      // Use isRowHiddenByUser(row) function to check if the row is hidden
      // The function returns 'true' if the row is hidden and 'false' otherwise
      var isRowHidden = empSheet.isRowHiddenByUser(firstRow);
  
      // If the row is NOT hidden, add the named range to the visibleNamedRanges dictionary
      if (!isRowHidden) {
        visibleNamedRanges[namedRange] = true;
      }
    }
    // At this point, visibleNamedRanges will contain the names of all unhidden named ranges
    Logger.log("Visible named ranges: " + JSON.stringify(visibleNamedRanges));

    /** VISIBLE NAMED RANGES END

    // Initialize selectedEmployee object
    var selectedEmployee = {}

    // Function to handle processing of response
    function processResponse(response, storeToSelectedEmployee = true) {
      if (response) {
        for (var namedRange in empTemplateHeaderCells) {  // Iterate over named ranges in empTemplateHeaderCells
      
          // Skip this named range if it's not in visibleNamedRanges
          if (!visibleNamedRanges.hasOwnProperty(namedRange)) {
            continue;
          }

          for (var key in response.response) {  // Iterate over response keys
            // Check if the key exists in the empTemplateHeaderCells dictionary under the namedRange
            if (empTemplateHeaderCells[namedRange].hasOwnProperty(key)) {
              // Get the cell details
              var cellDetails = empTemplateHeaderCells[namedRange][key];
          
              // Get the cell
              var cell = empSheet.getRange(cellDetails.row, cellDetails.column);

              // Check if the key includes 'date'
              if (key.toLowerCase().includes('date') && response.response[key] !== '') {
                // Parse the date from the response, format it, and set the cell value to the formatted date
                var date = new Date(response.response[key]);
                var formattedDate = Utilities.formatDate(date, "GMT", "dd/MM/yyyy");
                cell.setValue(formattedDate);
              } else {
                // Set the cell value to the value from the response
                cell.setValue(response.response[key]);
              }
              // Add value to selectedEmployee if storeToSelectedEmployee is true
              if (storeToSelectedEmployee) {
                selectedEmployee[key] = response.response[key];
              }
            }
          }
        }
      }
    }

    processResponse(matchingResponse);
    processResponse(filmDetails.find(response => response.objectName === "Film Details"), false);

    // Set cell background colors
    for (var namedRange in empTemplateHeaderCells) {
      for (var key in selectedEmployee) {  
        if (empTemplateHeaderCells[namedRange].hasOwnProperty(key)) {
          var cellDetails = empTemplateHeaderCells[namedRange][key];
          var cell = empSheet.getRange(cellDetails.row, cellDetails.column);
          backgroundColour(cell, empTemplateHeaderCells, selectedEmployee);
        }
      }
    }

    // Store the selectedEmployee dictionary in user properties
    userProperties.setProperty('selectedEmployee', JSON.stringify(selectedEmployee));

    Logger.log("Finished populateEmployeeDetails: " + value);
  }
}

/** ****************************************************************************************************

function loadingSection(value, row, column, dropdownHeader) {

  Logger.log("Starting loadingSection")

  Logger.log('Loading Section value = ' + value)
  Logger.log('Loading Section row = ' + row)
  Logger.log('Loading Section column = ' + column)
  Logger.log('Loading Section dropdownHeader = ' + dropdownHeader)

  // Retrieve the selectedEmployee dictionary from user properties
  var userProperties = PropertiesService.getUserProperties();
  var selectedEmployeeJSON = userProperties.getProperty(selectedEmployeeKey);
  var empTemplateHeaderCells = JSON.parse(userProperties.getProperty('empTemplateHeaderCells'));
  var selectedEmployee = JSON.parse(selectedEmployeeJSON);
  Logger.log('loadingSection savedSelectedEmployee = ' + JSON.stringify(selectedEmployee))
  var namedRangesObject = JSON.parse(userProperties.getProperty('empTemplateNamedRanges'));

  //covert the incoming value into the namedRange format.
  var valueNoSpace = value.replace(/\s+/g, '')
  Logger.log('valueNoSpace = ' + valueNoSpace)

  
  var namedRangeKeys = Object.keys(namedRangesObject);
  var matchingKey = namedRangeKeys.find(key => key.replace(/\s+/g, '') === valueNoSpace);
  var selectedNamedRange = namedRangesObject[matchingKey];
  Logger.log('loadingSection namedRangesObject = ' + JSON.stringify(namedRangesObject))
  Logger.log('loadingSection namedRangeKeys = ' + JSON.stringify(namedRangeKeys))
  Logger.log('loadingSection matchingKey = ' + matchingKey)
  Logger.log('loadingSection selectedNamedRange = ' + JSON.stringify(selectedNamedRange))


  // 1. Unhide the selected namedRange
  //var valueNoSpace = value.replace(/\s+/g, '')
  
  //var selectedNamedRange = namedRanges.find(namedRange => namedRange.getName() === valueNoSpace);

  //Logger.log('selectedNameRange = ' + JSON.stringify(selectedNamedRange))
  
  if (selectedNamedRange) {
    var startRow = selectedNamedRange.rowStart
    var endRow = selectedNamedRange.rowEnd
    var numRows = endRow - startRow + 1
    var targetRange = empSheet.getRange(startRow, endRow, numRows);
    Logger.log('targetRange = ' + targetRange)
    targetRange.activate();
    empSheet.unhideRow(targetRange);
    var row = targetRange.getRow();
    var column = targetRange.getColumn();
  }else{
    Logger.log('Not in the selectedNameRange')
    return
  }

  // 2. Populate data for the selected namedRange
  Logger.log('Section trying to populate ' + JSON.stringify(value))
  populateEmployeeDetails(value, row, column);

  // 3. Hide other namedRanges
  Object.keys(namedRangesObject).forEach(key => {
    if (key !== matchingKey && key !== 'mainDetails' && key !== 'hoursSection' && key !== 'employeeDetails') {
      let currentRangeObj = namedRangesObject[key];
      let startR = currentRangeObj.rowStart;
      let endR = currentRangeObj.rowEnd;
      let numR = endR - startR + 1;
      var targetRange = empSheet.getRange(startR, numR);
      empSheet.hideRow(hideRange);
    }
  });

}

function matchSelection(value,row,column) {

  Logger.log("Starting matchSelection");

  
  var userProperties = PropertiesService.getUserProperties();
  var empNameDictionary = JSON.parse(userProperties.getProperty('nameDictionary'));
  var filmDetails = JSON.parse(userProperties.getProperty('filmDetails'));
  var empTemplateHeaderCells = JSON.parse(userProperties.getProperty('empTemplateHeaderCells'));
  var empTemplateNamedRanges = JSON.parse(userProperties.getProperty('empTemplateNamedRanges'));
  var empTemplateDropdownCells = JSON.parse(userProperties.getProperty('empTemplateDropdownCells'));
  var storedMatchingResponse = JSON.parse(userProperties.getProperty('matchingResponse'));

  //Logger.log ('empNameNameDictionary = ' +JSON.stringify(empNameNameDictionary))

  // Find the entry in empNameNameDictionary where the empName matches the selected value
  var empNameMatchingEntry = empNameDictionary.find(entry => entry.objectName === value);
  
  Logger.log ('Employee Name Matching Entry = ' +JSON.stringify(empNameMatchingEntry))

  var empNameMatch = false

  /** MATCHING RESPONSE START
  // If there's a matching entry..
  if (empNameMatchingEntry) {

    empNameMatch = true

    // Get the relevant row from formSheet using the rowNumber from empNameMatchingEntry
    var formResponseRow = formSheet.getRange(empNameMatchingEntry.rowNumber, 1, 1, formSheet.getLastColumn()).getValues();
    var formHeaders = formSheet.getRange(1, 1, 1, formSheet.getLastColumn()).getValues()[0];

    // Convert the row data and headers into an object for easier processing
    var matchingResponse = {
      objectName: value,
      response: {}
    };
    //find and match any namedRanges that match the dropdown values within matchingResponse
    for (var i = 0; i < formHeaders.length; i++) {
      matchingResponse.response[formHeaders[i]] = formResponseRow[0][i];
    }

    /** Store matchingResponse to the userProperties
    userProperties.setProperty('matchingResponse', JSON.stringify(matchingResponse));

    Logger.log('empNameMatch =' + empNameMatch)
    Logger.log ('matchingResponse = ' + JSON.stringify(matchingResponse))

    Logger.log ('empTemplateDropdownCells = ' + JSON.stringify(empTemplateDropdownCells))

    /** MATCHING RESPONSE END

    unhideNamedRange(value,row,column,empNameMatch)

  } else {

    Logger.log("No matching entry in the main Employee Name Dropdown found for: " + value);


    /** Attempt to find the named range that matches the value
    //convert value to NamedRange format with no spaces - ONLY use this to find value in namedRange
    var valueNoSpace = value.replace(/\s+/g, '')

    Logger.log("valueNoSpace: " + valueNoSpace);

    var namedRangeMatch = empTemplateNamedRanges[valueNoSpace];
    
    // If a namedRangeMatch is found
    if (namedRangeMatch) {
        Logger.log("YES!, namedRangeMatch matches entry found in NamedRange for: " + JSON.stringify(valueNoSpace));

        // Get the header name using the provided row and column
        var dropdownHeader = getHeaderNameByCell(row, column, value);
        Logger.log('dropdownHeader = ' + JSON.stringify(dropdownHeader))
        
        // If dropdownHeader is not identified, log a message and proceed further.
        if (!dropdownHeader) {
        Logger.log("Unable to determine the dropdownHeader for the given row/column. Finishing populateEmployeeDetails");
        return;
    }

      // Check if the value is different to the storedMatchingResponse value
      if (storedMatchingResponse.response && storedMatchingResponse.response[dropdownHeader] !== value) {
        // This is where you'd handle the case where the values are different
        // For now, just logging it
        Logger.log(value + "(value) is different from "+ storedMatchingResponse.response[dropdownHeader] + "(the stored matching response!)");

        unhideNamedRange(value, row, column, empNameMatch, dropdownHeader)

      }else{
        
        Logger.log('values are the same. Finishing matchSelection')
        
        return;
      }

    } else {
      Logger.log('No namedRange matches this row or column. Please refresh the page and start again.')
      Logger.log("Finishing matchSelection");

      return;
    }
  }
    
}

function unhideNamedRange(value,row,column) {
  
  Logger.log("Starting unhideNamedRange");

  var userProperties = PropertiesService.getUserProperties();
  var objectNameDictionary = JSON.parse(userProperties.getProperty('nameDictionary'));
  var filmDetails = JSON.parse(userProperties.getProperty('filmDetails'));
  var empTemplateHeaderCells = JSON.parse(userProperties.getProperty('empTemplateHeaderCells'));
  var empTemplateNamedRanges = JSON.parse(userProperties.getProperty('empTemplateNamedRanges'));
  var empTemplateDropdownCells = JSON.parse(userProperties.getProperty('empTemplateDropdownCells'));
  var storedMatchingResponse = JSON.parse(userProperties.getProperty('matchingResponse'));

  //Logger.log ('objectNameDictionary = ' +JSON.stringify(objectNameDictionary))

  // Find the entry in objectNameDictionary where the objectName matches the selected value
  var matchingEntry = objectNameDictionary.find(entry => entry.objectName === value);
  
  Logger.log ('main Dropdown matchingEntry = ' +JSON.stringify(matchingEntry))

  /** MATCHING RESPONSE START
  // If there's a matching entry..
  if (matchingEntry) {

    // Get the relevant row from formSheet using the rowNumber from matchingEntry
    var formResponseRow = formSheet.getRange(matchingEntry.rowNumber, 1, 1, formSheet.getLastColumn()).getValues();
    var formHeaders = formSheet.getRange(1, 1, 1, formSheet.getLastColumn()).getValues()[0];

    // Convert the row data and headers into an object for easier processing
    var matchingResponse = {
      objectName: value,
      response: {}
    };
    //find and match any namedRanges that match the dropdown values within matchingResponse
    for (var i = 0; i < formHeaders.length; i++) {
      matchingResponse.response[formHeaders[i]] = formResponseRow[0][i];
    }

    /** Store matchingResponse to the userProperties
    userProperties.setProperty('matchingResponse', JSON.stringify(matchingResponse));
    Logger.log ('matchingResponse = ' + JSON.stringify(matchingResponse))

    Logger.log ('empTemplateDropdownCells = ' + JSON.stringify(empTemplateDropdownCells))

    /** MATCHING RESPONSE END

    /** UNHIDE SECTION START

    // Iterate over the empTemplateDropdownCells headers
      for (var dropdownHeader in empTemplateDropdownCells) {
        // Check if the header exists in matchingResponse.response
        if (matchingResponse.response.hasOwnProperty(dropdownHeader)) {
          // Get the corresponding value from matchingResponse
          var dropdownValue = matchingResponse.response[dropdownHeader].replace(/\s+/g, '');

          Logger.log('dropdownValue = ' + dropdownValue)

          // Check if the value matches a namedRange in empTemplateNamedRanges
          if (empTemplateNamedRanges.hasOwnProperty(dropdownValue)) {
            // Unhide the namedRange using the rowStart and rowEnd
            var namedRangeInfo = empTemplateNamedRanges[dropdownValue];
            empSheet.showRows(namedRangeInfo.rowStart, namedRangeInfo.rowEnd - namedRangeInfo.rowStart + 1);
        }
      }
    }

    /** UNHIDE SECTION END

    /** VISIBLE NAMED RANGES START

    // Initialize an empty dictionary to store the visibility status of each named range
    var visibleNamedRanges = {};

    // Iterate through each named range to check if its first row is hidden
    for (var namedRange in empTemplateNamedRanges) {
      var firstRow = empTemplateNamedRanges[namedRange].rowStart;
      // Use isRowHiddenByUser(row) function to check if the row is hidden
      // The function returns 'true' if the row is hidden and 'false' otherwise
      var isRowHidden = empSheet.isRowHiddenByUser(firstRow);
  
      // If the row is NOT hidden, add the named range to the visibleNamedRanges dictionary
      if (!isRowHidden) {
        visibleNamedRanges[namedRange] = true;
      }
    }
    // At this point, visibleNamedRanges will contain the names of all unhidden named ranges
    Logger.log("Visible named ranges: " + JSON.stringify(visibleNamedRanges));

    /** VISIBLE NAMED RANGES END

    // Initialize selectedEmployee object
    var selectedEmployee = {}

  } else {

    Logger.log("No matching entry in the main ObjectName Dropdown found for: " + value);


    /** Attempt to find the named range that matches the value 
    //convert value to NamedRange format with no spaces - ONLY use this to find value in namedRange
    var valueNoSpace = value.replace(/\s+/g, '')

    Logger.log("valueNoSpace: " + valueNoSpace);

    var namedRangeMatch = empTemplateNamedRanges[valueNoSpace];
    
    // If a namedRangeMatch is found
    if (namedRangeMatch) {
        Logger.log("YES!, namedRangeMatch matches entry found in NamedRange for: " + JSON.stringify(valueNoSpace));


        // Get the header name using the provided row and column
        var dropdownHeader = getHeaderNameByCell(row, column, value);
        Logger.log('dropdownHeader = ' + JSON.stringify(dropdownHeader))
        
        // If dropdownHeader is not identified, log a message and proceed further.
        if (!dropdownHeader) {
        Logger.log("Unable to determine the dropdownHeader for the given row/column. Finishing populateEmployeeDetails");
        
        return;
    }

      // Check if the value is different to the storedMatchingResponse value
      if (storedMatchingResponse.response && storedMatchingResponse.response[dropdownHeader] !== value) {
        // This is where you'd handle the case where the values are different
        // For now, just logging it
        Logger.log(value + "(value) is different from "+ storedMatchingResponse.response[dropdownHeader] + "(the stored matching response!)");

        loadingSection(value, row, column, dropdownHeader)

      }else{
        
        Logger.log('values are the same. Finishing populateEmployeeDetails')
        
        return;
      }

    } else {
      Logger.log('No namedRange matches this row or column. Please refresh the page and start again.')
      Logger.log("Finishing populateEmployeeDetails");

      return;
    }
  }
    
}

*/
