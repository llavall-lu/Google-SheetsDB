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
      rowNumber: i + 1 
    };

    nameDictionary.push(entry);
  }

  nameDictionary.sort(function(a, b) {
    return a.objectName.localeCompare(b.objectName);
  });

  var dropdownNames = nameDictionary.map(function(entry) {
    return entry.objectName;
  });
  dropdownNames.unshift('<Create New Employee>');

  var userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty('nameDictionary', JSON.stringify(nameDictionary));

  var dropdownRange = empSheet.getRange(dropdownRangeValue);
  dropdownRange.clearDataValidations();

  var rule = SpreadsheetApp.newDataValidation().requireValueInList(dropdownNames).build();

  dropdownRange.setDataValidation(rule);

  Logger.log('processFormResponses: nameDictionary = ' + JSON.stringify(nameDictionary))

  Logger.log("Finished processFormResponses");
}
function processFilmDetails() {

  Logger.log("Starting processFilmDetails");

  var filmDetailsRange = filmSheet.getDataRange(); 
  var filmDetailsValues = filmDetailsRange.getValues(); 

  var filmDetailsDictionary = {}; 

  for (var i = 0; i < filmDetailsValues.length; i++) {
    var header = filmDetailsValues[i][0]; 
    var value = filmDetailsValues[i][1]; 

    if (header !== "") {
      filmDetailsDictionary[header] = value; 
    }
  }

  var filmDetails = [{
    "objectName": "Film Details",
    "response": filmDetailsDictionary
  }];

  var userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty('filmDetails', JSON.stringify(filmDetails));

  Logger.log("Finished processFilmDetails");
}
function processEmployeeTemplateDetailsCells() {
  Logger.log("Starting processEmployeeTemplateDetailsCells");

  var formDataRange = formSheet.getDataRange();
  var formData = formDataRange.getValues();
  var headers = formData[0];
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var namedRanges = ss.getNamedRanges();

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

    empTemplateHeaderCells[rangeName] = {};

    empTemplateNamedRanges[rangeName] = {};
    var rangeRowEnd = rangeRowStart + range.getNumRows() - 1;
    empTemplateNamedRanges[rangeName]['rowStart'] = rangeRowStart;
    empTemplateNamedRanges[rangeName]['rowEnd'] = rangeRowEnd;

    for (var i = 0; i < rangeData.length; i++) {
      for (var j = 0; j < rangeData[i].length; j++) {
        var cellValue = rangeData[i][j];
        var row = rangeRowStart + i;
        var column = rangeColStart + j + 1; 

        if (headers.includes(cellValue)) {
          empTemplateHeaderCells[rangeName][cellValue] = {
            row: row,
            column: column
          };

          var dropdownValues = checkIfCellIsDropdown(cellValue, row, column);

          if (dropdownValues) {
            empTemplateDropdownCells[cellValue] = {
              dropdown: true,
              values: dropdownValues 
            };
          }
        }
      }
    }
  });

  var userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty('empTemplateHeaderCells', JSON.stringify(empTemplateHeaderCells));
  userProperties.setProperty('empTemplateNamedRanges', JSON.stringify(empTemplateNamedRanges));
  userProperties.setProperty('empTemplateDropdownCells', JSON.stringify(empTemplateDropdownCells));

  Logger.log('empTemplateHeaderCells = ' + JSON.stringify(empTemplateHeaderCells, null, 2));
  Logger.log('empTemplateNamedRanges = ' + JSON.stringify(empTemplateNamedRanges, null, 2));
  Logger.log('empTemplateDropdownCells = ' + JSON.stringify(empTemplateDropdownCells, null, 2));

  Logger.log("Finished processEmployeeTemplateDetailsCells");
}
  var selectedEmployee = {}
 function processResponse(response, storeToSelectedEmployee = true) {
      if (response) {
        for (var namedRange in empTemplateHeaderCells) { 

          if (!visibleNamedRanges.hasOwnProperty(namedRange)) {
            continue;
          }

          for (var key in response.response) { 

            if (empTemplateHeaderCells[namedRange].hasOwnProperty(key)) {

              var cellDetails = empTemplateHeaderCells[namedRange][key];

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
    }

    processResponse(matchingResponse);
    processResponse(filmDetails.find(response = >response.objectName === "Film Details"), false);

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

    Logger.log("Finished populateEmployeeDetails: " + value);
  }
}
function employeeOnEdit(e) {

  Logger.log("Starting employeeOnEdit");

  loadingIndicator();

  var range = e.range;
  var sheet = range.getSheet();
  //var sheetEndRow = empSheet.getLastRow();
  //var feeSectionRange = ss.getRangeByName('WeeklyDailyFee');
  //var hoursSectionRange = ss.getRangeByName('hoursSection')

  if (sheet.getName() === 'Employee Details') { 

    var row = range.getRow();
    var column = range.getColumn()

    var value = range.getValue();
    Logger.log('value = ' + value)

    /** Is it a dropdown menu? */   
    var hasDropdown = checkIfCellIsDropdown(value, row, column);
    /** Is it a dropdown menu? YES*/
    if (hasDropdown) {

      /**  Is it the main Employee dropdown menu? YES*/
      if (range.getA1Notation() === dropdownCellValue) {

        resetEmployeeDetailsLayout()
        /**  Is the value <CREATE NEW EMPLOYEE>? YES*/
        if (value === '<Create New Employee>') {
          Logger.log ('Setting <Create New Employee>')
          clearEmployeeDetails(empTemplateHeaderCells, selectedEmployeeKey);
          Logger.log ('Finishing Setting <Create New Employee>')

         
        } else {
          /**  Is the value <CREATE NEW EMPLOYEE>? NO*/
          // Populate the employee details if an existing employee is selected
          Logger.log('value trying to populate ' + JSON.stringify(value))

          var result = matchSelection(value, row, column);

          Logger.log("employeeOnEdit - empNameMatch: " + result.empNameMatch);
          unhideNamedRange(value,row,column,result.empNameMatch)
          
        }
      }else {
        /**  Is it the main Employee dropdown menu? NO*/ 
        Logger.log('value trying to populate is in a small dropdown: ' + JSON.stringify(value))
        var result = matchSelection(value, row, column);

        Logger.log("employeeOnEdit - empNameMatch: " + result.empNameMatch, result.dropdownHeader);
        unhideNamedRange(value,row,column,result.empNameMatch,dropdownHeader)
      }
    }

    
    // Retrieve and parse the empTemplateHeaderCells dictionary from user properties
    var userProperties = PropertiesService.getUserProperties();
    var empTemplateHeaderCellsJSON = userProperties.getProperty('empTemplateHeaderCells');
    var empTemplateHeaderCells = JSON.parse(empTemplateHeaderCellsJSON);
    var matchingResponse = JSON.parse(userProperties.getProperty('matchingResponse'));
    var filmDetails = JSON.parse(userProperties.getProperty('filmDetails'));

    populateEmployeeDetails(matchingResponse,result.empNameMatch)
    
    if (result.empNameMatch) {
      populateEmployeeDetails(filmDetails.find(response => response.objectName === "Film Details"), result.empNameMatch, true);
    }

    // Check if the selectedEmployee dictionary is complete
    var selectedEmployeeJSON = userProperties.getProperty(selectedEmployeeKey);
    var selectedEmployee = JSON.parse(selectedEmployeeJSON);

    Logger.log('selectedEmployee = ' + JSON.stringify(selectedEmployee))
    //Logger.log('selectedEmployeeKey = ' + JSON.stringify(selectedEmployeeKey))
  

    employeeCalc(e, value, row, column);


    if (selectedEmployee && range.getA1Notation() !== dropdownCellValue) {
      // Iterate over each cell in the edited range and apply the backgroundColour function
      range.getValues().forEach((row, i) => {
        row.forEach((value, j) => {
          var cell = range.getCell(i + 1, j + 1);
          backgroundColour( cell, empTemplateHeaderCells, selectedEmployee);
        });
      });
    }
  }

  loadingIndicator();

  Logger.log("Finished employeeOnEdit");
}
