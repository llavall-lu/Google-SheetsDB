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

