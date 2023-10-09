

function populateEmployeeDetailsOLD(value, row, column, focusSection = null) {
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

  /** MATCHING RESPONSE START */

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

    /** Store matchingResponse to the userProperties */
    userProperties.setProperty('matchingResponse', JSON.stringify(matchingResponse));
    Logger.log ('matchingResponse = ' + JSON.stringify(matchingResponse))

    Logger.log ('empTemplateDropdownCells = ' + JSON.stringify(empTemplateDropdownCells))

    /** MATCHING RESPONSE END */

    /** UNHIDE SECTION START */

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

    /** UNHIDE SECTION END */

    /** VISIBLE NAMED RANGES START */

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

    /** VISIBLE NAMED RANGES END */

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


function loadingSectionOLD(value, row, column, dropdownHeader) {

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
