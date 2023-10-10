function matchSelection(value,row,column) {

  Logger.log("Starting matchSelection");

  var empNameMatch = false

  var userProperties = PropertiesService.getUserProperties();
  var empNameDictionary = JSON.parse(userProperties.getProperty('nameDictionary'));
  var filmDetails = JSON.parse(userProperties.getProperty('filmDetails'));
  var empTemplateHeaderCells = JSON.parse(userProperties.getProperty('empTemplateHeaderCells'));
  var empTemplateNamedRanges = JSON.parse(userProperties.getProperty('empTemplateNamedRanges'));
  var empTemplateDropdownCells = JSON.parse(userProperties.getProperty('empTemplateDropdownCells'));
  var storedMatchingResponse = JSON.parse(userProperties.getProperty('matchingResponse'));

  //Logger.log ('empNameNameDictionary = ' +JSON.stringify(empNameNameDictionary))

  /**  Find the entry in empNameNameDictionary where the empName matches the selected value */
  var empNameMatchingEntry = empNameDictionary.find(entry => entry.objectName === value);
  
  Logger.log ('Employee Name Matching Entry = ' +JSON.stringify(empNameMatchingEntry))

  

  /** MATCHING RESPONSE START  */
  /**Is the value an employee name? YES */
  if (empNameMatchingEntry) {

    empNameMatch = true

    /** Get formResponse Row*/
    // Get the relevant row from formSheet using the rowNumber from empNameMatchingEntry
    var formResponseRow = formSheet.getRange(empNameMatchingEntry.rowNumber, 1, 1, formSheet.getLastColumn()).getValues();

    /** Get formHeaders */
    var formHeaders = formSheet.getRange(1, 1, 1, formSheet.getLastColumn()).getValues()[0];

    /** Convert the row data and headers into an object for easier processing */ 
    var matchingResponse = {
      objectName: value,
      response: {}
    };
    //find and match any namedRanges that match the dropdown values within matchingResponse
    for (var i = 0; i < formHeaders.length; i++) {
      matchingResponse.response[formHeaders[i]] = formResponseRow[0][i];
    }

    /** Store matchingResponse to the userProperties */
    userProperties.setProperty('matchingResponse', JSON.stringify(matchingResponse));

    Logger.log('empNameMatch = ' + empNameMatch)
    Logger.log ('matchingResponse = ' + JSON.stringify(matchingResponse))

    Logger.log ('empTemplateDropdownCells = ' + JSON.stringify(empTemplateDropdownCells))

    /** MATCHING RESPONSE END  */

      return {
        dropdownHeader: dropdownHeader,
        empNameMatch: empNameMatch
      }
  /**Is the value an employee name? NO */
  } else {

    Logger.log("No matching entry in the main Employee Name Dropdown found for: " + value);


    /** Attempt to find the named range that matches the value */ 
    /** convert value to NamedRange format with no spaces - ONLY use this to find value in namedRange */
    var valueNoSpace = value.replace(/\s+/g, '')
    Logger.log("valueNoSpace: " + valueNoSpace);


   /** Find and match a namedRanges that match the value */ 
    var namedRangeMatch = empTemplateNamedRanges[valueNoSpace];
    
    /** Find and match a namedRanges - YES */
    if (namedRangeMatch) {
        Logger.log("YES!, namedRangeMatch matches entry found in NamedRange for: " + JSON.stringify(valueNoSpace));

        /** Get the header name using the provided row and column */ 
        var dropdownHeader = getHeaderNameByCell(row, column, value);
        Logger.log('dropdownHeader = ' + JSON.stringify(dropdownHeader))
        
        /**Is ther a dropdownHeader? NO */
        if (!dropdownHeader) {
        Logger.log("Unable to determine the dropdownHeader for the given row/column. Finishing matchSelection");
        return;
    }
      /**Is ther a dropdownHeader? YES */
      /**Is the value different to the matchingResponse value? YES */
      if (storedMatchingResponse.response && storedMatchingResponse.response[dropdownHeader] !== value) {
        // This is where you'd handle the case where the values are different
        Logger.log(value + " (value) is different from "+ storedMatchingResponse.response[dropdownHeader] + "(the stored matching response!)");
        
        return {
        dropdownHeader: dropdownHeader,
        empNameMatch: empNameMatch
      }
      
      }else{
        /**Is the value different to the matchingResponse value? NO */
        Logger.log('values are the same. Finishing matchSelection')
        
        return;
      }
    /** Find and match a namedRanges - NO */
    } else {
      Logger.log('No namedRange matches this row or column. Please refresh the page and start again.')
      Logger.log("Finishing matchSelection");

      return;
    }
  }    
}

/** **************************************************************************************************** */

function unhideNamedRange(value, row, column, empNameMatch, dropdownHeader) {
  
  Logger.log("Starting unhideNamedRange");

  var userProperties = PropertiesService.getUserProperties();
  var empTemplateNamedRanges = JSON.parse(userProperties.getProperty('empTemplateNamedRanges'));
  var empTemplateDropdownCells = JSON.parse(userProperties.getProperty('empTemplateDropdownCells'));
  var matchingResponse = JSON.parse(userProperties.getProperty('matchingResponse'));

  if (empNameMatch){
    var unhiddenNamedRanges = [];
    var hiddenNamedRanges = []
    var hiddenNamedRanges = Object.keys(empTemplateNamedRanges).filter(range => range !== "mainDetails" && range !== "hoursSection");

    /**  Loop over the empTemplateDropdownCells headers */
    for (var dropdownTitle in empTemplateDropdownCells) {
      /** Check if the header exists in matchingResponse.response */
      if (matchingResponse.response.hasOwnProperty(dropdownTitle)) {
        /** Get the corresponding value from matchingResponse */
        var dropdownValue = matchingResponse.response[dropdownTitle].replace(/\s+/g, '');

        Logger.log('unhideNamedRange: dropdownValue = ' + dropdownValue)

        /** Check if the value matches a namedRange in empTemplateNamedRanges */
        if (empTemplateNamedRanges.hasOwnProperty(dropdownValue)) {
          /** Unhide the namedRange using the rowStart and rowEnd and save name to unhiddenNamedRanges variable */
          var namedRangeInfo = empTemplateNamedRanges[dropdownValue];
          empSheet.showRows(namedRangeInfo.rowStart, namedRangeInfo.rowEnd - namedRangeInfo.rowStart + 1);
          unhiddenNamedRanges.push(dropdownValue)

          // Remove the unhidden range from the hiddenNamedRanges list
          hiddenNamedRanges = hiddenNamedRanges.filter(range => range !== dropdownValue);
        }
      }
    }

    userProperties.setProperty('hiddenNamedRanges', JSON.stringify(hiddenNamedRanges));
    userProperties.setProperty('unhiddenNamedRanges', JSON.stringify(unhiddenNamedRanges));

    Logger.log ('This is an employee Name Match and the following namedRanges have been unhidden: ' + unhiddenNamedRanges.join(", "))
    Logger.log ('The following Ranges are hidden: ' + hiddenNamedRanges.join(", "))
    Logger.log("Finishing unhideNamedRange");
     return
  }

  if (!empNameMatch) {
    if (empTemplateDropdownCells.hasOwnProperty(dropdownHeader)) {
      // Assign the 'values' attribute of the corresponding dropdownHeader to selectedDropdown
      var selectedDropdown = {
      dropdown: empTemplateDropdownCells[dropdownHeader].dropdown,
      values: empTemplateDropdownCells[dropdownHeader].values
    };
      Logger.log('selectedDropdown = ' + JSON.stringify(selectedDropdown, null, 2));
    } else {
      Logger.log('No matching dropdownHeader found in empTemplateDropdownCells for: ' + dropdownHeader);
    }

    /**CLEAR DATA IN EXISITNG NAMEDRANGE, HIDE EXISTING NAMEDRANGE IN DROPDOWNHEADER, TAKE THE VALUE AND UNHIDE THE NEW NAMEDRAGE.  */
    Logger.log ('This is not an employee Name Match and needs namedRanges hidden')

  }

}

/** **************************************************************************************************** */

function populateEmployeeDetails(response, empNameMatch, storeToSelectedEmployee = true) {

  Logger.log('Starting populateEmployeeDetails')

  Logger.log('empNameMatch = ' + empNameMatch)

  var userProperties = PropertiesService.getUserProperties();
  var empTemplateHeaderCells = JSON.parse(userProperties.getProperty('empTemplateHeaderCells'));

  var unhiddenNamedRanges = JSON.parse(userProperties.getProperty('unhiddenNamedRanges'));



  if (empNameMatch) {
    var selectedEmployee = {}
    unhiddenNamedRanges.push ("mainDetails", "hoursSection")
    Logger.log ('The following namedRanges are unhidden: ' + unhiddenNamedRanges.join(", "))
  }

  if (response) {
    for (var namedRange in empTemplateHeaderCells) {  // Iterate over named ranges in empTemplateHeaderCells
      
      // Skip this named range if it's not in unhiddenNamedRanges
      if (!unhiddenNamedRanges.includes(namedRange)) {
        continue;
      }

      Logger.log('Checking namedRange: ' + namedRange);

      for (var key in response.response) {  // Iterate over response keys
        // Check if the key exists in the empTemplateHeaderCells dictionary under the namedRange
        if (empTemplateHeaderCells[namedRange].hasOwnProperty(key)) {
          // Get the cell details
          var cellDetails = empTemplateHeaderCells[namedRange][key];
          Logger.log('Checking key: ' + key);
          
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

  Logger.log('selectedEmployee = ' + JSON.stringify(selectedEmployee));

  Logger.log('Finishing populateEmployeeDetails')
}
