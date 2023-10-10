function matchSelection(value, row, column) {

  Logger.log("Starting matchSelection");

  var empNameMatch = false

  var userProperties = PropertiesService.getUserProperties();
  var empNameDictionary = JSON.parse(userProperties.getProperty('nameDictionary'));
  var filmDetails = JSON.parse(userProperties.getProperty('filmDetails'));
  var empTemplateHeaderCells = JSON.parse(userProperties.getProperty('empTemplateHeaderCells'));
  var empTemplateNamedRanges = JSON.parse(userProperties.getProperty('empTemplateNamedRanges'));
  var empTemplateDropdownCells = JSON.parse(userProperties.getProperty('empTemplateDropdownCells'));
  var storedMatchingResponse = JSON.parse(userProperties.getProperty('matchingResponse'));

  var empNameMatchingEntry = empNameDictionary.find(entry = >entry.objectName === value);

  Logger.log('Employee Name Matching Entry = ' + JSON.stringify(empNameMatchingEntry))

  if (empNameMatchingEntry) {

    empNameMatch = true

    var formResponseRow = formSheet.getRange(empNameMatchingEntry.rowNumber, 1, 1, formSheet.getLastColumn()).getValues();

    var formHeaders = formSheet.getRange(1, 1, 1, formSheet.getLastColumn()).getValues()[0];

    var matchingResponse = {
      objectName: value,
      response: {}
    };

    for (var i = 0; i < formHeaders.length; i++) {
      matchingResponse.response[formHeaders[i]] = formResponseRow[0][i];
    }

    userProperties.setProperty('matchingResponse', JSON.stringify(matchingResponse));

    Logger.log('empNameMatch = ' + empNameMatch) Logger.log('matchingResponse = ' + JSON.stringify(matchingResponse))

    Logger.log('empTemplateDropdownCells = ' + JSON.stringify(empTemplateDropdownCells))

    return {
      dropdownHeader: dropdownHeader,
      empNameMatch: empNameMatch
    }

  } else {

    Logger.log("No matching entry in the main Employee Name Dropdown found for: " + value);

    var valueNoSpace = value.replace(/\s+/g, '') Logger.log("valueNoSpace: " + valueNoSpace);

    var namedRangeMatch = empTemplateNamedRanges[valueNoSpace];

    if (namedRangeMatch) {
      Logger.log("YES!, namedRangeMatch matches entry found in NamedRange for: " + JSON.stringify(valueNoSpace));

      var dropdownHeader = getHeaderNameByCell(row, column, value);
      Logger.log('dropdownHeader = ' + JSON.stringify(dropdownHeader))

      if (!dropdownHeader) {
        Logger.log("Unable to determine the dropdownHeader for the given row/column. Finishing matchSelection");
        return;
      }

      if (storedMatchingResponse.response && storedMatchingResponse.response[dropdownHeader] !== value) {

        Logger.log(value + " (value) is different from " + storedMatchingResponse.response[dropdownHeader] + "(the stored matching response!)");

        return {
          dropdownHeader: dropdownHeader,
          empNameMatch: empNameMatch
        }

      } else {

        Logger.log('values are the same. Finishing matchSelection')

        return;
      }

    } else {
      Logger.log('No namedRange matches this row or column. Please refresh the page and start again.') Logger.log("Finishing matchSelection");

      return;
    }
  }
}
function loadingSectionOLD(value, row, column, dropdownHeader) {

  Logger.log("Starting loadingSection")

  Logger.log('Loading Section value = ' + value) Logger.log('Loading Section row = ' + row) Logger.log('Loading Section column = ' + column) Logger.log('Loading Section dropdownHeader = ' + dropdownHeader)

  var userProperties = PropertiesService.getUserProperties();
  var selectedEmployeeJSON = userProperties.getProperty(selectedEmployeeKey);
  var empTemplateHeaderCells = JSON.parse(userProperties.getProperty('empTemplateHeaderCells'));
  var selectedEmployee = JSON.parse(selectedEmployeeJSON);
  Logger.log('loadingSection savedSelectedEmployee = ' + JSON.stringify(selectedEmployee)) var namedRangesObject = JSON.parse(userProperties.getProperty('empTemplateNamedRanges'));

  var valueNoSpace = value.replace(/\s+/g, '') Logger.log('valueNoSpace = ' + valueNoSpace)

  var namedRangeKeys = Object.keys(namedRangesObject);
  var matchingKey = namedRangeKeys.find(key = >key.replace(/\s+/g, '') === valueNoSpace);
  var selectedNamedRange = namedRangesObject[matchingKey];
  Logger.log('loadingSection namedRangesObject = ' + JSON.stringify(namedRangesObject)) Logger.log('loadingSection namedRangeKeys = ' + JSON.stringify(namedRangeKeys)) Logger.log('loadingSection matchingKey = ' + matchingKey) Logger.log('loadingSection selectedNamedRange = ' + JSON.stringify(selectedNamedRange))

  if (selectedNamedRange) {
    var startRow = selectedNamedRange.rowStart
    var endRow = selectedNamedRange.rowEnd
    var numRows = endRow - startRow + 1
    var targetRange = empSheet.getRange(startRow, endRow, numRows);
    Logger.log('targetRange = ' + targetRange) targetRange.activate();
    empSheet.unhideRow(targetRange);
    var row = targetRange.getRow();
    var column = targetRange.getColumn();
  } else {
    Logger.log('Not in the selectedNameRange') return
  }

  Logger.log('Section trying to populate ' + JSON.stringify(value)) populateEmployeeDetails(value, row, column);

  Object.keys(namedRangesObject).forEach(key = >{
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
