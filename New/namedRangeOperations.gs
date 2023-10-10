function unhideNamedRange(value, row, column, empNameMatch, dropdownHeader) {

  Logger.log("Starting unhideNamedRange");

  var userProperties = PropertiesService.getUserProperties();
  var empTemplateNamedRanges = JSON.parse(userProperties.getProperty('empTemplateNamedRanges'));
  var empTemplateDropdownCells = JSON.parse(userProperties.getProperty('empTemplateDropdownCells'));
  var matchingResponse = JSON.parse(userProperties.getProperty('matchingResponse'));

  if (empNameMatch) {
    var unhiddenNamedRanges = [];
    var hiddenNamedRanges = []
    var hiddenNamedRanges = Object.keys(empTemplateNamedRanges).filter(range = >range !== "mainDetails" && range !== "hoursSection");

    for (var dropdownTitle in empTemplateDropdownCells) {

      if (matchingResponse.response.hasOwnProperty(dropdownTitle)) {

        var dropdownValue = matchingResponse.response[dropdownTitle].replace(/\s+/g, '');

        Logger.log('unhideNamedRange: dropdownValue = ' + dropdownValue)

        if (empTemplateNamedRanges.hasOwnProperty(dropdownValue)) {

          var namedRangeInfo = empTemplateNamedRanges[dropdownValue];
          empSheet.showRows(namedRangeInfo.rowStart, namedRangeInfo.rowEnd - namedRangeInfo.rowStart + 1);
          unhiddenNamedRanges.push(dropdownValue)

          hiddenNamedRanges = hiddenNamedRanges.filter(range = >range !== dropdownValue);
        }
      }
    }

    userProperties.setProperty('hiddenNamedRanges', JSON.stringify(hiddenNamedRanges));
    userProperties.setProperty('unhiddenNamedRanges', JSON.stringify(unhiddenNamedRanges));

    Logger.log('This is an employee Name Match and the following namedRanges have been unhidden: ' + unhiddenNamedRanges.join(", ")) Logger.log('The following Ranges are hidden: ' + hiddenNamedRanges.join(", ")) Logger.log("Finishing unhideNamedRange");
    return
  }

  if (!empNameMatch) {
    if (empTemplateDropdownCells.hasOwnProperty(dropdownHeader)) {

      var selectedDropdown = {
        dropdown: empTemplateDropdownCells[dropdownHeader].dropdown,
        values: empTemplateDropdownCells[dropdownHeader].values
      };
      Logger.log('selectedDropdown = ' + JSON.stringify(selectedDropdown, null, 2));
    } else {
      Logger.log('No matching dropdownHeader found in empTemplateDropdownCells for: ' + dropdownHeader);
    }

    Logger.log('This is not an employee Name Match and needs namedRanges hidden')

  }

}
