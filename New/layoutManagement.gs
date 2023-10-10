function resetEmployeeDetailsLayout() {

  Logger.log("Starting resetEmployeeDetailsLayout");

  var totalRows = empSheet.getMaxRows();
  if (totalRows > 3) {

    empSheet.deleteRows(4, totalRows - 3);

  }

  empSheet.insertRowAfter(3);

  var employeeDetails = ss.getRangeByName('employeeDetails');

  if (employeeDetails) {

    empSheet.insertRows(4, employeeDetails.getNumRows() - 1); 
    employeeDetails.copyTo(empSheet.getRange(4, 1));
  } else {
    Logger.log("'employeeDetails' named range not found.");
  }

  var namedRanges = ss.getNamedRanges();

  namedRanges.forEach(function(namedRange) {
    var rangeName = namedRange.getName();
    var range = namedRange.getRange();

    if (rangeName !== 'employeeDetails' && rangeName !== 'mainDetails' && rangeName !== 'hoursSection') {
      empSheet.hideRows(range.getRow(), range.getNumRows());
    }
  });

  Logger.log("Finished resetEmployeeDetailsLayout");

}

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

  userProperties.setProperty('autoDetails', autoDetails); 

  loadingIndicator()

  Logger.log("Finished autoDetailsSwitch");
}
