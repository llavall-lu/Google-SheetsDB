function employeeCalc(e, value, row, column) {

  Logger.log('Starting employeeCalc: value = ' + value + ' Row = ' + row + ' Column = ' + column)

  var userProperties = PropertiesService.getUserProperties();
  var autoDetails = userProperties.getProperty('autoDetails');

  if (autoDetails === 'ON') {
    Logger.log('AutoDetails is ON')

    var namedRanges = ss.getNamedRanges();

    var filteredNamedRanges = namedRanges.filter(function(namedRange) {
      return namedRange.getName() !== 'employeeDetails';
    });

    for (var i = 0; i < filteredNamedRanges.length; i++) {
      var namedRange = filteredNamedRanges[i];
     var range = namedRange.getRange();
if (range.getRow() <= row && row <= range.getLastRow() && range.getColumn() <= column && column <= range.getLastColumn()) {

      if (namedRange.getName() === 'WeeklyDailyFee') {
    Logger.log('FEECALCULATIONS GO!');
    feeCalculations(e);
} else if (namedRange.getName() === 'hoursSection') {
    Logger.log('WORKINGDATES GO!');
    workingDates(e);
} else {
    Logger.log('NO CALC MATCHES');
}

      }
    }

  } else if (autoDetails === 'OFF') {
    Logger.log('AutoDetails is OFF') 
    return

  }

  Logger.log('Finishing employeeCalc: value = ' + value + ' Row = ' + row + ' Column = ' + column)

}
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

  for (var key in empTemplateHeaderCells) {
    var cellInfo = empTemplateHeaderCells[key];
    if (cellInfo.row === activeCellRow && cellInfo.column === activeCellCol) {
      activeKey = key;
      break;
    }
  }

  if (! ['Daily Rate inc Hol Pay', 'Weekly Rate inc Hol Pay', 'Holiday Pay Percentage'].includes(activeKey)) {
    return;
  }

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

  var weeklyExclHolPayFormula = '=round((' + weeklyInclHolPayCell.getA1Notation() + '/' + holidayPayPercentageCellValue + '), 2)';
  var dailyExclHolPayFormula = '=round((' + dailyInclHolPayCell.getA1Notation() + '/' + holidayPayPercentageCellValue + '), 2)';
  weeklyExclHolPayCell.setFormula(weeklyExclHolPayFormula);
  dailyExclHolPayCell.setFormula(dailyExclHolPayFormula);

  backgroundColour(weeklyExclHolPayCell, empTemplateHeaderCells, selectedEmployee);
  backgroundColour(dailyExclHolPayCell, empTemplateHeaderCells, selectedEmployee);

  var weeklyHolPayValueFormula = '=round((' + weeklyInclHolPayCell.getA1Notation() + '-' + weeklyExclHolPayCell.getA1Notation() + '),2)';
  var dailyHolPayValueFormula = '=round((' + dailyInclHolPayCell.getA1Notation() + '-' + dailyExclHolPayCell.getA1Notation() + '),2)';
  weeklyHolPayValueCell.setFormula(weeklyHolPayValueFormula);
  dailyHolPayValueCell.setFormula(dailyHolPayValueFormula);

  backgroundColour(weeklyHolPayValueCell, empTemplateHeaderCells, selectedEmployee);
  backgroundColour(dailyHolPayValueCell, empTemplateHeaderCells, selectedEmployee);

  Logger.log("Finished feeCalculations");
}
