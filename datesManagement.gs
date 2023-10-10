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
        Logger.log('activeCategory = ' + activeCategory + ' and activeKey = ' + activeKey) break;
      }
    }
    if (activeKey) break; 
  }

  if (! ['Start Date', 'End Date', 'Weeks'].includes(activeKey)) {
    Logger.log('No working Dates selected. Finished workingDates') return;
  }

  var startDateInfo = empTemplateHeaderCells[activeCategory]['Start Date'];
  var startDateCell = empSheet.getRange(startDateInfo.row, startDateInfo.column);
  var startDateValue = new Date(startDateCell.getValue());
  var endDateInfo = empTemplateHeaderCells[activeCategory]['End Date'];
  var endDateCell = empSheet.getRange(endDateInfo.row, endDateInfo.column);
  var endDateValue = new Date(endDateCell.getValue());
  var weeksInfo = empTemplateHeaderCells[activeCategory]['Weeks'];
  var weeksCell = empSheet.getRange(weeksInfo.row, weeksInfo.column);
  var weeksValue = weeksCell.getValue();

  switch (activeKey) {
  case 'Start Date':
    if (startDateValue && endDateValue) {
      var timeDiff = endDateValue - startDateValue + (1000 * 3600 * 24); 
      var diffWeeks = Math.floor(timeDiff / (1000 * 3600 * 24 * 7));
      var diffDays = Math.ceil((timeDiff % (1000 * 3600 * 24 * 7)) / (1000 * 3600 * 24));
      var weekString = diffWeeks + ' weeks ' + diffDays + ' days';
      weeksCell.setValue(weekString);
      backgroundColour(weeksCell, empTemplateHeaderCells, selectedEmployee);
    } else {

      endDateCell.setValue('');
      weeksCell.setValue('');
    }
    break;
  case 'Weeks':
    if (weeksValue === '' || isNaN(weeksValue)) {

      endDateCell.clearContent();
      backgroundColour(endDateCell, empTemplateHeaderCells, selectedEmployee);
    } else {

      var endDate = new Date(startDateValue);
      endDate.setHours(0, 0, 0, 0); 
      endDate.setDate(endDate.getDate() + weeksValue * 7 - 1); 

      var totalDays = weeksValue * 7;
      var exactWeeks = Math.floor(totalDays / 7);
      var remainingDays = totalDays % 7;

      weeksCell.setValue(exactWeeks + ' weeks ' + remainingDays + ' days');

      endDateCell.setValue(endDate);
      backgroundColour(endDateCell, empTemplateHeaderCells, selectedEmployee);
    }
    break;
  case 'End Date':
    if (endDateValue) {
      var timeDiff = endDateValue - startDateValue + (1000 * 3600 * 24); 
      var diffWeeks = Math.floor(timeDiff / (1000 * 3600 * 24 * 7));
      var diffDays = Math.ceil((timeDiff % (1000 * 3600 * 24 * 7)) / (1000 * 3600 * 24));
      var weekString = diffWeeks + ' weeks ' + diffDays + ' days';
      weeksCell.setValue(weekString);
      backgroundColour(weeksCell, empTemplateHeaderCells, selectedEmployee);
    } else {

      weeksCell.setValue('');
    }
    break;
  }
  Logger.log("Finished workingDates");
}
