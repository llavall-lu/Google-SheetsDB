function getHeaderNameByCell(row, column, value) {
  Logger.log('Starting getHeaderNameByCell');
  Logger.log(value + '= row: ' + row + ', column: ' + column);

  var userProperties = PropertiesService.getUserProperties();
  var empTemplateHeaderCells = JSON.parse(userProperties.getProperty('empTemplateHeaderCells'));

  for (var namedRange in empTemplateHeaderCells) {
    Logger.log('namedRange = ' + namedRange) for (var header in empTemplateHeaderCells[namedRange]) {

      if (empTemplateHeaderCells[namedRange][header].row == row && empTemplateHeaderCells[namedRange][header].column == column) {
        Logger.log('Finishing getHeaderNameByCell and found header ' + header) return header;
      }
    }
  }
  Logger.log('Failed getHeaderNameByCell and returning "null"') return null;
}

function backgroundColour(cell, empTemplateHeaderCells, selectedEmployee) {

  var cellRow = cell.getRow();
  var cellCol = cell.getColumn();
  var cellValue = cell.getValue();

  var foundNamedRangeAndKey;
  for (var namedRange in empTemplateHeaderCells) {
    var key = Object.keys(empTemplateHeaderCells[namedRange]).find(key = >JSON.stringify(empTemplateHeaderCells[namedRange][key]) === JSON.stringify({
      row: cellRow,
      column: cellCol
    }));
    if (key) {
      foundNamedRangeAndKey = {
        namedRange: namedRange,
        key: key
      };
      break;
    }
  }

  if (!foundNamedRangeAndKey) {
    return;
  }

  var cellKey = foundNamedRangeAndKey.key;

  if (cellKey && cellKey.toLowerCase().includes('date')) {
    if (selectedEmployee[cellKey] !== '' && selectedEmployee[cellKey] !== null) {
      selectedEmployee[cellKey] = Utilities.formatDate(new Date(selectedEmployee[cellKey]), "GMT", "dd/MM/yyyy");
    }
  }

  if (cellValue === '' || cellValue === null) {
    if (selectedEmployee[cellKey] === '' || selectedEmployee[cellKey] === null) {
      cell.setBackground(emptyColour); 
    } else {
      cell.setBackground(editedColour); 
    }
  }

  else {
    var cellValueLowerCase = String(cellValue).toLowerCase();
    var selectedValueLowerCase = String(selectedEmployee[cellKey]).toLowerCase();

    if (cellKey && selectedValueLowerCase === cellValueLowerCase) {
      cell.setBackground(loadedColour); 
    } else {
      cell.setBackground(editedColour); 
    }
  }

}

function checkIfCellIsDropdown(cellValue, row, column) {

  var cell = empSheet.getRange(row, column) var dataValidation = cell.getDataValidation()

  if (dataValidation && dataValidation.getCriteriaType() === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
    var dropdownValues = dataValidation.getCriteriaValues();
    Logger.log('cellValue: ' + cellValue + ' is a dropdown menu with values: ' + dropdownValues.join(", "));
    return dropdownValues;
  } else {
    return false;
  }

}
