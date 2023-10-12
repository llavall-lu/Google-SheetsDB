


function resetLoadingIndicator() {
  var loadingCellPos = empSheet.getRange(loadingCellValue);
  loadingCellPos.setValue('');
  loadingCellPos.setBackground(whiteColour);
  loadingCellPos.setFontColor(blackColour);

}
function loadingIndicator() {
  var loadingCellPos = empSheet.getRange(loadingCellValue);
  var loadingCell = loadingCellPos.getValue();

  if (loadingCell !== 'Loading...') {
    loadingCellPos.setValue('Loading...');

    loadingCellPos.setBackground(loadingColour);
    loadingCellPos.setFontColor(whiteColour);

    SpreadsheetApp.flush();
  } else {

    SpreadsheetApp.flush();

    loadingCellPos.setValue('');
    loadingCellPos.setBackground(whiteColour);
    loadingCellPos.setFontColor(blackColour);
  }
}
