function addData(data,type, sheetName, ssId, startColumn, removeCsvHeader, headerRow, countRowsByCol) {
    var ss = ssId ? SpreadsheetApp.openById(ssId) : SpreadsheetApp.getActiveSpreadsheet()
    var sheet = ss.getSheetByName(sheetName)
    headerRow = headerRow || 0
    startColumn = startColumn || 1
    if (type == "r") {
        if (removeCsvHeader == true) {
            data.shift()
        }
        sheet.getRange(headerRow+1,startColumn,sheet.getMaxRows(),data[0].length).clearContent(),
        sheet.getRange(headerRow+1,startColumn,data.length,data[0].length).setValues(data)
    } else if (type = "a") {
        data.shift()
        var lastRow = countRowsByCol ? Sheet.getRange(countRowsByCol + "1",countRowsByCol).getValues().length : Sheet.getLastRow()
        sheet.getRange(lastRow+1,startColumn,data.length,data[0].length).setValues(data)
    }
}