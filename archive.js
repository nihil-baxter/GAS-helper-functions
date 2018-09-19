function archive (fromSheet,toSheet,condition,conditionColHeader,headerRow) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetFrom = ss.getSheetByName(fromSheet);
    var sheetTo = ss.getSheetByName(toSheet);
    var lastTo = sheetTo.getLastRow()+1;
    var rangeFrom = sheetFrom.getDataRange();
    var valuesFrom = rangeFrom.getValues();
    var fromLen = valuesFrom.length;
    var headers = valuesFrom[headerRow-1]
    var conditionCol = headers.indexOf(conditionColHeader);
    var arrTo = [];
    var toDelete = [];
    var i,j;
    for (i = fromLen-1; i>0; i--) {
        if (valuesFrom[i][conditionCol] == condition) {
            arrTo.push(valuesFrom[i]);
            toDelete.push(i+1);
        } else {
            continue;
        }
    }
    if (arrTo.length === 0) {
        return false
    }
    var rangeTo = sheetTo.getRange(lastTo, 1, arrTo.length, arrTo[0].length);
    rangeTo.setValues(arrTo);
    var delStart = toDelete[toDelete.length-1];
    var delCount = 1;
    for (j=toDelete.length-1;j>=0;j--) {
        if(j>0 && toDelete[j-1] == delStart-1) {
            delCount++;
            continue;
        } else {
            sheetFrom.deleteRows(toDelete[j],delCount);
            delCount = 1;
        }
        
    }
}

function archiveNew(fromSheet,toSheet,condition,headerRow,ssId) {
    if (typeof fromSheet != "string") {
        new Error("Please provide the name of the source Sheet")
    }
    if (typeof toSheet != "string") {
        new Error("Please provide the name of the archive sheet")
    }
    if (typeof condition != "object" && typeof condition != "function") {
        new Error("Condition must be an Object or a Function.")
    }
    if (typeof condition == "object") {
        if (!condition.type) {
            condition.type = "AND"
        } else {
            if (condition.type != "AND" && condition.type != "OR") {
                new Error("Condition type must be 'AND' or 'OR'")
            }
        }
    }
    headerRow = headerRow || 1
    var ss = ssId ? SpreadsheetApp.openById(ssId) : SpreadsheetApp.getActiveSpreadsheet()
    var sourceSheet = ss.getSheetByName(fromSheet)
    var archiveSheet = ss.getSheetByName(toSheet)
    var lastRowArchive = archiveSheet.getLastRow()
    var sourceValues = sourceSheet.getRange(headerRow,1,sourceSheet.getLastRow(),sourceSheet.getLastColumn()).getValues();
    var sourceHeader = sourceValues.shift()
    var arrTo = [];
    var toDelete = []
    sourceValues.forEach(function(row,index) {
        var bool
        if (typeof condition == "object") {
            var keys = Object.keys(condition)
            keys.splice(keys.indexOf("type"))
            if(condition.type == "AND") {
                bool = keys.every(function (key) {
                    if (Array.isArray(condition[key])) {
                        return condition[key].some(function(item) {
                            return item == row[sourceHeader.indexOf(key)]
                        })
                    } else {
                        return condition[key] == row[sourceHeader.indexOf(key)]
                    }
                    
                })
                
            }
            if (condition.type == "OR") {
                bool = keys.some(function (key) {
                    if (Array.isArray(condition[key])) {
                        return condition[key].some(function(item) {
                            return item == row[sourceHeader.indexOf(key)]
                        })
                    } else {
                        return condition[key] == row[sourceHeader.indexOf(key)]
                    }
                    
                })
            }
        }
        if (typeof condition === "function") {
            bool = condition(row)
        }
        if (bool === true) {
            arrTo.push(row)
            toDelete.push(index+headerRow+1)
        }
    })
    if (arrTo.length === 0) {
        return false
    }
    var archiveRange = archiveSheet.getRange(lastRowArchive+1, 1, arrTo.length, arrTo[0].length);
    archiveRange.setValues(arrTo);
    var delStart = toDelete[toDelete.length-1];
    var delCount = 1;
    var i;
    for (i = toDelete.length-1; i >= 0; i--) {
        if(i > 0 && toDelete[i-1] == delStart-1) {
            delCount++;
            continue;
        } else {
            sourceSheet.deleteRows(toDelete[i],delCount);
            delCount = 1;
        }
        
    }
}