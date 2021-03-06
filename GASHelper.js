//SCRIPT_NAME = GAS-Helper-Functions
//SCRIPT_VERSION = dev.v0_PL
//SCRIPT_ID = 1x4QZqy-MRtnWwgygqMPjokDo9amDbrGzn6hK_oceEHgsT4AlDTOEng5e


;(function (root,factory) {
    root.GASHelper = factory()
})(this, function() {
    var GASHelper = {}

    function permission() {
        GmailApp.getInboxUnreadCount()
    }

    function archive(fromSheet,toSheet,condition,headerRow,ssId) {
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
                keys.splice(keys.indexOf("type"),1)
                if(condition.type == "AND") {
                    bool = keys.every(function (key) {
                        if (typeof condition[key] == "function") {
                            return condition[key](row[sourceHeader.indexOf(key)])
                        }
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
                        if (typeof condition[key] == "function") {
                            return condition[key](row[sourceHeader.indexOf(key)])
                        }
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

    function listCompare(inputList,existingList,compObj) {
        var existingList = JSON.parse(JSON.stringify(existingList));
        var inputKeys = Object.keys(inputList).sort();
        var existingKeys = Object.keys(existingList).sort();
        var inputLen = inputKeys.length;
        var existingLen = existingKeys.length;
        var removedItems = {};
        var addedItems = {};
        var i = 0;
        var j = 0;
        var iFin = false;
        var jFin = false;
        var header = Object.keys(existingList[existingKeys[0]]);
        while (i<inputLen || j<existingLen) {
            if (inputKeys[i] == undefined ) {
                i = inputLen -1;
                iFin = true;
            }
            if (existingKeys[j] == undefined ) {
                j = existingLen -1;
                jFin = true;
            }
            if (iFin && jFin) {
                break; 
            }
            if (inputKeys[i] < existingKeys[j] || jFin) {
                if (existingKeys.indexOf(inputKeys[i]) === -1) {
                    
                if (compObj && typeof compObj === "object") {
                    if (compObj.hasOwnProperty("compare") && typeof compObj["compare"] === "object") {
                        for (objHeader in compObj["compare"]) {
                            if (compObj["compare"].hasOwnProperty(objHeader)) {
                                var context = {};
                                var firstArgument = inputList[inputKeys[i]];
                                for (result in compObj["compare"][objHeader]) {
                                    if (compObj["compare"][objHeader].hasOwnProperty(result)) {
                                        var func = compObj["compare"][objHeader][result].bind(context,firstArgument);
                                        if (func() === true) {
                                            inputList[inputKeys[i]][objHeader] = result;
                                            break;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                    header.forEach(function(item) {
                        if (Object.keys(inputList[inputKeys[i]]).indexOf(item) === -1) {
                            inputList[inputKeys[i]][item] = "";
                        }
                    });
                    existingList[inputKeys[i]] = JSON.parse(JSON.stringify(inputList[inputKeys[i]]));
                    addedItems[inputKeys[i]] = JSON.parse(JSON.stringify(inputList[inputKeys[i]]));
                }
                i++;
            }
            if (existingKeys[j] < inputKeys[i] || iFin) {
                if (inputKeys.indexOf(existingKeys[j]) === -1 && existingList[existingKeys[j]] != undefined) {
                    removedItems[existingKeys[j]] = JSON.parse(JSON.stringify(existingList[existingKeys[j]]));
                    delete existingList[existingKeys[j]];
                }
                j++;
            };
            if (existingKeys[j] === inputKeys[i] && (!iFin && !jFin)) {
                if (compObj && typeof compObj === "object") {
                    if (compObj.hasOwnProperty("update") && Array.isArray(compObj["update"]) && compObj["update"].length > 0) {
                        for (k=0;k<compObj["update"].length;k++) {
                            var headerName = compObj["update"][k];
                            if (inputKeys[i] != undefined) {
                                existingList[inputKeys[i]][headerName] = inputList[inputKeys[i]][headerName];
                            }
                        }
                    }
                    if (compObj.hasOwnProperty("compare") && typeof compObj["compare"] === "object") {
                        for (objHeader in compObj["compare"]) {
                            if (compObj["compare"].hasOwnProperty(objHeader)) {
                                var context = {};
                                var firstArgument = inputList[inputKeys[i]];
                                for (result in compObj["compare"][objHeader]) {
                                    if (compObj["compare"][objHeader].hasOwnProperty(result)) {
                                        var func = compObj["compare"][objHeader][result].bind(context,firstArgument);
                                        if (func() === true) {
                                            existingList[inputKeys[i]][objHeader] = result;
                                            break;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                i++;
                j++;
            }
        }
            return {
                updated: existingList,
                removed: removedItems,
                added: addedItems
                }
        
    }

    function importCsv(searchString, searchPlace, fileName, parseConfig, deleteMailFile, safeMode, encoding, manipulateData) {
        if (searchString == undefined || typeof searchString != "string") {
            return Error("searchString is mandatory and has to be a String");
            };
        parseConfig = parseConfig || {delimiter: ",", skipEmptyLines: true};
        fileName = fileName || "";
        searchPlace = searchPlace.toLowerCase() || "m";
        //deleteMail = deleteMail || true;
        if (deleteMailFile === undefined) {
            deleteMailFile = true;
        }
        if (safeMode === undefined) {
            safeMode = false;
        }
        encoding = encoding === undefined ? "cp1252" : encoding
        if (searchPlace !== "m" && searchPlace !== "d") {
            return Error("searchPlace must be a String with the value 'm' or 'd'");
        }
        if (searchPlace === "m") {
            var threads, message, attachments, attName, parseString, parsedCsv, data;
            var attFound = false;
            threads = GmailApp.search(searchString);
            if (threads.length == 0) {
                return Error("No Email matched the search string")
            }
            message = threads[0].getMessages()[0];
            attachments = message.getAttachments();
            if (attachments.length == 0) {
                return Error("No Attachments");
            }
            if (attachments.length > 1 && fileName === "") {
                return Error("More than one attachment, please provide a filename.");
            }
            var i;
            for(i=0;i<attachments.length;i++) {
                attName = attachments[i].getName();
                if(attName.slice(-3).toUpperCase() === 'CSV' && attName.indexOf(fileName) > -1) {
                    attFound = true;
                    parseString = attachments[i].getDataAsString(encoding)
                    if (safeMode === true) {
                        parseString = parseString.replace(/(?:\\"|")+/g, '')
                    }
                    break;
                };
            };
            if (!attFound) {
                return Error("No matching attachment found, please check filename and/or -type");
            }
        } else if (searchPlace === "d") {
            var files, file , blob, driveFileName, parseString;
            var fileFound = false;
            files = DriveApp.searchFiles(searchString);
            while (files.hasNext() && !fileFound) {
                file = files.next();
                driveFileName = file.getName();
                if(driveFileName.slice(-3).toUpperCase() === 'CSV' && driveFileName.indexOf(fileName) > -1) {
                    fileFound = true;
                    blob = file.getBlob();
                    parseString = blob.getDataAsString(encoding)
                    if (safeMode === true) {
                        parseString = parseString.replace(/(?:\\"|")+/g, '')
                    }
                }
            }
        };
        parsedCsv = Papa.parse(parseString,parseConfig);
        /*if (parsedCsv.errors.length > 0) {
            
        }*/
        data = parsedCsv.data;
        if (manipulateData != undefined && typeof manipulateData === "function") {
            data = manipulateData(data);
        };
        if (typeof deleteMailFile === "boolean") {
            if (deleteMailFile && searchPlace === "m") {
                threads[0].moveToTrash();
            };
            if (deleteMailFile && searchPlace === "d") {
                file.setTrashed();
            };
        } else if (typeof deleteMailFile === "object" && searchPlace === "m") {
            if (deleteMailFile.hasOwnProperty("applyLabel")) {
                if (Array.isArray(deleteMailFile["applyLabel"])) {
                    deleteMailFile["applyLabel"].forEach(function(x) {
                        var label = GmailApp.getUserLabelByName(x);
                        threads[0].addLabel(label);
                    })
                } else if (typeof deleteMailFile["applyLabel"] === "string") {
                    var label = GmailApp.getUserLabelByName(deleteMailFile["applyLabel"]);
                    threads[0].addLabel(label);
                }
            };
            if (deleteMailFile.hasOwnProperty("markRead") && deleteMailFile["markRead"] === true) {
                threads[0].markRead();
            };
            if (deleteMailFile.hasOwnProperty("removeInbox") && deleteMailFile["removeInbox"] === true) {
                threads[0].moveToArchive();
            }
        }
        return data
    };

    function addDataToSheet(data, sheetName, ssId, type, startColumn, removeCsvHeader, headerRow, countRowsByCol, useAdvanced) {
        useAdvanced = useAdvanced || false
        if (useAdvanced) {
            if (typeof Sheets === "undefined") {
                throw new Error("Please enable Advanced Sheets Services.")
            }
            ssId = ssId ? ssId : SpreadsheetApp.getActiveSpreadsheet().getId()
        }
        if (!useAdvanced) {
            var ss = ssId ? SpreadsheetApp.openById(ssId) : SpreadsheetApp.getActiveSpreadsheet()
            var sheet = sheet ? ss.getSheetByName(sheetName) : ss.getActiveSheet()
        }
        removeCsvHeader = removeCsvHeader || false
        type = type || "r"
        headerRow = headerRow || 0
        startColumn = startColumn || 1
        if (typeof startColumn === "string" && !useAdvanced) {
            startColumn = a1ToColNum(startColumn)
        } else if (typeof startColumn === "number" && useAdvanced) {
            startColumn = colNumToA1(startColumn)
        }
        if (type == "r") {
            if (removeCsvHeader == true) {
                data.shift()
            }
            if (useAdvanced) {
                var rangeString = "'" + sheetName + "'!" + startColumn + (headerRow+1)   
                Sheets.Spreadsheets.Values.clear({}, ssId, rangeString + ":ZZZ")
                writeDataAdv(data, rangeString, ssId, "RAW", sheetName)
            } else {
                sheet.getRange(headerRow+1,startColumn,sheet.getMaxRows(),data[0].length).clearContent(),
                sheet.getRange(headerRow+1,startColumn,data.length,data[0].length).setValues(data)
            }
            
        } else if (type = "a") {
            if (removeCsvHeader == true) {
                data.shift()
            }
            if (useAdvanced) {
                writeDataAdv(data, undefined, ssId, "RAW", sheetName)
            } else {
                var lastRow = countRowsByCol ? sheet.getRange(countRowsByCol + "1",countRowsByCol).getValues().length : sheet.getLastRow()
                sheet.getRange(lastRow+1,startColumn,data.length,data[0].length).setValues(data)
            }
        }
    }
    
    /**
     * 
     * 
     * @param {Object} event 
     * @param {String} condSheet 
     * @param {String} condColHeader 
     * @param {any} cond 
     * @param {any} textObj 
     * @param {any} subjObj 
     * @param {any} emailRecColHeader 
     * @param {any} emailObj 
     * @param {any} headerRow 
     * @returns 
     */
    function sendMailOnEdit(event, condSheet, condColHeader, cond, textObj, subjObj, emailRecColHeader, emailObj, headerRow) {
        if (arguments.length < 8) {
            throw Error("Not enough arguments")
        }
        if (arguments.length > 9) {
            throw Error("Too many arguments")
        }
        if (typeof cond === "string") {
            if (typeof condColHeader != "string") {
                throw Error("if Condition is a string Headers must also need to be a string")
            }
            cond = [cond];
            condColHeader = [condColHeader]
        }
        if (Array.isArray(cond)) {
            if (!Array.isArray(condColHeader)) {
                throw Error("if Condition is an array Headers must also need to be an array")
            }
            var checkRun = cond[0];
            var checkRun2 = condColHeader[0]
        }
        if (typeof event != "object") {
            throw Error("No event Object");
        }
        if (typeof condSheet != "string") {
            throw Error("Sheet name must be a string")
        }
        if (typeof textObj === "string") {
            textObj = {end: textObj}
        }
        if (typeof subjObj === "string") {
            subjObj = {end: subjObj}
        }
        if (textObj.hasOwnProperty("header")) {
            if (!Array.isArray(textObj.header)) {
                throw Error("Header must be an Array")
            }
            textObj.header.forEach(function(item) {
                if (!textObj.hasOwnProperty(item)) {
                    throw Error("Object doesn't have property: "+item);
                }
            })
        }
        if (subjObj.hasOwnProperty("header")) {
            if (!Array.isArray(subjObj.header)) {
                throw Error("Header must be an Array")
            }
            subjObj.header.forEach(function(item) {
                if (!subjObj.hasOwnProperty(item)) {
                    throw Error("Object doesn't have property: "+item);
                }
            })
        }
        emailObj = emailObj || {};
        if (typeof emailObj != "object") {
            throw Error("Email Options need to be an object, see GAS documentation")
        }
        headerRow =  headerRow || 1;
        if (typeof headerRow != "number") {
            throw Error("Header row needs to be a number.")
        }
        
        if(event.value != checkRun || event.source.getActiveSheet().getName() != condSheet || event.source.getActiveSheet().getRange(headerRow, event.range.getColumn()).getValue() != checkRun2) {
            return false;
        }
        var patt = /^[A-Z0-9._%+-]{1,64}@(?:[A-Z0-9-]{1,63}\.){1,125}[A-Z]{2,63}$/i;
        var recipient; 
        var ss = event.source;
        var sheet = ss.getActiveSheet();
        var range = sheet.getDataRange();
        var values = range.getValues();
        var headers = values[headerRow-1];
        var i; 
        var text = "";
        var subject = "";
        var row = event.range.getRow()-1;
        var stop = false;
        cond.forEach(function(item, index) {
            if(values[row][headers.indexOf(condColHeader[index])] != item) {
                stop = true;
            }
        });
        if (stop) return false;
        if (textObj.hasOwnProperty("header")) {
            var textDataHeader = textObj.header;
            for (i = 0; i < textDataHeader.length; i++) {
                text += textObj[textDataHeader[i]];
                text += values[row][headers.indexOf(textDataHeader[i])];
            }
        }
        if (textObj.hasOwnProperty("end")) {
            text += textObj.end;
        }
        if (subjObj.hasOwnProperty("header")) {
            var subjDataHeader = subjObj.header;
            for (i = 0; i < subjDataHeader.length; i++) {
                subject += subjObj[subjDataHeader[i]];
                subject += values[row][headers.indexOf(subjDataHeader[i])];
            }
        }
        if (subjObj.hasOwnProperty("end")) {
            subject += subjObj.end;
        }
        if (patt.test(emailRecColHeader)) {
            recipient = emailRecColHeader
        } else {
            recipient = values[row][headers.indexOf(emailRecColHeader)];
        }
        GmailApp.sendEmail(recipient, subject, text, emailObj);
    }

    function moveValues(sourceName,targetName,conObj,headerRows,inputStartCol,targetArr,sourceSs,targetSs) {
        headerRows = headerRows == undefined ? [1,1] : typeof headerRows == "number" ? [headerRows,1] : headerRows;
        inputStartCol = inputStartCol == undefined ? 1 : inputStartCol;
        sourceSs = sourceSs == undefined ? SpreadsheetApp.getActiveSpreadsheet() : SpreadsheetApp.openById(sourceSs)
        targetSs = targetSs == undefined ? SpreadsheetApp.getActiveSpreadsheet() : SpreadsheetApp.openById(targetSs)
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var sourceSheet = sourceSs.getSheetByName(sourceName);
        var targetSheet = targetSs.getSheetByName(targetName);
        var sourceData = sourceSheet.getDataRange().getValues();
        var sourceHeader = sourceData[headerRows[0]-1];
        sourceData.splice(0,headerRows[0]);
        var targetArrType = "whole";
        if (targetArr != undefined) {
            if(targetArr.every(function(item) {return typeof item == "number"})) {
                targetArrType = "number"
            } else if (targetArr.every(function(item) {return typeof item == "string"})) {
                if (targetArr.every(function(item) {return sourceHeader.indexOf(item) > -1})) {
                    targetArrType = "header"
                } else {
                    targetArrType = "column"
                }
            }
        }
        var orKey, andKey;
        var newData = [];
        sourceData.forEach(function(row) {
            var addRowBool = false 
            if (typeof conObj == "function") {
                addRowBool = conObj(row)
            }
            if (Array.isArray(conObj)) {
                addRowBool = conObj.some(function(orObj) {
                    return Object.keys(orObj).every(function(andKey) {
                        return row[sourceHeader.indexOf(andKey)].toString().trim().toUpperCase() == conObj[orKey][andKey].trim().toUpperCase()
                    })
                })
            }
            if (typeof conObj == "object") {
                addRowBool = Object.keys(conObj).every(function(key) {
                    var test = conObj[key]
                    return row[sourceHeader.indexOf(key)].toString().trim().toUpperCase() == conObj[key].trim().toUpperCase()
                })
            }
            if (addRowBool === true) {
                switch (targetArrType) {
                    case "whole":
                        newData.push(row);
                        break;
                    case "number":
                        var tempArr = [];
                        targetArr.forEach(function(num) {
                            tempArr.push(row[num-1]);
                        });
                        newData.push(tempArr)
                        break;
                    case "header":
                        var tempArr = [];
                        targetArr.forEach(function(header) {
                            tempArr.push(row[sourceHeader.indexOf(header)]);
                        });
                        newData.push(tempArr)
                        break;
                    case "column":
                        var tempArr = [];
                        targetArr.forEach(function(letters) {
                            var sum = 0;
                            letters.split("").forEach(function(letter,index) {
                                sum *= 26;
                                sum += (letters.charCodeAt(index) - ("A".charCodeAt(0)-1));  
                            })
                            tempArr.push(row[sum-1]);
                        });
                        newData.push(tempArr)
                        break;
                    default:
                        throw new Error("The format of targetArr doesn't fit the specifications")
                }
            }
        });
        targetSheet.getRange(targetSheet.getLastRow()+1,inputStartCol,newData.length,newData[0].length).setValues(newData);
    }

    function a1ToColNum(char) {
        var num = 0, len = char.length, pos = len;
        while (--pos > -1) {
            num += (char.charCodeAt(pos) - 64) * Math.pow(26, len - 1 - pos);
        }
        return num; 
    }

    function colNumToA1(num) {
        for (var ret = '', a = 1, b = 26; (num -= a) >= 0; a = b, b *= 26) {
                ret = String.fromCharCode(parseInt((num % b) / a) + 65) + ret;
        }
        return ret;
        
    }

    function writeDataAdv(data,ranges,ssId,type,sheet) {
        type = type == undefined ? "RAW" : type
        ssId = ssId == undefined ? SpreadsheetApp.getActiveSpreadsheet().getId() : ssId
        if (!Array.isArray(ranges) && ranges != undefined) {
            var valueRange = Sheets.newValueRange()
            valueRange.values = data
            valueRange.range = ranges
            Sheets.Spreadsheets.Values.update(valueRange,ssId,ranges,{valueInputOption: type})
        } else if (ranges == undefined) {
            var valueRange = Sheets.newRowData()
            valueRange.values = data
            var appendRequest = Sheets.newAppendCellsRequest();
            appendRequest.sheetId = ssId;
            appendRequest.rows = [valueRange];
            Sheets.Spreadsheets.Values.append(valueRange, ssId, sheet, {valueInputOption: type});
        } else {
            requests = []
            ranges.forEach(function(range,idx) {
                var valueRange = Sheets.newValueRange()
                valueRange.range = range
                valueRange.values = data[idx]
                requests.push(valueRange)
            })
            var updateRequest = Sheets.newBatchUpdateValuesRequest()
            updateRequest.data = requests
            updateRequest.valueInputOption = type
            Sheets.Spreadsheets.Values.batchUpdate(updateRequest,ssId)
        }
    }
    
    GASHelper.archive = archive
    GASHelper.compare = listCompare
    GASHelper.importCsv = importCsv
    GASHelper.addDataToSheet = addDataToSheet
    GASHelper.sendMailOnEdit = sendMailOnEdit
    GASHelper.moveValues = moveValues
    GASHelper.permission = permission

    return GASHelper

})