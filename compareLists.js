


/**
 * 
 * 
 * @param {Object} inputList 
 * @param {Object} existingList 
 * @param {Object} compObj 
 * @returns 
 */
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
            if (inputKeys.indexOf(existingKeys[j]) === -1) {
                removedItems[existingKeys[j]] = JSON.parse(JSON.stringify(existingList[existingKeys[j]]));
                delete existingList[existingKeys[j]];
            }
            j++;
        };
        if (existingKeys[j] === inputKeys[i]) {
            if (compObj && typeof compObj === "object") {
                if (compObj.hasOwnProperty("update") && Array.isArray(compObj["update"]) && compObj["update"].length > 0) {
                    for (k=0;k<compObj["update"].length;k++) {
                        var headerName = compObj["update"][k];
                        existingList[inputKeys[i]][headerName] = inputList[inputKeys[i]][headerName];
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
        return {updated: existingList,
                removed: removedItems,
                added: addedItems
               }
    }
}