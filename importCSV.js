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

