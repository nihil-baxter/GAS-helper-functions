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
