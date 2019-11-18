var webAppUrl = '##';
var spreadsheetUrl = '##';
var mainSheetName = 'Micro-Purchase Request';
var pdfSheetName = 'PdfReportTemplate';
var emailTemplatesSheetName = 'Templates';
var emailStatusSheetName = 'Email Status';
var configurationSheetName = 'Configuration';
var approvedValue = 'Yes';
var notApprovedValue = 'No';
var debugMode = false;
var emailDebugPrefix = debugMode ? "MP.DEBUG.3:" : "";
var firstDataRowNum = 3;

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Scripts')
      .addItem('Send D7 Email to PDE', 'sendD7PdeEmail')
      .addToUi();
}

function menuItem1() {
  SpreadsheetApp.getUi(); // Or DocumentApp or FormApp.
     sendD7PdeEmail;
}

function rowToDict(sheet, rownumber){
    var columns = sheet.getRange(2,1,1, sheet.getMaxColumns()).getValues()[0];
    var data = sheet.getDataRange().getValues()[rownumber-1];
    var dict_data = {};
  dict_data["row"] = rownumber;
  dict_data["spreadsheetUrl"] = sheet.getParent().getUrl()+ "#gid=" + sheet.getSheetId();
  dict_data["spreadSheetName"] = sheet.getName();
    for(var keys in columns){
      var key = columns[keys];
      dict_data[key] = data[keys];
    }
    return dict_data;
}

function getDefaultTemplateTokens(ss, row) {
    var sheet = ss.getSheetByName(mainSheetName);
    var DictRowObject = rowToDict(sheet,row);  
  return {
        "row": DictRowObject["row"],
        "spreadsheetUrl": DictRowObject["spreadsheetUrl"],
        "spreadSheetName": DictRowObject["spreadSheetName"],
        "refNum": DictRowObject["Reference #"], /* Form Filled */
        "requestedBy": DictRowObject["Requested By"],
        "requestorSubmissionTime": DictRowObject["Timestamp"],
        "purchaseType": DictRowObject["Type Of Purchase"],
        "description": DictRowObject["Description"],
        "requestedAmount": DictRowObject["Requested Amount"],
        "vendor": DictRowObject["Vendor"],
        "vendorDuns": DictRowObject["DUNs/TIN"],
        "orgCode": DictRowObject["Organization Code"],
        "budgetAct": DictRowObject["Budget Activity"],
        "buildingNumber": DictRowObject["Building Number"],
        "rwaNum": DictRowObject["RWA Number"],
        "projectNum": DictRowObject["Project Number"],
        "workItem": DictRowObject["Work Item"],
        "managerEmail": DictRowObject["Manager Email"], /* Manager */
        "managerName": DictRowObject["Manager"],
        "managerApproval": DictRowObject["Manager Approval"],
        "managerApprovalDate": DictRowObject["Date of Supervisor Approval"], /* Form Filled */
        "approvingOfficialEmail": DictRowObject["Approving Official Email"], /* Approving Official */
        "approvingOfficialName": DictRowObject["Approving Official"],
        "approvingOfficialApproval": DictRowObject["Approving Official Approval"],
        "dateofApprovingOfficialApproval": DictRowObject["Date of Approving Official Approval"], /* Budget Analyst */
        "budgetAnalystEmail": DictRowObject["Budget Analyst Email"],
        "budgetAnalystName": DictRowObject["Fund Certifier"],
        "budgetAnalystApproval": DictRowObject["Budget Analyst Approval"],
        "budgetAnalystApprovalDate": DictRowObject["Date of Budget Analyst Approval"],
        "fiscalYear": DictRowObject["Fiscal Year"],
        "fundCode": DictRowObject["Fund Code"],
        "fc": DictRowObject["FC"],
        "soc": DictRowObject["SOC"],
        "assetNum": DictRowObject["Asset Number"],
        "assetType": DictRowObject["Asset Type"],
        "cardHolderName": DictRowObject["Card Holder Name"],
        "cardHolderEmail": DictRowObject["Card Holder Email"],
        "pegasysUserId": DictRowObject["Pegasys UserID"],
        "inputPegasysBy": DictRowObject["Input into Pegasys by:"],
        "pegasysInputDate": DictRowObject["Pegasys Input Date"],
        "pdn": DictRowObject["PDN #"],
        "dateOrdered": DictRowObject["Date Ordered"],
        "remarks": DictRowObject["Remarks"],
        "generatePDFReport": DictRowObject["Generate PDF Report"],
        "emailAddress": DictRowObject["Email Address"]
    };
}

function createOuput(ss, templateName, t1, t2) {
    var t = fillTemplate(ss, templateName, t1, t2);
    return HtmlService.createHtmlOutput(t.htmlBody);
}

function getValForColumnName(ss, columnName, row) {
    var colNum = findColumnNumberForColumnName(ss, columnName);
    if (colNum == null) return null;
    var sheet = ss.getSheetByName(mainSheetName);
    return sheet.getRange(row, colNum).getValue();
}


function setValForColumnName(ss, columnName, row, val) {
    var colNum = findColumnNumberForColumnName(ss, columnName);
    if (colNum == null) return null;
    var sheet = ss.getSheetByName(mainSheetName);
    sheet.getRange(row, colNum).setValue(val);
}

var columnNumberByColumnNameBySheetId = {};

function findColumnNumberForColumnName(ss, columnName) {
    if (columnName == null || columnName == "") return null;
    var sheetId = ss.getId();
    var columnNumberByColumnName = columnNumberByColumnNameBySheetId[sheetId];
    columnName = columnName.trim().toLowerCase();
    if (columnNumberByColumnName==null)
    {
      columnNumberByColumnName = {};

      var sheet = ss.getSheetByName(mainSheetName);
      var columnNames = sheet.getRange("A2:AX2").getValues();
      
      for (var x in columnNames) {
        for (var y in columnNames[x]) {
          var sample = columnNames[x][y];
          sample = ('' + sample).trim().toLowerCase();
          if (sample.length>0)
          {
          columnNumberByColumnName[sample] = (y * 1) + 1;
          }
        }
      }

      columnNumberByColumnNameBySheetId[sheetId]=columnNumberByColumnName;
    }

    var colNum = columnNumberByColumnName[columnName];
    return colNum;
}

//http://www.andrewroberts.net/2017/03/apps-script-create-pdf-multi-sheet-google-sheet/
function convertSpreadsheetToPdf(spreadsheet, sheetId) {
    var spreadsheetId = spreadsheet.getId();
    var url_base = spreadsheet.getUrl().replace(/edit$/, '');

    var url_ext = 'export?exportFormat=pdf&format=pdf' //export as pdf

    // Print either the entire Spreadsheet or the specified sheet if optSheetId is provided
    + (sheetId ? ('&gid=' + sheetId) : ('&id=' + spreadsheetId))
    // following parameters are optional...
    + '&size=letter' // paper size
    + '&portrait=true' // orientation, false for landscape
    + '&fitw=true' // fit to width, false for actual size
    + '&sheetnames=false&printtitle=false&pagenumbers=false' //hide optional headers and footers
    + '&gridlines=false' // hide gridlines
    + '&fzr=false'; // do not repeat row headers (frozen rows) on each page

    var options = {
        headers: {
            'Authorization': 'Bearer ' + ScriptApp.getOAuthToken(),
        },
      muteHttpExceptions : true
    }

    var response = UrlFetchApp.fetch(url_base + url_ext, options);
  
    return response.getContent();
}

function createAndSendSheetAsPdf(ss, rowNumber) {
    var tempSheetName = "pdf" + Math.floor(Math.random() * 10000);

    var sheet = ss.getSheetByName(mainSheetName);
    var refNum = getValForColumnName(ss, 'Reference #', rowNumber);
    var pdn = getValForColumnName(ss, 'PDN #', rowNumber);

    sheet = ss.getSheetByName(pdfSheetName);
    sheet.getRange("O4").setValue(rowNumber);
    sheet.copyTo(ss);
    sheet = ss.getSheetByName("Copy of " + pdfSheetName);
    sheet.activate();
    ss.renameActiveSheet(tempSheetName);
    sheet = ss.getSheetByName(tempSheetName);
    sheet.showSheet();

    var pdf = convertSpreadsheetToPdf(ss, sheet.getSheetId());
    var pdfRecipient = Session.getActiveUser();


    var attach = {
      fileName: 'Micro-Purchase Report for PDN: ' + pdn + ' Ref Number: ' + refNum,
        content: pdf,
        mimeType: 'application/pdf'
    };

    var t = fillTemplate(ss,
        'PDF Report',
        getDefaultTemplateTokens(ss, rowNumber), {});
    MailApp.sendEmail({
        to: pdfRecipient,
        subject: t.subject,
        body: t.body,
        htmlBody: t.htmlBody,
        attachments: attach,
        noReply: true,
    });

    ss.deleteSheet(sheet);

    return createOuput(ss, 'PDF Created Message', getDefaultTemplateTokens(ss, rowNumber), {});

}

function createAndSendSheetAsPdfForD7(ss, rowNumber) {
    var tempSheetName = "pdf" + Math.floor(Math.random() * 10000);

    var sheet = ss.getSheetByName(mainSheetName);
    var refNum = getValForColumnName(ss, 'Reference #', rowNumber);
  
    sheet = ss.getSheetByName(pdfSheetName);
    sheet.getRange("O4").setValue(rowNumber);
    sheet.copyTo(ss);
    sheet = ss.getSheetByName("Copy of " + pdfSheetName);
    sheet.activate();
    ss.renameActiveSheet(tempSheetName);
    sheet = ss.getSheetByName(tempSheetName);
    sheet.showSheet();

    var pdf = convertSpreadsheetToPdf(ss, sheet.getSheetId());
    var pdfRecipient = Session.getActiveUser();


    var attach = {
      fileName: 'Micro-Purchase Report for Ref Number: ' + refNum + '.pdf',
        content: pdf,
        mimeType: 'application/pdf'
    };

    ss.deleteSheet(sheet);

    return attach;
}
function doGet(request) {
    var ssid = request.parameters.zssid;
    var ss = SpreadsheetApp.openById(ssid);
    if (request.parameters.zcmd == 'pdf') {
        return createAndSendSheetAsPdf(ss, request.parameters.zr);
    } else {
        var row = request.parameters.zr;
        var col = request.parameters.zc;
        var value = request.parameters.zt;
        var level = request.parameters.zl;
        var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(mainSheetName);
        var sheetname = sheet.getSheetName();
        var approverNameCol;

        var columnName = sheet.getRange(2, col).getValue();

        switch (columnName) {
            case 'Manager Approval':
                level = 'Manager';
                approverNameCol = 'Manager';
                break;
            case 'Approving Official Approval':
                level = 'Approving Official';
                approverNameCol = 'Approving Official';
                break;
            case 'Budget Analyst Approval':
                level = 'Budget Analyst';
                approverNameCol = 'Fund Certifier';
                break;
        }
        if (sheet.getRange(row, col).getValue() != "") {
            return createOuput(ss, 'Already Submitted Message', getDefaultTemplateTokens(ss, row));
        }
        sheet.getRange(row, col).setValue(value);
        dateTime(sheet, row, (col * 1 + 1));


        var currentApprover = getValForColumnName(ss, approverNameCol, row);
        if (approverNameCol != null && (currentApprover == null || currentApprover.trim() == "")) {
            var u = Session.getActiveUser();
            setValForColumnName(ss, approverNameCol, row, u.getEmail());
        }


        if (value == approvedValue) {
            NextLevel(sheet, level, row)
            return createOuput(ss, 'Thank You For Submitting Message (Approval)', getDefaultTemplateTokens(ss, row));
        } else if (value == notApprovedValue) {
            var timeCol = findColumnNumberForColumnName(ss, 'Requested By');
            var initialSubmissionTimeCol = findColumnNumberForColumnName(ss, 'Requested By');
            var requestorSubmissionTime = sheet.getRange(row, (timeCol * 1 + 1)).getValue();
            var approverName = sheet.getRange(row, findColumnNumberForColumnName(ss, approverNameCol)).getValue();
            var recipient = getValForColumnName(ss, 'Email Address', row);

            var approverName = getValForColumnName(ss, approverNameCol, row);
            var t = fillTemplate(ss,
                'Disapproval',
                getDefaultTemplateTokens(ss, row), {
                    "level": level,
                    "requestorSubmissionTime": requestorSubmissionTime,
                    "approverName": approverName,
                });
            sendAndTrackEmail(ss, row, {
                to: recipient,
                subject: emailDebugPrefix + t.subject,
                body: t.body,
                htmlBody: t.htmlBody,
                noReply: true,
            });

        }
        return createOuput(ss, 'Thank You For Submitting Message (Disapproval)', getDefaultTemplateTokens(ss, row));

    }
}

function sendAndTrackEmail(ss, row, message) {
    var refNum = getValForColumnName(ss, 'Reference #', row);
    var errorMessage = null;
    var success;
    try {
        MailApp.sendEmail(message);
        success = true;
    } catch (error) {
        errorMessage = error;
        success = false;
    }

    var sheet = ss.getSheetByName(emailStatusSheetName);
  
    sheet.appendRow([refNum,new Date(),message.to,success,errorMessage]);
}

function getBudgetAnalystEmail(ss, row) {
    var baCode = ('' + getValForColumnName(ss, 'Budget Activity', row)).trim().toLowerCase();
    var sheet = ss.getSheetByName(configurationSheetName);
    var vals = sheet.getRange('BACodeEmailMap').getValues();
    for (var y = 0; y < vals.length; ++y) {
        if (('' + vals[y][0]).trim().toLowerCase() == baCode) {
            return vals[y][1];
        }
    }
    return sheet.getRange('BACodeDefaultEmail').getValue();
}

function right(s, len) {
    s = "" + s;
    var l = s.length;
    return s.substr(l - len);
}

function currentFY(){
    var d = new Date();
    var m = d.getMonth() + 1; // in JS, months are 1 based
    var y = d.getFullYear();
    if (m>=10) {
        ++y;
    }
    return right(y,2);
}

function createReferenceNumber(ss, sheet, row){
	var zoneNumber = sheet.getRange('Zone').getValue();
	var refNum = (""+sheet.getRange('SheetUID').getValue()+"000000").substr(0,6).toLowerCase() +
        "-" +
        currentFY() +  
        "-";
    if (zoneNumber == 1){
		refNum += right("0000"+(row),4);
	}
	else {
		var col = findColumnNumberForColumnName(ss, 'Reference #');
		var colLetter = returnColumnToLetter(col);
		var sel = colLetter+firstDataRowNum+":"+colLetter+row;
		var range = sheet.getRange(sel).getValues();
		var i = 0;
		for (var y=0;y<range.length-1;++y){
			var val = range[y][0];
		  if (val.substr(0,refNum.length)==refNum){
		  ++i;
		  }
		}
		refNum += right("0000"+(i+1),4);
	}
    return refNum;
}

/*function createReferenceNumber(ss, sheet, row){
    var refNum = (""+sheet.getRange('SheetUID').getValue()+"000000").substr(0,6).toLowerCase() +
        "-" +
        currentFY() +  
        "-";
    var col = findColumnNumberForColumnName(ss, 'Reference #');
    var colLetter = returnColumnToLetter(col);
    var sel = colLetter+firstDataRowNum+":"+colLetter+row;
    var range = sheet.getRange(sel).getValues();
    var i = 0;
    for (var y=0;y<range.length-1;++y){
        var val = range[y][0];
      if (val.substr(0,refNum.length)==refNum){
      ++i;
      }
    }
    refNum += right("0000"+(i+1),4);
    return refNum;
}*/

function onSubmit(e) {
    var row = e.range.getRow();
    var sheet = e.range.getSheet();
    var ss = sheet.getParent();

    setValForColumnName(ss, 'Budget Analyst Email', row, getBudgetAnalystEmail(ss, row));
    setValForColumnName(ss, 'Reference #', row, createReferenceNumber(ss, sheet, row));

    var col = findColumnNumberForColumnName(ss, 'Generate PDF Report');
    sheet.getRange(row, col).setFormula('=HYPERLINK("' + webAppUrl + '?zssid=' + ss.getId() + '&zr="&row()&"&zcmd=pdf","PDF Report")');

    var level;

    var startingLevelValue = sheet.getRange('SkipManagerApproval').getValue();
    var D7Val = IsD7(ss,row);
    if (startingLevelValue == 0 || D7Val == true) {
        level = 'Beginning';
    } else {
        level = 'Manager';
    }
    NextLevel(sheet, level, row);

}

function dateTime(sheet, row, col) {
    sheet.getRange(row, col).setValue(new Date()).setNumberFormat('mm/dd/yyyy hh:mm:ss');
}


function approvalTrigger(e) {
    var sheet = e.source.getActiveSheet();
    var ss = sheet.getParent();
    var row = sheet.getActiveRange().getRow();
    var col = sheet.getActiveRange().getColumn();
    var columnName = sheet.getRange(2, col).getValue();
    var level, approverNameCol;
    var value = null;

    switch (columnName) {
        case 'Manager Approval':
            level = 'Manager';
            approverNameCol = 'Manager';
            value = getValForColumnName(ss, columnName, row);
            break;
        case 'Approving Official Approval':
            level = 'Approving Official';
            approverNameCol = 'Approving Official';
            value = getValForColumnName(ss, columnName, row);
            break;
        case 'Budget Analyst Approval':
            level = 'Budget Analyst';
            approverNameCol = 'Fund Certifier';
            value = getValForColumnName(ss, columnName, row);
            break;
        case 'PDN #':
            var val = sheet.getActiveCell().getValue();
            if (("" + val).length > 1) {
                NextLevel(sheet, 'Card Holder Name', row);
            }
            return;
        default:
            return;
    }

    var currentApprover = getValForColumnName(ss, approverNameCol, row);
    if (approverNameCol != null && (currentApprover == null || currentApprover.trim() == "")) {
        var u = Session.getActiveUser();
        setValForColumnName(ss, approverNameCol, row, u.getEmail());
    }

    if (value == approvedValue) {
        NextLevel(sheet, level, row);
        dateTime(sheet, row, (col * 1 + 1));
    } else if (value == notApprovedValue) {
        dateTime(sheet, row, (col * 1 + 1));
        var approverName = getValForColumnName(ss, approverNameCol, row);
        var timeCol = findColumnNumberForColumnName(ss, 'Requested By');
        var requestorSubmissionTime = sheet.getRange(row, (timeCol * 1 + 1)).getValue();
        var recipient = getValForColumnName(ss, 'Email Address', row);

        var t = fillTemplate(ss,
            'Disapproval',
            getDefaultTemplateTokens(ss, row), {
                "level": level,
                "requestorSubmissionTime": requestorSubmissionTime,
                "approverName": approverName,
            });
        sendAndTrackEmail(ss, row, {
            to: recipient,
            subject: emailDebugPrefix + t.subject,
            body: t.body,
            htmlBody: t.htmlBody,
            noReply: true,
        });

    }
}

function NextLevel(sheet, level, row) {
    var nextLevel, nextApprovalColumnName, nextEmailColumnName, nextEmailColumnName2, nextEmailColumnName3;
    var D7Val = IsD7(sheet.getParent(), row);
  
  var SkipAO = sheet.getRange('SkipApprovingOfficialApproval').getValue();
    //forcing to skip Approving Official Level after Manager level if the row is a D7.
  if (level == 'Manager' && D7Val == true && SkipAO == 1){
    level = 'Approving Official';
  }
  
    switch (level) {
        case 'Beginning':
            nextLevel = 'Manager';
            nextApprovalColumnName = 'Manager Approval';
            nextEmailColumnName = 'Manager Email';
            nextEmailColumnName2 = '';
            nextEmailColumnName3 = '';
            break;
        case 'Manager':
            nextLevel = 'Approving Official';
            nextApprovalColumnName = 'Approving Official Approval';
            nextEmailColumnName = 'Approving Official Email';
            nextEmailColumnName2 = '';
            nextEmailColumnName3 = '';
            break;
        case 'Approving Official':
            nextLevel = 'Budget Analyst';
            nextApprovalColumnName = 'Budget Analyst Approval';
            nextEmailColumnName = 'Budget Analyst Email';
            nextEmailColumnName2 = '';
            nextEmailColumnName3 = '';
            break;
        case 'Budget Analyst':
            nextLevel = 'Card Holder Name';
            nextApprovalColumnName = 'Card Holder Name';
            nextEmailColumnName = 'Card Holder Email';
            nextEmailColumnName2 = '';
            nextEmailColumnName3 = '';
            break;
        case 'Card Holder Name':
            nextLevel = 'Requested By';
            nextApprovalColumnName = null;
            nextEmailColumnName = 'Email Address';
            nextEmailColumnName2 = 'Manager Email';
            if(D7Val == true && SkipAO == 1)
            {
              nextEmailColumnName3 = '';
            }
            else
            {
              nextEmailColumnName3 = 'Approving Official Email';
            }
            break;
        default:
            return;
    }

    var ss = sheet.getParent();

    var recipient = sheet.getRange(row, findColumnNumberForColumnName(ss, nextEmailColumnName)).getValue();
    
    if (nextEmailColumnName2 != '') {
      recipient = recipient + ', ' + sheet.getRange(row, findColumnNumberForColumnName(ss, nextEmailColumnName2)).getValue();
    }
    
    if (nextEmailColumnName3 != '') {
      recipient = recipient + ', ' + sheet.getRange(row, findColumnNumberForColumnName(ss, nextEmailColumnName3)).getValue();
    }

    var timeCol = findColumnNumberForColumnName(ss, 'Requested By');
    var requestorSubmissionTime = sheet.getRange(row, (timeCol * 1 + 1)).getValue();

    var emailTemplateName;

    if (level == 'Beginning' || level == 'Manager') {
        var approvalCol = findColumnNumberForColumnName(ss, nextApprovalColumnName);
        var columnLetter = returnColumnToLetter(approvalCol);
        sendApprovalEmail(ss,
            row,
            approvalCol,
            level,
            createUrls(ss, row, approvalCol, nextLevel, approvedValue),
            createUrls(ss, row, approvalCol, nextLevel, notApprovedValue),
            recipient);
        return;
    } else if (level == 'Approving Official') {
        emailTemplateName = 'Approving Official';
        var approvalCol = findColumnNumberForColumnName(ss, nextApprovalColumnName);
        var columnLetter = returnColumnToLetter(approvalCol);
    } else if (level == 'Budget Analyst') {
        emailTemplateName = 'Budget Analyst';
        var approvalCol = findColumnNumberForColumnName(ss, nextApprovalColumnName);
        var columnLetter = returnColumnToLetter(approvalCol);
    } else if (level == 'Card Holder Name') {
        emailTemplateName = 'Card Holder';
    } else {
        return;
    }

    var e = fillTemplate(ss,
        emailTemplateName,
        getDefaultTemplateTokens(ss, row), {
            "requestorSubmissionTime": requestorSubmissionTime,
            "columnLetter": columnLetter,
        });


    if (debugMode) {
        e.htmlBody = e.htmlBody + '<hr/><div>' + JSON.stringify({
            sheet: sheet,
            level: level,
            row: row,
            nextLevel: nextLevel,
            nextApprovalColumnName: nextApprovalColumnName,
            nextEmailColumnName: nextEmailColumnName,
            approvalCol: approvalCol,
        }) + '</div>';
    }

    sendAndTrackEmail(ss, row, {
        to: recipient,
        subject: emailDebugPrefix + e.subject,
        body: e.body,
        htmlBody: e.htmlBody,
        noReply: true,
    });
}

function replaceAll(original, search, replacement) {
    return original.replace(new RegExp(escapeRegExp(search), 'g'), replacement);
}

function escapeRegExp(str) {
    return str.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"); // $& means the whole matched string
}

function getTemplate(ss, templateName) {
    var sheet = ss.getSheetByName(emailTemplatesSheetName);
    var cells = sheet.getRange("A5:E50").getValues();

    templateName = templateName.trim().toLowerCase();

    for (var y in cells) {
        var v = cells[y][0];
        v = ('' + v).trim().toLowerCase();
        if (v == templateName) {
            return {
                templateName: cells[y][0],
                templateType: cells[y][1],
                subject: cells[y][2],
                body: cells[y][3],
                htmlBody: cells[y][4]
            };
        }
    }
    return null;
}

function assign(o1, o2) {
    if (!(o2 == null)) {
        for (x in o2) {
            o1[x] = o2[x];
        }
    }
    return o1;
}

function fillTemplate(ss, emailTemplateName, t1, t2) {
    var tokens = assign(t1, t2);
    var e = getTemplate(ss, emailTemplateName);
    for (x in tokens) {
        var v = tokens[x];
        var t = "{" + x + "}";
        e.subject = replaceAll(e.subject, t, v);
        e.body = replaceAll(e.body, t, v);
        e.htmlBody = replaceAll(e.htmlBody, t, v);
    }
    return e;
}

function createUrls(ss, r, c, level, type) {
    var url = webAppUrl + '?';
    url += 'zr=' + r;
    url += '&zc=' + c;
    url += '&zl=' + level;
    url += '&zt=' + type;
    url += '&zssid=' + ss.getId();
    return url;
}

function returnColumnToLetter(column) {
    var temp, letter = '';
    while (column > 0) {
        temp = (column - 1) % 26;
        letter = String.fromCharCode(temp + 65) + letter;
        column = (column - temp - 1) / 26;
    }
    return letter;
}

function sendApprovalEmail(ss, row, col, level, url1, url2, recipient) {
    var columnLetter = returnColumnToLetter(col);

    var emailTemplateName;

    if (level == 'Beginning') {
        emailTemplateName = 'Beginning';
    } else if (level == 'Manager') {
        emailTemplateName = 'Manager';
    }

    var e = fillTemplate(ss,
        emailTemplateName,
        getDefaultTemplateTokens(ss, row), {
            "columnLetter": columnLetter,
            "url1": url1,
            "url2": url2,
        });
    sendAndTrackEmail(ss, row, {
        to: recipient,
        subject: emailDebugPrefix + e.subject,
        body: e.body,
        htmlBody: e.htmlBody,
        noReply: true,
        });
    }


function IsD7(ss, row) {
    var DocumentType = getValForColumnName(ss, 'Document Type', row);
    if (DocumentType == 'D7'){
      return true;
    }
  else {
    return false;
  }
}

function D7RunType(sheet) {
  var D7RunType = sheet.getRange('SendD7EmailRunType').getValue();
  
  if (D7RunType == 'Sandbox') {
    var ToEmail = sheet.getRange('PDESandboxEmail').getValue();
    return ToEmail;
  }
  if (D7RunType == 'UAT') {
    var ToEmail = sheet.getRange('PDEUatEmail').getValue();
    return ToEmail;
  }
  if (D7RunType == 'Production') {
    var ToEmail = sheet.getRange('PDEProductionEmail').getValue();
    return ToEmail;
  }  
}

function sendD7PdeEmail(){
  var MPName = SpreadsheetApp.getActiveSpreadsheet().getName(); //Get name of the Google Sheet
  var d7sheetName = 'Micro-Purchase Request'; //Active sheet name
  var ssObject = SpreadsheetApp.openByUrl(spreadsheetUrl); //opens spreadsheet
  var lastrow = ssObject.getSheetByName(d7sheetName).getLastRow();	
  var ToEmailAddress = D7RunType(ssObject); //get production email address
  var ccRecipients = ssObject.getRange('ErrorRecipients').getValue();
  var fiscalyear = ssObject.getRange('FiscalYear').getValue();

  var ssRange = ssObject.getSheetByName(d7sheetName).getRange("AL2:AO"+ lastrow).getValues(); //opens micro-purchase sheet and gets value for populated range
  var PdfCounter = 0;
  var ProcessingDate = new Date(ssObject.getRange('SendD7EmailProcessingDate').getValue()); //from the configuration file
  //var errorlist = [];

//currently pulling 3 columns; need to expand to more columns
  try {
    for (var i = ssRange.length-1; i>1; i--){
      var iValue = i;
      var rowNum = i+2;
      var rowItem = ssRange[i]; //puts the row item for every row in micropurchase sheet
      var docType = rowItem[0];
      if(docType != "D7") {
        continue
      }
      var inputIntoPegasys = rowItem[3];
      var rowObject = getDefaultTemplateTokens(ssObject,rowNum,d7sheetName);
      var D7Timestamp = new Date(rowObject["requestorSubmissionTime"]); //from the 

      // && D7Timestamp >= ProcessingDate && 
      if (i != 0  &&  (inputIntoPegasys == "" || inputIntoPegasys == "Data Validation Error") && rowObject["budgetAnalystApproval"] == "Yes" && rowObject["budgetAnalystEmail"] != "" && rowObject["budgetAnalystName"] != ""){
        var col = findColumnNumberForColumnName(ssObject, 'PDN #'); //for the url
        var colLetter = returnColumnToLetter(col);
        
        var remarksmsg = "";     
        var errorFlag = false;  
        var FY = rowObject["fiscalYear"];
        var priorremark = rowObject["remarks"];
          
          if(rowObject["refNum"] ==  "" )  
          
            {  remarksmsg =  remarksmsg.concat(" Reference Num cannot be blank. ");
             setValForColumnName(ssObject, "Remarks", rowNum, remarksmsg);   errorFlag=true;}
       
          if(rowObject["description"] ==  "" )   
          
            {  remarksmsg =  remarksmsg.concat(" Description cannot be blank. ");
            setValForColumnName(ssObject, "Remarks", rowNum, remarksmsg);  errorFlag=true;}
        
          if(rowObject["vendor"] ==  "" )  
            {  remarksmsg =  remarksmsg.concat(" Vendor cannot be blank. ");
            setValForColumnName(ssObject, "Remarks", rowNum, remarksmsg);  errorFlag=true;}
          
        //This needs to be updated per fiscal year
          if(parseInt(FY,10) != fiscalyear)   
          {  remarksmsg =  remarksmsg.concat(" Fiscal Year must be a number and in the current fiscal year. ");
          setValForColumnName(ssObject, "Remarks", rowNum, remarksmsg);  errorFlag=true;} 
          
          if( FY.toString().length != 4)  
            {  remarksmsg =  remarksmsg.concat(" Fiscal Year must be 4 digits long. ");
            setValForColumnName(ssObject, "Remarks", rowNum, remarksmsg);  errorFlag=true;}
            
          if(rowObject["fundCode"] ==  "" )  
            {  remarksmsg =  remarksmsg.concat("Fund Code cannot be blank. ");
            setValForColumnName(ssObject, "Remarks", rowNum, remarksmsg);  errorFlag=true;}
            
          if(rowObject["orgCode"] ==  "" )   
            {  remarksmsg =  remarksmsg.concat(" Org Code cannot be blank. ");
              setValForColumnName(ssObject, "Remarks", rowNum, remarksmsg);  errorFlag=true;}
              
          if(rowObject["budgetAct"] ==  "" )   
          { remarksmsg =  remarksmsg.concat(" Budget Activity cannot be blank. ");
            setValForColumnName(ssObject, "Remarks", rowNum, remarksmsg);  errorFlag=true;}
            
          if(rowObject["soc"] ==  "" )   
            { remarksmsg =  remarksmsg.concat(" SOC cannot be blank.");  
              setValForColumnName(ssObject, "Remarks", rowNum, remarksmsg);  errorFlag=true;}
              
          if(rowObject["fc"] ==  "" )   
            {  remarksmsg =  remarksmsg.concat(" FC cannot be blank. ");
              setValForColumnName(ssObject, "Remarks", rowNum, remarksmsg);  errorFlag=true;}
            
          if(rowObject["requestedAmount"] ==  "" || rowObject["requestedAmount"] < 0) 
          {  remarksmsg =  remarksmsg.concat(" Requested Amount cannot be blank or a negative number. ");
          setValForColumnName(ssObject, "Remarks", rowNum, remarksmsg);  errorFlag=true;}
          
          if (rowObject["budgetAct"] == "PG80" && rowObject["rwaNum"] == "") 
          { remarksmsg =  remarksmsg.concat(" RWA Number cannot be blank when Budget Activity is PG80. ");
          setValForColumnName(ssObject, "Remarks", rowNum, remarksmsg); errorFlag=true;}
          
          rowObject = getDefaultTemplateTokens(ssObject,rowNum,d7sheetName);
                       
        if(errorFlag)
        { 
          setValForColumnName(ssObject, "Input into Pegasys by:", rowNum, "Data Validation Error");
          if(remarksmsg != priorremark){
            var emailSubject2 = MPName + " Error Notice";
            var emailBody2 = '<div><div style3D"padding:15px"><p>Please review the following entries that had errors.' +' <br><br></p><table><tbody>' +
              '<tr><td><b>Pegasys Document Type: </b></td><td>D7</td></tr>' +
                '<tr><td><b>Reference Num: </b></td><td><a href="'+spreadsheetUrl+'">'+rowObject["refNum"]+'</a></td></tr>' +
                  '<tr><td><b>Remarks: </b></td><td>'+rowObject["remarks"]+'</td></tr>' + 
                    '<tr><td><b>Requested By: </b></td><td>'+rowObject["requestedBy"]+'</td></tr>' + 
                      '<tr><td><b>Fund Certifier: </b></td><td>'+rowObject["budgetAnalystName"]+'</td></tr>';
            
            GmailApp.sendEmail(rowObject["budgetAnalystName"]+","+ccRecipients , emailSubject2, emailBody2,
                               { name: 'Automatic Emailer Script', 
                               htmlBody: emailBody2} );
          }
        }
        else {
        var pdf = createAndSendSheetAsPdfForD7(ssObject, rowNum);
        PdfCounter = PdfCounter + 1;
        if(PdfCounter > 4) {
          throw new Error("Too much pdfs are being created in a short amount of time. Please wait a while");
        }
        
        var emailSubject = "D7 transfer from Micropurchase Tool: Reference Number - " + rowObject["refNum"];
        var emailBody = '<div><div style3D"padding:15px"><p>Please process this D7 from the micropurchase sheet: '+ spreadsheetUrl + 'range=' + colLetter + rowObject["row"] + '. <br><br></p><table><tbody>' +
        '<tr><td><b>Pegasys Document Type: </b></td><td>D7</td></tr>' +
        '<tr><td><b>Award Title: </b></td><td>'+rowObject["description"]+'</td></tr>' +
        '<tr><td><b>Reference Num: </b></td><td>'+rowObject["refNum"]+'</td></tr>' +
        '<tr><td><b>Owner: </b></td><td>'+rowObject["requestedBy"]+'</td></tr>' + 
        '<tr><td><b>Requested By: </b></td><td>'+rowObject["requestedBy"]+'</td></tr>' +
        '<tr><td><b>Budget Analyst Email: </b></td><td>'+rowObject["budgetAnalystEmail"]+'</td></tr>' +
        '<tr><td><b>Fund Certifier: </b></td><td>'+rowObject["budgetAnalystName"]+'</td></tr>' +          
        '<tr><td><b>Obligated Value: </b></td><td>'+rowObject["requestedAmount"]+'</td></tr>' +
        '<tr><td><b>Description: </b></td><td>'+rowObject["description"]+'</td></tr>' +
        '<tr><td><b>Vendor: </b></td><td>'+rowObject["vendorDuns"].slice(1)+'</td></tr>' +
        '<tr><td><b>Award Mod Form: </b></td><td>SF33</td></tr>' +
        '<tr><td><b>Invoice Number: </b></td><td>[no value yet]</td></tr>' +
        '<tr><td><b>Log Date: </b></td><td>[no value yet]</td></tr>' +          
        '<tr><td><b>Pegasys User ID: </b></td><td>'+rowObject["pegasysUserId"]+'</td></tr>' +
        '<tr><td><b>Fiscal Year: </b></td><td>'+rowObject["fiscalYear"]+'</td></tr>' +
        '<tr><td><b>Region: </b></td><td>'+rowObject["orgCode"].substr(1,2)*1+'</td></tr>' +
        '<tr><td><b>Fund Code: </b></td><td>'+rowObject["fundCode"]+'</td></tr>' +      
        '<tr><td><b>Org Code: </b></td><td>'+rowObject["orgCode"]+'</td></tr>' +
        '<tr><td><b>Budget Activity: </b></td><td>'+rowObject["budgetAct"]+'</td></tr>' +
        '<tr><td><b>SOC: </b></td><td>'+rowObject["soc"]+'</td></tr>' +
        '<tr><td><b>FC: </b></td><td>'+rowObject["fc"]+'</td></tr>' +
        '<tr><td><b>Building Number: </b></td><td>'+rowObject["buildingNumber"]+'</td></tr>' +
        '<tr><td><b>RWA Number: </b></td><td>'+rowObject["rwaNum"]+'</td></tr>' +
        '<tr><td><b>Cardholder: </b></td><td>'+rowObject["cardHolderName"]+'</td></tr>' +
        '<tr><td><b>Requested By Email: </b></td><td>'+rowObject["emailAddress"]+'</td></tr>' +
        '<tr><td><b>Project Number: </b></td><td>'+rowObject["projectNum"]+'</td></tr>' +
        '<tr><td><b>Work Item: </b></td><td>'+rowObject["workItem"]+'</td></tr>' +
        '<tr><td><b>Asset Number: </b></td><td>'+rowObject["assetNum"]+'</td></tr>' +
        '<tr><td><b>Asset Type: </b></td><td>'+rowObject["assetType"]+'</td></tr>';
      
        GmailApp.sendEmail(ToEmailAddress, emailSubject, emailBody,
                           {
                             attachments: pdf,
                             name: 'Automatic Emailer Script',
                             htmlBody: emailBody
                           }
                          );
        var d = new Date();
        var datetimeString = Utilities.formatDate(d,'America/New_York', 'MM/dd/yyyy HH:mm:ss'); //d.toLocaleTimeString();        
        
        setValForColumnName(ssObject, "Input into Pegasys by:",rowNum,"D7 Robot 1"); //rowObject["inputPegasysBy"]);
        setValForColumnName(ssObject, "Pegasys Input Date", rowNum, datetimeString); //rowObject["pegasysInputDate"]);
        setValForColumnName(ssObject, "Remarks", rowNum, "");
        }
      } //bracket for if statement
      } //bracket for loop
      } //bracket for try catch
  
      catch (e) {
        throw new Error("An error occurred during processing the MP sheet. " + e.message);
      }

} 

/*
      finally{    
        if(errorlist.length != 0) {
        var emailSubject3 = MPName+" Error Summary";
        var emailBody3 = '<div><div style3D"padding:15px"><p>Please review the following entries that had errors.' +' <br><br></p><table><tbody>' +
          '<tr><td><b>Pegasys Document Type, Reference Num, Remarks, Requested By, Fund Certifier</b></td></tr>';
          while(0 != errorlist.length) {
            var tempArray = errorlist.pop();
            var tempArrayRest = tempArray.slice(1,tempArray.length);
            emailBody3 = emailBody3+'<tr><td><a href="'+spreadsheetUrl+'">'+tempArray[0]+'</a>,'+tempArrayRest.join()+'</td></tr>';
          }
          GmailApp.sendEmail(ccAaron+","+ccLynnie,  emailSubject3, emailBody3, {name: 'Automatic Emailer Script', htmlBody: emailBody3}); 
       }
      }
*/


/*--------------------------TESTS------------------------------*/
function openTestSpreadsheet() {
    return SpreadsheetApp.openByUrl('##');
}

function assert(test, msg){
  if (!test){
    msg = msg==null ? test : msg;
    throw msg;
  }
}

function assertEquals(expected, actual){
  assert(expected==actual, "Expected=["+expected+"] != Actual=["+actual+"]");
}

function testCurrentFY(){
  var fy = currentFY();
  assertEquals(18, fy);
}

function testCreateReferenceNumber(){
    var ss = openTestSpreadsheet();
    var sheet = ss.getSheetByName(mainSheetName);
    var refNum = createReferenceNumber(ss, sheet, 3);
    assertEquals(14, refNum.length);
    assertEquals(1,right(refNum,4));
    refNum = createReferenceNumber(ss, sheet, 15);
    assertEquals(14, refNum.length);
}

function testNextLevel() {
    var ss = openTestSpreadsheet();
    var sheet = ss.getActiveSheet();
    var row = 111;
    var level = 'Beginning';
    NextLevel(sheet, level, row);
}

function testOnSubmit() {
    var ss = openTestSpreadsheet();
    var sheet = ss.getSheetByName(mainSheetName);
    var r = sheet.getRange("A111");
    var e = {
        range: r
    }
    onSubmit(e);
}

function testApprovalTrigger() {
    var ss = openTestSpreadsheet();
    var sheet = ss.getSheetByName(mainSheetName);
    var r = sheet.getRange("Q113");
    ss.setActiveRange(r);
    approvalTrigger({
        user: Session.getActiveUser().getEmail(),
        source: ss,
        range: r,
        value: r.getValue(),
        authMode: "LIMITED"
    });
}

function testDoGet() {
    var ss = openTestSpreadsheet();
    var requestMock = {
        parameters: {
            zr: 55,
            zc: 17,
            zl: "Manager",
            zt: "Yes",
            zcmd: "No",
            zssid: ss.getId()
        }
    };
    doGet(requestMock);
}

function testCurrentUser() {
    var u = Session.getActiveUser();
    var myEmail = u.getEmail();
    u = Session.getEffectiveUser();
    myEmail = u.getEmail();
    var props = PropertiesService.getUserProperties();
    var ss = openTestSpreadsheet();
}

function testEmail() {
    var ss = openTestSpreadsheet();
    sendAndTrackEmail(ss, 97, {
        to: 'ashleig.lesnaigsa.gov',
        subject: 'Invalid email',
        body: 'This is a test for a bad email',
        noReply: true,
    });

    sendAndTrackEmail(ss, 97, {
        to: 'jonathan.chen@gsa.gov',
        subject: 'Good email address',
        body: 'This is a test for a good email',
        noReply: true,
    });

    sendAndTrackEmail(ss, 97, {
        to: 'doesnotexist@gsa.gov',
        subject: 'Email does not exisit',
        body: 'This is a test for a fake email',
        noReply: true,
    });

}

function testIsD7(){
  var ss = openTestSpreadsheet();
  var row = 11
  
  var responseValue = IsD7(ss,row);
  Logger.log(responseValue);
}
