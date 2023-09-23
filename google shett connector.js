let emailNotification = false;
let emailAddress = "Change_to_your_Email";

let isNewSheet = false;
let postedData = [];
const EXCLUDE_PROPERTY = 'e_gs_exclude';
const ORDER_PROPERTY = 'e_gs_order';
const SHEET_NAME_PROPERTY = 'e_gs_SheetName';

function doGet(e) {
    return HtmlService.createHtmlOutput("Yepp this is the webhook URL, request received");
}

function doPost(e) {
    let params = JSON.stringify(e.parameter);
    params = JSON.parse(params);
    postedData = params;
    insertToSheet(params);

    return HtmlService.createHtmlOutput("post request received");
}

const flattenObject = (ob) => {
    let toReturn = {};
    for (let i in ob) {
        if (!ob.hasOwnProperty(i)) continue;
        if ((typeof ob[i]) === 'object') {
            let flatObject = flattenObject(ob[i]);
            for (let x in flatObject) {
                if (!flatObject.hasOwnProperty(x)) continue;
                toReturn[i + '.' + x] = flatObject[x];
            }
        } else {
            toReturn[i] = ob[i];
        }
    }
    return toReturn;
}

const getHeaders = (formSheet) => {
    let headers = ["First Name", "Email", "Last Name", "Mobile", "Message"];
    if (!isNewSheet) {
        let existingHeaders = formSheet.getRange(1, 1, 1, formSheet.getLastColumn()).getValues()[0];
        existingHeaders.forEach(header => {
            if (!headers.includes(header)) {
                headers.push(header);
            }
        });
    }

    headers = excludeColumns(headers);
    return headers;
};

const getValues = (headers, flat) => {
    const values = [];
    headers.forEach((h) => values.push(flat[h] || ""));
    return values;
}

const insertRowData = (sheet, row, values, bold = false) => {
    const currentRow = sheet.getRange(row, 1, 1, values.length);
    currentRow.setValues([values])
        .setFontWeight(bold ? "bold" : "normal")
        .setHorizontalAlignment("center");
}

const setHeaders = (sheet, values) => insertRowData(sheet, 1, values, true);
const setValues = (sheet, values) => {
    const lastRow = Math.max(sheet.getLastRow(), 1);
    sheet.insertRowAfter(lastRow);
    insertRowData(sheet, lastRow + 1, values);
}

const getFormSheet = (sheetName) => {
    const activeSheet = SpreadsheetApp.getActiveSpreadsheet();
    if (!activeSheet.getSheetByName(sheetName)) {
        const formSheet = activeSheet.insertSheet();
        formSheet.setName(sheetName);
        isNewSheet = true;
    }
    return activeSheet.getSheetByName(sheetName);
}

const insertToSheet = (data) => {
    const flat = flattenObject(data);
    const formSheet = getFormSheet(getSheetName(data));
    const headers = getHeaders(formSheet);
    const values = getValues(headers, flat);

    setHeaders(formSheet, headers);
    setValues(formSheet, values);

    if (emailNotification) {
        sendNotification(data, getSheetURL());
    }
}

const getSheetName = (data) => data[SHEET_NAME_PROPERTY] || data["form_name"];
const getSheetURL = () => SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getUrl();

const stringToArray = (str) => str && str.split ? str.split(",").map(el => el.trim()) : [];

const excludeColumns = (headers) => {
    if (!postedData || !postedData[EXCLUDE_PROPERTY]) {
        return headers;
    }
    const columnsToExclude = stringToArray(postedData[EXCLUDE_PROPERTY]);
    return headers.filter(header => !columnsToExclude.includes(header));
}

const sendNotification = (data, url) => {
    MailApp.sendEmail(
        emailAddress,
        "A new Elementor Pro Forms submission has been inserted to your sheet",
        `A new submission has been received via ${data['form_name']} form and inserted into your Google sheet at: ${url}`,
        {
            name: 'Automatic Emailer Script'
        }
    );
};
