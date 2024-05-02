///////////////////////////////////////////////////////
// EXPORT JSON to S3 ///////////////////////////
////////////////////////////////////////////

function onOpen() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var menuEntries = [
        { name: "Publish to DMN data store", functionName: "exportS3" },
    ];
    ss.addMenu("Publish Data", menuEntries);
}


function exportS3() {

    var s3 = S3.getInstance("<<AWS AccessKey>>", "<<AWS SecretKey>>");
    var doc = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = doc.getActiveSheet();
    var json = getRowsData_(sheet);

    year = (new Date()).getFullYear();

    s3.putObject("interactives.dallasnews.com", "data-store/" + year + "/" + doc.getName() + ".json", json, { logRequests: true });

    var upload_addy = "https://interactives.dallasnews.com/data-store/" + year + "/" + doc.getName() + ".json";

    SpreadsheetApp.getUi().alert("Your data was posted to " + upload_addy + " .");

}



// getRowsData iterates row by row in the input range and returns an array of objects.
// Each object contains all the data for a given row, indexed by its normalized column name.
// Arguments:
//   - sheet: the sheet object that contains the data to be processed
//   - range: the exact range of cells where the data is stored
//   - columnHeadersRowIndex: specifies the row number where the column names are stored.
//       This argument is optional and it defaults to the row immediately above range; 
// Returns an Array of objects.
function getRowsData_(sheet) {
    var headersRange = sheet.getRange(1, 1, sheet.getFrozenRows(), sheet.getMaxColumns());
    var headers = headersRange.getValues()[0];
    var dataRange = sheet.getRange(sheet.getFrozenRows() + 1, 1, sheet.getMaxRows(), sheet.getMaxColumns());
    var objects = getObjects_(dataRange.getValues(), normalizeHeaders(headers));

    return objects;

}

// getColumnsData iterates column by column in the input range and returns an array of objects.
// Each object contains all the data for a given column, indexed by its normalized row name.
// Arguments:
//   - sheet: the sheet object that contains the data to be processed
//   - range: the exact range of cells where the data is stored
//   - rowHeadersColumnIndex: specifies the column number where the row names are stored.
//       This argument is optional and it defaults to the column immediately left of the range; 
// Returns an Array of objects.
function getColumnsData_(sheet, range, rowHeadersColumnIndex) {
    rowHeadersColumnIndex = rowHeadersColumnIndex || range.getColumnIndex() - 1;
    var headersTmp = sheet.getRange(range.getRow(), rowHeadersColumnIndex, range.getNumRows(), 1).getValues();
    var headers = normalizeHeaders(arrayTranspose(headersTmp)[0]);
    return getObjects(arrayTranspose(range.getValues()), headers);
}


// For every row of data in data, generates an object that contains the data. Names of
// object fields are defined in keys.
// Arguments:
//   - data: JavaScript 2d array
//   - keys: Array of Strings that define the property names for the objects to create
function getObjects_(data, keys) {
    var objects = [];
    for (var i = 0; i < data.length; ++i) {
        var object = {};
        var hasData = false;
        for (var j = 0; j < data[i].length; ++j) {
            var cellData = data[i][j];
            if (isCellEmpty(cellData)) {
                continue;
            }
            object[keys[j]] = cellData;
            hasData = true;
        }
        if (hasData) {
            objects.push(object);
        }
    }
    return objects;
}

// Returns an Array of normalized Strings.
// Arguments:
//   - headers: Array of Strings to normalize
function normalizeHeaders(headers) {
    return headers.map(normalizeHeader);
}

// Normalizes a string, by removing all alphanumeric characters and using mixed case
// to separate words. The output will always start with a lower case letter.
// This function is designed to produce JavaScript object property names.
// Arguments:
//   - header: string to normalize
// Examples:
//   "First Name" -> "firstName"
//   "Market Cap (millions) -> "marketCapMillions
//   "1 number at the beginning is ignored" -> "numberAtTheBeginningIsIgnored"
function normalizeHeader(header) {
    let parts = header.split(/\W+/g);

    while (/^\d/.test(parts[0]) && parts.length > 0) {
        parts.shift();
    }

    return parts.map(s => (s.charAt(0).toUpperCase() + s.slice(1))).join('');
}

// Returns true if the cell where cellData was read from is empty.
// Arguments:
//   - cellData: string
function isCellEmpty(cellData) {
    return typeof (cellData) == "string" && cellData == "";
}

// Given a JavaScript 2d Array, this function returns the transposed table.
// Arguments:
//   - data: JavaScript 2d Array
// Returns a JavaScript 2d Array
// Example: arrayTranspose([[1,2,3],[4,5,6]]) returns [[1,4],[2,5],[3,6]].
function arrayTranspose(data) {
    let rows = data.length;
    if (rows === 0) return null;

    let cols = data[0].length;
    if (cols === 0) return null;

    return (new Array(cols))
        .map((_, i) => ((new Array(rows))
            .map((_, j) => data[j][i])));
}
