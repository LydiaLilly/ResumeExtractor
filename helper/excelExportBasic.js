const xl = require('excel4node');
const createSheet = (sheetName, columnNames, data, headingColumnMap) => {
    const wb = new xl.Workbook();
    const ws = wb.addWorksheet(sheetName);
    createHeading(columnNames, ws);
    addRows(data, ws, columnNames, headingColumnMap);
    return wb;
};
const createHeading = (columnNames, ws) => {
    let headingColumnIndex = 1;
    columnNames.forEach(heading => {
        ws.cell(1, headingColumnIndex++)
            .string(heading)
    });
};
const addRows = (data, ws, columnNames, headingColumnMap) => {
    let rowIndex = 2;
    data.forEach( record => {
        let columnIndex = 1;
        columnNames.forEach(heading => {
            ws.cell(rowIndex, columnIndex++)
                .string((record[headingColumnMap[heading]]).toString());
        });
        rowIndex++;
    });
};
module.exports = {
    createSheet
};