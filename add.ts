
function main(workbook: ExcelScript.Workbook) {
    const selectedSheet = workbook.getActiveWorksheet();
    const selectedRange = workbook.getSelectedRange();
    const firstRow = selectedRange.getRowIndex()+1;
    const numRows = selectedRange.getRowCount();

    const table = selectedSheet.getTables()[0];
    const summaryRowIndex = table.getRowCount()+2;
    table.addRow();
    const summaryRow = selectedSheet.getRange(`A${summaryRowIndex}:I${summaryRowIndex}`);

    const formulas = [
        [`=SUMA.WARUNKÓW(C${firstRow}:C${firstRow + numRows - 1}; E${firstRow}:E${firstRow + numRows - 1}; \"IT\"; I${firstRow}:I${firstRow + numRows - 1}; \"<>wolne\")`,
        `=SUMA.WARUNKÓW(C${firstRow}:C${firstRow + numRows - 1}; E${firstRow}:E${firstRow + numRows - 1}; \"Helpdesk/Tel/Mail\"; I${firstRow}:I${firstRow + numRows - 1}; \"<>wolne\")`,
        `=SUMA.WARUNKÓW(C${firstRow}:C${firstRow + numRows - 1}; E${firstRow}:E${firstRow + numRows - 1}; \"Spotkania i Organizacja\"; I${firstRow}:I${firstRow + numRows - 1}; \"<>wolne\")`,
        `=SUMA.WARUNKÓW(C${firstRow}:C${firstRow + numRows - 1}; E${firstRow}:E${firstRow + numRows - 1}; \"IT\"; I${firstRow}:I${firstRow + numRows - 1}; \"wolne\")`,
        `=SUMA.WARUNKÓW(C${firstRow}:C${firstRow + numRows - 1}; E${firstRow}:E${firstRow + numRows - 1}; \"Helpdesk/Tel/Mail\"; I${firstRow}:I${firstRow + numRows - 1}; \"wolne\")`,
        `=SUMA.WARUNKÓW(C${firstRow}:C${firstRow + numRows - 1}; E${firstRow}:E${firstRow + numRows - 1}; \"Spotkania i Organizacja\"; I${firstRow}:I${firstRow + numRows - 1}; \"wolne\")`,
        `=SUMA(A${summaryRowIndex}:C${summaryRowIndex})`,
        `=SUMA(D${summaryRowIndex}:F${summaryRowIndex})`,
        `=SUMA(G${summaryRowIndex}:H${summaryRowIndex})`]
    ];

    summaryRow.setFormulasLocal(formulas);
    summaryRow.getFormat().getFill().setColor("yellow");
    const results = summaryRow.getValues();
    summaryRow.clear(ExcelScript.ClearApplyTo.contents)
    selectedSheet.getRange(`E${summaryRowIndex}`).setValue(`Podsumowano wiersze: ${firstRow}:${firstRow+numRows-1}`);
    selectedSheet.getRange(`F${summaryRowIndex}`).setValue(`IT: ${results[0][0]}, HD: ${results[0][1]}, SO: ${results[0][2]}, IT_W: ${results[0][3]}, HD_W: ${results[0][4]}, SO_W: ${results[0][5]}, RAZEM_BW: ${results[0][6]}, RAZEM_W: ${results[0][7]}, RAZEM: ${results[0][8]}`);
}
