function main(workbook: ExcelScript.Workbook) {

	const selectedCell = workbook.getActiveCell();
	const selectedSheet = workbook.getActiveWorksheet();
	const sourceRowIndex = selectedCell.getRowIndex()+1;

	const mandatoryCells = selectedSheet.getRange(`A${sourceRowIndex}:F${sourceRowIndex}`).getValues();
	for (const cell of mandatoryCells[0]) {
		if(cell.toString().length === 0) {
			throw("Wybrano nieuzupełniony wiersz");
		}
	}

	const alreadyEntered = selectedSheet.getRange(`J${sourceRowIndex}`);
	if(alreadyEntered.getValue().toString() != "") {
		throw("Wpis został już wprowadzony");
	}

	const clientName = selectedSheet.getRange(`A${sourceRowIndex}`).getValue().toString();
	const employeeName = selectedSheet.getName().toString();
	const sourceValues = selectedSheet.getRange(`B${sourceRowIndex}:I${sourceRowIndex}`);

	let destinationSheet = workbook.getWorksheet(clientName);
	if(destinationSheet === undefined) {
		destinationSheet = workbook.getWorksheet("FV Firmy bez umowy");
		const destinationRowIndex = destinationSheet.getTable("BezUmowy").getRowCount()+2;
		destinationSheet.getRange(`C${destinationRowIndex}`).copyFrom(sourceValues, ExcelScript.RangeCopyType.values, false, false);
		destinationSheet.getRange(`A${destinationRowIndex}`).setValue(clientName);
		destinationSheet.getRange(`B${destinationRowIndex}`).setValue(employeeName);
	}
	else {
		const destinationRowIndex = destinationSheet.getTable(clientName).getRowCount()+2;
		destinationSheet.getRange(`B${destinationRowIndex}`).copyFrom(sourceValues, ExcelScript.RangeCopyType.values, false, false);
		destinationSheet.getRange(`A${destinationRowIndex}`).setValue(employeeName);
	}
	alreadyEntered.setValue("TAK");
}
