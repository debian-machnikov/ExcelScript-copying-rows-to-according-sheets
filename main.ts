function main(workbook: ExcelScript.Workbook) {

	const selectedCell = workbook.getActiveCell();
	const selectedSheet = workbook.getActiveWorksheet();

	const sourceRowIndex = selectedCell.getRowIndex()+1;

	const clientName = selectedSheet.getRange(`A${sourceRowIndex}`).getValue().toString();
	const employeeName = selectedSheet.getName().toString();

	if(clientName.length !== 0 && employeeName.length !== 0) {
		const mandatoryCells = selectedSheet.getRange(`B${sourceRowIndex}:F${sourceRowIndex}`).getValues();
		for (const cell of mandatoryCells[0]) {
			if(cell.toString().length === 0) {
				throw("Wybrano nieuzupełniony wiersz");
			}
		}

		const destinationSheet = workbook.getWorksheet(clientName);
		if(destinationSheet === undefined) {
			const defaultDestination = workbook.getWorksheet("FV Firmy bez umowy");
			const destinationRowIndex = defaultDestination.getTable("BezUmowy").getRowCount()+2;
			defaultDestination.getRange(`C${destinationRowIndex}`).copyFrom(selectedSheet.getRange(`B${sourceRowIndex}:I${sourceRowIndex}`), ExcelScript.RangeCopyType.values, false, false);
			defaultDestination.getRange(`A${destinationRowIndex}`).setValue(clientName);
			defaultDestination.getRange(`B${destinationRowIndex}`).setValue(employeeName);
		}
		else {
			const destinationRowIndex = destinationSheet.getTable(clientName).getRowCount()+2;
			destinationSheet.getRange(`B${destinationRowIndex}`).copyFrom(selectedSheet.getRange(`B${sourceRowIndex}:I${sourceRowIndex}`), ExcelScript.RangeCopyType.values, false, false);
			destinationSheet.getRange(`A${destinationRowIndex}`).setValue(employeeName);
		}
		


	}
	else {
		throw("Wybrano pusty/nieuzupełniony wiersz");
	}
}
