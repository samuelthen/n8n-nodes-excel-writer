import ExcelJS from 'exceljs';

export function findOrCreateColumn(
	sheet: ExcelJS.Worksheet,
	title: string,
	headerRow: number,
): number {
	const header = sheet.getRow(headerRow);
	for (let i = 1; i <= sheet.columnCount; i++) {
		if (header.getCell(i).value === title) {
			return i;
		}
	}
	const newCol = sheet.columnCount + 1;
	header.getCell(newCol).value = title;
	header.getCell(newCol).font = { bold: true };
	return newCol;
}
