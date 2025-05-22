import type { IExecuteFunctions, INodeExecutionData } from 'n8n-workflow';
import ExcelJS from 'exceljs';
import { findOrCreateColumn } from '../utils/findOrCreateColumn';

export async function writeTextToExcel(this: IExecuteFunctions, items: INodeExecutionData[]) {
	const returnData: INodeExecutionData[] = [];

	for (let i = 0; i < items.length; i++) {
		const sheetName = this.getNodeParameter('sheetName', i) as string;
		const serialNumber = this.getNodeParameter('serialNumber', i) as number;
		const headerTitle = this.getNodeParameter('headerTitle', i) as string;

		const binaryExcel = items[i].binary?.data;
		const binaryText = items[i].binary?.text;

		if (!binaryExcel) {
			throw new Error('Binary Excel file (data) not found in input.');
		}
		if (!binaryText) {
			throw new Error('Binary text file (text) not found in input.');
		}

		const excelBuffer = await this.helpers.getBinaryDataBuffer(i, 'data');
		const textBuffer = await this.helpers.getBinaryDataBuffer(i, 'text');
		const textData = textBuffer.toString('utf-8');

		const workbook = new ExcelJS.Workbook();
		await workbook.xlsx.load(excelBuffer);

		const sheet = workbook.getWorksheet(sheetName);
		if (!sheet) {
			throw new Error(`Sheet "${sheetName}" not found in Excel file.`);
		}

		const rowOffset = 1;
		const rowNum = serialNumber + rowOffset;
		const colIndex = findOrCreateColumn(sheet, headerTitle, rowOffset);

		const cell = sheet.getCell(rowNum, colIndex);
		cell.value = textData;
		cell.alignment = { wrapText: true };
		sheet.getColumn(colIndex).width = 50;

		const updatedBuffer = await workbook.xlsx.writeBuffer();

		returnData.push({
			json: { success: true },
			binary: {
				data: await this.helpers.prepareBinaryData(updatedBuffer as Buffer, 'updated.xlsx'),
			},
		});
	}

	return [returnData];
}
