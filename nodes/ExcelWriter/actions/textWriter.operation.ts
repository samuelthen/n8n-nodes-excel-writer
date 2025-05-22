import type { IExecuteFunctions, INodeExecutionData } from 'n8n-workflow';
import ExcelJS from 'exceljs';
import { findOrCreateColumn } from '../utils/findOrCreateColumn';

export async function writeTextToExcel(this: IExecuteFunctions, items: INodeExecutionData[]) {
	const returnData: INodeExecutionData[] = [];

	for (let i = 0; i < items.length; i++) {
		const binary = items[i].binary ?? {};
		if (!binary.data) {
			throw new Error('No binary Excel file found in input "data".');
		}

		const textBinaryKey = binary.text ? 'text' : undefined;
		if (!textBinaryKey) {
			throw new Error('No binary text file found in input "text".');
		}

		const excelBuffer = await this.helpers.getBinaryDataBuffer(i, 'data');
		const textBuffer = await this.helpers.getBinaryDataBuffer(i, textBinaryKey);
		const textContent = textBuffer.toString('utf-8');

		const sheetName = this.getNodeParameter('sheetName', i) as string;
		const serialNumber = this.getNodeParameter('serialNumber', i) as number;
		const headerTitle = this.getNodeParameter('headerTitle', i) as string;

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

		cell.value = textContent;
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
