import type { IExecuteFunctions, INodeExecutionData } from 'n8n-workflow';
import ExcelJS from 'exceljs';
import { findOrCreateColumn } from '../utils/findOrCreateColumn';

export async function writeJsonToExcel(this: IExecuteFunctions, items: INodeExecutionData[]) {
	const returnData: INodeExecutionData[] = [];

	for (let i = 0; i < items.length; i++) {
		const binaryData = items[i].binary?.data;
		if (!binaryData) {
			throw new Error('No binary Excel file found in input "data"');
		}

		const buffer = await this.helpers.getBinaryDataBuffer(i, 'data');
		const jsonData = items[i].json;
		const sheetName = this.getNodeParameter('sheetName', i) as string;
		const serialNumber = this.getNodeParameter('serialNumber', i) as number;

		const workbook = new ExcelJS.Workbook();
		await workbook.xlsx.load(buffer);

		const sheet = workbook.getWorksheet(sheetName);
		if (!sheet) {
			throw new Error(`Sheet "${sheetName}" not found in Excel file.`);
		}

		const rowOffset = 1;
		const rowNum = serialNumber + rowOffset;

		for (const [key, value] of Object.entries(jsonData)) {
			const colIndex = findOrCreateColumn(sheet, key, rowOffset);
			const cell = sheet.getCell(rowNum, colIndex);

			if (
				typeof value === 'string' ||
				typeof value === 'number' ||
				typeof value === 'boolean' ||
				value === null
			) {
				cell.value = value;
			} else if (typeof value === 'undefined') {
				cell.value = '';
			} else {
				cell.value = JSON.stringify(value);
			}

			cell.alignment = { wrapText: true };
			sheet.getColumn(colIndex).width = 50;
		}

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
