import type { IExecuteFunctions, INodeExecutionData } from 'n8n-workflow';
import ExcelJS from 'exceljs';
import { findOrCreateColumn } from '../utils/findOrCreateColumn';

export async function writeTextToExcel(this: IExecuteFunctions, items: INodeExecutionData[]) {
	const returnData: INodeExecutionData[] = [];

	for (let i = 0; i < items.length; i++) {
		const { binary = {}, json } = items[i];

		// ── 1) Parameters ──────────────────────────────────────────────────────
		const excelField     = this.getNodeParameter('excelField',     i) as string;
		const dataField      = this.getNodeParameter('dataField',      i) as string;
		const headerTitle    = this.getNodeParameter('headerTitle',    i) as string;
		const sheetName      = this.getNodeParameter('sheetName',      i) as string;
		const serialNumber   = this.getNodeParameter('serialNumber',   i) as number;
		const outputFileName = this.getNodeParameter('outputFileName', i) as string;

		// ── 2) Validate Excel ──────────────────────────────────────────────────
		if (!binary[excelField]?.data) {
			throw new Error(`No Excel binary found in field "${excelField}".`);
		}
		const excelBuffer = await this.helpers.getBinaryDataBuffer(i, excelField);

		// ── 3) Get value from JSON[dataField] ──────────────────────────────────
		const textValue = json[dataField];
		if (typeof textValue !== 'string') {
			throw new Error(`Expected a string in field "${dataField}", but got "${typeof textValue}".`);
		}

		// ── 4) Load Excel and locate cell ──────────────────────────────────────
		const workbook = new ExcelJS.Workbook();
		await workbook.xlsx.load(excelBuffer);
		const sheet = workbook.getWorksheet(sheetName);
		if (!sheet) {
			throw new Error(`Sheet "${sheetName}" not found in Excel file.`);
		}

		const rowOffset = 1;
		const rowNum    = serialNumber + rowOffset;
		const colIndex  = findOrCreateColumn(sheet, headerTitle, rowOffset);
		const cell      = sheet.getCell(rowNum, colIndex);

		cell.value = textValue;
		cell.alignment = { wrapText: true };
		sheet.getColumn(colIndex).width = 50;

		// ── 5) Output updated Excel ────────────────────────────────────────────
		const updatedBuffer = await workbook.xlsx.writeBuffer();
		returnData.push({
			json: { success: true },
			binary: {
				[excelField]: await this.helpers.prepareBinaryData(updatedBuffer as Buffer, outputFileName),
			},
		});
	}

	return [returnData];
}
