import type { IExecuteFunctions, INodeExecutionData } from 'n8n-workflow';
import ExcelJS from 'exceljs';
import sharp from 'sharp';
import { findOrCreateColumn } from '../utils/findOrCreateColumn';

export async function writeImageToExcel(this: IExecuteFunctions, items: INodeExecutionData[]) {
	const returnData: INodeExecutionData[] = [];

	for (let i = 0; i < items.length; i++) {
		const sheetName = this.getNodeParameter('sheetName', i) as string;
		const serialNumber = this.getNodeParameter('serialNumber', i) as number;
		const headerTitle = this.getNodeParameter('headerTitle', i) as string;

		const binaryExcel = items[i].binary?.data;
		const binaryImage = items[i].binary?.image;

		if (!binaryExcel) {
			throw new Error('Binary Excel file (data) not found in input.');
		}

		if (!binaryImage) {
			throw new Error('Binary image (image) not found in input.');
		}

		const excelBuffer = await this.helpers.getBinaryDataBuffer(i, 'data');
		const imageBuffer = await this.helpers.getBinaryDataBuffer(i, 'image');

		const ext = (binaryImage.fileExtension || 'png').toLowerCase();
		if (!['png', 'jpeg', 'jpg', 'gif'].includes(ext)) {
			throw new Error(`Unsupported image extension: ${ext}`);
		}

		const workbook = new ExcelJS.Workbook();
		await workbook.xlsx.load(excelBuffer);

		const sheet = workbook.getWorksheet(sheetName);
		if (!sheet) {
			throw new Error(`Sheet "${sheetName}" not found in Excel file.`);
		}

		const rowOffset = 1;
		const rowNum = serialNumber + rowOffset;
		const colIndex = findOrCreateColumn(sheet, headerTitle, rowOffset);

		const resizedBuffer = await sharp(imageBuffer)
			.resize({ width: 700, height: 467, fit: 'inside' })
			.toBuffer();

		const imageId = workbook.addImage({
			buffer: resizedBuffer as Buffer,
			extension: ext as 'png' | 'jpeg' | 'gif',
		});

		sheet.addImage(imageId, {
			tl: { col: colIndex - 1, row: rowNum - 1 },
			ext: { width: 700, height: 467 },
			editAs: 'oneCell',
		});

		sheet.getColumn(colIndex).width = 100;
		sheet.getRow(rowNum).height = 350;

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
