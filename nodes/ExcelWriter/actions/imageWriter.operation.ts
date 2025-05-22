import type { IExecuteFunctions, INodeExecutionData } from 'n8n-workflow';
import ExcelJS from 'exceljs';
import sharp from 'sharp';
import { findOrCreateColumn } from '../utils/findOrCreateColumn';

export async function writeImageToExcel(this: IExecuteFunctions, items: INodeExecutionData[]) {
	const returnData: INodeExecutionData[] = [];

	for (let i = 0; i < items.length; i++) {
		const { binary = {} } = items[i];

		// ── 1) Parameters ──────────────────────────────────────────────────────
		const excelField     = this.getNodeParameter('excelField',     i) as string;
		const dataField      = this.getNodeParameter('dataField',      i) as string; // image binary field
		const sheetName      = this.getNodeParameter('sheetName',      i) as string;
		const headerTitle    = this.getNodeParameter('headerTitle',    i) as string;
		const serialNumber   = this.getNodeParameter('serialNumber',   i) as number;
		const outputFileName = this.getNodeParameter('outputFileName', i) as string;

		// ── 2) Validate binary inputs ──────────────────────────────────────────
		if (!binary[excelField]?.data) {
			throw new Error(`No Excel binary found in field "${excelField}".`);
		}
		if (!binary[dataField]?.data) {
			throw new Error(`No image binary found in field "${dataField}".`);
		}

		// ── 3) Read Excel and Image Buffers ────────────────────────────────────
		const excelBuffer = await this.helpers.getBinaryDataBuffer(i, excelField);

		let imageBuffer: Buffer;
		if (binary[dataField]?.id) {
			imageBuffer = await this.helpers.getBinaryDataBuffer(i, dataField);
		} else {
			imageBuffer = Buffer.from(binary[dataField].data, 'base64');
		}

		// ── 4) Validate image extension ────────────────────────────────────────
		const extRaw = (binary[dataField].fileExtension || 'png').toLowerCase();
		const ext = extRaw === 'jpg' ? 'jpeg' : extRaw;

		if (!['png', 'jpeg', 'gif'].includes(ext)) {
			throw new Error(`Unsupported image extension: ${extRaw}`);
		}

		// ── 5) Load Excel and locate position ──────────────────────────────────
		const workbook = new ExcelJS.Workbook();
		await workbook.xlsx.load(excelBuffer);

		const sheet = workbook.getWorksheet(sheetName);
		if (!sheet) {
			throw new Error(`Sheet "${sheetName}" not found in Excel file.`);
		}

		const rowOffset = 1;
		const rowNum    = serialNumber + rowOffset;
		const colIndex  = findOrCreateColumn(sheet, headerTitle, rowOffset);

		// ── 6) Resize image and insert ─────────────────────────────────────────
		const resizedBuffer = await sharp(imageBuffer)
			.resize({ width: 700, height: 467, fit: 'inside' })
			.toBuffer();

		const imageId = workbook.addImage({
			buffer: resizedBuffer,
			extension: ext as 'png' | 'jpeg' | 'gif',
		});

		sheet.addImage(imageId, {
			tl: { col: colIndex - 1, row: rowNum - 1 },
			ext: { width: 700, height: 467 },
			editAs: 'oneCell',
		});

		sheet.getColumn(colIndex).width = 100;
		sheet.getRow(rowNum).height = 350;

		// ── 7) Output updated Excel ────────────────────────────────────────────
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
