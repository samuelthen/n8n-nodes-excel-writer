import type { IExecuteFunctions, INodeExecutionData } from 'n8n-workflow';
import ExcelJS from 'exceljs';
import { findOrCreateColumn } from '../utils/findOrCreateColumn';

export async function writeJsonToExcel(this: IExecuteFunctions, items: INodeExecutionData[]) {
  const returnData: INodeExecutionData[] = [];

  for (let i = 0; i < items.length; i++) {
    const { binary = {}, json } = items[i];

    // 1) Parameters
    const excelField     = this.getNodeParameter('excelField',     i) as string; // e.g. "excel"
    const dataField      = this.getNodeParameter('dataField',      i) as string; // e.g. "data"
    const sheetName      = this.getNodeParameter('sheetName',      i) as string;
    const serialNumber   = this.getNodeParameter('serialNumber',   i) as number;
    const outputFileName = this.getNodeParameter('outputFileName', i) as string;

    // 2) Load Excel buffer
    if (!binary[excelField]?.data) {
      throw new Error(`No Excel binary found in field "${excelField}".`);
    }
    const excelBuffer = await this.helpers.getBinaryDataBuffer(i, excelField);

    // 3) Load raw payload
    let rawPayload: unknown;
    if (binary[dataField]?.data) {
      // Binary slot (e.g. a .json file)
      const buf = await this.helpers.getBinaryDataBuffer(i, dataField);
      const txt = buf.toString('utf8');
      try {
        rawPayload = JSON.parse(txt);
      } catch {
        throw new Error(`Binary field "${dataField}" did not contain valid JSON.`);
      }
    } else if (json[dataField] !== undefined) {
      // JSON field on the item
      rawPayload = json[dataField];
    } else {
      throw new Error(`Data field "${dataField}" not found in item.`);
    }

    // 4) Ensure it's an object
    let jsonData: Record<string, any>;
    if (typeof rawPayload === 'object' && rawPayload !== null) {
      jsonData = rawPayload as Record<string, any>;
    } else {
      throw new Error(`Payload must be a JSON object, got "${typeof rawPayload}".`);
    }

    // 5) Write into Excel
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(excelBuffer);
    const sheet = workbook.getWorksheet(sheetName);
    if (!sheet) throw new Error(`Sheet "${sheetName}" not found.`);

    const rowOffset = 1;
    const rowNum    = serialNumber + rowOffset;

    for (const [key, value] of Object.entries(jsonData)) {
      const colIndex = findOrCreateColumn(sheet, key, rowOffset);
      const cell     = sheet.getCell(rowNum, colIndex);
      cell.value     = ['string','number','boolean'].includes(typeof value) || value === null
                        ? value
                        : JSON.stringify(value);
      cell.alignment = { wrapText: true };
      sheet.getColumn(colIndex).width = 50;
    }

    // 6) Return updated Excel
    const updatedBuffer = await workbook.xlsx.writeBuffer();
    returnData.push({
      json: { success: true },
      binary: {
        [excelField]: await this.helpers.prepareBinaryData(
          updatedBuffer as Buffer,
          outputFileName,
        ),
      },
    });
  }

  return [returnData];
}
