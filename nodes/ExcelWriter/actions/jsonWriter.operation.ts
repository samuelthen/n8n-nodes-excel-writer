import type { IExecuteFunctions, INodeExecutionData } from 'n8n-workflow';
import ExcelJS from 'exceljs';
import { findOrCreateColumn } from '../utils/findOrCreateColumn';

export async function writeJsonToExcel(this: IExecuteFunctions, items: INodeExecutionData[]) {
  const returnData: INodeExecutionData[] = [];

  for (let i = 0; i < items.length; i++) {
    const { binary = {}, json } = items[i];

    // 1) Load Excel buffer
    if (!binary.data) {
      throw new Error('No binary Excel file found in input "data".');
    }
    const excelBuffer = await this.helpers.getBinaryDataBuffer(i, 'data');

    // 2) Determine rawPayload: prefer binary.text, else items[i].json
    let rawPayload: unknown;
    if (binary.text) {
      const textBuffer = await this.helpers.getBinaryDataBuffer(i, 'text');
      rawPayload = JSON.parse(textBuffer.toString('utf-8'));
    } else {
      rawPayload = json;
    }

    // 3) If rawPayload is an object with exactly one key, unwrap it
    let jsonData: Record<string, any>;
    if (
      typeof rawPayload === 'object' &&
      rawPayload !== null &&
      !Array.isArray(rawPayload) &&
      Object.keys(rawPayload).length === 1
    ) {
      const inner = (rawPayload as Record<string, any>)[Object.keys(rawPayload)[0]];
      // if it’s a JSON string, parse it
      if (typeof inner === 'string') {
        jsonData = JSON.parse(inner);
      } else if (typeof inner === 'object' && inner !== null) {
        jsonData = inner;
      } else {
        throw new Error('Wrapped value is not an object or JSON string.');
      }
    } else if (typeof rawPayload === 'object' && rawPayload !== null) {
      jsonData = rawPayload as Record<string, any>;
    } else {
      throw new Error('Payload is not a JSON object.');
    }

    // 4) Write jsonData to Excel
    const sheetName = this.getNodeParameter('sheetName', i) as string;
    const serialNumber = this.getNodeParameter('serialNumber', i) as number;

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(excelBuffer);

    const sheet = workbook.getWorksheet(sheetName);
    if (!sheet) {
      throw new Error(`Sheet "${sheetName}" not found.`);
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
