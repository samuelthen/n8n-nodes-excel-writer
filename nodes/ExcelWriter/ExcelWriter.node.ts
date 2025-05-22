import type {
	IExecuteFunctions,
	INodeExecutionData,
	INodeType,
	INodeTypeDescription,
} from 'n8n-workflow';
import {
	NodeConnectionType,
	NodeOperationError,
} from 'n8n-workflow';

import { writeJsonToExcel } from './actions/jsonWriter.operation';
import { writeTextToExcel } from './actions/textWriter.operation';
import { writeImageToExcel } from './actions/imageWriter.operation';

export class ExcelWriter implements INodeType {
	description: INodeTypeDescription = {
		displayName: 'Excel Writer',
		name: 'excelWriter',
		icon: {
			light: 'file:excel.svg',
			dark: 'file:convertToFile.excel.svg',
		},
		group: ['transform'],
		version: 1,
		description: 'Writes JSON, text, or images into a binary Excel file',
		defaults: {
			name: 'Excel Writer',
		},
		inputs: [NodeConnectionType.Main, NodeConnectionType.Main],
		inputNames: ['Excel File', 'Data File'],
		outputs: [NodeConnectionType.Main],
		properties: [
			{
				displayName: 'Operation',
				name: 'operation',
				type: 'options',
				options: [
					{ name: 'JSON to Excel',  value: 'json'  },
					{ name: 'Text to Excel',  value: 'text'  },
					{ name: 'Image to Excel', value: 'image' },
				],
				default: 'json',
				description: 'Select how you want to write your data into the Excel file',
			},
			{
				displayName: 'Excel Field Name',
				name: 'excelField',
				type: 'string',
				default: 'excel',
				description: 'Name of the binary field that contains your Excel file',
			},
			{
				displayName: 'Data Field Name',
				name: 'dataField',
				type: 'string',
				default: 'data',
				description: 'Name of the binary (or JSON) field that contains your payload',
			},
			{
				displayName: 'Sheet Name',
				name: 'sheetName',
				type: 'string',
				default: 'Sheet1',
				description: 'The name of the worksheet to write into',
			},
			{
				displayName: 'Serial Number',
				name: 'serialNumber',
				type: 'number',
				default: 1,
				description:
					'The row index where data will start (row 1 is reserved for column headers)',
			},
			{
				displayName: 'Output File Name',
				name: 'outputFileName',
				type: 'string',
				default: 'updated.xlsx',
				description: 'Filename for the Excel file that will be generated',
			},
			{
				displayName: 'Header Title',
				name: 'headerTitle',
				type: 'string',
				default: 'Data',
				description: 'Column header to use when writing plain text or images',
				displayOptions: {
					show: { operation: ['text', 'image'] },
				},
			},
			{
				displayName: 'Save All Images in Folder',
				name: 'saveAll',
				type: 'boolean',
				default: false,
				description:
					'When writing images, save each one in its own folder inside the sheet',
				displayOptions: {
					show: { operation: ['image'] },
				},
			},
		],
	};

	async execute(this: IExecuteFunctions): Promise<INodeExecutionData[][]> {
		// 1) Read the field-name parameters:
		const excelField  = this.getNodeParameter('excelField', 0) as string;
		const dataField   = this.getNodeParameter('dataField',  0) as string;
		const operation   = this.getNodeParameter('operation',  0) as 'json' | 'text' | 'image';

		// 2) Grab both input streams:
		const excelItems = this.getInputData(0);
		const dataItems  = this.getInputData(1) ?? [];

		if (!excelItems.length) {
			throw new NodeOperationError(this.getNode(), 'First input (Excel) is empty.');
		}

		// 3) Merge them into one array of items,
		//    attaching both binaries under the correct keys:
		const mergedItems: INodeExecutionData[] = excelItems.map((excelItem, index) => {
			const dataItem = dataItems[index] ?? {};

			// Ensure the Excel binary is there:
			if (!excelItem.binary?.[excelField]) {
				throw new NodeOperationError(
					this.getNode(),
					`No Excel binary found in field "${excelField}". ` +
					`Available keys: ${Object.keys(excelItem.binary || {}).join(', ')}`,
				);
			}

			// Build the merged item:
			return {
				json: dataItem.json ?? {},
				binary: {
					// Excel file under its configured key:
					[excelField]: excelItem.binary![excelField],
					// If there's a data binary, attach it under its key:
					...(dataItem.binary?.[dataField]
						? { [dataField]: dataItem.binary[dataField] }
						: {}),
				},
			};
		});

		try {
			// 4) Delegate to the correct writer:
			switch (operation) {
				case 'json':
					return await writeJsonToExcel.call(this, mergedItems);
				case 'text':
					return await writeTextToExcel.call(this, mergedItems);
				case 'image':
					return await writeImageToExcel.call(this, mergedItems);
			}
		} catch (err) {
			if (this.continueOnFail()) {
				return [[{ json: { error: (err as Error).message } }]];
			}
			throw err;
		}

		// (unreachable, but TypeScript needs it)
		return [ [] ];
	}
}
