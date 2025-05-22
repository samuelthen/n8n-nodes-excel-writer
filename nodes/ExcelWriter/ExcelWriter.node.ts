import type {
	IExecuteFunctions,
	INodeExecutionData,
	INodeType,
	INodeTypeDescription,
} from 'n8n-workflow';
import { NodeConnectionType, NodeOperationError } from 'n8n-workflow';

import { writeJsonToExcel } from './actions/jsonWriter.operation';
import { writeTextToExcel } from './actions/textWriter.operation';
import { writeImageToExcel } from './actions/imageWriter.operation';

export class ExcelWriter implements INodeType {
	description: INodeTypeDescription = {
		displayName: 'Excel Writer',
		name: 'excelWriter',
		icon: { light: 'file:excel.svg', dark: 'file:convertToFile.excel.svg' },
		group: ['transform'],
		version: 1,
		description: 'Writes JSON, text, or images into a binary Excel file',
		defaults: {
			name: 'Excel Writer',
		},
		inputs: [NodeConnectionType.Main, NodeConnectionType.Main], // Two input streams
		inputNames: ['Excel File', 'Data File'],
		outputs: [NodeConnectionType.Main],
		properties: [
			{
				displayName: 'Operation',
				name: 'operation',
				type: 'options',
				options: [
					{ name: 'Write JSON', value: 'json' },
					{ name: 'Write Text', value: 'text' },
					{ name: 'Write Image', value: 'image' },
				],
				default: 'json',
			},
			{
				displayName: 'Sheet Name',
				name: 'sheetName',
				type: 'string',
				default: 'Sheet1',
			},
			{
				displayName: 'Serial Number',
				name: 'serialNumber',
				type: 'number',
				default: 1,
			},
			{
				displayName: 'Header Title',
				name: 'headerTitle',
				type: 'string',
				default: 'Data',
				displayOptions: {
					show: {
						operation: ['text', 'image'],
					},
				},
			},
			{
				displayName: 'Save All Images in Folder',
				name: 'saveAll',
				type: 'boolean',
				default: false,
				displayOptions: {
					show: {
						operation: ['image'],
					},
				},
			},
		],
	};

	async execute(this: IExecuteFunctions): Promise<INodeExecutionData[][]> {
		const inputs = [this.getInputData(0), this.getInputData(1)];

		if (inputs[0] === undefined || inputs[0].length === 0) {
			throw new NodeOperationError(this.getNode(), 'Input 1 is empty or missing.');
		}

		// Support single or dual input by defaulting to input[0] if input[1] is missing
		const inputA = inputs[0];
		const inputB = inputs[1] ?? [];

		// Use whichever input contains the Excel file (binary.data)
		const mergedItems: INodeExecutionData[] = inputA.map((itemA, index) => {
			const itemB = inputB[index] ?? itemA;

			// Determine which item contains the Excel file
			const hasExcelInA = itemA.binary?.data !== undefined;
			const hasExcelInB = itemB.binary?.data !== undefined;

			let excelItem: INodeExecutionData;
			let dataItem: INodeExecutionData;

			if (hasExcelInA && !hasExcelInB) {
				excelItem = itemA;
				dataItem = itemB;
			} else if (hasExcelInB && !hasExcelInA) {
				excelItem = itemB;
				dataItem = itemA;
			} else if (hasExcelInA && hasExcelInB) {
				throw new NodeOperationError(this.getNode(), 'Both inputs contain Excel binary. Only one should.');
			} else {
				throw new NodeOperationError(this.getNode(), 'No Excel binary found in either input.');
			}

			return {
				json: dataItem.json,
				binary: {
					...(dataItem.binary ?? {}),
					data: excelItem.binary!.data,
				},
			};
		});

		const operation = this.getNodeParameter('operation', 0) as 'json' | 'text' | 'image';

		try {
			switch (operation) {
				case 'json':
					return await writeJsonToExcel.call(this, mergedItems);
				case 'text':
					return await writeTextToExcel.call(this, mergedItems);
				case 'image':
					return await writeImageToExcel.call(this, mergedItems);
				default:
					throw new NodeOperationError(this.getNode(), `Unsupported operation "${operation}"`);
			}
		} catch (error) {
			if (this.continueOnFail()) {
				return [[{ json: { error: (error as Error).message } }]];
			}
			throw new NodeOperationError(this.getNode(), error);
		}
	}

}
