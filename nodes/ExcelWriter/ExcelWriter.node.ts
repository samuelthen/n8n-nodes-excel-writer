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
		inputs: [NodeConnectionType.Main],
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
		const items = this.getInputData();
		const operation = this.getNodeParameter('operation', 0) as 'json' | 'text' | 'image';

		try {
			switch (operation) {
				case 'json':
					return await writeJsonToExcel.call(this, items);
				case 'text':
					return await writeTextToExcel.call(this, items);
				case 'image':
					return await writeImageToExcel.call(this, items);
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
