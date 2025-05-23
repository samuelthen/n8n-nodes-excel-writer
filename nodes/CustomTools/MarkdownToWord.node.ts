import type {
	IExecuteFunctions,
	INodeExecutionData,
	INodeType,
	INodeTypeDescription,
} from 'n8n-workflow';
import { NodeConnectionType } from 'n8n-workflow';
import MarkdownIt from 'markdown-it';
import { JSDOM } from 'jsdom';
import * as htmlToDocx from 'html-docx-js-typescript';

export class MarkdownToWord implements INodeType {
	description: INodeTypeDescription = {
		displayName: 'Markdown to Word',
		name: 'markdownToWord',
		group: ['transform'],
		version: 1,
		description: 'Converts Markdown text to a Word (.docx) file and returns it as binary',
		defaults: {
			name: 'Markdown to Word',
		},
		inputs: [NodeConnectionType.Main],
		outputs: [NodeConnectionType.Main],
		properties: [
			{
				displayName: 'Markdown Text',
				name: 'markdownText',
				type: 'string',
				typeOptions: {
					rows: 10,
				},
				default: '',
			},
			{
				displayName: 'Filename',
				name: 'filename',
				type: 'string',
				default: 'document.docx',
			},
		],
	};

	async execute(this: IExecuteFunctions): Promise<INodeExecutionData[][]> {
		const returnData: INodeExecutionData[] = [];

		for (let i = 0; i < this.getInputData().length; i++) {
			const markdownText = this.getNodeParameter('markdownText', i) as string;
			const filename = this.getNodeParameter('filename', i) as string;

			// Convert Markdown to HTML
			const md = new MarkdownIt({ html: true });
			const htmlContent = md.render(markdownText);

			// Wrap HTML in a full DOM
			const dom = new JSDOM(htmlContent);
			const htmlString = dom.window.document.documentElement.outerHTML;

			// Convert HTML to DOCX
			const output = await htmlToDocx.asBlob(htmlString);

			// Handle Buffer or fallback to arrayBuffer
			const finalBuffer = Buffer.isBuffer(output)
				? output
				: Buffer.from(await (output as any).arrayBuffer());

			// Return binary data
			returnData.push({
				json: {
					success: true,
				},
				binary: {
					data: await this.helpers.prepareBinaryData(
						finalBuffer,
						filename,
						'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
					),
				},
			});
		}

		return [returnData];
	}
}
