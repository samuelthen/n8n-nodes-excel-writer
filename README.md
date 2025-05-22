# n8n-nodes-excel-writer

This is an n8n community node. It lets you use Excel file manipulation in your n8n workflows.

This node enables you to write text, JSON objects, and insert images into `.xlsx` Excel files using the `exceljs` and `sharp` libraries, entirely in memory. It's ideal for automated report generation, templated document filling, and embedding media into spreadsheets.

[n8n](https://n8n.io) is a [fair-code licensed](https://docs.n8n.io/reference/license/) workflow automation platform.

---

## Installation

Follow the [installation guide](https://docs.n8n.io/integrations/community-nodes/installation/) in the n8n community nodes documentation.

In your self-hosted n8n instance directory, run:

```bash
npm install n8n-nodes-excel-writer
```

Then restart your n8n instance. Make sure `N8N_COMMUNITY_PACKAGES_ENABLED=true` is set in your `.env` file.

---

## Operations

This node supports the following operations:

### 📝 Write Text to Excel
- Writes a single string into a specific cell based on column header and row number.

### 📄 Write JSON to Excel
- Maps a JSON object to multiple columns in one row.
- Can read input from a JSON field or `.json` file in a binary field.

### 🖼️ Write Image to Excel
- Inserts a binary image (JPG, PNG, or GIF) into a specific cell.
- Supports auto-resizing and column/row adjustments.

---

## Credentials

❌ No authentication is required. This node works entirely with local or provided binary `.xlsx` and `.json` files. It does not connect to an external service.

---

## Compatibility

- ✅ Minimum required n8n version: `1.60.0`
- ✅ Tested on n8n `1.64.3`
- ✅ Compatible with local, Docker, and cloud self-hosted n8n environments
- ⚠️ Node must receive a valid `.xlsx` binary file

---

## Usage

To use this node:

1. Use an upstream node to provide a binary Excel file:
   - `Read Binary File`
   - `HTTP Request` (Response: file)
   - `Google Drive → Download`, etc.

2. Add the **Excel Writer** node and select an operation:
   - For text or JSON, use a JSON input or structured object.
   - For image insertion, make sure a binary image is present.

3. Fill in the following key fields:
   - `excelField` — name of the binary field with the Excel file (default: `data`)
   - `dataField` — text/JSON key or image binary field name
   - `sheetName`, `headerTitle`, `serialNumber` — used to place content
   - `outputFileName` — name of the updated Excel file to return

Refer to the [Try it out](https://docs.n8n.io/try-it-out/) documentation if you're new to n8n workflows.

---

## Resources

- [n8n community nodes documentation](https://docs.n8n.io/integrations/community-nodes/)
- [exceljs documentation](https://github.com/exceljs/exceljs)
- [sharp documentation](https://sharp.pixelplumbing.com/)

---

## Version history

| Version | Changes                                      |
|---------|----------------------------------------------|
| 0.1.0   | Initial release: supports text, JSON, and image writing into Excel |

---

## License

This project is licensed under the [MIT License](./LICENSE).  
Note that n8n is distributed under the [Sustainable Use License](https://docs.n8n.io/reference/license/).
