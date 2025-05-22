# n8n-nodes-excel-writer

This is an n8n community node that allows you to **write text, JSON objects, and images into Excel files** as part of your workflows.

It uses the powerful [`exceljs`](https://github.com/exceljs/exceljs) and [`sharp`](https://sharp.pixelplumbing.com/) libraries to dynamically edit `.xlsx` files in memory — no external services required.

[n8n](https://n8n.io/) is a [fair-code licensed](https://docs.n8n.io/reference/license/) workflow automation platform.

---

## Installation

Follow the [community node installation guide](https://docs.n8n.io/integrations/community-nodes/installation/) to install this node into your self-hosted n8n instance.

---

## Operations

This node supports the following operations:

### 📝 Write Text to Excel

- Inserts a single text value into a specified cell based on column header and row number.
- Automatically creates the column if it doesn't exist.
- Supports custom sheet name and output file name.

### 📄 Write JSON to Excel

- Accepts a JSON object and writes each key-value pair into the corresponding column.
- Automatically maps or creates column headers based on the object keys.
- Supports structured input from either JSON or binary `.json` file.

### 🖼️ Write Image to Excel

- Inserts an image (PNG, JPG, or GIF) into a specific cell.
- Automatically resizes and places the image inside the target cell.
- Works with binary image fields from sources like HTTP, file upload, or `Read Binary File`.

---

## Parameters

Each operation supports these customizable parameters:

| Name             | Type    | Description                                                                 |
|------------------|---------|-----------------------------------------------------------------------------|
| `excelField`     | string  | Binary field name that contains the Excel file (default: `data`)           |
| `dataField`      | string  | For text/JSON: JSON key. For images: binary field name containing the image |
| `sheetName`      | string  | Name of the sheet in the Excel file                                        |
| `headerTitle`    | string  | (For text/image) Column header to match or create                          |
| `serialNumber`   | number  | Row number (1-based) to write into (excluding header row)                  |
| `outputFileName` | string  | File name for the updated Excel file                                       |

---

## Compatibility

- ✅ n8n `v1.60.0+`
- ✅ Fully tested on `v1.64.3`
- ⚠️ Node must receive a valid `.xlsx` binary file

---

## Usage Tips

- Use `Read Binary File` or `HTTP Request (Response: file)` to load your `.xlsx` file into the binary field.
- For images, use nodes like `Read Binary File`, `Move Binary Data`, or `Webhook` with `Binary` input.
- `dataField` is overloaded by design to accept either:
  - A **JSON key** (for text and JSON writing)
  - A **binary field name** (for image insertion)

---

## Resources

- [n8n Community Nodes Documentation](https://docs.n8n.io/integrations/#community-nodes)
- [exceljs on GitHub](https://github.com/exceljs/exceljs)
- [sharp image library](https://sharp.pixelplumbing.com/)

---

## Version History

| Version | Changes                                      |
|---------|----------------------------------------------|
| 0.1.0   | Initial release: supports text, JSON, and image writing into Excel |

---

## License

This project follows the [n8n fair-code license](https://docs.n8n.io/reference/license/).
