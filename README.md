# JSspread 📊

[English](README.md) | [Português](README.pt-br.md)

A modern, high-performance **Google Apps Script** library written in **TypeScript** and bundled with **Vite**. Inspired by Python's `gspread`, **JSspread** simplifies reading, writing, filtering, and manipulating Google Sheets data with strong typing, robust error handling, and blazing-fast in-memory operations.

---

## 🚀 Features

- **Object-Oriented API**: Intuitive classes for `Spreadsheet`, `Worksheet`, and `Cell`.
- **Fuzzy & Dynamic Mapping**: `getMappedRecords` automatically maps columns even if columns change order, contain typos, or have title rows above them.
- **In-Memory ETL Tools**:
  - `extract()`: Slice specific columns from raw 2D matrices or object lists by name or index.
  - `merge()`: Lightning-fast in-memory **VLOOKUP / LEFT JOIN** ($O(N + M)$) across sheets without slowing down your spreadsheet with heavy formulas.
- **Robust Date & Text Parsers**:
  - `ensureDate()`: Safely handles Google Sheets numeric serial dates (e.g., `45400`), Brazilian formats (`DD/MM/YYYY`), and ISO strings.
  - `fmtData()`: Ultra-fast date formatter optimized for loops (avoids slow `Intl` and API round-trips).
  - `normalizeTxt()`: Normalizes accents, casing, duplicate spaces, and delimiters (`-`, `_`, `/`).
- **First-Class TypeScript Support**: Automatically generates client & library `.d.ts` definitions via `@antunes_s/clasp-types`.

---

## 📦 Installation & Setup

### Prerequisites
- [Node.js](https://nodejs.org/) (v18+)
- [pnpm](https://pnpm.io/)
- [Google Clasp](https://github.com/google/clasp) logged in (`clasp login`)

### Install Dependencies
```bash
pnpm install
```

---

## 🛠️ Available Commands

| Command | Description |
| :--- | :--- |
| `pnpm run build` | Bundles TypeScript files into production-ready `.gs` scripts in `dist/`. |
| `pnpm run deploy` | Runs `types`, builds with Vite, and pushes code to Google Apps Script via Clasp. |
| `pnpm run types` | Generates `.d.ts` definitions using `@antunes_s/clasp-types`. |
| `pnpm run dev` | Watches for changes and compiles automatically. |

---

## 📖 Usage Examples

### 1. Opening a Spreadsheet & Accessing Worksheets
```typescript
import { Spreadsheet } from "./Spreadsheet";

// Open the active spreadsheet or specify an ID: Spreadsheet.openById("...")
const ss = new Spreadsheet();

// Get or automatically create a sheet if it doesn't exist:
const ws = ss.getOrCreate("Vendas");
```

### 2. Reading Data as Objects
```typescript
// Read all rows as an array of objects based on header row (default: row 1):
const records = ws.getAllRecords();
// Output: [ { ID: 1, Produto: "Teclado", Preco: 150 }, ... ]

// Or group records by a specific column:
const grouped = ws.getAllRecords(1, "Turno");
// Output: { T1: [ ... ], T2: [ ... ] }
```

### 3. Dynamic Header Mapping (`getMappedRecords`)
Extract only the columns you need, resilient to column shifts and empty lines:

```typescript
const mapped = ws.getMappedRecords([
  { id: "nome", terms: ["Nome", "Colaborador", "Operador"], required: true },
  { id: "total", terms: ["Total", "Qtd", "Quantidade"], required: true },
  { id: "turno", terms: ["Turno"], required: false }
]);
```

### 4. In-Memory ETL Operations

#### `extract` (Slice columns by index or header name):
```typescript
import { extract } from "./assets/Utils";

const rawData = ws.getAllValues();

// Extract only columns A (0), C (2), and E (4):
const compactMatrix = extract(rawData, [0, 2, 4]);

// Or extract by header names (case & accent insensitive):
const report = extract(rawData, ["Data", "Colaborador", "Total"], 1);
```

#### `merge` (In-Memory VLOOKUP / LEFT JOIN):
Cross-reference two sheets in milliseconds without using slow `=VLOOKUP` formulas:

```typescript
import { merge } from "./assets/Utils";

const vendas = ss.worksheet("Vendas").getAllValues();     // [ID_Venda, Cliente_ID, Valor]
const clientes = ss.worksheet("Clientes").getAllValues(); // [Cliente_ID, Nome, Cidade]

// Merge by Cliente_ID:
const relatorio = merge(vendas, clientes, "Cliente_ID");
// Output header: [ID_Venda, Cliente_ID, Valor, Nome, Cidade]
```

### 5. Blazing-Fast Date & Text Helpers
```typescript
import { ensureDate, fmtData } from "./assets/Date";
import { normalizeTxt } from "./assets/Text";

// Safe date parsing (handles serial 45400, "15/04/2026", etc.):
const dateObj = ensureDate("15/04/2026");

// Fast formatting (DD/MM/YYYY):
const formatted = fmtData(dateObj); // "15/04/2026"
const withTime  = fmtData(dateObj, true); // "15/04/2026 14:30"

// Fuzzy text normalization:
normalizeTxt("  CÓDIGO - PRODUTO_A / 01  "); // "codigo produto a 01"
```

---

## 🚢 Using as a Google Apps Script Library

1. Run `pnpm run deploy` to upload the code to your Apps Script project.
2. In the Google Apps Script editor, click **Deploy > New deployment**, select type **Library**, and publish a version.
3. In any other Google Sheet script, go to **Libraries (+)**, enter your **Script ID**, and start using `JSspread`!

---

## 📄 License
MIT © [Gustavo Antunes](https://github.com/antune-ss)