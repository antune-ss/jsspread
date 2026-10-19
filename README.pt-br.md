# JSspread 📊

[English](README.md) | [Português](README.pt-br.md)

Uma biblioteca para o **Google Apps Script**, desenvolvida em **TypeScript** e empacotada com **Vite**. Fortemente inspirada no `gspread` do Python, a **JSspread** simplifica a leitura, escrita, filtragem e manipulação de dados de planilhas do Google Sheets com tipagem estrita, alta tolerância a falhas e operações em memória extremamente rápidas.

---

## 🚀 Recursos Principais

- **API Orientada a Objetos**: Classes intuitivas para `Spreadsheet`, `Worksheet` e `Cell`.
- **Mapeamento Dinâmico e Flexível**: `getMappedRecords` localiza colunas automaticamente mesmo se a ordem das colunas mudar, se houver erros de digitação ou linhas de título no topo.
- **Ferramentas de ETL em Memória**:
  - `extract()`: Recorta colunas específicas de matrizes 2D brutas ou listas de objetos por índice ou nome.
  - `merge()`: Executa **PROCV / LEFT JOIN** instantâneo em memória ($O(N + M)$) entre tabelas sem precisar de fórmulas lentas que travam a planilha.
- **Parsers Robustos de Data e Texto**:
  - `ensureDate()`: Converte com segurança números seriais de datas do Sheets (ex: `45400`), datas brasileiras (`DD/MM/YYYY`) e strings ISO.
  - `fmtData()`: Formatador de datas ultrarrápido otimizado para grandes matrizes (evita a lentidão do `Intl` e chamadas à API do Google).
  - `normalizeTxt()`: Normaliza acentuação, maiúsculas, espaços duplos e separadores (`-`, `_`, `/`).
- **Suporte de Primeira Classe ao TypeScript**: Gera automaticamente as definições de tipagem `.d.ts` para cliente e biblioteca via `@antunes_s/clasp-types`.

---

## 📦 Instalação e Configuração

### Pré-requisitos
- [Node.js](https://nodejs.org/) (v18+)
- [pnpm](https://pnpm.io/)
- [Google Clasp](https://github.com/google/clasp) com login ativo (`clasp login`)

### Instalação das Dependências
```bash
pnpm install
```

---

## 🛠️ Comandos Disponíveis

| Comando | Descrição |
| :--- | :--- |
| `pnpm run build` | Compila os arquivos TypeScript em scripts `.gs` prontos para produção na pasta `dist/`. |
| `pnpm run deploy` | Executa o gerador de tipos, compila com Vite e faz o upload para o Google Apps Script via Clasp. |
| `pnpm run types` | Gera as definições de tipos `.d.ts` usando o `@antunes_s/clasp-types`. |
| `pnpm run dev` | Observa alterações nos arquivos e compila automaticamente. |

---

## 📖 Exemplos Práticos de Uso

### 1. Abrindo uma Planilha e Acessando Abas
```typescript
import { Spreadsheet } from "./Spreadsheet";

// Abre a planilha ativa ou especifica um ID: Spreadsheet.openById("...")
const ss = new Spreadsheet();

// Obtém ou cria automaticamente a aba se ela não existir:
const ws = ss.getOrCreate("Vendas");
```

### 2. Lendo Dados como Objetos
```typescript
// Lê todas as linhas como um array de objetos usando o cabeçalho (padrão: linha 1):
const registros = ws.getAllRecords();
// Saída: [ { ID: 1, Produto: "Teclado", Preco: 150 }, ... ]

// Ou agrupa os registros por uma coluna específica (ex: Turno):
const agrupado = ws.getAllRecords(1, "Turno");
// Saída: { T1: [ ... ], T2: [ ... ] }
```

### 3. Mapeamento Dinâmico de Cabeçalhos (`getMappedRecords`)
Extraia apenas as colunas necessárias com tolerância a mudanças de posição ou linhas em branco no topo:

```typescript
const mapeado = ws.getMappedRecords([
  { id: "nome", terms: ["Nome", "Colaborador", "Operador"], required: true },
  { id: "total", terms: ["Total", "Qtd", "Quantidade"], required: true },
  { id: "turno", terms: ["Turno"], required: false }
]);
```

### 4. Operações de ETL em Memória

#### `extract` (Recorta colunas por índice ou nome):
```typescript
import { extract } from "./assets/Utils";

const dadosBrutos = ws.getAllValues();

// Extrai apenas as colunas A (0), C (2) e E (4):
const matrizReduzida = extract(dadosBrutos, [0, 2, 4]);

// Ou extrai buscando pelos nomes das colunas (sem sensibilidade a maiúsculas/acentos):
const relatorio = extract(dadosBrutos, ["Data", "Colaborador", "Total"], 1);
```

#### `merge` (PROCV / LEFT JOIN em Memória):
Cruze duas matrizes em milissegundos sem sobrecarregar a planilha com fórmulas `=PROCV()`:

```typescript
import { merge } from "./assets/Utils";

const vendas = ss.worksheet("Vendas").getAllValues();     // [ID_Venda, Cliente_ID, Valor]
const clientes = ss.worksheet("Clientes").getAllValues(); // [Cliente_ID, Nome, Cidade]

// Cruze as duas tabelas pelo Cliente_ID:
const relatorio = merge(vendas, clientes, "Cliente_ID");
// Cabeçalho da nova matriz consolidada: [ID_Venda, Cliente_ID, Valor, Nome, Cidade]
```

### 5. Utilitários Ultrarrápidos de Data e Texto
```typescript
import { ensureDate, fmtData } from "./assets/Date";
import { normalizeTxt } from "./assets/Text";

// Conversão segura de datas (trata serial 45400, "15/04/2026", objetos Date, etc.):
const dataObj = ensureDate("15/04/2026");

// Formatação veloz (DD/MM/YYYY):
const formatada = fmtData(dataObj); // "15/04/2026"
const comHora   = fmtData(dataObj, true); // "15/04/2026 14:30"

// Normalização flexível de texto:
normalizeTxt("  CÓDIGO - PRODUTO_A / 01  "); // "codigo produto a 01"
```

---

## 🚢 Usando como Biblioteca no Google Apps Script

1. Execute `pnpm run deploy` para enviar o código para o seu projeto no Apps Script.
2. No editor do Google Apps Script, clique em **Deploy > Nova implantação (New deployment)**, selecione o tipo **Biblioteca (Library)** e publique uma versão.
3. Em qualquer outra planilha sua, vá em **Bibliotecas (+)**, cole o seu **Script ID** e comece a usar a `JSspread`!

---

## 📄 Licença
MIT © [Gustavo Antunes](https://github.com/antune-ss)
