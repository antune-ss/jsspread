import { Spreadsheet } from './Spreadsheet.js';
import { Cell } from './Cell.js';
import { ExportFormat, ValueInputOption, InsertDataOption, ValueRenderOption } from './assets/Schemas.js';
import { normalizeTxt } from './assets/Text.js';

/** @public */
export interface InsertOptions {
  valueInputOption?: ValueInputOption;
  insertDataOption?: InsertDataOption;
}

/** @public */
export interface UpdateOptions {
  valueInputOption?: ValueInputOption;
}

/** @public */
export interface RenderOptions {
  valueRenderOption?: InsertDataOption;
}

/** @public */
export interface FilterRule {
  /** O índice da coluna (começa em 0). Ex: Coluna A = 0 */
  col: number;
  /** O valor exato para filtrar, ou um Array de valores aceitos. */
  val: string | number | boolean | (string | number | boolean)[];
  /** Se true, o filtro será parcial (ex: procurar "João" dentro de "João Silva"). Padrão: false */
  partial?: boolean;
}

/** @public */
export interface HeaderConfig {
  /** O nome da chave que vai aparecer no seu Objeto final (Ex: 'idade') */
  id: string;
  /** Lista de palavras aceitas para achar essa coluna (Ex: ['Idade', 'Idades', 'Anos']) */
  terms: string[];
  /** Se true, a função aborta caso não encontre esta coluna na planilha */
  required?: boolean;
}

/**
 * Inicializa o objeto da aba.
 * @class
 * @public
 */
export class Worksheet {
  /** @public */
  spreadsheet: Spreadsheet;
  /** @public */
  sheet: GoogleAppsScript.Spreadsheet.Sheet;

  /**
   * Inicializa o objeto da Aba.
   * 
   * OBSERVAÇÃO: Assim como a Spreadsheet, esta classe é um wrapper. 
   * A aba nativa do Google Apps Script fica armazenada no atributo `this.sheet`.
   * Você pode acessar os métodos nativos usando `ws.sheet.getRange(...)`, etc.
   * 
   * @param spreadsheet - A instância da classe Spreadsheet principal.
   * @param gasSheet - O objeto nativo da aba do Google Apps Script.
   * @public
   */
  constructor(spreadsheet: Spreadsheet, gasSheet: GoogleAppsScript.Spreadsheet.Sheet) {
    this.spreadsheet = spreadsheet;
    this.sheet = gasSheet;
  }

  // ==========================================
  // Propriedades da Worksheet
  // ==========================================
  /**
   * Retorna a URL direta da aba.
   * @public
   */
  get url(): string {
    const spreadsheetUrl = this.sheet.getParent().getUrl();
    const sheetId = this.sheet.getSheetId();
    return `${spreadsheetUrl}#gid=${sheetId}`;
  }

  /**
   * Retorna o ID único (GID) da aba.
   * @public
   */
  get id(): string {
    return this.sheet.getSheetId().toString();
  }

  /**
   * Retorna o nome da aba (como se fosse um atributo).
   * @public
   */
  get title(): string {
    return this.sheet.getName();
  }

  /**
   * Retorna o índice numérico da aba na planilha (começa em 0 ou 1 dependendo da API, 
   * no GAS getIndex() começa em 1).
   * @public
   */
  get index(): number {
    return this.sheet.getIndex();
  }

  /**
   * Retorna o número total de linhas alocadas na aba (Não conta as linhas em branco).
   * @public
   */
  get row_count(): number {
    return this.sheet.getLastRow();
  }

  /**
   * Retorna o número total de colunas alocadas na aba (Não conta as colunas em branco).
   * @public
   */
  get column_count(): number {
    return this.sheet.getLastColumn();
  }
  
  // ==========================================
  // Métodos Públicos
  // ==========================================
  /**
   * Exporta exclusivamente esta aba no formato especificado.
   * 
   * @param format - O formato de exportação (use ExportFormat).
   * @returns O arquivo exportado.
   * @public
   */
  export(format: string = ExportFormat.pdf): GoogleAppsScript.Base.Blob {
    return this.spreadsheet.export(format, this);
  }

  /**
   * Retorna todos os valores de uma linha específica.
   * @param index - O número da linha (começa em 1).
   * @public
   */
  rowValues(index: number): any[] {
    const lastCol = this.sheet.getLastColumn();
    if (lastCol === 0) return [];

    const data = this.sheet.getRange(index, 1, 1, lastCol).getValues();
    return data[0];
  }

  /**
   * Retorna todos os valores de uma coluna específica.
   * @param index - O número da coluna (começa em 1).
   * @public
   */
  colValues(index: number): any[] {
    const lastRow = this.sheet.getLastRow();
    if (lastRow === 0) return [];

    const data = this.sheet.getRange(1, index, lastRow, 1).getValues();
    return data.map(row => row[0]);
  }

  /**
   * Retorna os valores de um intervalo específico da planilha.
   * 
   * @param range - Intervalo em notação A1 (ex: "A1:C5").
   * @returns Uma matriz bidimensional com os valores do intervalo.
   * @public
   */
  getWithRange(range: string): any[][] {
    return this.sheet.getRange(range).getValues();
  }

  /**
   * Retorna todos os valores da planilha como uma lista de listas.
   * @public
   */
  getAllValues(): any[][] {
    const data = this.sheet.getDataRange().getValues();
    if (data.length === 0 || (data.length === 1 && data[0][0] === "")) return [];

    return data;
  }

  /**
   * Retorna todos os valores da planilha como uma lista de objetos mapeados pelo cabeçalho.
   * Se a chave de agrupamento (groupKey) for fornecida, retorna um Dicionário contendo Arrays de objetos.
   * 
   * @param head - A linha que contém o cabeçalho (começa em 1). Padrão: 1.
   * @param groupKey - (Opcional) Nome do cabeçalho para agrupar as linhas (Ex: "Turno"). Padrão: null.
   * @param empty2zero - (Opcional) Se true, converte as células em branco mapeadas para o número 0. Padrão: false.
   * @param default_blank - (Opcional) Valor padrão a ser inserido caso a célula esteja em branco (se empty2zero for false). Padrão: "".
   * @public
   */
  getAllRecords(head: number = 1, groupKey: string | null = null, empty2zero: boolean = false, default_blank: any = ""): Record<string, any>[] | Record<string, Record<string, any>[]> {
    const data = this.sheet.getDataRange().getValues();
    if (data.length < head) return [];

    const headers = data[head - 1];
    const rows = data.slice(head);

    let groupKeyIdx = -1;
    if (groupKey) {
      groupKeyIdx = headers.indexOf(groupKey);
      if (groupKeyIdx === -1) {
        throw new Error(`Worksheet Error: Cabeçalho '${groupKey}' não encontrado na linha de cabeçalho.`);
      }
    }

    const arrayResult: Record<string, any>[] = [];
    const groupResult: Record<string, Record<string, any>[]> = {};

    rows.forEach(row => {
      if (row.every(cell => cell === "" || cell === null)) return;
      
      let record: Record<string, any> = {};
      headers.forEach((header, index) => {
        if (header !== "") {
          let value = row[index];

          if (value === "") {
            value = empty2zero ? 0 : default_blank;
          }

          record[header] = value
        }
      });

      if (groupKeyIdx !== -1) {
        const keyValue = String(row[groupKeyIdx]).trim().toUpperCase();
        if (keyValue) {
          if (!groupResult[keyValue]) {
            groupResult[keyValue] = [];
          }

          groupResult[keyValue].push(record)
        }
      } else {
        arrayResult.push(record);
      }
    });

    return groupKeyIdx !== -1 ? groupResult : arrayResult;
  }

  /**
   * Mapeia os dados da planilha de forma dinâmica e inteligente, buscando pelos cabeçalhos configurados.
   * Super resiliente a mudanças de posição das colunas e a linhas em branco no topo da planilha.
   * 
   * @param config - Array contendo as configurações das colunas (id e termos) a serem extraídas.
   * @param groupKey - (Opcional) O 'id' configurado que servirá como chave para agrupar o resultado em um Dicionário. Padrão: null.
   * @param empty2zero - (Opcional) Se true, converte as células em branco mapeadas para o número 0. Padrão: false.
   * @param default_blank - (Opcional) Valor padrão a ser inserido caso a célula esteja em branco (se empty2zero for false). Padrão: "".
   * @param searchLimit - (Opcional) Quantidade máxima de linhas para varrer no topo da planilha procurando os cabeçalhos. Padrão: 40.
   * @returns Retorna uma Lista de Objetos ou, se um groupKey for informado, um Dicionário contendo Listas de Objetos.
   * @throws Lança um Erro caso os cabeçalhos obrigatórios ou a chave de agrupamento não sejam encontrados na planilha.
   * @public
   */
  getMappedRecords(config: HeaderConfig[], groupKey: string | null = null, empty2zero: boolean = false, default_blank: any = "", searchLimit: number = 40): Record<string, any>[] | Record<string, Record<string, any>[]> {
    const data = this.getAllValues();
    if (!data || data.length === 0) return [];

    const limit = Math.min(data.length, searchLimit);
    let headerRowIdx = -1;
    let cols: Record<string, number> = {};

    for(let i = 0; i < limit; i++) {
      const norm = data[i].map(c => normalizeTxt(c));
      let findedCols: Record<string, number> = {};
      let findAll = true;

      for(const conf of config) {
        let index = norm.findIndex(cell => conf.terms.some(term => cell === normalizeTxt(term)));

        if (index === -1) {
          index = norm.findIndex(cell => conf.terms.some(term => cell.includes(normalizeTxt(term))));
        }

        findedCols[conf.id] = index;

        if (conf.required && index === -1) {
          findAll = false;
          break;
        }
      }

      if (findAll) {
        headerRowIdx = i;
        cols = findedCols;
        break;
      }
    }

    if (headerRowIdx === -1) {
      throw new Error(`Worksheet Error (getMappedRecords): Cabeçalhos obrigatórios não foram encontrados nas primeiras ${searchLimit} linhas.`);
    }

    if (groupKey && (cols[groupKey] === undefined || cols[groupKey] === -1)) {
      throw new Error(`Worksheet Error (getMappedRecords): A coluna chave '${groupKey}' não foi localizada.`);
    }

    const arrayResult: Record<string, any>[] = [];
    const groupResult: Record<string, Record<string, any>[]> = {};

    for (let i = headerRowIdx + 1; i < data.length; i++) {
      const row = data[i];
      if (row.every(cell => cell === "" || cell === null)) continue;

      let record: Record<string, any> = {};

      for (const conf of config) {
        const colIdx = cols[conf.id];
        let value = (colIdx !== undefined && colIdx !== -1) ? row[colIdx] : null;

        if (value === "") {
          value = empty2zero ? 0 : default_blank;
        }

        record[conf.id] = value;
      }

      if (groupKey) {
        const keyIdx = cols[groupKey];
        const keyValue = String(row[keyIdx]).trim().toUpperCase();

        if (keyValue) {
          if (!groupResult[keyValue]) groupResult[keyValue] = [];
          groupResult[keyValue].push(record);
        }
      } else {
        arrayResult.push(record);
      }
    }

    return groupKey ? groupResult : arrayResult;
  }

  /**
   * Insere uma linha de dados em uma posição específica, empurrando as demais para baixo.
   * 
   * @param index - A posição onde a linha será inserida (começa em 1).
   * @param values - Uma lista simples com os dados da linha.
   * @param options - Configurações opcionais.
   * @public
   */
  insertRow(index: number, values: any[], options: UpdateOptions = {}): Worksheet {
    this._isSimpleArray(values, "TypeError: insertRow aceita apenas uma lista simples. Para múltiplas linhas, use insertRows()");
    const valueInputOption = this._verifyInputOption(options.valueInputOption);

    this.sheet.insertRowBefore(index);

    const targetRange = this.sheet.getRange(index, 1, 1, values.length);
    this._applyValues(targetRange, [values], valueInputOption);

    return this;
  }

  /**
   * Insere múltiplas linhas de dados em uma posição específica.
   * 
   * @param index - A posição inicial onde as linhas serão inseridas (começa em 1).
   * @param rows - Matriz (lista de listas) com os dados.
   * @param options - Configurações opcionais.
   * @public
   */
  insertRows(index: number, rows: any[][], options: UpdateOptions = {}): Worksheet {
    this._isValidMatrix(rows, "TypeError: insertRows exige uma lista de listas (matriz). Para uma única linha, use insertRow()");
    const valueInputOption = this._verifyInputOption(options.valueInputOption);

    const rowCount = rows.length;
    const maxColumns = Math.max(...rows.map(r => Array.isArray(r) ? r.length : 0));

    this.sheet.insertRowsBefore(index, rowCount);

    const targetRange = this.sheet.getRange(index, 1, rowCount, maxColumns);
    const normalizedMatrix = this._normalizeMatrix(rows, maxColumns);

    this._applyValues(targetRange, normalizedMatrix, valueInputOption);

    return this;
  }

  /**
   * Adiciona uma única linha de dados ao final da planilha.
   * 
   * @param values - Uma lista simples com os dados da linha (ex: ['João', 25]).
   * @param options - Configurações opcionais.
   * @public
   */
  appendRow(values: any[], options: InsertOptions = {}): Worksheet {
    this._isSimpleArray(values, "TypeError: appendRow aceita apenas uma lista simples. Para múltiplas linhas, use appendRows().");

    const valueInputOption = this._verifyInputOption(options.valueInputOption);
    const insertDataOption = this._verifyInsertOption(options.insertDataOption);

    const lastRow = this.sheet.getLastRow();

    if (insertDataOption === InsertDataOption.insert_rows && lastRow > 0) {
      this.sheet.insertRowAfter(lastRow);
    }

    const targetRow = lastRow + 1;
    const targetRange = this.sheet.getRange(targetRow, 1, 1, values.length);

    this._applyValues(targetRange, [values], valueInputOption);

    return this;
  }

  /**
   * Adiciona múltiplas linhas de dados de uma só vez ao final da planilha.
   * 
   * @param rows - Uma lista de listas contendo as linhas (ex: [['Maria', 30], ['Pedro', 22]]).
   * @param options - Configurações opcionais.
   * @public
   */
  appendRows(rows: any[][], options: InsertOptions = {}): Worksheet {
    this._isValidMatrix(rows, "TypeError: appendRows exige uma lista de listas (matriz). Para uma única linha, use appendRow().");

    const valueInputOption = this._verifyInputOption(options.valueInputOption);
    const insertDataOption = this._verifyInsertOption(options.insertDataOption);

    const rowCount = rows.length;
    const maxColumns = Math.max(...rows.map(r => Array.isArray(r) ? r.length : 0));
    const lastRow = this.sheet.getLastRow();

    if (insertDataOption === InsertDataOption.insert_rows && lastRow > 0) {
      this.sheet.insertRowsAfter(lastRow, rowCount);
    }

    const startRow = lastRow + 1;
    const targetRange = this.sheet.getRange(startRow, 1, rowCount, maxColumns);

    const normalizedMatrix = this._normalizeMatrix(rows, maxColumns);

    this._applyValues(targetRange, normalizedMatrix, valueInputOption);

    return this;
  }

  /**
   * Atualiza um intervalo de células com novos valores.
   * 
   * @param range - Intervalo em notação A1 (ex: "A1" ou "A1:B2")
   * @param values - Matriz de valores (lista de listas).
   * @param options - Configurações opcionais.
   * @public
   */
  update(range: string, values: any[][], options: UpdateOptions = {}): Worksheet {
    this._isValidMatrix(values, "TypeError: O método update exige uma matriz de valores (lista de listas).");

    const valueInputOption = this._verifyInputOption(options.valueInputOption);

    const numRows = values.length;
    const maxCols = Math.max(...values.map(r => Array.isArray(r) ? r.length : 0));

    const initialRange = this.sheet.getRange(range);
    const startRow = initialRange.getRow();
    const startCol = initialRange.getColumn();

    const targetRange = this.sheet.getRange(startRow, startCol, numRows, maxCols);

    const normalizedMatrix = this._normalizeMatrix(values, maxCols);
    
    this._applyValues(targetRange, normalizedMatrix, valueInputOption);

    return this;
  }

  /**
   * Atualiza o valor de uma célula específica passando linha e coluna.
   * 
   * @param row - O número da linha (começa em 1).
   * @param col - O número da coluna (começa em 1).
   * @param value - O valor a ser inserido na célula.
   * @param options - Configurações opcionais.
   * @public
   */
  updateCell(row: number, col: number, value: any, options: UpdateOptions = {}): Worksheet {
    const valueInputOption = this._verifyInputOption(options.valueInputOption);
    const cellRange = this.sheet.getRange(row, col);

    this._applyValues(cellRange, value, valueInputOption);

    return this;
  }

  /**
   * Atualiza o valor de uma célula específica usando notação A1.
   * 
   * @param range - Intervalo em notação A1 (ex: "A1" ou "A1:B2")
   * @param value - O valor a ser inserido na célula.
   * @param options - Configurações opcionais.
   * @public
   */
  updateACell(range: string, value: any, options: UpdateOptions = {}): Worksheet {
    const valueInputOption = this._verifyInputOption(options.valueInputOption);
    const cellRange = this.sheet.getRange(range);
    
    this._applyValues(cellRange, value, valueInputOption);

    return this;
  }

  /**
   * Retorna um objeto Cell correspondente a uma célula específica usando notação A1.
   * 
   * @param label - A notação A1 da célula (ex: "B1").
   * @param options - Configurações opcionais.
   * @public
   */
  acell(label: string, options: RenderOptions = {}): Cell {
    const renderOption = this._verifyRenderOption(options.valueRenderOption);

    const range = this.sheet.getRange(label);
    const row = range.getRow();
    const col = range.getColumn();

    let cellValue: any;

    if (renderOption === ValueRenderOption.formula) {
      cellValue = range.getFormula() || range.getValue();
    } else {
      cellValue = renderOption === ValueRenderOption.formatted 
                  ? range.getDisplayValue() 
                  : range.getValue();
    }

    return new Cell(row, col, cellValue);
  }

  /**
   * Retorna um objeto Cell correspondente a uma célula específica usando linha e coluna.
   * 
   * @param row - O número da linha (começa em 1).
   * @param col - O número da coluna (começa em 1).
   * @param options - Configurações opcionais.
   * @public
   */
  cell(row: number, col: number, options: RenderOptions = {}): Cell {
    const renderOption = this._verifyRenderOption(options.valueRenderOption);
    
    const range = this.sheet.getRange(row, col);
    
    let cellValue: any;

    // Ajuste para suportar tanto o enum quanto a string bruta devido à transição de JS
    if (renderOption === ValueRenderOption.formula || renderOption === 'FORMULA') {
      cellValue = range.getFormula() || range.getValue();
    } else {
      cellValue = (renderOption === ValueRenderOption.formatted || renderOption === 'FORMATTED_VALUE')
                  ? range.getDisplayValue() 
                  : range.getValue();
    }
    
    return new Cell(row, col, cellValue);
  }

  /**
   * Encontra a primeira célula que contém o valor ou que corresponda ao Regex.
   * 
   * @param query - O valor ou expressão regular a ser buscada.
   * @public
   */
  find(query: string | number | RegExp | any): Cell | null {
    if (query === undefined || query === null) return null;

    const isRegex = query instanceof RegExp;

    // Se for Regex, pega apenas o padrão em texto (a propriedade .source)
    const searchString = isRegex ? query.source : query.toString();

    const textFinder = this.sheet.createTextFinder(searchString);

    if (isRegex) {
      textFinder.useRegularExpression(true);
      textFinder.matchCase(!query.ignoreCase);
    }

    const result = textFinder.findNext();

    if (!result) {
      return null;
    }
    
    return new Cell(result.getRow(), result.getColumn(), result.getValue());
  }

  /**
   * Encontra todas as células que contêm o valor ou que correspondam ao Regex.
   * 
   * @param query - O valor ou expressão regular a ser buscada.
   * @public
   */
  findAll(query: string | number | RegExp | any): Cell[] {
    if (query === undefined || query === null) return [];

    const isRegex = query instanceof RegExp;
    const searchString = isRegex ? query.source : query.toString();

    const textFinder = this.sheet.createTextFinder(searchString);

    if (isRegex) {
      textFinder.useRegularExpression(true);
      textFinder.matchCase(!query.ignoreCase);
    }

    const results = textFinder.findAll();
    
    if (results.length === 0) {
      return [];
    }
    
    return results.map(range => 
      new Cell(range.getRow(), range.getColumn(), range.getValue())
    );
  }

  /**
   * Exclui uma ou mais linhas da planilha fisicamente.
   * 
   * @param index - O número da linha a ser excluída (começa em 1).
   * @param numRows - A quantidade de linhas a excluir a partir do index.
   * @public
   */
  deleteRows(index: number, numRows: number = 1): Worksheet {
    this.sheet.deleteRows(index, numRows);
    return this;
  }

  /**
   * Exclui uma ou mais colunas da planilha fisicamente.
   * 
   * @param index - O número da coluna a ser excluída (começa em 1).
   * @param numCols - A quantidade de colunas a excluir a partir do index.
   * @public
   */
  deleteColumns(index: number, numCols: number = 1): Worksheet {
    this.sheet.deleteColumns(index, numCols);
    return this;
  }

  /**
   * Limpa várias áreas específicas da planilha de uma só vez.
   * 
   * @param ranges - Lista de strings em notação A1 (ex: ["A1:B2", "D5:E10"])
   * @param keepFormat - Se true, apaga apenas os valores. Se false, apaga valores e formatação.
   * @public
   */
  batchClear(ranges: string[], keepFormat: boolean = true): Worksheet {
    if (!Array.isArray(ranges) || ranges.length === 0) {
      return this; 
    }

    const rangeList = this.sheet.getRangeList(ranges);

    if (keepFormat) {
      rangeList.clearContent(); 
    } else {
      rangeList.clear(); 
    }

    return this;
  }

  /**
   * Limpa a aba inteira.
   * 
   * @param keepFormat - Se true, apaga apenas os valores. Se false, apaga valores e formatação.
   * @public
   */
  clear(keepFormat: boolean = true): Worksheet {
    if (keepFormat) {
      this.sheet.clearContents();
    } else {
      this.sheet.clear();
    }
    
    return this;
  }

  /**
   * Redimensiona a planilha para ter exatamente o número de linhas e colunas especificado.
   * 
   * @param rows - Número total de linhas desejado.
   * @param cols - Número total de colunas desejado.
   * @public
   */
  resize(rows: number, cols: number): Worksheet {
    const currentRows = this.sheet.getMaxRows();
    const currentCols = this.sheet.getMaxColumns();

    if (rows < currentRows) {
      this.sheet.deleteRows(rows + 1, currentRows - rows);
    } else if (rows > currentRows) {
      this.sheet.insertRowsAfter(currentRows, rows - currentRows);
    }

    if (cols < currentCols) {
      this.sheet.deleteColumns(cols + 1, currentCols - cols);
    } else if (cols > currentCols) {
      this.sheet.insertColumnsAfter(currentCols, cols - currentCols);
    }

    return this;
  }

  /** @public */
  getData(range: string | null = null, getDisplayValues: boolean = false): void {
    // Método mantido vazio como no código original
  }

  /**
   * Copia todos os dados desta worksheet para uma worksheet de destino.
   * Se filtros forem fornecidos, apenas as linhas correspondentes serão copiadas.
   * Os dados atuais da aba de destino serão completamente apagados antes da colagem.
   * 
   * @param destWs - A aba (Worksheet) de destino para onde os dados serão transferidos.
   * @param filters - (Opcional) Um array de regras (FilterRule) para filtrar quais linhas devem ser copiadas. Padrão: [].
   * @param head - (Opcional) O número da linha que atua como cabeçalho (começa em 1). Linhas acima do cabeçalho serão ignoradas. Padrão: 1.
   * @returns Retorna a própria instância da aba de origem (Worksheet) para permitir o encadeamento de métodos.
   * @public
   */
  copyTo(destWs: Worksheet,  filters: FilterRule[] = [], head: number = 1): Worksheet {
    const data = this.getAllValues();
    
    if (!data || data.length === 0) {
      destWs.clear();
      return this;
    }

    let result = data;
    if (filters && filters.length > 0) {
      const headers = data[head - 1];
      const rows = data.slice(head);

      const filtered = rows.filter(row => {
        return filters.every(f => {
          const cell = row[f.col];
          if (cell === "" || cell === null || cell === undefined) return false;

          if (f.val !== null && f.val !== undefined) {
            const cellStr = String(cell).trim();
            const isPartial = f.partial === true;

            if (Array.isArray(f.val)) {
              return f.val.some(v => {
                const vStr = String(v).trim();
                return isPartial ? cellStr.includes(vStr) : cellStr === vStr;
              });
            } else {
              const vStr = String(f.val).trim();
              return isPartial ? cellStr.includes(vStr) : cellStr === vStr;
            }
          }

          return true;
        });
      });

      result = [headers, ...filtered];
    }

    destWs.clear();
    if (result.length > 0 && result[0].length > 0) {
      destWs.sheet.getRange(1, 1, result.length, result[0].length).setValues(result)
    }

    return this;
  }

  // ==========================================
  // Métodos Internos (Auxiliares)
  // ==========================================
  /**
   * Verifica se o valor fornecido é uma matriz unidimensional (lista simples).
   */
  private _isSimpleArray(array: any, errorMsg: string | null = null): boolean {
    if (!Array.isArray(array) || array.length === 0) {
      if (errorMsg) throw new Error(errorMsg);
      return false;
    }
    
    if (Array.isArray(array[0])) {
      if (errorMsg) throw new Error(errorMsg);
      return false;
    }

    return true;
  }

  /**
   * Verifica se o valor fornecido é uma matriz bidimensional (lista de listas).
   */
  private _isValidMatrix(matrix: any, errorMsg: string | null = null): boolean {
    if (!Array.isArray(matrix) || matrix.length === 0) {
      if (errorMsg) throw new Error(errorMsg);
      return false;
    }
    
    if (!Array.isArray(matrix[0])) {
      if (errorMsg) throw new Error(errorMsg);
      return false;
    }

    return true;
  }

  /**
   * Padroniza uma matriz bidimensional, garantindo que todas as linhas tenham o mesmo comprimento.
   */
  private _normalizeMatrix(matrix: any[][], maxColumns: number): any[][] {
    return matrix.map(row => {
      const paddedRow = Array.isArray(row) ? [...row] : [row];
      while (paddedRow.length < maxColumns) {
        paddedRow.push("");
      }
      return paddedRow;
    });
  }

  /**
   * Aplica dados a um intervalo, lidando com a formatação RAW ou USER_ENTERED.
   */
  private _applyValues(range: GoogleAppsScript.Spreadsheet.Range, data: any | any[][], valueInputOption: string): void {
    const isMatrix = Array.isArray(data) && Array.isArray(data[0]);

    if (valueInputOption === ValueInputOption.raw) {
      range.setNumberFormat('@');

      if (isMatrix) {
        const stringMatrix = (data as any[][]).map(row => 
          row.map(val => (val !== null && val !== undefined) ? val.toString() : "")
        );
        range.setValues(stringMatrix);
      } else {
        const stringValue = (data !== null && data !== undefined) ? data.toString() : "";
        range.setValue(stringValue);
      }
    } else {
      if (isMatrix) {
        range.setValues(data as any[][]);
      } else {
        range.setValue(data);
      }
    }
  }

  private _verifyInputOption(valueInputOption?: string): string {
    if (!valueInputOption) {
      return ValueInputOption.user_entered;
    }

    const option = valueInputOption.toUpperCase();

    if (!Object.values(ValueInputOption).includes(option as any)) {
      throw new Error(`ValueError: Opção recebida '${valueInputOption}' é inválida. Use o objeto ValueInputOption.`);
    }

    return option;
  }

  private _verifyInsertOption(insertDataOption?: string): string {
    if (!insertDataOption) {
      return InsertDataOption.insert_rows;
    }

    const option = insertDataOption.toUpperCase();

    if (!Object.values(InsertDataOption).includes(option as any)) {
      throw new Error(`ValueError: Opção recebida '${insertDataOption}' é inválida. Use o objeto InsertDataOption.`);
    }

    return option;
  }

  private _verifyRenderOption(valueRenderOption?: string): string {
    if (!valueRenderOption) {
      return ValueRenderOption.formatted;
    }

    const option = valueRenderOption.toUpperCase();

    if (!Object.values(ValueRenderOption).includes(option as any) && option !== 'FORMULA' && option !== 'FORMATTED_VALUE' && option !== 'UNFORMATTED_VALUE') {
      throw new Error(`ValueError: Opção recebida '${valueRenderOption}' é inválida. Use o objeto ValueRenderOption.`);
    }

    return option;
  }
}
