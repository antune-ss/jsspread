import { Worksheet } from './Worksheet.js';
import { ExportFormat } from './assets/Schemas.js';

/**
 * Inicializa o objeto da Planilha.
 * @class
 * @public
 */
export class Spreadsheet {
  /** @public */
  id: string | null;
  /** @public */
  file: GoogleAppsScript.Spreadsheet.Spreadsheet;

  /**
   * Inicializa o objeto da Planilha.
   * 
   * OBSERVAÇÃO: Esta classe atua como um "wrapper" (camada de abstração) que oferece 
   * métodos customizados inspirados na biblioteca gspread do Python. Como o objeto 
   * nativo do Google Apps Script é armazenado no atributo `this.file`, você ainda pode 
   * utilizar qualquer método original da engine do Google acessando-o diretamente 
   * (exemplo: `sh.file.addEditor('email@teste.com')`).
   * 
   * @param id - O ID da planilha. Se não fornecido, usa a planilha ativa.
   * @public
   */
  constructor(id: string | null = null) {
    this.id = id;
    this.file = id 
      ? SpreadsheetApp.openById(id)
      : SpreadsheetApp.getActiveSpreadsheet();
  }

  // ==========================================
  // Propriedades da Spreadsheet
  // ==========================================
  /**
   * Retorna a URL direta da planilha.
   * @public
   */
  get url(): string {
    return this.file.getUrl();
  }

  /**
   * Retorna o título (nome) do arquivo da planilha.
   * @public
   */
  get title(): string {
    return this.file.getName();
  }

  /**
   * Retorna a 1ª worksheet do arquivo da planilha.
   * @public
   */
  get sheet1(): Worksheet {
    const sheets = this.file.getSheets();
    if (sheets.length === 0) throw new Error("WorksheetNotFound: A planilha não possui abas.");
    return new Worksheet(this, sheets[0]);
  }

  // ==========================================
  // Métodos Públicos
  // ==========================================
  /** @public */
  worksheet(title: string): Worksheet {
    const sheet = this.file.getSheetByName(title);
    if (!sheet) {
      throw new Error(`WorksheetNotFound: Aba '${title}' não encontrada.`);
    }
    return new Worksheet(this, sheet);
  }

  /**
   * Retorna uma lista com todas as abas da planilha.
   * 
   * @param exclude_hidden - Se true, ignora as abas ocultas.
   * @public
   */
  worksheets(exclude_hidden: boolean = false): Worksheet[] {
    let sheets = this.file.getSheets();
    
    if (exclude_hidden) {
      sheets = sheets.filter(sheet => !sheet.isSheetHidden());
    }

    return sheets.map(sheet => new Worksheet(this, sheet));
  }

  /** @public */
  getWorksheet(index: number): Worksheet {
    const sheets = this.file.getSheets();
    
    if (index < 0 || index >= sheets.length) {
      throw new Error(`WorksheetNotFound: Não existe aba no índice ${index}. A planilha tem ${sheets.length} aba(s).`);
    }
    
    return new Worksheet(this, sheets[index]);
  }

  /** @public */
  getWorksheetById(id: number): Worksheet {
    const sheets = this.file.getSheets();
    const sheet = sheets.find(s => s.getSheetId() === id);

    if (!sheet) {
      throw new Error(`WorksheetNotFound: Aba com Id '${id}' não encontrada.`);
    }

    return new Worksheet(this, sheet);
  }

  /** @public */
  addWorksheet(title: string, rows: number = 1000, cols: number = 26, index: number | null = null): Worksheet {
    let sheet: GoogleAppsScript.Spreadsheet.Sheet;

    try {
      if (index !== null) {
        sheet = this.file.insertSheet(title, index);
      } else {
        sheet = this.file.insertSheet(title);
      }
    } catch (e) {
      throw new Error(`Erro ao criar aba: Já existe uma aba com o nome '${title}'.`);
    }

    // --- Ajusta a quantidade de linhas ---
    const currentRows = sheet.getMaxRows();
    if (rows && rows < currentRows) {
      sheet.deleteRows(rows + 1, currentRows - rows);
    } else if (rows && rows > currentRows) {
      sheet.insertRowsAfter(currentRows, rows - currentRows);
    }

    // --- Ajusta a quantidade de colunas ---
    const currentCols = sheet.getMaxColumns();
    if (cols && cols < currentCols) {
      sheet.deleteColumns(cols + 1, currentCols - cols);
    } else if (cols && cols > currentCols) {
      sheet.insertColumnsAfter(currentCols, cols - currentCols);
    }

    return new Worksheet(this, sheet);
  }

  /** @public */
  delWorksheet(worksheet: Worksheet | string): void {
    let sheet: GoogleAppsScript.Spreadsheet.Sheet | null = null;
    if (worksheet instanceof Worksheet) {
      sheet = worksheet.sheet;
    } else if (typeof worksheet === 'string') {
      sheet = this.file.getSheetByName(worksheet);
      
      if (!sheet) {
        throw new Error(`WorksheetNotFound: Aba '${worksheet}' não encontrada para exclusão.`);
      }
    } else {
      throw new Error("TypeError: O parâmetro deve ser uma string (nome da aba) ou um objeto Worksheet.");
    }

    this.file.deleteSheet(sheet);
  }

  /**
   * Tenta acessar uma aba pelo título. Se ela não existir, cria e a retorna.
   * 
   * @param title - O nome da aba desejada.
   * @param rows - (Opcional) Número de linhas caso a aba seja criada.
   * @param cols - (Opcional) Número de colunas caso a aba seja criada.
   * @param index - (Opcional) Posição da aba caso seja criada.
   * @public
   */
  getOrCreate(title: string, rows: number = 1000, cols: number = 26, index: number | null = null): Worksheet {
    try {
      return this.worksheet(title);
    } catch (e) {
      return this.addWorksheet(title, rows, cols, index);
    }
  }

  /**
   * Exporta a planilha (ou uma aba específica) no formato especificado.
   * 
   * @param format - O formato de exportação (use ExportFormat).
   * @param worksheet - (Opcional) A aba específica que deseja exportar.
   * @returns O arquivo exportado.
   * @public
   */
  export(format: string = ExportFormat.pdf, worksheet: Worksheet | null = null): GoogleAppsScript.Base.Blob {
    if (!Object.values(ExportFormat).includes(format as ExportFormat)) {
      throw new Error(`ValueError: Formato '${format}' inválido.`);
    }
    
    let url = `https://docs.google.com/spreadsheets/d/${this.id}/export?format=${format}`;
    
    if (worksheet) {
      url += `&gid=${worksheet.id}`;
    }

    const options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
      headers: { Authorization: `Bearer ${ScriptApp.getOAuthToken()}` },
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(url, options);

    if (response.getResponseCode() !== 200) {
      throw new Error(`Erro ao exportar planilha: ${response.getContentText()}`);
    }

    const blob = response.getBlob();
    
    const fileName = worksheet ? `${this.file.getName()} - ${worksheet.title}` : this.file.getName();
    blob.setName(`${fileName}.${format}`);
    
    return blob;
  }
}
