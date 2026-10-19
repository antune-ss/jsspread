// Type definitions for JSspread
// Generated using clasp-types

declare namespace jsspread {

    /**
     * The main entry point to interact with JSspread
     * 
     * Script ID: **1mJmoX64rhmKabUDVYfPVt5xqaXEbS1JRTJCCMaTcqr9A5v_cmLiIo8ci**
     */
    export interface JSspread {

        ValueRenderOption: typeof ValueRenderOption;

        ValueInputOption: typeof ValueInputOption;

        InsertDataOption: typeof InsertDataOption;

        ExportFormat: typeof ExportFormat;

        /**
         * Converte de forma segura qualquer formato de data e hora para um objeto Date válido.
         * Utiliza tolerância para números seriais do Sheets e strings BR/ISO.
         *
         * @returns O objeto Date validado ou null se for impossível converter.
         */
        ensureDate(value: any): Date | null;

        /**
         * Formata datas de para strings de forma rápida.
         * Não utiliza a API do Google e nem o Intl do JS.
         *
         * @returns A data formatada em String no formato br, ou null em caso de entrada inválida.
         */
        fmtData(value: any, includeTime: boolean): string | null;

        /**
         * Normaliza textos para comparação e busca.
         * Remove acentos, capitalização e reduz múltiplos espaços.
         *
         * @returns O texto normalizado.
         */
        normalizeTxt(text: any, cleanSeparators: boolean): string;

        /**
         * Sobrecarga 1: Extrai propriedades específicas de uma Lista de Objetos.
         */
        extract<T extends Record<string, any>>(data: T[], columns: (string | keyof T)[]): Partial<T>[];

        /**
         * Sobrecarga 2: Extrai colunas específicas de uma Matriz 2D (por índices ou nomes de cabeçalho).
         */
        extract(data: any[][], columns: (string | number)[], head?: number): any[][];

        /**
         * Sobrecarga 1: Mescla duas Listas de Objetos usando uma propriedade chave (LEFT JOIN).
         */
        merge<T extends Record<string, any>, U extends Record<string, any>>(target: T[], source: U[], keyTarget: string | keyof T, keySource?: string | keyof U): (T & Partial<U>)[];

        /**
         * Sobrecarga 2: Mescla duas Matrizes 2D anexando as colunas da matriz secundária (PROCV em memória).
         */
        merge(target: any[][], source: any[][], keyTarget: string | number, keySource?: string | number, headTarget?: number, headSource?: number): any[][];

    }

    /**
     * Inicializa o objeto da célula.
     */
    export interface Cell {

        col: number;

        row: number;

        value: any;

        /**
         * Converte a célula em um objeto JSON puro.
         */
        toJSON(): object;

        /**
         * Retorna uma representação em string da célula para debug.
         */
        toString(): string;

    }

    /**
     * Inicializa o objeto da Planilha.
     */
    export interface Spreadsheet {

        file: GoogleAppsScript.Spreadsheet.Spreadsheet;

        id: string | null;

        addWorksheet(title: string, rows: number, cols: number, index: number | null): Worksheet;

        delWorksheet(worksheet: string | Worksheet): void;

        /**
         * Exporta a planilha (ou uma aba específica) no formato especificado.
         *
         * @returns O arquivo exportado.
         */
        export(format: string, worksheet: Worksheet | null): GoogleAppsScript.Base.Blob;

        /**
         * Tenta acessar uma aba pelo título. Se ela não existir, cria e a retorna.
         */
        getOrCreate(title: string, rows: number, cols: number, index: number | null): Worksheet;

        getWorksheet(index: number): Worksheet;

        getWorksheetById(id: number): Worksheet;

        worksheet(title: string): Worksheet;

        /**
         * Retorna uma lista com todas as abas da planilha.
         */
        worksheets(exclude_hidden: boolean): Worksheet[];

    }

    /**
     * Inicializa o objeto da aba.
     */
    export interface Worksheet {

        sheet: GoogleAppsScript.Spreadsheet.Sheet;

        spreadsheet: Spreadsheet;

        /**
         * Retorna um objeto Cell correspondente a uma célula específica usando notação A1.
         */
        acell(label: string, options: RenderOptions): Cell;

        /**
         * Adiciona uma única linha de dados ao final da planilha.
         */
        appendRow(values: any[], options: InsertOptions): Worksheet;

        /**
         * Adiciona múltiplas linhas de dados de uma só vez ao final da planilha.
         */
        appendRows(rows: any[][], options: InsertOptions): Worksheet;

        /**
         * Limpa várias áreas específicas da planilha de uma só vez.
         */
        batchClear(ranges: string[], keepFormat: boolean): Worksheet;

        /**
         * Retorna um objeto Cell correspondente a uma célula específica usando linha e coluna.
         */
        cell(row: number, col: number, options: RenderOptions): Cell;

        /**
         * Limpa a aba inteira.
         */
        clear(keepFormat: boolean): Worksheet;

        /**
         * Retorna todos os valores de uma coluna específica.
         */
        colValues(index: number): any[];

        /**
         * Copia todos os dados desta worksheet para uma worksheet de destino.
         * Se filtros forem fornecidos, apenas as linhas correspondentes serão copiadas.
         * Os dados atuais da aba de destino serão completamente apagados antes da colagem.
         *
         * @returns Retorna a própria instância da aba de origem (Worksheet) para permitir o encadeamento de métodos.
         */
        copyTo(destWs: Worksheet, filters: FilterRule[], head: number): Worksheet;

        /**
         * Exclui uma ou mais colunas da planilha fisicamente.
         */
        deleteColumns(index: number, numCols: number): Worksheet;

        /**
         * Exclui uma ou mais linhas da planilha fisicamente.
         */
        deleteRows(index: number, numRows: number): Worksheet;

        /**
         * Exporta exclusivamente esta aba no formato especificado.
         *
         * @returns O arquivo exportado.
         */
        export(format: string): GoogleAppsScript.Base.Blob;

        /**
         * Encontra a primeira célula que contém o valor ou que corresponda ao Regex.
         */
        find(query: any): Cell | null;

        /**
         * Encontra todas as células que contêm o valor ou que correspondam ao Regex.
         */
        findAll(query: any): Cell[];

        /**
         * Retorna todos os valores da planilha como uma lista de objetos mapeados pelo cabeçalho.
         * Se a chave de agrupamento (groupKey) for fornecida, retorna um Dicionário contendo Arrays de objetos.
         */
        getAllRecords(head: number, groupKey: string | null, empty2zero: boolean, default_blank: any): Record<string, any>[] | Record<string, Record<string, any>[]>;

        /**
         * Retorna todos os valores da planilha como uma lista de listas.
         */
        getAllValues(): any[][];

        getData(range: string | null, getDisplayValues: boolean): void;

        /**
         * Mapeia os dados da planilha de forma dinâmica e inteligente, buscando pelos cabeçalhos configurados.
         * Super resiliente a mudanças de posição das colunas e a linhas em branco no topo da planilha.
         *
         * @returns Retorna uma Lista de Objetos ou, se um groupKey for informado, um Dicionário contendo Listas de Objetos.
         *
         * @throws Lança um Erro caso os cabeçalhos obrigatórios ou a chave de agrupamento não sejam encontrados na planilha.
         */
        getMappedRecords(config: HeaderConfig[], groupKey: string | null, empty2zero: boolean, default_blank: any, searchLimit: number): Record<string, any>[] | Record<string, Record<string, any>[]>;

        /**
         * Retorna os valores de um intervalo específico da planilha.
         *
         * @returns Uma matriz bidimensional com os valores do intervalo.
         */
        getWithRange(range: string): any[][];

        /**
         * Insere uma linha de dados em uma posição específica, empurrando as demais para baixo.
         */
        insertRow(index: number, values: any[], options: UpdateOptions): Worksheet;

        /**
         * Insere múltiplas linhas de dados em uma posição específica.
         */
        insertRows(index: number, rows: any[][], options: UpdateOptions): Worksheet;

        /**
         * Redimensiona a planilha para ter exatamente o número de linhas e colunas especificado.
         */
        resize(rows: number, cols: number): Worksheet;

        /**
         * Retorna todos os valores de uma linha específica.
         */
        rowValues(index: number): any[];

        /**
         * Atualiza um intervalo de células com novos valores.
         */
        update(range: string, values: any[][], options: UpdateOptions): Worksheet;

        /**
         * Atualiza o valor de uma célula específica usando notação A1.
         */
        updateACell(range: string, value: any, options: UpdateOptions): Worksheet;

        /**
         * Atualiza o valor de uma célula específica passando linha e coluna.
         */
        updateCell(row: number, col: number, value: any, options: UpdateOptions): Worksheet;

    }

    export interface FilterRule {

        /**
         * O índice da coluna (começa em 0). Ex: Coluna A = 0
         */
        col: number;

        /**
         * Se true, o filtro será parcial (ex: procurar "João" dentro de "João Silva"). Padrão: false
         */
        partial?: boolean;

        /**
         * O valor exato para filtrar, ou um Array de valores aceitos.
         */
        val: string | number | boolean | (string | number | boolean)[];

    }

    export interface HeaderConfig {

        /**
         * O nome da chave que vai aparecer no seu Objeto final (Ex: 'idade')
         */
        id: string;

        /**
         * Se true, a função aborta caso não encontre esta coluna na planilha
         */
        required?: boolean;

        /**
         * Lista de palavras aceitas para achar essa coluna (Ex: ['Idade', 'Idades', 'Anos'])
         */
        terms: string[];

    }

    export interface InsertOptions {

        insertDataOption?: InsertDataOption;

        valueInputOption?: ValueInputOption;

    }

    export interface RenderOptions {

        valueRenderOption?: InsertDataOption;

    }

    export interface UpdateOptions {

        valueInputOption?: ValueInputOption;

    }

    /**
     * Opções de formato para exportação.
     */
    export enum ExportFormat {

        csv,

        excel,

        open_office_sheet,

        pdf,

        tsv,

        zipped_html,

    }

    /**
     * Opções para o comportamento de inserção física de linhas.
     */
    export enum InsertDataOption {

        insert_rows,

        overwrite,

    }

    /**
     * Opções para a forma como os dados inseridos devem ser interpretados.
     */
    export enum ValueInputOption {

        raw,

        user_entered,

    }

    /**
     * Opções para a forma como os dados lidos devem ser renderizados.
     */
    export enum ValueRenderOption {

        formatted,

        formula,

        unformatted,

    }

}

declare var JSspread: jsspread.JSspread;