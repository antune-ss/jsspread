/**
 * Opções de formato para exportação.
 * @public
 */
export enum ExportFormat {
  pdf = 'pdf',
  excel = 'xlsx',
  csv = 'csv',
  open_office_sheet = 'ods',
  tsv = 'tsv',
  zipped_html = 'zip'
}

/**
 * Opções para a forma como os dados inseridos devem ser interpretados.
 * @public
 */
export enum ValueInputOption {
  raw = 'RAW',
  user_entered = 'USER_ENTERED'
}

/**
 * Opções para o comportamento de inserção física de linhas.
 * @public
 */
export enum InsertDataOption {
  overwrite = 'OVERWRITE',
  insert_rows = 'INSERT_ROWS'
}

/**
 * Opções para a forma como os dados lidos devem ser renderizados.
 * @public
 */
export enum ValueRenderOption {
  formatted = 'FORMATTED_VALUE',
  unformatted = 'UNFORMATTED_VALUE',
  formula = 'FORMULA'
}
