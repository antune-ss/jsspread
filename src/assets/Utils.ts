import { normalizeTxt } from "./Text.js";

/**
 * Sobrecarga 1: Extrai propriedades específicas de uma Lista de Objetos.
 * @public
 */
export function extract<T extends Record<string, any>>(
  data: T[],
  columns: (keyof T | string)[]
): Partial<T>[];

/**
 * Sobrecarga 2: Extrai colunas específicas de uma Matriz 2D (por índices ou nomes de cabeçalho).
 * @public
 */
export function extract(
  data: any[][],
  columns: (string | number)[],
  head?: number
): any[][];

/**
 * Extrai colunas selecionadas de uma matriz de dados ou de uma lista de objetos.
 * 
 * @param data - Matriz 2D bruta ou Array de Objetos.
 * @param columns - Array com os índices das colunas [0, 2] ou nomes ['Data', 'Turno'].
 * @param head - (Opcional) Índice da linha de cabeçalho se usar nomes na matriz 2D (começa em 1). Padrão: 1.
 * @returns Nova matriz ou lista contendo apenas as colunas extraídas.
 * @public
 */
export function extract(data: any[], columns: any[], head: number = 1): any[] {
  if (!data || data.length === 0 || !columns || columns.length === 0) return [];

  // CENÁRIO A: Array de Objetos [{ col1: 'a', col2: 'b' }]
  if (!Array.isArray(data[0]) && typeof data[0] === "object") {
    return data.map(item => {
      const row: Record<string, any> = {};
      for (const col of columns) {
        const key = String(col);
        row[key] = item[key];
      }
      return row;
    });
  }

  // CENÁRIO B: Matriz 2D [ [ 'A', 'B' ], [ 1, 2 ] ]
  const matrix = data as any[][];
  let targetIdx: number[] = [];

  const hasStringColumns = columns.some(c => typeof c === "string");

  if (hasStringColumns) {
    const headerRow = matrix[head - 1] || [];
    const headerNorm = headerRow.map(cell => normalizeTxt(cell));

    targetIdx = columns.map(col => {
      if (typeof col === "number") return col;
      return headerNorm.indexOf(normalizeTxt(col));
    });
  } else {
    targetIdx = columns as number[];
  }

  return matrix.map(row => {
    return targetIdx.map(idx => {
      if (idx === -1 || idx === undefined || idx >= row.length) {
        return ""; // Coluna não encontrada ou fora do intervalo
      }
      return row[idx];
    });
  });
}

/**
 * Sobrecarga 1: Mescla duas Listas de Objetos usando uma propriedade chave (LEFT JOIN).
 * @public
 */
export function merge<T extends Record<string, any>, U extends Record<string, any>>(
  target: T[],
  source: U[],
  keyTarget: keyof T | string,
  keySource?: keyof U | string
): (T & Partial<U>)[];

/**
 * Sobrecarga 2: Mescla duas Matrizes 2D anexando as colunas da matriz secundária (PROCV em memória).
 * @public
 */
export function merge(
  target: any[][],
  source: any[][],
  keyTarget: string | number,
  keySource?: string | number,
  headTarget?: number,
  headSource?: number
): any[][];

/**
 * Mescla duas fontes de dados (Matrizes 2D ou Listas de Objetos) com base em uma coluna chave em comum.
 * 
 * @param target - A tabela principal (base).
 * @param source - A tabela secundária de onde virão as novas informações.
 * @param keyTarget - Nome ou índice da coluna chave na tabela principal.
 * @param keySource - (Opcional) Nome ou índice da coluna chave na secundária (padrão: igual a keyTarget).
 * @param headTarget - (Opcional) Linha do cabeçalho da tabela principal (começa em 1). Padrão: 1.
 * @param headSource - (Opcional) Linha do cabeçalho da tabela secundária. Padrão: igual ao headTarget.
 * @returns Nova matriz ou lista com os dados consolidados.
 * @public
 */
export function merge(
  target: any[], 
  source: any[], 
  keyTarget: any, 
  keySource: any = keyTarget,
  headTarget: number = 1,  
  headSource: number = headTarget
): any[] {
  if (!target || target.length === 0) return [];
  if (!source || source.length === 0) return target;

  // CENÁRIO A: Array de Objetos [{ col1: 'a', col2: 'b' }]
  if (!Array.isArray(target[0]) && typeof target[0] === "object") {
    const lookup = new Map<string, any>();
    for (const item of source) {
      const val = item[keySource];
      if (val !== undefined && val !== null && val !== "") {
        lookup.set(normalizeTxt(val), item);
      }
    }

    return target.map(item => {
      const val = item[keyTarget];
      const k = val !== undefined && val !== null && val !== "" ? normalizeTxt(val) : "";
      const match = lookup.get(k);
      return match ? Object.assign({}, item, match) : Object.assign({}, item);
    })
  }

  // CENÁRIO B: Matriz 2D [ [ 'A', 'B' ], [ 1, 2 ] ]
  const targetMatrix = target as any[][];
  const sourceMatrix = source as any[][];

  const headerTarget = targetMatrix[headTarget - 1] || [];
  const headerSource = sourceMatrix[headSource - 1] || [];

  let targetKeyIdx: number;
  if (typeof keyTarget === "number") {
    targetKeyIdx = keyTarget;
  } else {
    const norm = headerTarget.map(c => normalizeTxt(c));
    targetKeyIdx = norm.indexOf(normalizeTxt(keyTarget));
  }

  let sourceKeyIdx: number;
  if (typeof keySource === "number") {
    sourceKeyIdx = keySource;
  } else {
    const norm = headerSource.map(c => normalizeTxt(c));
    sourceKeyIdx = norm.indexOf(normalizeTxt(keySource));
  }

  if (targetKeyIdx === -1 || sourceKeyIdx === -1) {
    throw new Error(`Utils Error (merge): A coluna chave não foi localizada em uma das matrizes.`);
  }

  const sourceColsToAppend: number[] = [];
  for (let c = 0; c < headerSource.length; c++) {
    if (c !== sourceKeyIdx) {
      sourceColsToAppend.push(c);
    }
  }

  const lookup = new Map<string, any[]>();
  for (let i = headSource; i < sourceMatrix.length; i++) {
    const row = sourceMatrix[i];
    const val = row[sourceKeyIdx]
    if (val !== undefined && val !== null && val !== "") {
      const k = normalizeTxt(val);
      lookup.set(k, sourceColsToAppend.map(colIdx => row[colIdx]));
    }
  }

  const emptyFiller = new Array(sourceColsToAppend.length).fill("");
  const sourceHeaderToAppend = sourceColsToAppend.map(colIdx => headerSource[colIdx]);

  const result: any[][] = [];

  for (let i = 0; i < targetMatrix.length; i++) {
    const row = [...targetMatrix[i]];

    if (i < headTarget - 1) {
      result.push(row.concat(emptyFiller));
    } else if (i === headTarget -1) {
      result.push(row.concat(sourceHeaderToAppend));
    } else {
      const keyVal = row[targetKeyIdx];
      const k = keyVal !== undefined && keyVal !== null && keyVal !== "" ? normalizeTxt(keyVal) : "";
      const match = lookup.get(k);

      if (match) {
        result.push(row.concat(match));
      } else {
        result.push(row.concat(emptyFiller));
      }
    }
  }

  return result;
}