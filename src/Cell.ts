/**
 * Inicializa o objeto da célula.
 * @class
 * @public
 */
export class Cell {
  public row: number;
  /** @public */
  col: number;
  /** @public */
  value: any;

  /**
   * Representa uma célula individual na planilha.
   * 
   * @param row - O número da linha (começa em 1).
   * @param col - O número da coluna (começa em 1).
   * @param value - O valor contido na célula.
   * @public
   */
  constructor(row: number, col: number, value: any) {
    this.row = row;
    this.col = col;
    this.value = value;
  }

  /**
   * Retorna a notação A1 da célula (ex: linha 1, coluna 1 vira "A1").
   * @public
   */
  get address(): string {
    let tempCol = this.col;
    let letter = '';
    
    while (tempCol > 0) {
      let remainder = (tempCol - 1) % 26;
      letter = String.fromCharCode(65 + remainder) + letter;
      tempCol = Math.floor((tempCol - 1) / 26);
    }
    
    return `${letter}${this.row}`;
  }

  /**
   * Retorna o valor da célula convertido para número, ou null se não for possível.
   * @public
   */
  get numeric_value(): number | null {
    if (this.value === null || this.value === undefined || this.value === '') {
      return null;
    }
    const parsed = Number(this.value);
    return isNaN(parsed) ? null : parsed;
  }

  /**
   * Cria uma nova instância de Cell a partir de uma notação A1 e um valor.
   * Método estático (chamado direto na classe: Cell.fromAddress('B5', 10)).
   * 
   * @param label - A notação A1 (ex: "B5").
   * @param value - O valor da célula.
   */
  static fromAddress(label: string, value: any): Cell {
    const match = label.toUpperCase().match(/^([A-Z]+)([0-9]+)$/);

    if (!match) {
      throw new Error(`ValueError: Notação A1 inválida '${label}'.`);
    }

    const colStr = match[1];
    const row = parseInt(match[2], 10);

    let col = 0;
    for (let i = 0; i < colStr.length; i++) {
      col = col * 26 + (colStr.charCodeAt(i) - 64);
    }

    return new Cell(row, col, value);
  }

  /**
   * Retorna uma representação em string da célula para debug.
   * @public
   */
  toString(): string {
    return `Cell(${this.address}, '${this.value}')`;
  }

  /**
   * Converte a célula em um objeto JSON puro.
   * @public
   */
  toJSON(): object {
    return {
      row: this.row,
      col: this.col,
      address: this.address,
      value: this.value
    };
  }
}
