/**
 * Normaliza textos para comparação e busca.
 * Remove acentos, capitalização e reduz múltiplos espaços.
 * 
 * @param text - Qualquer texto ou valor de célula.
 * @param cleanSeparators - Se true, substitui hífens, barras e underscores por espaços. Padrão: true.
 * @returns O texto normalizado.
 * @public
 */
export function normalizeTxt(text: any, cleanSeparators: boolean = false): string {
  if (text === null || text === undefined) return "";

  const shouldClean = typeof cleanSeparators === 'boolean' ? cleanSeparators : false;

  let str = String(text)
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, ""); // Remove acentos

  if (shouldClean) {
    str = str.replace(/[-_\/\s]+/g, " "); // Troca -, _, / e espaços repetidos por 1 espaço
  } else {
    str = str.replace(/\s+/g, " ");      // Reduz apenas espaços repetidos, preserva símbolos
  }

  return str.trim();
}