/**
 * Converte de forma segura qualquer formato de data e hora para um objeto Date válido.
 * Utiliza tolerância para números seriais do Sheets e strings BR/ISO.
 * 
 * @param value - O valor bruto recebido da planilha.
 * @returns O objeto Date validado ou null se for impossível converter.
 * @public
 */
export function ensureDate(value: any): Date | null {
  if (value === null || value === undefined || value === "") return null;

  if (value instanceof Date) {
    if (!isNaN(value.getTime())) return value;
  }

  if (typeof value === 'number') {
    // 20.000 dias dá o ano de 1954, e 100.000 dias dá o ano de 2173
    if (value > 20000 && value < 100000) {
      // Sheets começou a contar em 1899 e o JS em 1970. A diferença exata entre essas duas datas de início são 25.569 dias. Ao subtrair isso do valor, "sincronizamos" o calendário da planilha com o do JavaScript.
      // Como o valor da planilha está em Dias, multiplicamos por 86.400 (segundos em um dia) e depois por 1.000 (para transformar em milissegundos, que é a língua nativa do JavaScript).
      const d = new Date(Math.round((value - 25569) * 86400 * 1000));

      // A data cai no fuso horário universal (UTC). Dependendo de onde o código rodar, o relógio poderia atrasar 3 horas (fuso do Brasil) e a data 15/04 viraria 14/04 às 21h00. 
      // Somar o TimezoneOffset força a data a ficar exatamente no dia em que deveria estar, não importa de qual país esteja rodando.
      d.setMinutes(d.getMinutes() + d.getTimezoneOffset());
      return d;
    }

    const d = new Date(value);
    if (!isNaN(d.getTime())) return d;
  }

  if (typeof value === 'string') {
    const clean = value.trim();

    const br = clean.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})(?:\s+(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?)?/);
    if (br) {
      const d = new Date(
        parseInt(br[3], 10),                   // Ano
        parseInt(br[2], 10) - 1,               // Mês (0-11)
        parseInt(br[1], 10),                   // Dia
        br[4] ? parseInt(br[4], 10) : 0,       // Horas (se não tiver, é 0)
        br[5] ? parseInt(br[5], 10) : 0,       // Minutos
        br[6] ? parseInt(br[6], 10) : 0        // Segundos
      );
      if (!isNaN(d.getTime())) return d;
    }

    const us = clean.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})(?:\s+(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?)?/);
    if (us) {
      const d = new Date(
        parseInt(us[1], 10),                   // Ano
        parseInt(us[2], 10) - 1,               // Mês (0-11)
        parseInt(us[3], 10),                   // Dia
        us[4] ? parseInt(us[4], 10) : 0,       // Horas (se não tiver, é 0)
        us[5] ? parseInt(us[5], 10) : 0,       // Minutos
        us[6] ? parseInt(us[6], 10) : 0        // Segundos
      );
      if (!isNaN(d.getTime())) return d;
    }

    const d = new Date(value);
    if (!isNaN(d.getTime())) return d;
  }

  return null;
}

/**
 * Formata datas de para strings de forma rápida.
 * Não utiliza a API do Google e nem o Intl do JS.
 * 
 * @param value - Qualquer valor de data (String, Serial Numérico ou Date).
 * @param includeTime - Se true, inclui a hora (DD/MM/YYYY HH:mm). Padrão: false.
 * @returns A data formatada em String no formato br, ou null em caso de entrada inválida.
 * @public
 */
export function fmtData(value: any, includeTime: boolean = false): string | null {
  const d = ensureDate(value);
  if (!d) return null;

  const day = String(d.getDate()).padStart(2, '0');
  const month = String(d.getMonth() + 1).padStart(2, '0');
  const year = d.getFullYear();

  if (!includeTime) return `${day}/${month}/${year}`;

  const hours = String(d.getHours()).padStart(2, '0');
  const minutes = String(d.getMinutes()).padStart(2, '0');

  return `${day}/${month}/${year} ${hours}:${minutes}`;
} 