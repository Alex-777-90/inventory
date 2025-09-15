// lib/excel.ts
import ExcelJS from 'exceljs';
import { Buffer } from 'node:buffer';

/** Tipos binários aceitos */
type Bin = ArrayBuffer | Uint8Array | Buffer;

/** Type guard p/ Buffer (evita “never” no TS) */
function isNodeBuffer(x: any): x is Buffer {
  return (
    x != null &&
    typeof x === 'object' &&
    typeof x.byteLength === 'number' &&
    typeof x.byteOffset === 'number' &&
    x.constructor?.name === 'Buffer'
  );
}

/** Converte entrada (Buffer/Uint8Array/ArrayBuffer) para Uint8Array */
function toUint8(input: Bin): Uint8Array {
  if (input instanceof Uint8Array) return input;            // Buffer também herda de Uint8Array
  if (input instanceof ArrayBuffer) return new Uint8Array(input);
  if (isNodeBuffer(input)) {
    return new Uint8Array(input.buffer, input.byteOffset, input.byteLength);
  }
  // fallback seguro (caso exótico)
  return new Uint8Array(Buffer.from(input as any));
}

/** Converte saída (Buffer/Uint8Array/ArrayBuffer) para Buffer (Node) */
function toNodeBuffer(x: Bin): Buffer {
  if (isNodeBuffer(x)) return x;
  if (x instanceof ArrayBuffer) return Buffer.from(x);
  // Uint8Array (ou Buffer-like)
  return Buffer.from(x.buffer, x.byteOffset, x.byteLength);
}

/** Carrega workbook a partir de ArrayBuffer/Uint8Array/Buffer */
export async function loadWorkbook(input: Bin): Promise<ExcelJS.Workbook> {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(toUint8(input)); // exceljs espera Uint8Array/ArrayBuffer
  return wb;
}

/** Serializa workbook para Buffer (usado p/ enviar download) */
export async function workbookToBuffer(wb: ExcelJS.Workbook): Promise<Buffer> {
  const out = (await wb.xlsx.writeBuffer()) as ArrayBuffer | Uint8Array | Buffer;
  return toNodeBuffer(out);
}

/* ========= utilitários que você já usa ========= */

/** Depósitos do SAP que devem ser excluídos (filtro) */
export const filtrosExcluirSAP = new Set<string>([
  'CHEM WIP', 'V_CHEWP', 'SC_Nest', 'SC_DSB', 'SC_DSB_2', 'TST', 'MS WIP', 'EM DSB',
]);

/** Carimbo simples data/hora (YYYY-MM-DD HH-mm) */
export function timestamp(): string {
  const d = new Date();
  const pad = (n: number) => String(n).padStart(2, '0');
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())} ${pad(d.getHours())}-${pad(d.getMinutes())}`;
}

/** Lista nomes de abas do workbook */
export function getSheetNames(wb: ExcelJS.Workbook): string[] {
  return wb.worksheets.map(w => w.name);
}
