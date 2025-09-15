// lib/excel.ts
import ExcelJS from 'exceljs';
import * as XLSX from 'xlsx';

export async function loadWorkbook(buffer: Buffer) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer);
  return wb;
}

export function getSheetNames(buffer: Buffer) {
  const wb = XLSX.read(buffer, { type: 'buffer' });
  return wb.SheetNames;
}

// 🔧 importante: retornar Buffer Node (não ArrayBuffer)
export async function workbookToBuffer(wb: ExcelJS.Workbook): Promise<Buffer> {
  const ab = await wb.xlsx.writeBuffer();              // ArrayBuffer
  return Buffer.from(ab as ArrayBuffer);               // -> Buffer Node
}

export function timestamp() {
  const d = new Date();
  const pad = (n:number)=>`${n}`.padStart(2,'0');
  return `${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())} ${pad(d.getHours())}-${pad(d.getMinutes())}-${pad(d.getSeconds())}`;
}

export const filtrosExcluirSAP = new Set([
  'CHEM WIP','V_CHEWP','SC_Nest','SC_DSB','SC_DSB_2','TST','MS WIP','EM DSB'
]);
