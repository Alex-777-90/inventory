import { NextRequest, NextResponse } from 'next/server';
import JSZip from 'jszip';
import ExcelJS from 'exceljs';
import {
  loadWorkbook,
  workbookToBuffer,
  filtrosExcluirSAP,
} from '@/lib/excel';
import { carregarMapa } from '@/lib/mapDepositos';

export const runtime = 'nodejs';
export const dynamic = 'force-dynamic';

/* ======================= utils ======================= */

function norm(s: any) {
  return String(s ?? '')
    .toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/[^\w]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function findColSynAtRow(ws: ExcelJS.Worksheet, headerRow: number, names: string[]) {
  const head = ws.getRow(headerRow);
  const want = names.map(norm);
  for (let c = 1; c <= head.cellCount; c++) {
    const v = norm(head.getCell(c).value);
    if (want.includes(v)) return c;
  }
  return -1;
}

function detectHeaderRow(ws: ExcelJS.Worksheet, groups: string[][], minScore = 3): number {
  const maxScan = Math.min(30, ws.rowCount);
  let bestRow = 1;
  let bestScore = 0;
  for (let r = 1; r <= maxScan; r++) {
    const row = ws.getRow(r);
    if (!row || row.cellCount === 0) continue;
    const seen = new Set<number>();
    for (let c = 1; c <= row.cellCount; c++) {
      const v = norm(row.getCell(c).value);
      if (!v) continue;
      groups.forEach((syns, idx) => {
        if (seen.has(idx)) return;
        if (syns.some(s => s === v)) seen.add(idx);
      });
    }
    const score = seen.size;
    if (score > bestScore) { bestScore = score; bestRow = r; }
    if (score >= minScore) return r;
  }
  return bestRow;
}

function num(cellVal: any): number {
  if (cellVal && typeof cellVal === 'object' && (cellVal as any).result != null) {
    return Number((cellVal as any).result);
  }
  return Number(cellVal ?? 0);
}

/** Exporta uma cópia “plana” (sem fórmulas/estilos) da planilha para evitar avisos do Excel */
function toPlainWorkbook(ws: ExcelJS.Worksheet, headerRow: number, name = 'SAP') {
  const out = new ExcelJS.Workbook();
  const outWS = out.addWorksheet(name);
  const colCount = Math.max(ws.columnCount, ws.getRow(headerRow).cellCount);

  // cabeçalho
  for (let c = 1; c <= colCount; c++) {
    outWS.getRow(1).getCell(c).value = ws.getRow(headerRow).getCell(c).value;
  }

  const plain = (v: any) => (v && typeof v === 'object' && 'result' in v ? (v as any).result : v);

  // linhas
  let rr = 2;
  for (let r = headerRow + 1; r <= ws.rowCount; r++, rr++) {
    const src = ws.getRow(r);
    const dst = outWS.getRow(rr);
    for (let c = 1; c <= colCount; c++) {
      dst.getCell(c).value = plain(src.getCell(c).value);
    }
  }
  return out;
}

const LABEL_SAP = 'SAP';
const LABEL_FISICO = 'FÍSICO';

function normalizeSapDepositToCode(depRaw: string | undefined): string {
  const v = (depRaw || '').toUpperCase().replace(/\s+/g, '').replace(/-/g, '');
  if (!v) return '';
  if (v.includes('DESCUDL') || v === 'DESCUDL' || v === 'DESC_UDL') return 'DESC_UDL';
  if (v.includes('AVUDL')   || v === 'AVUDL'   || v === 'AV_UDL')   return 'AV_UDL';
  if (v.includes('AMUDL')   || v === 'AMUDL'   || v === 'AM_UDL')   return 'AM_UDL';
  if (v.includes('VUDL')    || v === 'VUDL'    || v === 'V_UDL')    return 'V_UDL';
  if (v.includes('UDLMP')   || v === 'UDLMP'   || v === 'UDL_MP')   return 'UDL_MP';
  if (v.includes('UDLOG') || v === 'UDL' || v === 'UD_LOG') return 'UDL';
  return 'UDL';
}
function codeToShortLabel(code: string): string {
  switch ((code || '').toUpperCase()) {
    case 'UDL':     return 'UDL';
    case 'AV_UDL':  return 'AV UDL';
    case 'DESC_UDL':return 'DESC UDL';
    case 'AM_UDL':  return 'AM UDL';
    case 'V_UDL':   return 'V UDL';
    case 'UDL_MP':  return 'UDL MP';
    default:        return (code || '').toUpperCase().replace(/_/g, ' ');
  }
}
function composeDepositoCell(currentCode: string, targetCode: string): string {
  const cur = codeToShortLabel(currentCode);
  const tar = codeToShortLabel(targetCode);
  if (!currentCode && targetCode) return tar;
  if (!targetCode) return cur || '';
  if (currentCode.toUpperCase() === targetCode.toUpperCase()) return cur;
  return `${cur} TRANSFERIR PARA ${tar}`;
}

/* ======================= handler ======================= */

export const POST = async (req: NextRequest) => {
  const form = await req.formData();
  const exportType = ((form.get('exportType') as string) || 'zip') as 'fisico'|'sap'|'zip';
  const fisicoSheet = (form.get('fisicoSheet') as string) || '';
  const sapSheet = (form.get('sapSheet') as string) || 'Planilha1';

  const fisicoFile = form.get('fisico') as File | null;
  const sapFile    = form.get('sap') as File | null;

  if (!fisicoFile || !sapFile) {
    return NextResponse.json({ error: 'Envie os dois arquivos (físico e SAP).' }, { status: 400 });
  }

  // abre workbooks
  const fisicoWB = await loadWorkbook(Buffer.from(await fisicoFile.arrayBuffer()));
  const sapWB    = await loadWorkbook(Buffer.from(await sapFile.arrayBuffer()));

  const fisicoWS = fisicoWB.getWorksheet(fisicoSheet);
  const sapWS    = sapWB.getWorksheet(sapSheet);

  if (!fisicoWS) return NextResponse.json({ error: `Aba FÍSICO '${fisicoSheet}' não encontrada` }, { status: 400 });
  if (!sapWS)    return NextResponse.json({ error: `Aba SAP '${sapSheet}' não encontrada` }, { status: 400 });

  /* ---- mapeamentos/headers ---- */

  const fisicoGroups = [
    ['codigo','código'],
    ['descricao','descrição'],
    ['qtd disponivel','qtde disponivel','qtd. disponivel','qtd disponivel','qtde disponivel','qtd. disponível'],
    ['lote','(a) lote'],
    ['ds fsa','ds. fsa','ds_fsa'],
    ['unid','unid.','unidade'],
    ['(a) fabricacao','fabricacao','fabricação'],
    ['(a) validade','validade'],
    ['observacoes','observações'],
    ['deposito sap','depósito sap']
  ];
  const sapGroups = [
    ['nº do item','n° do item','no do item','nº do item','n do item'],
    ['item'],
    ['deposito','depósito','codigo de deposito','codigo de depósito','código de deposito','código de depósito'],
    ['lote'],
    ['qtde por lote','qtd por lote','quantidade por lote'],
    ['detalhes analise','detalhes análise','detalhes'],
  ];

  const fisHeader = detectHeaderRow(fisicoWS, fisicoGroups, 3);
  const sapHeader = detectHeaderRow(sapWS,    sapGroups,    2);

  const c = {
    fisico: {
      codigo:      findColSynAtRow(fisicoWS, fisHeader, ['CÓDIGO','CODIGO','Código','Codigo']),
      descricao:   findColSynAtRow(fisicoWS, fisHeader, ['DESCRIÇÃO','DESCRICAO','Descrição','Descricao']),
      qtd:         findColSynAtRow(fisicoWS, fisHeader, ['QTD. DISPONÍVEL','QTD DISPONIVEL','QTDE DISPONIVEL','QTDE. DISPONIVEL','Qtde disponível','QTD DISPONIVEL']),
      unid:        findColSynAtRow(fisicoWS, fisHeader, ['UNID.','UNID','UNIDADE']),
      dsfsa:       findColSynAtRow(fisicoWS, fisHeader, ['DS. FSA','DS FSA','DS_FSA']),
      fabricacao:  findColSynAtRow(fisicoWS, fisHeader, ['(A) FABRICACAO','(A) FABRICAÇÃO','FABRICAÇÃO','FABRICACAO','Fabricação']),
      validade:    findColSynAtRow(fisicoWS, fisHeader, ['(A) VALIDADE','VALIDADE']),
      lote:        findColSynAtRow(fisicoWS, fisHeader, ['(A) LOTE','LOTE']),
      observacoes: findColSynAtRow(fisicoWS, fisHeader, ['OBSERVAÇÕES','OBSERVACOES','Observações','Observacoes']),
      depositoSap: findColSynAtRow(fisicoWS, fisHeader, ['depósito SAP','DEPOSITO SAP','DEPÓSITO SAP','deposito sap']),
    },
    sap: {
      numItem:   findColSynAtRow(sapWS, sapHeader, ['Nº do item','N° do item','No do item','Nº do Item','n do item']),
      item:      findColSynAtRow(sapWS, sapHeader, ['Item']),
      deposito:  findColSynAtRow(sapWS, sapHeader, ['Depósito','Codigo de depósito','Código de depósito','Codigo de deposito','Código de deposito','Deposito','deposito']),
      lote:      findColSynAtRow(sapWS, sapHeader, ['Lote']),
      qtde:      findColSynAtRow(sapWS, sapHeader, ['Qtde por lote','Qtd por lote','Quantidade por lote']),
      detalhes:  findColSynAtRow(sapWS, sapHeader, ['detalhes analise','detalhes análise','detalhes']),
    }
  };

  // garante colunas auxiliares
  if (c.fisico.depositoSap < 0) {
    const col = fisicoWS.getRow(fisHeader).cellCount + 1;
    fisicoWS.getRow(fisHeader).getCell(col).value = 'depósito SAP';
    c.fisico.depositoSap = col;
  }
  if (c.fisico.observacoes < 0) {
    const col = fisicoWS.getRow(fisHeader).cellCount + 1;
    fisicoWS.getRow(fisHeader).getCell(col).value = 'OBSERVAÇÕES';
    c.fisico.observacoes = col;
  }
  if (c.sap.detalhes < 0) {
    const col = sapWS.getRow(sapHeader).cellCount + 1;
    sapWS.getRow(sapHeader).getCell(col).value = 'detalhes analise';
    c.sap.detalhes = col;
  } else {
    const v = String(sapWS.getRow(sapHeader).getCell(c.sap.detalhes).value ?? '');
    if (norm(v) === 'detalhes') {
      sapWS.getRow(sapHeader).getCell(c.sap.detalhes).value = 'detalhes analise';
    }
  }

  // Excluir depósitos indesejados (4.1 do fluxo)
  if (c.sap.deposito >= 0) {
    for (let r = sapWS.rowCount; r >= sapHeader + 1; r--) {
      const dep = String(sapWS.getRow(r).getCell(c.sap.deposito).value ?? '').trim();
      if (filtrosExcluirSAP.has(dep)) sapWS.spliceRows(r, 1);
    }
  }

  /* ---- índices/agrupamentos ---- */

  // SAP — agrega por (código,lote) e identifica depósito “dominante”
  const sapInfo = new Map<string, { total:number; byDep:Record<string, number>; topDep:string }>();
  for (let r = sapHeader + 1; r <= sapWS.rowCount; r++) {
    const row = sapWS.getRow(r);
    const codigo = String(row.getCell(c.sap.numItem).value ?? '').trim();
    const lote   = String(row.getCell(c.sap.lote).value ?? '').trim();
    if (!codigo || !lote) continue;

    const qtde = num(row.getCell(c.sap.qtde).value);
    const depRaw = c.sap.deposito >= 0 ? String(row.getCell(c.sap.deposito).value ?? '').trim() : '';
    const depKey = (depRaw || 'SAP').replace(/\s+/g, '_').toUpperCase();

    const key = `${codigo}||${lote}`;
    const curr = sapInfo.get(key) ?? { total: 0, byDep: {}, topDep: '' };
    curr.total += qtde;
    curr.byDep[depKey] = (curr.byDep[depKey] ?? 0) + qtde;
    curr.topDep = Object.entries(curr.byDep).sort((a,b)=>b[1]-a[1])[0]?.[0] ?? depKey;
    sapInfo.set(key, curr);
  }

  // Físico — totais por (código,lote) e conjunto de LOTES (para a regra “OK” no SAP)
  const fisicoTotals = new Map<string, number>();
  const fisicoLotes = new Set<string>();

  for (let r = fisHeader + 1; r <= fisicoWS.rowCount; r++) {
    const row = fisicoWS.getRow(r);
    const codigo = String(row.getCell(c.fisico.codigo).value ?? '').trim();
    const lote   = String(row.getCell(c.fisico.lote).value ?? '').trim();
    if (!codigo || !lote) continue;

    const q = num(row.getCell(c.fisico.qtd).value);
    const key = `${codigo}||${lote}`;

    fisicoTotals.set(key, (fisicoTotals.get(key) ?? 0) + q);
    fisicoLotes.add(lote.toUpperCase());
  }

  /* ---- aplicar análise no FÍSICO ---- */

  const mapa = carregarMapa();

  for (let r = fisHeader + 1; r <= fisicoWS.rowCount; r++) {
    const row = fisicoWS.getRow(r);
    const codigo = String(row.getCell(c.fisico.codigo).value ?? '').trim();
    const lote   = String(row.getCell(c.fisico.lote).value ?? '').trim();
    if (!codigo || !lote) continue;

    const dsfsa = String(row.getCell(c.fisico.dsfsa).value ?? '');
    const key = `${codigo}||${lote}`;

    const targetCode = (mapa.decidirDepositoPorDSFSA(dsfsa) || '').toUpperCase();
    const info = sapInfo.get(key);

    // coluna “depósito SAP” — onde está + se transfere
    if (!info) {
      row.getCell(c.fisico.depositoSap).value = codeToShortLabel(targetCode);
    } else {
      const currentCode = normalizeSapDepositToCode(info.topDep);
      row.getCell(c.fisico.depositoSap).value = composeDepositoCell(currentCode, targetCode);
    }

    // OBSERVAÇÕES (comparação por TOTAIS)
    if (!info) {
      row.getCell(c.fisico.observacoes).value = 'Lote não Localizado no SAP';
    } else {
      const totalSAP = Math.round(info.total);
      const totalFIS = Math.round(fisicoTotals.get(key) ?? 0);
      const diffAbs  = Math.abs(totalFIS - totalSAP);

      if (diffAbs === 0) {
        row.getCell(c.fisico.observacoes).value = 'OK';
      } else {
        row.getCell(c.fisico.observacoes).value =
          `DIFERENÇA DE ${diffAbs} KG - ESTOQUE ${LABEL_SAP} ${totalSAP} E ESTOQUE ${LABEL_FISICO} ${totalFIS}`;
      }
    }
  }

  /* ---- preencher “detalhes analise” no SAP (por LOTE) ---- */

  for (let r = sapHeader + 1; r <= sapWS.rowCount; r++) {
    const row = sapWS.getRow(r);
    const loteSAP = String(row.getCell(c.sap.lote).value ?? '').trim().toUpperCase();
    row.getCell(c.sap.detalhes).value = loteSAP && fisicoLotes.has(loteSAP) ? 'OK' : '';
  }

  /* ---- exportação ---- */

  if (exportType === 'fisico') {
    const buf = await workbookToBuffer(fisicoWB);
    return new Response(buf, {
      headers: {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Disposition': 'attachment; filename="ESTOQUE FISICO (analisado).xlsx"'
      }
    });
  }

  if (exportType === 'sap') {
    const sapOutWB = toPlainWorkbook(sapWS, sapHeader, 'SAP');
    const buf = await workbookToBuffer(sapOutWB);
    return new Response(buf, {
      headers: {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Disposition': 'attachment; filename="ESTOQUE SAP (analisado).xlsx"'
      }
    });
  }

  // ZIP com os dois
  const fisBuf = await workbookToBuffer(fisicoWB);
  const sapOutWB = toPlainWorkbook(sapWS, sapHeader, 'SAP');
  const sapBuf = await workbookToBuffer(sapOutWB);

  const zip = new JSZip();
  zip.file('ESTOQUE FISICO (analisado).xlsx', fisBuf, { binary: true });
  zip.file('ESTOQUE SAP (analisado).xlsx',    sapBuf, { binary: true });
  const zipBuf = await zip.generateAsync({ type: 'nodebuffer' });

  return new Response(zipBuf, {
    headers: {
      'Content-Type': 'application/zip',
      'Content-Disposition': 'attachment; filename="resultado-analise.zip"'
    }
  });
};
