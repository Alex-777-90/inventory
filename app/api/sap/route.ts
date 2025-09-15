// app/api/sap/route.ts
import { NextRequest, NextResponse } from 'next/server';
import ExcelJS from 'exceljs';
import {
  loadWorkbook,
  workbookToBuffer,
  filtrosExcluirSAP,
} from '@/lib/excel';


export const runtime = 'nodejs';
export const dynamic = 'force-dynamic';

/* ======================= helpers ======================= */

function norm(s: any) {
  return String(s ?? '')
    .toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/[^\w]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function findColSynAtRow(
  ws: ExcelJS.Worksheet,
  headerRow: number,
  names: string[],
): number {
  const head = ws.getRow(headerRow);
  const want = names.map(norm);
  for (let c = 1; c <= head.cellCount; c++) {
    const v = norm(head.getCell(c).value);
    if (want.includes(v)) return c;
  }
  return -1;
}

/** tenta “adivinhar” a linha de cabeçalho olhando sinônimos comuns */
function detectHeaderRow(
  ws: ExcelJS.Worksheet,
  groups: string[][],
  minScore = 2,
): number {
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
        if (syns.some((s) => s === v)) seen.add(idx);
      });
    }
    const score = seen.size;
    if (score > bestScore) {
      bestScore = score;
      bestRow = r;
    }
    if (score >= minScore) return r;
  }
  return bestRow;
}

/* ======================= handler ======================= */

export const POST = async (req: NextRequest) => {
  try {
    const form = await req.formData();

    // arquivo e aba
    const sapFile = form.get('sap') as File | null;
    const sapSheet =
      (form.get('sapSheet') as string) ||
      (form.get('sheet') as string) ||
      'Planilha1';

    if (!sapFile) {
      return NextResponse.json(
        { error: 'Envie a planilha do SAP.' },
        { status: 400 },
      );
    }

    // abre workbook
    const sapWB = await loadWorkbook(Buffer.from(await sapFile.arrayBuffer()));
    const sapWS =
      sapWB.getWorksheet(sapSheet) || sapWB.worksheets[0] || null;

    if (!sapWS) {
      return NextResponse.json(
        { error: `Aba SAP '${sapSheet}' não encontrada` },
        { status: 400 },
      );
    }

    /* ---------- localizar cabeçalho e colunas ---------- */

    // grupos de sinônimos típicos do SAP
    const sapGroups = [
      ['nº do item', 'n° do item', 'no do item', 'n do item'],
      ['item'],
      ['deposito', 'depósito', 'codigo de deposito', 'código de depósito'],
      ['lote'],
      ['qtde por lote', 'qtd por lote', 'quantidade por lote'],
      ['detalhes analise', 'detalhes análise', 'detalhes'],
    ];

    const headerRow = detectHeaderRow(sapWS, sapGroups, 2);

    // colunas
    const colDeposito = findColSynAtRow(sapWS, headerRow, [
      'depósito',
      'deposito',
      'codigo de deposito',
      'código de depósito',
    ]);
    let colDetalhes = findColSynAtRow(sapWS, headerRow, [
      'detalhes analise',
      'detalhes análise',
      'detalhes',
    ]);

    // garante coluna "detalhes analise"
    if (colDetalhes < 0) {
      const newCol = sapWS.getRow(headerRow).cellCount + 1;
      sapWS.getRow(headerRow).getCell(newCol).value = 'detalhes analise';
      colDetalhes = newCol;
    } else {
      // se o cabeçalho for "detalhes", renomeia para "detalhes analise"
      const v = String(
        sapWS.getRow(headerRow).getCell(colDetalhes).value ?? '',
      );
      if (norm(v) === 'detalhes') {
        sapWS.getRow(headerRow).getCell(colDetalhes).value =
          'detalhes analise';
      }
    }

    /* ---------- filtro de depósitos indesejados ---------- */
    if (colDeposito >= 0) {
      for (let r = sapWS.rowCount; r >= headerRow + 1; r--) {
        const dep = String(
          sapWS.getRow(r).getCell(colDeposito).value ?? '',
        ).trim();
        if (filtrosExcluirSAP.has(dep)) {
          sapWS.spliceRows(r, 1);
        }
      }
    }

    // (opcional) limpamos qualquer conteúdo em "detalhes analise" — aqui só consolidamos;
    // a marcação "OK" por LOTE será feita no /api/analisar
    for (let r = headerRow + 1; r <= sapWS.rowCount; r++) {
      sapWS.getRow(r).getCell(colDetalhes).value = '';
    }

    /* ---------- resposta (xlsx processado) ---------- */

    const buf = await workbookToBuffer(sapWB);
    return new Response(buf, {
      headers: {
        'Content-Type':
          'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        // fetch não baixa automaticamente; o front decide o que fazer com o blob
        'Content-Disposition':
          'inline; filename="ESTOQUE SAP (consolidado).xlsx"',
      },
    });
  } catch (err: any) {
    console.error('api/sap error:', err);
    return NextResponse.json(
      { error: 'Falha ao consolidar a planilha do SAP.' },
      { status: 500 },
    );
  }
};
