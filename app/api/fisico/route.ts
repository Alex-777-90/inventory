import { NextRequest, NextResponse } from 'next/server';
import {
  loadWorkbook,
  workbookToBuffer,
  getSheetNames,
  timestamp,
} from '@/lib/excel';

export const runtime = 'nodejs';

const isPreview = (req: NextRequest) =>
  new URL(req.url).searchParams.get('preview') === '1';

export const POST = async (req: NextRequest) => {
  const form = await req.formData();
  const file = form.get('file') as File | null;
  const sheetName = (form.get('sheet') as string) || '';

  if (!file) return NextResponse.json({ error: 'Arquivo não enviado' }, { status: 400 });
  const buffer = Buffer.from(await file.arrayBuffer());

  // sem sheet → devolve lista de abas
  if (!sheetName) {
    const sheets = getSheetNames(buffer);
    return NextResponse.json({ sheets });
  }

  const wb = await loadWorkbook(buffer);
  const ws = wb.getWorksheet(sheetName);
  if (!ws) return NextResponse.json({ error: `Aba '${sheetName}' não encontrada` }, { status: 400 });

  // adiciona "depósito SAP" se não existir
  const header = ws.getRow(1);
  let has = false;
  for (let c = 1; c <= header.cellCount; c++) {
    const v = String(header.getCell(c).value ?? '').trim().toLowerCase();
    if (v === 'depósito sap' || v === 'deposito sap') { has = true; break; }
  }
  if (!has) header.getCell(header.cellCount + 1).value = 'depósito SAP';

  // preview: só confirma
  if (isPreview(req)) return NextResponse.json({ ok: true });

  // download (opcional)
  const out = await workbookToBuffer(wb);
  const name = `ESTOQUE FISICO ${timestamp()}.xlsx`;
  return new Response(out, {
    headers: {
      'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'Content-Disposition': `attachment; filename="${name}"`
    }
  });
};
