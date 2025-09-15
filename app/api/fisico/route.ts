import { NextRequest, NextResponse } from 'next/server';
import { loadWorkbook, getSheetNames } from '@/lib/excel';

export const runtime = 'nodejs';
export const dynamic = 'force-dynamic';

export async function POST(req: NextRequest) {
  try {
    const sp = req.nextUrl.searchParams;
    const isPreview = sp.get('preview') === '1';

    const form = await req.formData();

    // Aceita tanto 'fisico' quanto 'file' (seu front usa 'file')
    const fisico = (form.get('fisico') || form.get('file')) as File | null;

    if (!(fisico instanceof File)) {
      return NextResponse.json({ error: 'Envie a planilha do Físico.' }, { status: 400 });
    }

    const sheet = (form.get('sheet') || '').toString().trim();
    const ab = await fisico.arrayBuffer();
    const wb = await loadWorkbook(ab);

    // Se 'preview=1' -> apenas valida (arquivo + aba)
    if (isPreview) {
      if (!sheet) {
        return NextResponse.json({ error: 'Selecione a aba do Físico.' }, { status: 400 });
      }
      if (!wb.getWorksheet(sheet)) {
        return NextResponse.json({ error: `Aba "${sheet}" não encontrada no Físico.` }, { status: 400 });
      }
      return NextResponse.json({ ok: true });
    }

    // Sem preview:
    // - Se NÃO veio 'sheet' -> listar abas
    if (!sheet) {
      const sheets = getSheetNames(wb);
      if (!sheets.length) {
        return NextResponse.json({ error: 'Nenhuma aba encontrada no arquivo do Físico.' }, { status: 400 });
      }
      return NextResponse.json({ sheets });
    }

    // - Se veio 'sheet' -> validar e confirmar
    if (!wb.getWorksheet(sheet)) {
      return NextResponse.json({ error: `Aba "${sheet}" não encontrada no Físico.` }, { status: 400 });
    }
    return NextResponse.json({ ok: true });
  } catch (err: any) {
    console.error('[api/fisico] erro:', err);
    return NextResponse.json(
      { error: err?.message || 'Falha ao processar a planilha do Físico.' },
      { status: 500 },
    );
  }
}
