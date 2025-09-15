'use client';

import { useEffect, useState } from 'react';

type ExportType = 'fisico' | 'sap' | 'zip';

export default function Page() {
  // ===== arquivos em memória =====
  const [fisicoFile, setFisicoFile] = useState<File | null>(null);
  const [sapFile, setSapFile] = useState<File | null>(null);

  // ===== abas / seleção =====
  const [fisicoSheets, setFisicoSheets] = useState<string[]>([]);
  const [fisicoSheet, setFisicoSheet] = useState<string>('');
  const [sapSheet, setSapSheet] = useState<string>('Planilha1');

  // ===== status =====
  const [fisicoReady, setFisicoReady] = useState(false);
  const [sapReady, setSapReady] = useState(false);

  // ===== exportação =====
  const [exportType, setExportType] = useState<ExportType>('zip');
  const [loading, setLoading] = useState(false);

  // ===== tema (cards/dark) =====
  const [theme, setTheme] = useState<'cards' | 'dark'>('cards');
  useEffect(() => {
    const saved = typeof window !== 'undefined' ? localStorage.getItem('inv-theme') : null;
    if (saved === 'dark' || saved === 'cards') setTheme(saved);
  }, []);
  useEffect(() => {
    document.documentElement.setAttribute('data-theme', theme);
    localStorage.setItem('inv-theme', theme);
  }, [theme]);

  // ===== utils =====
  function downloadBlob(blob: Blob, filename: string) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = filename; a.style.display = 'none';
    document.body.appendChild(a); a.click(); a.remove();
    URL.revokeObjectURL(url);
  }

  // ===== chamadas API =====
  async function listarAbasFisico() {
    if (!fisicoFile) return alert('Selecione a planilha do Estoque Físico.');
    try {
      const fd = new FormData();
      fd.append('file', fisicoFile);
      // sheet vazio => backend responde {sheets: []}
      const res = await fetch('/api/fisico', { method: 'POST', body: fd });
      if (!res.ok) {
        alert(await res.text());
        return;
      }
      const data = await res.json() as { sheets: string[] };
      setFisicoSheets(data.sheets || []);
      if (data.sheets?.length) {
        setFisicoSheet(data.sheets[0]);
        setFisicoReady(false);
      }
    } catch (e: any) {
      alert(e?.message || 'Falha ao listar abas');
    }
  }

  // Botão verde 1 — não baixa: apenas valida (preview) e marca “✔️ Pronto”
  async function consolidarFisico() {
    if (!fisicoFile || !fisicoSheet) return alert('Selecione a planilha e a aba do Físico.');
    const fd = new FormData();
    fd.append('file', fisicoFile);
    fd.append('sheet', fisicoSheet);
    const res = await fetch('/api/fisico?preview=1', { method: 'POST', body: fd });
    if (!res.ok) return alert(await res.text());
    setFisicoReady(true);
  }

  // Botão verde 2 — não baixa: apenas valida (preview) e marca “✔️ Pronto”
async function consolidarSAP() {
  if (!sapFile) return alert('Selecione a planilha do SAP.');
  const fd = new FormData();
  fd.append('sap', sapFile);             // << aqui é o ajuste
  fd.append('sheet', sapSheet);
  fd.append('colDeposito', 'Código de depósito'); // opcional
  const res = await fetch('/api/sap?preview=1', { method: 'POST', body: fd });
  if (!res.ok) return alert(await res.text());
  setSapReady(true);
}

  // Analisar e exportar (Físico | SAP | ZIP)
  async function analisar() {
    if (!fisicoFile || !sapFile || !fisicoSheet) {
      return alert('Envie Físico, SAP e escolha a aba do Físico.');
    }
    setLoading(true);
    try {
      const fd = new FormData();
      fd.append('fisico', fisicoFile);
      fd.append('sap', sapFile);
      fd.append('fisicoSheet', fisicoSheet);
      fd.append('sapSheet', sapSheet);
      fd.append('exportType', exportType);

      const res = await fetch('/api/analisar', { method: 'POST', body: fd });
      setLoading(false);
      if (!res.ok) {
        const msg = await res.text();
        alert(msg || 'Falha na análise');
        return;
      }
      const blob = await res.blob();
      const name =
        exportType === 'fisico'
          ? 'ESTOQUE FISICO (analisado).xlsx'
          : exportType === 'sap'
          ? 'ESTOQUE SAP (analisado).xlsx'
          : 'resultado-analise.zip';
      downloadBlob(blob, name);
    } catch (e: any) {
      setLoading(false);
      alert(e?.message || 'Erro ao gerar análise');
    }
  }

  return (
    <main className="container">
      {/* Cabeçalho */}
      <div className="header">
        <h1 className="title">Conferência de Inventário — Físico × SAP</h1>
        <button className="btn theme-switch" onClick={() => setTheme(theme === 'dark' ? 'cards' : 'dark')}>
          {theme === 'dark' ? '🌗 Cards' : '🌙 Dark'}
        </button>
      </div>

      {/* 1) Físico */}
      <section className="card">
        <h2>1) Consolidação — Estoque Físico</h2>
        <div className="row center">
          <input
            type="file"
            accept=".xlsx,.xls"
            onChange={(e) => {
              setFisicoFile(e.target.files?.[0] ?? null);
              setFisicoReady(false);
              setFisicoSheets([]);
              setFisicoSheet('');
            }}
          />
          <button className="btn btn-ghost" onClick={listarAbasFisico}>Listar abas</button>
          <select
            value={fisicoSheet}
            onChange={(e) => { setFisicoSheet(e.target.value); setFisicoReady(false); }}
          >
            <option value="">— escolha a aba —</option>
            {fisicoSheets.map((s) => <option key={s} value={s}>{s}</option>)}
          </select>
          <button className="btn btn-primary" onClick={consolidarFisico}>
            Adicionar coluna “depósito SAP”
          </button>
          {fisicoReady && (
            <span className="btn btn-ghost" style={{ pointerEvents: 'none' }}>✔️ Pronto</span>
          )}
        </div>
      </section>

      {/* 2) SAP */}
      <section className="card">
        <h2>2) Consolidação — SAP</h2>
        <div className="row center">
          <input
            type="file"
            accept=".xlsx,.xls"
            onChange={(e) => { setSapFile(e.target.files?.[0] ?? null); setSapReady(false); }}
          />
          <input
            type="text"
            value={sapSheet}
            onChange={(e) => { setSapSheet(e.target.value); setSapReady(false); }}
            placeholder="Planilha1"
          />
          <button className="btn btn-primary" onClick={consolidarSAP}>
            Adicionar “detalhes” + filtrar depósitos
          </button>
          {sapReady && (
            <span className="btn btn-ghost" style={{ pointerEvents: 'none' }}>✔️ Pronto</span>
          )}
        </div>
      </section>

      {/* 3) Análise */}
      <section className="card">
        <h2>3) Análise — lotes, quantidades, observações, depósito SAP</h2>
        <div className="row center">
          <div className="segmented">
            <input id="exp_fisico" type="radio" name="exportar" checked={exportType === 'fisico'} onChange={() => setExportType('fisico')} />
            <label htmlFor="exp_fisico">Físico</label>

            <input id="exp_sap" type="radio" name="exportar" checked={exportType === 'sap'} onChange={() => setExportType('sap')} />
            <label htmlFor="exp_sap">SAP</label>

            <input id="exp_zip" type="radio" name="exportar" checked={exportType === 'zip'} onChange={() => setExportType('zip')} />
            <label htmlFor="exp_zip">ZIP (2)</label>
          </div>

          <button className="btn btn-accent" onClick={analisar} disabled={loading}>
            {loading ? 'Gerando…' : 'Analisar e Exportar'}
          </button>
        </div>

        <p className="note">
          OBS.: consolidado das analises das planilhas do físico com o SAP
        </p>
      </section>
    </main>
  );
}
