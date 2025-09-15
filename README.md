# Conferência Inventário – Vercel (Next.js)

Fluxo implementado conforme especificação do Alex:
- Consolidação FÍSICO: escolher aba + adicionar coluna **"depósito SAP"** e salvar "ESTOQUE FISICO + data + hora".
- Consolidação SAP: adicionar **"detalhes"** + excluir depósitos {CHEM WIP, V_CHEWP, SC_Nest, SC_DSB, SC_DSB_2, TST, MS WIP, EM DSB} e salvar "ESTOQUE SAP + data + hora".
- Análise por lote (apenas lotes do Físico), comparar quantidades e preencher:
  - **OBSERVAÇÕES (Físico)**: divergências; se faltante no SAP → "Lote não Localizado no SAP".
  - **depósito SAP (Físico)**: mapeado por **DS. FSA** via **DEPÓSITOS - PROCEDIMENTO.json** (UDL/AV_UDL/DESC_UDL/AM_UDL/UDL_MP/V_UDL).
  - **detalhes (SAP)**: se lote não existir no Físico (mesmo Nº do item) → "Lote não Localizado no Fisico".

## Rodar local
pnpm i
pnpm dev

## Deploy
- Suba no GitHub → "New Project" na Vercel → Import.
- Build Command: `next build` ; Output: `.next`
