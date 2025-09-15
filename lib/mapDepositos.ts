import { ProcedimentoDepositos, DepositoCode } from './types';

// Padrões de palavra-chave em DS. FSA → depósito SAP
const defaultRules: [RegExp, DepositoCode][] = [
  [/AVARIA/i, 'AV_UDL'],
  [/DESCARTE|DESC\.*:/i, 'DESC_UDL'],
  [/AMOSTRA|AMOSTRAS|AMOSTRA COLETA/i, 'AM_UDL'],
  [/BLENDA|PRODUÇÃO|PRODUCAO|MP/i, 'UDL_MP'],
  [/VENCID/i, 'V_UDL'],
];

export function carregarMapa(json?: ProcedimentoDepositos) {
  // Se o JSON vier do upload do usuário, priorizamos ele.
  // (Ele define os depósitos e descrições oficiais.)
  // :contentReference[oaicite:3]{index=3}
  const lista = json?.deposits?.map(d => d.code) ?? ['UDL','AV_UDL','DESC_UDL','AM_UDL','UDL_MP','V_UDL'];
  return {
    decidirDepositoPorDSFSA(dsfsa?: string): DepositoCode {
      const s = (dsfsa ?? '').toString();
      for (const [rx, dep] of defaultRules) if (rx.test(s)) return dep;
      return 'UDL'; // vazio/OK => UDL
    },
    validCodes: new Set(lista as DepositoCode[]),
  };
}
