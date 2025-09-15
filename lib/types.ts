export type DepositoCode = 'UDL'|'AV_UDL'|'DESC_UDL'|'AM_UDL'|'UDL_MP'|'V_UDL';

export type ProcedimentoDepositos = {
  documento: string;
  deposits: { name: string; code: DepositoCode; description: string; temporary?: boolean; }[];
  post_validation_rule?: any;
};

export type FluxoInputs = {
  // nomes de colunas “canônicos” conforme seus padrões
  fisico: {
    codigo: string; descricao: string; qtd: string; unid: string; dsfsa: string;
    fabricacao: string; validade: string; lote: string; observacoes: string;
    depositoSap: string; // coluna criada
  };
  sap: {
    numItem: string; item: string; deposito: string; lote: string;
    qtde: string; custo: string; fabricacao: string; admissao: string; vencimento: string;
    detalhes: string; col1?: string;
  };
};
