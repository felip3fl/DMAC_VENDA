select * from MovimentoCaixa where MC_Grupo = '11006' and mc_loja = '364'

insert into dmac874.dmac_loja.dbo.movimentocaixa(MC_NumeroECF,MC_CodigoOperador,MC_Loja,MC_Data,MC_Grupo,MC_SubGrupo,MC_Documento,MC_Serie,MC_Valor,MC_Banco,MC_Agencia,MC_ContaCorrente,MC_NumeroCheque,MC_BomPara,MC_Parcelas,MC_Remessa,MC_SituacaoEnvio,MC_ControleAVR,MC_DataBaixaAVR,MC_Protocolo,MC_NroCaixa,MC_GrupoAuxiliar,MC_Situacao,MC_Pedido,MC_DataProcesso,MC_TipoNota,MC_SequenciaTEF) 
select MC_NumeroECF,MC_CodigoOperador,'874','2015/01/23',MC_Grupo,MC_SubGrupo,MC_Documento,MC_Serie,1665.45,MC_Banco,MC_Agencia,MC_ContaCorrente,MC_NumeroCheque,MC_BomPara,MC_Parcelas,MC_Remessa,MC_SituacaoEnvio,MC_ControleAVR,MC_DataBaixaAVR,'8',MC_NroCaixa,MC_GrupoAuxiliar,MC_Situacao,MC_Pedido,MC_DataProcesso,MC_TipoNota,MC_SequenciaTEF 
from movimentocaixa where MC_Grupo = '11006' and mc_loja = '364'

