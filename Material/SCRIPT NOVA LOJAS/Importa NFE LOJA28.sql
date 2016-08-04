use [DMAC_Loja]

select ECFNF, lojaorigem, nf, * from nfcapa where serie in ('CF','NE') and tiponota = 'V' and dataemi = '2014/07/18' and lojaorigem = '28'
select lojaorigem, * from nfitens where serie in ('CF','NE') and tiponota = 'V' and dataemi = '2014/07/18' and lojaorigem = '28'
select MC_Loja, * from movimentocaixa where mc_data = '2014/07/18' and mc_serie in ('CF','NE') and mc_tiponota = 'V'

insert into [demeoserver].[demeo].[dbo].nfcapa select * from nfcapa where serie in ('CF','NE') and tiponota = 'V' and dataemi = '2014/07/18' and lojaorigem = '28'
insert into [demeoserver].[demeo].[dbo].nfitens select * from nfitens where serie in ('CF','NE') and tiponota = 'V' and dataemi = '2014/07/18' and lojaorigem = '28'
insert into [demeoserver].[demeo].[dbo].movimentocaixa
(MC_NumeroECF,
MC_CodigoOperador,
MC_Loja,
MC_Data,
MC_Grupo,
MC_Documento,
MC_Serie,
MC_Valor,
MC_Banco,
MC_Agencia,
MC_ContaCorrente,
MC_NumeroCheque,
MC_BomPara,
MC_Parcelas,
MC_Remessa,
MC_Sequencia,
MC_SituacaoEnvio,
MC_ControleAVR,
MC_DataBaixaAVR)
select 
MC_NumeroECF,
MC_CodigoOperador,
MC_Loja,
MC_Data,
MC_Grupo,
MC_Documento,
MC_Serie,
MC_Valor,
MC_Banco,
MC_Agencia,
MC_ContaCorrente,
MC_NumeroCheque,
MC_BomPara,
MC_Parcelas,
MC_Remessa,
MC_Sequencia,
MC_SituacaoEnvio,
MC_ControleAVR,
MC_DataBaixaAVR 
from MovimentoCaixa  
where mc_data = '2014/07/18' 
and mc_serie IN ('CF','NE')
and mc_tiponota = 'V'
and MC_loja = '28'

/*

update nfcapa set ECFNF = '' where serie in ('CF','NE') and tiponota = 'V' and dataemi = '2014/07/18' 

sp_help nfcapa
sp_help nfitens
sp_help movimentocaixa
*/