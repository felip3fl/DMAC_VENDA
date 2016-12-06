use [DMAC_Loja]

/*
insert into [serverwin28].[loja028].[dbo].nfcapa (
NUMEROPED,
DATAEMI,
VENDEDOR,
VLRMERCADORIA,
DESCONTO,
SUBTOTAL,
LOJAORIGEM,
TIPONOTA,
CONDPAG,
AV,
CLIENTE,
CODOPER,
DATAPAG,
PGENTRA,
LOJAT,
QTDITEM,
PEDCLI,
TM,
PESOBR,
PESOLQ,
VALFRETE,
FRETECOBR,
OUTRALOJA,
OUTROVEND,
NF,
TOTALNOTA,
NATOPERACAO,
DATAPED,
BASEICMS,
ALIQICMS,
VLRICMS,
SERIE,
HORA,
TOTALIPI,
ECF,
NUMEROSF,
NOMCLI,
FONECLI,
CGCCLI,
INSCRICLI,
ENDCLI,
UFCLIENTE,
MUNICIPIOCLI,
BAIRROCLI,
CEPCLI,
PESSOACLI,
REGIAOCLI,
CFOAUX,
AnexoAUx,
PAGINANF,
ECFNF,
Carimbo1,
Carimbo2,
Carimbo3,
Carimbo4,
Carimbo5,
CustoMedioLiquido,
VendaLiquida,
MargemContribuicao,
ValorTotalCodigoZero,
TotalNotaAlternativa,
ValorMercadoriaAlternativa,
SituacaoEnvio,
VendedorLojaVenda,
LojaVenda,
NotaCredito,
NfDevolucao,
SerieDevolucao,
EmiteDataSaida,
CancelarNota,
HoraManutencao,
DataProcessamento,
SituacaoFec,
ObsSituacaoFec,
Parcelas,
ModalidadeVenda,
SenhaDesconto,
Volume,
TipoFrete,
ParcelasTEF,
AutorizacaoTEF,
NroResidencia,
CompleResidencia,
GarantiaEstendida,
TotalGarantia,
SeguroPremiado,
CertificadoSorte,
numeroSorte,
sp_premioLiquido,
sp_IOF,
sp_valorRemuneracao,
sp_percentualRemuneracao,
sp_valorRepasse,
vendedorGarantia)
select 
NUMEROPED,
DATAEMI,
VENDEDOR,
VLRMERCADORIA,
DESCONTO,
SUBTOTAL,
LOJAORIGEM,
TIPONOTA,
CONDPAG,
AV,
CLIENTE,
CODOPER,
DATAPAG,
PGENTRA,
LOJAT,
QTDITEM,
PEDCLI,
TM,
PESOBR,
PESOLQ,
VALFRETE,
FRETECOBR,
OUTRALOJA,
OUTROVEND,
NF,
TOTALNOTA,
NATOPERACAO,
DATAPED,
BASEICMS,
ALIQICMS,
VLRICMS,
SERIE,
HORA,
TOTALIPI,
ECF,
NUMEROSF,
NOMCLI,
FONECLI,
CGCCLI,
INSCRICLI,
ENDCLI,
UFCLIENTE,
MUNICIPIOCLI,
BAIRROCLI,
CEPCLI,
PESSOACLI,
REGIAOCLI,
CFOAUX,
AnexoAUx,
PAGINANF,
ECFNF,
Carimbo1,
Carimbo2,
Carimbo3,
Carimbo4,
Carimbo5,
CustoMedioLiquido,
VendaLiquida,
MargemContribuicao,
ValorTotalCodigoZero,
TotalNotaAlternativa,
ValorMercadoriaAlternativa,
SituacaoEnvio,
VendedorLojaVenda,
LojaVenda,
NotaCredito,
NfDevolucao,
SerieDevolucao,
EmiteDataSaida,
CancelarNota,
HoraManutencao,
DataProcessamento,
SituacaoFec,
ObsSituacaoFec,
Parcelas,
ModalidadeVenda,
SenhaDesconto,
Volume,
TipoFrete,
ParcelasTEF,
AutorizacaoTEF,
NroResidencia,
CompleResidencia,
GarantiaEstendida,
TotalGarantia,
SeguroPremiado,
CertificadoSorte,
numeroSorte,
sp_premioLiquido,
sp_IOF,
sp_valorRemuneracao,
sp_percentualRemuneracao,
sp_valorRepasse,
vendedorGarantia
from nfcapa where
serie in ('CF','NE')
and tiponota = 'V' 
and dataemi = '2014/07/18' 
and lojaorigem = '28'

insert into [serverwin28].[loja028].[dbo].nfitens (
NUMEROPED,
DATAEMI,
REFERENCIA,
QTDE,
VLUNIT,
VLUNIT2,
VLTOTITEM,
DESCRAT,
ICMS,
ITEM,
VLIPI,
DESCONTO,
PLISTA,
COMISSAO,
VALORICMS,
BCOMIS,
CSPROD,
LINHA,
SECAO,
VBUNIT,
ICMPDV,
CODBARRA,
NF,
SERIE,
LOJAORIGEM,
CLIENTE,
VENDEDOR,
ALIQIPI,
TIPONOTA,
REDUCAOICMS,
BASEICMS,
TIPOMOVIMENTACAO,
DETALHEIMPRESSAO,
SerieProd1,
SerieProd2,
CustoMedioLiquido,
VendaLiquida,
MargemContribuicao,
EncargosVendaLiquida,
EncargosCustoMedioLiquido,
PrecoUnitAlternativa,
ValorMercadoriaAlternativa,
ReferenciaAlternativa,
SituacaoEnvio,
DescricaoAlternativa,
Tributacao,
IcmsMargem,
PisCofins,
DeducoesVendas,
EncargosFinanceiros,
EstoqueAntes,
EstoqueDepois,
GarantiaEstendida,
PlanoGarantia,
CoeficientePlano,
QtdeGarantia,
ValorGarantia,
CertificadoInicio,
CertificadoFim,
ge_premioLiquido,
ge_IOF,
ge_dataInicioVigencia,
ge_dataFinalVigencia,
ge_valorCustoSeguradora
)
select 
NUMEROPED,
DATAEMI,
REFERENCIA,
QTDE,
VLUNIT,
VLUNIT2,
VLTOTITEM,
DESCRAT,
ICMS,
ITEM,
VLIPI,
DESCONTO,
PLISTA,
COMISSAO,
VALORICMS,
BCOMIS,
CSPROD,
LINHA,
SECAO,
VBUNIT,
ICMPDV,
CODBARRA,
NF,
SERIE,
LOJAORIGEM,
CLIENTE,
VENDEDOR,
ALIQIPI,
TIPONOTA,
REDUCAOICMS,
BASEICMS,
TIPOMOVIMENTACAO,
DETALHEIMPRESSAO,
SerieProd1,
SerieProd2,
CustoMedioLiquido,
VendaLiquida,
MargemContribuicao,
EncargosVendaLiquida,
EncargosCustoMedioLiquido,
PrecoUnitAlternativa,
ValorMercadoriaAlternativa,
ReferenciaAlternativa,
SituacaoEnvio,
DescricaoAlternativa,
Tributacao,
IcmsMargem,
PisCofins,
DeducoesVendas,
EncargosFinanceiros,
EstoqueAntes,
EstoqueDepois,
GarantiaEstendida,
PlanoGarantia,
CoeficientePlano,
QtdeGarantia,
ValorGarantia,
CertificadoInicio,
CertificadoFim,
ge_premioLiquido,
ge_IOF,
ge_dataInicioVigencia,
ge_dataFinalVigencia,
ge_valorCustoSeguradora
from nfitens where
serie in ('CF','NE')
and tiponota = 'V' 
and dataemi = '2014/07/18' 
and lojaorigem = '28'
*/


insert into [serverwin28].[loja028].[dbo].movimentocaixa (
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
MC_SituacaoEnvio,
MC_ControleAVR,
MC_DataBaixaAVR
)
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
MC_SituacaoEnvio,
MC_ControleAVR,
MC_DataBaixaAVR
from movimentocaixa where
mc_data = '2014/07/18' 
and mc_serie in ('CF','NE')
and mc_tiponota = 'V' 
and mc_loja = '28'

/*
select * from nfcapa where serie IN ('CF','NE') and tiponota = 'V' and dataemi >= '2014/07/18' and lojaorigem = '28' and AnexoAUx is not null
select * from nfitens where serie IN ('CF','NE') and tiponota = 'V' and dataemi >= '2014/07/18' and lojaorigem = '28' and EncargosCustoMedioLiquido is not null

select * from movimentocaixa

select * from movimentocaixa  where
mc_data = '2014/07/17' 
and mc_serie in ('CF')
and mc_tiponota = 'V' 
and mc_loja = '28'

update nfitens set ReferenciaAlternativa = 0 where serie IN ('CF','NE') and tiponota = 'V' and dataemi >= '2014/07/18' and ReferenciaAlternativa is null

sp_help nfcapa
sp_help nfitens
sp_help movimentocaixa
*/