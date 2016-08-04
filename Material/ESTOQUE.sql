USE DMAC

-- Ultimo Acerto = '23/10/2014'
--- COMPRARACAO -------------------------------------------------------------------------------------------------------------------------------

select es_loja,es_referencia,es_estoque,el_estoque from [dmac28].[dmac_loja].[dbo].estoqueloja,
estoque where el_referencia=es_referencia and es_estoque<>el_estoque and es_loja=el_loja

select es_referencia as referencia,es_estoque as estoqueDEMEO,el_estoqueanterior as estoqueAnteriorLOJA28 
from [dmac28].[dmac_loja].[dbo].estoqueloja,
[DemeoServer].[Demeo].[dbo].estoque
where el_referencia=es_referencia and es_estoque<>el_estoqueanterior and es_loja=el_loja 

select es_referencia as referencia, es_estoque as estoqueDEMEO, es_estoque as estoqueAnteriorLOJA28 , ES_Transito as Transito
from [DemeoServer].[Demeo].[dbo].estoque
where ES_Transito < 0 and ES_Loja = 28


--- LOJA / DMAC ------------------------------------------------------------------------------------------------------------------------------

--Update dmac..Loja set LO_Conexao='S' where LO_Loja='28'
--update [dmac28].[dmac_loja].[dbo].NFCAPA set dataprocesso = '2014/10/23' where nf = 2407 and serie = 'ne' and tiponota = 'T'
--update [dmac28].[dmac_loja].[dbo].nfitens set dataprocesso = '2014/10/23' where nf = 2407 and serie = 'ne' and tiponota = 'T'
--SELECT * FROM [dmac28].[dmac_loja].[dbo].nfcapa where nf = 2685 and serie = 'NE' and tiponota = 'T'
--SELECT NF,SERIE,* FROM nfcapa where nf = 2407 and serie = 'ne' and tiponota = 'T'
--SELECT NF,SERIE,* FROM nfitens where nf = 2407 and serie = 'ne' and tiponota = 'T'
select nf,tiponota,* from [dmac28].[dmac_loja].[dbo].nfitens where Referencia in (0600222) and DataEmi > '2014/12/01' 
select nf,tiponota,* from nfitens where Referencia in (0600222) and DataEmi > '2014/12/01' 
select * from ajuste where aj_loja = '28' and AJ_Referencia in(0600222)

select nf,* from nfcapa where dataemi = '2014/11/14' and serie = 'NE'

select VC_Observacao, * from ItemNFVenda, capanfvenda 
where VI_Referencia = (4330067) and VI_DataEmissao > '2014/10/01' and VI_NotaFiscal = VC_NotaFiscal and VI_Serie = VC_Serie and VC_LojaOrigem = VI_LojaOrigem
and VC_LojaDestino = '28'

select VC_Observacao,* from capanfvenda where VC_Notafiscal = '2534'
select VC_Observacao,* from capanfvenda where VC_Observacao = '0' and vc_tiponota = 'T' and VC_DataEmissao > '2014/10/01' and VC_LojaDestino = '28'
select * from itemnfvenda where Vi_Notafiscal = '2534' 

--- DEMEOSERVER --------------------------------------------------------------------------------------------------------------------------------------

select * from [demeoserver].[demeo].[dbo].ItemNFVenda where VI_Referencia = (1880261) and VI_DataEmissao > '2014/11/01' and VI_LojaOrigem = '28'

select VC_NotaFiscal,VC_Observacao, * from [demeoserver].[demeo].[dbo].ItemNFVenda, [demeoserver].[demeo].[dbo].capanfvenda 
where VI_Referencia = (1880261) and VI_DataEmissao > '2014/11/01' and VI_NotaFiscal = VC_NotaFiscal and VI_Serie = VC_Serie and VC_LojaOrigem = VI_LojaOrigem
and VC_LojaDestino = '28'

-----------------------------------------------------------------------------------------------------------------------------------------------

select ES_Estoque,es_estoqueAnterior,ES_Transito, * from estoque where es_loja = '28' and ES_Referencia in ('0618029') order by ES_Referencia
--update estoque set ES_Estoque = 1, es_transito = 0, es_estoqueAnterior = 1  where es_loja = '28' and ES_Referencia = '0618029'
select El_Estoque,el_estoqueAnterior, * from [dmac28].[dmac_loja].[dbo].estoqueloja where el_loja = '28' and El_Referencia in ('0618029') order by el_referencia
--update [dmac28].[dmac_loja].[dbo].estoqueloja set El_Estoque = 0, el_estoqueAnterior = 0 where el_loja = '28' and El_Referencia = '0618029'
select ES_Estoque,es_estoqueAnterior,ES_Transito, * from [demeoserver].[demeo].[dbo].estoque where es_loja = '28' and ES_Referencia in ('0618029') order by ES_Referencia
--update [demeoserver].[demeo].[dbo].estoque set ES_Estoque = 0, es_transito = 0, es_estoqueAnterior = 26 where eS_loja = '28' and ES_Referencia = '7150017'

-----------------------------------------------------------------------------------------------------------------------------------------------

/*

update [dmac28].[dmac_loja].[dbo].estoqueloja set el_estoque = ES_Estoque, el_estoqueAnterior = es_estoqueAnterior
from estoque, [dmac28].[dmac_loja].[dbo].estoqueloja 
where es_loja = '28' and ES_Referencia in (select aj_referencia from ajuste where aj_data >= '2014/09/06' and aj_loja = '28' ) and es_referencia = el_referencia and el_loja = es_loja

select * from ajuste where aj_loja = '28' and AJ_data in('2014/11/06')

UPDATE nf

*/

select es_loja,es_referencia,es_estoque,el_estoque from [dmac28].[dmac_loja].[dbo].estoqueloja,
estoque where el_referencia=es_referencia and es_estoque<>el_estoque and es_loja=el_loja

select es_referencia,es_estoque,el_estoqueanterior from [dmac28].[dmac_loja].[dbo].estoqueloja,
[DemeoServer].[Demeo].[dbo].estoque 
where el_referencia=es_referencia and es_estoque<>el_estoqueanterior and es_loja=el_loja





-----------------------------------------------------------------------------------------------------------------------------------------------

select es_referencia as referencia,es_estoque as estoqueDEMEO,el_estoqueanterior as estoqueAnteriorLOJA28 
from [dmac28].[dmac_loja].[dbo].estoqueloja,
[DemeoServer].[Demeo].[dbo].estoque
where el_referencia=es_referencia and es_estoque<>el_estoqueanterior and es_loja=el_loja 
AND ES_Referencia not in (select REFERENCIA from [dmac28].[dmac_loja].[dbo].nfitens where DATAEMI = '2014/11/06' and tiponota in ('V','T'))

--drop table [dmac28].[dmac_loja].[dbo].estoquelojaTEMP0611
--select * into [dmac28].[dmac_loja].[dbo].estoquelojaTEMP0611 from [dmac28].[dmac_loja].[dbo].estoqueloja

--drop table estoqueTEMP0611
--select * into estoqueTEMP0611 from estoque

/*

select * from ajuste where aj_loja = '28' and AJ_data in('2014/11/06')
update [dmac28].[dmac_loja].[dbo].estoqueloja set el_estoqueAnterior = el_estoqueAnterior - aj_quantidade from ajuste where 
AJ_data in('2014/11/06') and aj_loja = '28' and aj_referencia = el_referencia and el_loja = '28' and el_referencia = '0130053'


update [dmac28].[dmac_loja].[dbo].estoqueloja 
set EL_Estoque = es_estoque, EL_EstoqueAnterior = es_estoque 
from [dmac28].[dmac_loja].[dbo].estoqueloja, [DemeoServer].[Demeo].[dbo].estoque 
where ES_Referencia = EL_Referencia 
and ES_Loja = EL_Loja 
and es_estoque<>el_estoqueanterior
AND ES_Referencia not in (select REFERENCIA from [dmac28].[dmac_loja].[dbo].nfitens where DATAEMI = '2014/12/04' and tiponota in ('V','T'))


update [dmac28].[dmac_loja].[dbo].estoqueloja 
set EL_Estoque = es_estoque - (select top 1 QTDE from [dmac28].[dmac_loja].[dbo].nfitens where DATAEMI = '2014/12/04' and tiponota in ('V','T') and referencia = es_referencia and referencia = EL_Referencia)
from [dmac28].[dmac_loja].[dbo].estoqueloja, [DemeoServer].[Demeo].[dbo].estoque 
where ES_Referencia = EL_Referencia 
and ES_Loja = EL_Loja 
and es_estoque<>el_estoque
AND ES_Referencia in (select referencia from [dmac28].[dmac_loja].[dbo].nfitens where DATAEMI = '2014/12/04' and tiponota in ('V','T'))

update estoque
set es_estoque = EL_Estoque, ES_EstoqueAnterior = EL_EstoqueAnterior
from [dmac28].[dmac_loja].[dbo].estoqueloja, estoque 
where ES_Referencia = EL_Referencia 
and ES_Loja = EL_Loja 
and es_estoque<>el_estoque
AND ES_Referencia not in (select REFERENCIA from [dmac28].[dmac_loja].[dbo].nfitens where DATAEMI = '2014/12/04' and tiponota in ('V','T'))

update estoque
set ES_EstoqueAnterior = demeo.ES_Estoque
from [DemeoServer].[Demeo].[dbo].estoque as demeo, estoque as dmac
where demeo.ES_Referencia = dmac.ES_Referencia 
and demeo.ES_Loja = dmac.ES_Loja 
AND dmac.ES_Referencia not in (select REFERENCIA from [dmac28].[dmac_loja].[dbo].nfitens where DATAEMI = '2014/11/06' and tiponota in ('V','T'))

update [DemeoServer].[Demeo].[dbo].estoque
set ES_Transito = 0,
es_estoque = es_estoque + (es_transito) 
where  es_loja = '28' 
and es_transito < 0 
and es_referencia = '6800982'

SELECT * FROM estoque, [dmac28].[dmac_loja].[dbo].estoqueloja where el_referencia = es_referencia and el_loja = es_loja and es_estoqueanterior <> el_estoqueanterior


SP_EST_Movimentacao_Estoque_Direto

*/