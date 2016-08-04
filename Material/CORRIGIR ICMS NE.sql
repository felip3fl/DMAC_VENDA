declare @nf varchar(10)
declare @serie varchar(10)
declare @loja varchar(10)

set @nf = '12254'
set @serie = 'NE'
set @loja = 'CD'

/*
DELETE CAPANFVENDA WHERE VC_NotaFiscal = @nf and VC_Serie = @serie and VC_LojaOrigem = @loja
DELETE ITEMNFVENDA WHERE VI_NotaFiscal = @nf and VI_Serie = @serie and VI_LojaOrigem = @loja
EXEC SP_VDA_Cria_Capa_Item_NFVenda '2016/07/01','2016/07/27'
*/


update NfItens set BASEICMS = round(VLTOTITEM,2) where 
nf = @nf and serie = @serie and LOJAORIGEM = @loja and BASEICMS = 0
update NfItens set VALORICMS = round(((BASEICMS * ICMSAplicado) / 100),2) where 
nf = @nf and serie = @serie and LOJAORIGEM = @loja

update itemnfvenda set VI_BaseICMS = round(VI_ValorMercadoria,2) where 
VI_NotaFiscal = @nf and VI_Serie = @serie and VI_LojaOrigem = @loja and VI_BaseICMS = 0
update itemnfvenda set VI_ValorICMS = round(((VI_BaseICMS * VI_AliquotaICMS) / 100),2) where 
VI_NotaFiscal = @nf and VI_Serie = @serie and VI_LojaOrigem = @loja

       --       -       -       -       -       -       -       -       -       -       -       -       -       -       -       -

update NfCapa set vlrICMS = round((select SUM(VALORICMS) as total from NfItens where nf = @nf and serie = @serie AND LojaOrigem = @loja),2) where 
nf = @nf and serie = @serie and LOJAORIGEM = @loja
update NfCapa set BASEICMS = round((select SUM(BASEICMS) as total from NfItens where nf = @nf and serie = @serie AND LojaOrigem = @loja),2) where 
nf = @nf and serie = @serie and LOJAORIGEM = @loja

update capanfvenda set VC_BaseICMS = round((select SUM(VI_BaseICMS) as total from itemnfvenda where VI_NotaFiscal = @nf and VI_Serie = @serie AND VI_LojaOrigem = @loja),2) where 
VC_NotaFiscal = @nf and VC_Serie = @serie and VC_LojaOrigem = @loja

update capanfvenda set VC_ValorICMS = round((select SUM(VI_ValorICMS) as total from itemnfvenda where VI_NotaFiscal = @nf and VI_Serie = @serie AND VI_LojaOrigem = @loja),2) where 
VC_NotaFiscal = @nf and VC_Serie = @serie and VC_LojaOrigem = @loja


/*
select VALORICMS, * from nfitens where nf = 12254 and serie = 'NE' and lojaorigem = 'CD'
select vlrICMS,* from nfcapa where nf = 12254 and serie = 'NE' and lojaorigem = 'CD'


select VC_ValorICMS, * from CAPANFVENDA where vc_notafiscal = 12254 and vc_serie = 'NE' AND vc_LOJAORIGEM = 'CD'
select VI_ValorICMS, VI_AliquotaICMS,VI_BaseICMS, * from itemnfvenda where vi_notafiscal in (12254) and vi_serie = 'NE' AND vi_LOJAORIGEM = 'CD'

SELECT * FROM PRODUTO WHERE PR_REFERENCIA = '9920720'




*/