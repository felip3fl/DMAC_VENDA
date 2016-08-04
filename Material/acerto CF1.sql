select * from dmac85.dmac_loja.dbo.controleserie
update dmac85.dmac_loja.dbo.ControleSerie set CS_SerieCF = 'CF1'

select * from dmac85.dmac_loja.dbo.nfitens

select serie, nf, * from dmac85.dmac_loja.dbo.nfcapa where dataemi > '2015/08/20' and tiponota = 'V' and SERIE = 'CF' and nf between 1 and 1000 and lojaorigem = '85'
--UPDATE dmac85.dmac_loja.dbo.nfcapa set serie = 'CF1' where dataemi > '2015/08/20' and tiponota = 'V' and SERIE = 'CF' and nf between 1 and 1000 and lojaorigem = '85'

select serie, nf, * from dmac85.dmac_loja.dbo.nfitens where dataemi > '2015/08/20' and tiponota = 'V' and SERIE = 'CF' and nf between 1 and 1000 and lojaorigem = '85'
--UPDATE dmac85.dmac_loja.dbo.nfitens set serie = 'CF1' where dataemi > '2015/08/20' and tiponota = 'V' and SERIE = 'CF' and nf between 1 and 1000 and lojaorigem = '85'

select * from CapaNFVenda where VC_serie = 'CF1' and VC_dataemissao > '2015/08/20' and VC_tiponota = 'V' and VC_SERIE = 'CF' and VC_notafiscal between 1 and 1000 and VC_lojaorigem = '85'
--select * from CapaNFVenda where VC_serie = 'CF1' and VC_dataemissao > '2015/08/20' and VC_tiponota = 'V' and VC_SERIE = 'CF' and VC_notafiscal between 1 and 1000 and VC_lojaorigem = '85'

select * from itemnfvenda where Vi_serie = 'CF1' and Vi_dataemissao > '2015/08/20' and Vi_tiponota = 'V' and Vi_SERIE = 'CF' and Vi_notafiscal between 1 and 1000 and Vi_lojaorigem = '85'
--select * from CapaNFVenda where Vi_serie = 'CF1' and Vi_dataemissao > '2015/08/20' and Vi_tiponota = 'V' and Vi_SERIE = 'CF' and Vi_notafiscal between 1 and 1000 and Vi_lojaorigem = '85'


EXEC SP_Atualiza_Lojas '535','''ALTER TABLE nfcapa ALTER COLUMN SerieDevolucao nvarchar(3)'''
EXEC SP_Atualiza_Lojas '535','''ALTER TABLE nfcapa ALTER COLUMN SerieDevolucao nvarchar(3)'''