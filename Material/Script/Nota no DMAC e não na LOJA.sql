Nota no DMAC e não na LOJA

INSERT INTO DMAC535.DMAC_LOJA.DBO.nfitens 
select * from nfitens as dmac where dataemi between '2016/07/13' and '2016/07/13' and lojaorigem = '535' and not exists 
(select NF from DMAC535.DMAC_LOJA.DBO.nfitens as loja where dmac.nf = loja.nf and dmac.lojaorigem = loja.lojaorigem and
dmac.DATAEMI = loja.DATAEMI)

delete DMAC535.DMAC_LOJA.DBO.nfitens where numeroped in (select numeroped from nfitens as dmac where dataemi between '2016/07/13' and '2016/07/13' and lojaorigem = '535' and not exists 
(select NF from DMAC535.DMAC_LOJA.DBO.nfitens as loja where dmac.nf = loja.nf and dmac.lojaorigem = loja.lojaorigem and
dmac.DATAEMI = loja.DATAEMI) group by numeroped) and tiponota = 'PD'