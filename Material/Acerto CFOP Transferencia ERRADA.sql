select * from capanfvenda where vc_dataemissao between '2015/03/01' and '2015/03/20' and vc_tiponota = 'T' AND vc_codigooperacao in ('5102','5405')
select * from itemnfvenda where vi_dataemissao between '2015/03/01' and '2015/03/20' and vi_tiponota = 'T' AND vi_cfop in ('5102','5405')

select situacaoProcesso,CFOAUX,codoper,nf, * from nfcapa where tiponota = 'T' and dataemi between '2015/03/01' and '2015/03/20' and CFOAUX in ('5102','5405')
select nf, * from nfitens where tiponota = 'T' and dataemi between '2015/03/01' and '2015/03/20' and cfop in ('5102','5405')

select * from nfitens where dataemi = '2015/03/19' and numeroped = 2942

select CFOAUX, CODOPER, * from nfcapa WHERE NUMEROPED = '2963' and LOJAORIGEM = '646' and DATAEMI = '2015/03/19'


select * from codigooperacaonovo where cn_codigooperacaonovo in ('5405','5102')
select * from codigooperacaonovo where cn_codigooperacaonovo in ('5409','5152')

UPDATE nfcapa set CFOAUX = '5409', CODOPER = '5409' where tiponota = 'T' and dataemi between '2015/03/01' and '2015/03/20' and CFOAUX in ('5405')
UPDATE nfcapa set CFOAUX = '5152', CODOPER = '5152' where tiponota = 'T' and dataemi between '2015/03/01' and '2015/03/20' and CFOAUX in ('5102')

UPDATE nfitens set CFOP = '5409' where tiponota = 'T' and dataemi between '2015/03/01' and '2015/03/20' and CFOP in ('5405')
UPDATE nfitens set CFOP = '5152' where tiponota = 'T' and dataemi between '2015/03/01' and '2015/03/20' and CFOP in ('5102')

--update capanfvenda set vc_codigoOperacao = '5152' where vc_dataemissao between '2015/03/01' and '2015/03/20' and vc_tiponota = 'T' and vc_lojaorigem = '646' AND vc_codigooperacao in ('5102')