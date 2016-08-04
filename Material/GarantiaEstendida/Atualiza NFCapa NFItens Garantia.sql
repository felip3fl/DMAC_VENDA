use demeo
declare @nota as char(6)
declare @serie char(2)
declare @tiponota char(20) 
declare @dataemi CHAR(10)
declare @lojaorigem char(5)

set @nota = '1204'
set @serie  = 'NE'
set @tiponota = 'V'
set @dataemi = '2014/05/26'
set @lojaorigem = '48'

--select garantiaEstendida, TotalGarantia, nf,serie,TIPONOTA, * from nfcapa where lojaorigem = '48' AND dataemi >  '2014/05/25' and nf = '1204' 
--select CERTIFICADOINICIO, nf,serie,TIPONOTA, * from NFITENS where lojaorigem = '48' AND dataemi >  '2014/05/25' and nf = '1204' ORDER BY CERTIFICADOINICIO
--garantiaEstendida = l.garantiaEstendida,
--TotalGarantia = l.TotalGarantia    

update nfcapa set
numeroSorte = null
from [serverwin48].[loja048].[dbo].nfcapa as l, nfcapa as d
where d.serie = l.serie collate sql_latin1_general_cp1_ci_as
and d.tiponota = l.tiponota collate sql_latin1_general_cp1_ci_as
and d.dataemi = l.dataemi 
and d.lojaorigem = l.lojaorigem collate sql_latin1_general_cp1_ci_as
and l.nf = d.nf 
and d.serie = @serie 
and d.tiponota = @tiponota 
and d.dataemi = @dataemi 
and d.lojaorigem = @lojaorigem 
and d.nf = @nota 



select garantiaEstendida, planoGarantia, coeficientePlano, QtdeGarantia, ValorGarantia, certificadoinicio, certificadoFIM, ge_premioLiquido, ge_iof, ge_dataInicioVigencia, ge_dataFinalVigencia, ge_valorCustoSeguradora, nf,serie,TIPONOTA, referencia, item, * from NFITENS 
where lojaorigem = '48' and garantiaEstendida in ('S','C') AND dataemi >  '2014/05/10' and nf = 1204 ORDER BY CERTIFICADOINICIO
select garantiaEstendida, totalGarantia, nf,serie,TIPONOTA, * from nfcapa where lojaorigem = '48' and garantiaEstendida in ('S','C') AND dataemi >  '2014/05/10' and nf = 1204

select * from nfcapa where lojaorigem = '48' and garantiaEstendida in ('S','C') and dataemi  '2014/05/26'
select CERTIFICADOINICIO, * from NFITENS where lojaorigem = '48' and garantiaEstendida in ('S','C') AND dataemi >  '2014/05/10'  ORDER BY CERTIFICADOINICIO

/*
update nfcapa set 
garantiaEstendida = 'S', 
totalGarantia = 29.17
where nf = 1204 
and serie = 'NE' 
and tiponota = 'V' 
and lojaorigem = '48'

update nfitens set 
garantiaEstendida = 'S', 
planoGarantia = '24', 
coeficientePlano = 29.17, 
QtdeGarantia = 1, 
ValorGarantia = 29.17, 
certificadoinicio = '217', 
certificadoFIM = '217', 
ge_premioLiquido = 27.17, 
ge_iof = 2, 
ge_dataInicioVigencia = '2015-05-27 00:00:00.000', 
ge_dataFinalVigencia = '2016-05-26 00:00:00.000', 
ge_valorCustoSeguradora = 29.17
where nf = 1204 
and serie = 'NE' 
and tiponota = 'V' 
and lojaorigem = '48'
and referencia = '0618063'
*/