/* CANCELAMENTO DE GARANTIA ESTENDIDA */
/* CANCELAR GARANTIA ESTENDIDA */

use demeo
declare @nota as char(6)
declare @serie char(2)
declare @dataemi CHAR(10)
declare @lojaorigem char(5)
declare @seqCancelamento int
declare @referencia char(7)
set @nota = '38040'
set @serie  = 'S2'
set @dataemi = '2014/06/20'
set @lojaorigem = '28'
set @referencia = '1583851'

set @seqCancelamento = (select max(ge_seqCancelamento) + 1 from nfitens where garantiaEstendida = 'C')
--set @seqCancelamento = 8
--print @seqCancelamento

update nfitens 
set garantiaEstendida = 'C', 
ge_seqCancelamento = @seqCancelamento,
ge_dataCancelamento = DATAEMI
where garantiaEstendida = 'S'
and @nota = nf
and @serie  = serie
and @dataemi = dataemi
and @lojaorigem = lojaorigem
and @referencia = referencia

update nfcapa 
set garantiaEstendida = 'C' 
where garantiaEstendida = 'S'
and @nota = nf
and @serie  = serie
and @dataemi = dataemi
and @lojaorigem = lojaorigem

update [serverwin28].[loja028].[dbo].nfitens
set garantiaEstendida = 'C'
where garantiaEstendida = 'S'
and @nota = nf
and @serie  = serie
and @dataemi = dataemi
and @lojaorigem = lojaorigem
and @referencia = referencia

update [serverwin28].[loja028].[dbo].NFCAPA
set garantiaEstendida = 'C' 
where garantiaEstendida = 'S'
and @nota = nf
and @serie  = serie
and @dataemi = dataemi
and @lojaorigem = lojaorigem

SELECT nf, serie, dataemi,lojaorigem, garantiaEstendida, ge_seqCancelamento, ge_dataCancelamento, * FROM nfitens where @nota = nf and @dataemi = dataemi and @lojaorigem = lojaorigem --and @referencia = referencia -- and @serie  = serie
SELECT nf, serie, dataemi,lojaorigem, garantiaEstendida, * FROM [serverwin28].[loja028].[dbo].nfitens where @nota = nf and @dataemi = dataemi and @lojaorigem = lojaorigem --and @referencia = referencia -- and @serie  = serie
SELECT nf, serie, dataemi,lojaorigem, garantiaEstendida, * FROM nfcapa where @nota = nf and @dataemi = dataemi and @lojaorigem = lojaorigem -- and @serie  = serie
SELECT nf, serie, dataemi,lojaorigem, garantiaEstendida, * FROM [serverwin28].[loja028].[dbo].nfcapa where @nota = nf and @dataemi = dataemi and @lojaorigem = lojaorigem -- and @serie  = serie

/* CONSULTA



*/

--sp_help nfitens
--sp_help nfCAPA