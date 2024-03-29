USE [DMAC]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

/*

exec [SP_VDA_Conexao_DEMEOSERVER2]

*/

--select * from nfcapa where 

alter    Procedure [dbo].[SP_VDA_Conexao_DEMEOSERVER2]
            --@data CHAR(12)                     
As

Begin
	Declare @SQL Char(4000)
	declare @nota as char(6)
	declare @serie char(100)
	declare @tiponota char(100) 
	declare @dataemi CHAR(12)
	declare @lojas char(100)

set @tiponota = '(''V'',''C'',''T'')'
set @SERIE = '(''NE'',''CF'')'
set @dataemi = CONVERT (date,getdate(), 111)
--set @dataemi = @data
--set @dataemi = '2014/07/29'
set @lojas = '(''28'')'

	--select top 10 * from [DEMEOSERVER].[Desenv_Demeo].[dbo].nfcapa where numeroped = 121
	--select top 10 * from [DEMEOSERVER].[Desenv_Demeo].[dbo].nfcapa where numeroped = 121

	-- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- 

	Select @SQL = 
	'select * into nfcapaSicronizacao from nfcapa as dmac where NOT EXISTS 
	(SELECT * FROM [DEMEOSERVER].[Demeo].[dbo].nfcapa as demeo 
	where dmac.nf = demeo.nf 
	and dmac.LOJAORIGEM = demeo.lojaorigem 
	and dmac.serie = demeo.serie
	and dmac.dataemi = demeo.DATAEMI) 
	and dmac.dataemi = ''' + @dataemi  + '''
	and dmac.LOJAORIGEM in ' + @lojas + '
	and dmac.TIPONOTA in '+ @tiponota +'
	and dmac.serie in ' + @SERIE + ''
	print(@sql)
	Execute (@SQL)

	--delete [DEMEOSERVER].[DESENV_DEMEO].[dbo].nfcapa from nfcapaSicronizacao as dmac where lojaorigem = '28' and dataemi = '2014/07/31' and dmac.nf = nf and dmac.serie = serie
	insert into [DEMEOSERVER].[Demeo].[dbo].nfcapa select * from nfcapaSicronizacao

	drop table nfcapaSicronizacao

	-- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- 

	Select @SQL = 
	'select * into nfitensSicronizacao from nfitens as dmac where NOT EXISTS 
	(SELECT * FROM [DEMEOSERVER].[Demeo].[dbo].nfitens as demeo 
	where dmac.nf = demeo.nf 
	and dmac.LOJAORIGEM = demeo.lojaorigem 
	and dmac.serie = demeo.serie
	and dmac.dataemi = demeo.DATAEMI) 
	and dmac.dataemi = ''' + @dataemi  + '''
	and dmac.LOJAORIGEM in ' + @lojas + '
	and dmac.TIPONOTA in '+ @tiponota +'
	and dmac.serie in ' + @SERIE + ''
	--print(@sql)
	EXecute (@SQL)

	insert into [DEMEOSERVER].[Demeo].[dbo].nfitens select * from nfitensSicronizacao

	drop table nfitensSicronizacao

	-- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- 

	Select @SQL = 
	'select MC_NumeroECF,MC_CodigoOperador,MC_Loja,
	MC_Data,MC_Grupo,MC_Documento,MC_Serie,
	MC_Valor,MC_Banco,MC_Agencia,MC_ContaCorrente,
	MC_NumeroCheque,MC_BomPara,MC_Parcelas,MC_Remessa,
	MC_Sequencia,MC_SituacaoEnvio,MC_ControleAVR,MC_DataBaixaAVR 
	into movimentocaixaSicronizacao 
	from movimentocaixa as dmac where NOT EXISTS 
	(SELECT * FROM [demeoserver].[Demeo].[dbo].movimentocaixa as demeo
	where dmac.MC_Data = demeo.mc_data 
	and dmac.MC_Documento = demeo.MC_Documento
	and dmac.mc_loja = demeo.mc_loja
	and dmac.mc_serie = demeo.mc_serie)
	and dmac.mc_data = ''' + @dataemi  + '''
	and dmac.mc_serie IN ' + @SERIE + '
	and dmac.mc_tiponota in ' + @tiponota + '
	and dmac.mc_loja in ' + @lojas + ''
	--print(@sql)
	Execute (@SQL)

	insert into [DEMEOSERVER].[Demeo].[dbo].movimentocaixa select * from movimentocaixaSicronizacao

	drop table movimentocaixaSicronizacao

	-- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- 

	If @@ERROR <> 0
	   Begin	
	   	Rollback Transaction		
	   	Return
	   End
	

End
