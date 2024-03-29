USE [DMAC]
GO
/****** Object:  StoredProcedure [dbo].[SP_VDA_Conexao_DEMEOSERVER]    Script Date: 23/09/2014 10:26:52 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

/*

exec SP_VDA_Conexao_DEMEOSERVER

*/


ALTER    Procedure [dbo].[SP_VDA_Conexao_DEMEOSERVER]
            --@data CHAR(12)                     
As

Begin
	Declare @SQL Char(4000)
	declare @nota as char(6)
	declare @serie char(100)
	declare @tiponota char(100) 
	declare @dataemi CHAR(12)
	declare @lojas char(100)
	declare @dataAtual char(10)

	declare @numeroNota char(6)
	declare @serieNota char(2)

set @tiponota = '(''V'',''C'',''T'',''E'',''S'')'
set @SERIE = '(''NE'',''CF'')'
set @dataemi = CONVERT (date,getdate(), 111)
set @dataAtual = CONVERT (date,getdate(), 111)
--set @dataemi = @data
--set @dataemi = '2014/09/15'
set @lojas = '(''28'')'

	--select top 10 * from [DEMEOSERVER].[Demeo].[dbo].nfcapa where numeroped = 121
	--select top 10 * from [DEMEOSERVER].[Demeo].[dbo].nfcapa where numeroped = 121
	--select * from nfcapa where dataemi = '2014/07/31' and TIPONOTA in ('C') and lojaorigem  = '28' order by NUMEROPED
	--select * from [DEMEOSERVER].[Demeo].[dbo].nfcapa where dataemi = '2014/07/31' and TIPONOTA in ('C') and lojaorigem  = '28' order by NUMEROPED

	-- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- 
	print 'INSERINDO CAPA'

	IF object_id('nfcapaSicronizacao') IS NOT NULL 
	BEGIN
		drop table nfcapaSicronizacao
	END

	Select @SQL = 
	'select * into nfcapaSicronizacao from nfcapa as dmac where NOT EXISTS 
	(SELECT * FROM [DEMEOSERVER].[Demeo].[dbo].nfcapa as demeo 
	where dmac.nf = demeo.nf 
	and dmac.LOJAORIGEM = demeo.lojaorigem 
	and dmac.serie = demeo.serie
	and dmac.dataemi = demeo.DATAEMI
	and dmac.tiponota = demeo.tiponota) 
	and dmac.dataemi = ''' + @dataemi  + '''
	and dmac.LOJAORIGEM in ' + @lojas + '
	and dmac.TIPONOTA in '+ @tiponota +'
	and dmac.serie in ' + @SERIE + ''
	--print(@sql)
	Execute (@SQL)

	while (select count(nf) from nfcapaSicronizacao) <> 0
	Begin

		select top 1 @numeroNota = nf, @serieNota = serie  from nfcapaSicronizacao order by nf
		--set @serieNota = (select top 1 serie from nfcapaSicronizacao order by nf)

		Select @sql = 'delete [DEMEOSERVER].[Demeo].[dbo].nfcapa 
		where lojaorigem in ' + @lojas + ' 
		and dataemi = ''' + @dataAtual  + ''' 
		and TIPONOTA in '+ @tiponota + ' 
		and serie = ''' + @serieNota + '''
		and nf = ''' + rtrim(@numeroNota) + ''''
		--print @sql
		Execute (@SQL)

		Select @sql = 'delete [DEMEOSERVER].[Demeo].[dbo].nfitens 
		where lojaorigem in ' + @lojas + ' 
		and dataemi = ''' + @dataAtual  + ''' 
		and TIPONOTA in '+ @tiponota + ' 
		and serie = ''' + @serieNota + '''
		and nf = ''' + rtrim(@numeroNota) + ''''
		--print @sql
		Execute (@SQL)

		Select @sql = 'delete [DEMEOSERVER].[Demeo].[dbo].movimentocaixa 
		where mc_data = ''' + @dataAtual  + '''
		and mc_serie = ''' + @serieNota + '''
		and mc_loja in ' + @lojas + '
		and mc_documento = ''' + rtrim(@numeroNota) + ''''
		--print @sql
		Execute (@SQL)

		insert into [DEMEOSERVER].[Demeo].[dbo].nfcapa 
		select * from nfcapaSicronizacao 
		where nf = @numeroNota 
		and serie = @serieNota

		--------------------------------
		-- ACERTOS --
		--------------------------------

		Select @sql = 'update [DEMEOSERVER].[Demeo].[dbo].nfcapa 
		set cgccli = CPFNFP
		where lojaorigem in ' + @lojas + ' 
		and dataemi = ''' + @dataAtual  + ''' 
		and TIPONOTA = ''V'' 
		and serie = ''' + @serieNota + '''
		and nf = ''' + rtrim(@numeroNota) + '''
		and cpfnfp <> '''' 
		and cpfnfp is not null
		and cliente like ''999999'''
		--print @sql
		Execute (@SQL)

		--------------------------------

		delete nfcapaSicronizacao 
		where nf = @numeroNota 
		and serie = @serieNota

	end

	drop table nfcapaSicronizacao

	-- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- 
	print char(13) + char(13) + 'INSERINDO ITENS'

	IF object_id('nfitensSicronizacao') IS NOT NULL 
	BEGIN
		drop table nfitensSicronizacao
	END

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
	print char(13) + char(13) + 'INSERINDO MOVIMENTO CAIXA'

	IF object_id('movimentocaixaSicronizacao') IS NOT NULL 
	BEGIN
		drop table movimentocaixaSicronizacao
	END

	Select @SQL = 
	'select MC_NumeroECF,MC_CodigoOperador,MC_Loja,
	MC_Data,MC_Grupo,MC_Documento,MC_Serie,
	MC_Valor,MC_Banco,MC_Agencia,MC_ContaCorrente,
	MC_NumeroCheque,MC_BomPara,MC_Parcelas,MC_Remessa,
	MC_Sequencia,MC_SituacaoEnvio,MC_ControleAVR,MC_DataBaixaAVR 
	into movimentocaixaSicronizacao 
	from movimentocaixa as dmac where NOT EXISTS 
	(SELECT * FROM [DEMEOSERVER].[Demeo].[dbo].movimentocaixa as demeo
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

	--If @@ERROR <> 0
	   --Begin	
	   	--Rollback Transaction		
	   	--Return
	   --End
	
End


/*

exec SP_VDA_Conexao_DEMEOSERVER

*/