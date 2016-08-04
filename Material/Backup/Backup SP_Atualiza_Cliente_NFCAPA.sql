USE [DMAC]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

/*

select top 10 nf, * from [dmac28].[dmac_loja].[dbo].nfcapa where lojaorigem = '28' and dataemi = '2014/07/30' and serie = 'CF' and nf = 399
select top 10 nf, * from nfcapa where lojaorigem = '28' and dataemi = '2014/07/30' and serie = 'CF' and nf = 399

exec SP_Atualiza_Cliente_NFCAPA '28',717,'NE'

*/

--select * from nfcapa where 

alter    Procedure [dbo].[SP_Atualiza_Cliente_NFCAPA]
		@Loja				Char(5),
		@NF					char(6),
		@Serie				Char(2)                    
As

Begin
	Declare @SQL Char(4000)
	Declare @tiponota char(1)
	declare @NomeServidor varchar(50)

--set @tiponota = '(''V'',''C'')'
--set @SERIE = 'CF'
--set @Loja = '271'
--set @nf = '74'

	--select top 10 * from [DEMEOSERVER].[Desenv_Demeo].[dbo].nfcapa where numeroped = 121
	--select top 10 * from [DEMEOSERVER].[Desenv_Demeo].[dbo].nfcapa where numeroped = 121

	Select @NomeServidor = (Select LO_NomeServidor from Loja where LO_Loja=@Loja)

	select @tiponota = (select top 1 tiponota from nfcapa where LOJAORIGEM = @loja and nf = @NF and serie = @Serie)

	-- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- 

	if @tiponota = 'V' 
	BEGIN

		Select @sql = 'update nfcapa 
		set NOMCLI = CE_Razao,
		FONECLI = CE_Telefone,
		CGCCLI = CE_CGC,
		INSCRICLI = CE_InscricaoEstadual,
		ENDCLI = CE_Endereco,
		UFCLIENTE = CE_Estado,
		MUNICIPIOCLI = CE_Municipio,
		BAIRROCLI = CE_Bairro,
		CEPCLI = CE_CEP, 
		codmun = '''',
		CompleResidencia = '''',
		NroResidencia = CE_Numero
		from nfcapa, ' + @NomeServidor + 'fin_cliente where 
		tiponota = ''' + @tiponota + '''
		and SERIE = ''' + @Serie  + '''
		and LOJAORIGEM = ''' + @Loja  + '''
		and nf = ''' + @nf + '''
		and cliente = CE_CodigoCliente'

		--print @sql
		Execute (@SQL)

	end 

	-- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- 

	if @tiponota = 'T'
	BEGIN
	
		Select @sql = 'update nfcapa 
		set NOMCLI = LO_Razao,
		FONECLI = lo_telefone,
		CGCCLI = lo_cgc,
		INSCRICLI = lo_inscricaoEstadual,
		ENDCLI = lo_endereco,
		UFCLIENTE = lo_uf,
		MUNICIPIOCLI = lo_municipio,
		BAIRROCLI = lo_bairro,
		CEPCLI = lo_cep, 
		codmun = lo_codigoMunicipio,
		CompleResidencia = '''',
		NroResidencia = lo_endereconronfe
		from nfcapa,loja where 
		tiponota = ''' + @tiponota + '''
		and SERIE = ''' + @Serie  + '''
		and LOJAORIGEM = ''' + @Loja + '''
		and nf = ''' + @nf + '''
		and cliente = lo_loja
		and lo_gruporegiao <  80'

		--print @sql
		Execute (@sql)

	end 

	-- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- 


	If @@ERROR <> 0
	   Begin	
	   	Rollback Transaction		
	   	Return
	   End
	

End
