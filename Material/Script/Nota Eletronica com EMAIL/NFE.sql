USE [DMAC_Loja]
GO
/****** Object:  StoredProcedure [dbo].[SP_VDA_Cria_NFe]    Script Date: 17/11/2015 09:56:59 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO


/*
exec SP_VDA_Cria_NFe 'CD','3590','NE',''

SELECT nf,* FROM nfcapa where dataemi = '2015/09/21' nf = 3590

select * from fin_cliente  where ce_codigocliente = 192

select * from nfe_controle

DROP TABLE NFE_ESTRUTURA


*/

--exec SP_VDA_Cria_NFe 'CD','3661','NE',''

ALTER PROCEDURE [dbo].[SP_VDA_Cria_NFe]

	@Loja		Char(5),
	@NF		    Numeric,
	@Serie		Char(2),
    @Carimbo    varchar(MAX)

AS

	DECLARE	@SQL        	char(4000),
			@CondPagto		Char(2),
			@CondPagtoNF	Char(2),
			@Parcelas       Char(2),
			@NroNF_NFe		Char(10),
			@Referencia		Char(7),
			@UFCliente		Char(2),
			@IDDEST			char(1),
			@finNFe			char(1),
            @CEPCliente     Char(8),
            @NomeServidor   char(40),
            @Cliente        char(6),
			@ClienteT       char(6),
			@IE				char(13),
			@Pessoa         char(1),
			@TipoEmissao    Char(1),
			@QtdeVolume     float,
			@TotalFrete     numeric,
			@PercFrete		float,
			@DiferencaFrete float,
			@Item			numeric,
			@tiponota		char(4),
			@Operacao		char(60),
			@cfop			numeric(18,0),
            @Hora           char(12),
            @Chave          char(8),
            @UFLoja         char(2),
			@EntradaSaida   char(1)


                 
BEGIN

	exec sp_delete_nfe @loja, @nf, @Serie
	delete NFE_NFLojas 
	 where NFL_NroNFE = @nf
	
	Select @tiponota = (Select top 1 tiponota 
	                      from nfcapa 
	                     where LojaOrigem = @Loja 
	                       And NF = @NF 
	                       And Serie = @Serie)

	
	-- -- ACERTOS NFCAPA -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- 
	
	update NfItens set 
		   VALORICMS = round(((BASEICMS * ICMSAplicado) / 100),2) 
	 where nf = @nf 
	   and serie = @Serie
	   and @tiponota <> 'S'
	
	update NfCapa set 
		   vlrICMS = round((select SUM(VALORICMS) as total 
						      from NfItens 
						     where nf = @nf 
						       and serie = @Serie),2) 
	 where nf = @nf 
	   and serie = @Serie
	   and @tiponota <> 'S'       

	-- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- --


	print 'OK 1'
	Select @CondPagtoNF = (Select TOP 1 CondPag 
	                         from NFcapa 
	                        where LojaOrigem = @Loja 
	                          And NF = @NF 
	                          And Serie = @Serie)

	Select @Parcelas = (Select TOP 1 CP_parcelas 
	                      from CondicaoPagamento 
	                     Where CP_Codigo = @CondPagtoNF)
	                     
	SELECT @cfop = (Select TOP 1 CODOPER 
	                        from NFcapa 
	                       where LojaOrigem = @Loja 
	                         And NF = @NF 
	                         And Serie = @Serie)
	
	--Update ControleSup set CS_NumeroNFe = (CS_NumeroNFe + 1)
	
	Select @NroNF_NFe = @NF
	
	print 'OK 2'
	Select @UFCliente = (select ce_Estado 
	                       from NFCapa,FIN_cliente 
	                      where ce_codigocliente = cliente 
	                        and lojaorigem = @Loja 
	                        and NF = @Nf 
	                        and Serie = @serie)

	Select @Pessoa = (Select CE_TipoPessoa 
	                    from NFCapa,FIN_cliente 
	                   where ce_codigocliente = cliente 
	                     and lojaorigem = @Loja 
	                     and NF = @Nf 
	                     and Serie = @serie)
               

    Select @CEPCliente = (Select replicate('0',8 - len(CE_Cep)) + CE_Cep 
                            from NFCapa,FIN_cliente 
                           where ce_codigocliente = cliente 
                             and lojaorigem = @Loja 
                             and NF = @Nf 
                             and Serie = @serie)

	print 'OK 3'
    Select @QtdeVolume = (Select sum(qtde) 
                            from nfItens 
                           where LojaOrigem = @Loja 
                             And NF = @NF 
                             And Serie = @Serie)

	--select @EntradaSaida = (Select top 1 substring(codoper,1,1) from nfcapa where LojaOrigem = @Loja And NF = @NF And Serie = @Serie)
	
	print 'OK 3-1'	
	Select @Cliente = (Select top 1 cliente 
	                     from nfcapa 
	                    where LojaOrigem = @Loja 
	                      And NF = @NF 
	                      And Serie = @Serie 
	                      and tiponota <> 'T')
	
	print 'OK 3-2'
	Select @ClienteT = (Select top 1 lojat 
	                      from nfcapa 
	                     where LojaOrigem = @Loja 
	                       And NF = @NF 
	                       And Serie = @Serie 
	                       and TIPONOTA = 'T')
	
	print 'OK 3-3'	
    Select @TotalFrete = (Select fretecobr 
                            from NFCapa 
                           where lojaorigem = @Loja 
                             and NF = @Nf 
                             and Serie = @serie)
    
    print 'OK 3-4'    
    Select @PercFrete = (Select ((fretecobr * 100)/ vlrmercadoria) 
                           from NFCapa 
                          where lojaorigem = @Loja 
                            and NF = @Nf 
                            and Serie = @serie)
	print 'OK 3-5'	
	select @DiferencaFrete = (select ( @TotalFrete - (sum(((vltotitem - desconto) * @PercFrete) / 100))) 
	                            from NFitens
		                       where lojaorigem = @Loja 
		                         and NF = @Nf 
		                         and Serie = @serie)
	print 'OK 3-6'	                         
	Select @Item = (select top 1 Item 
	                  from nfitens 
	                 where lojaorigem = @Loja 
	                   and NF = @Nf 
	                   and Serie = @serie 
	                 order by Item)
    
    print 'OK 3-7'
    Select @UFLoja = (select distinct substring(convert(nvarchar(9),lo_codigoMunicipio),1,2)
                        from Loja,nfcapa 
                       where lojaorigem = @Loja 
                         and NF = @Nf 
                         and Serie = @serie 
                         and lojaorigem = lo_loja)
                         
    Select @Hora = CONVERT(varchar,GETDATE(),114)
    Select @Hora = replace(@Hora,':','')
    Select @Chave = substring(@hora,5,2) + substring(@hora,3,2) + substring(@hora,1,2) + substring(@hora,8,2)
      

-- SELECT Name + REPLICATE('*', 20 - LEN(Name)) FROM Employee
--	update nfcapa set fonecli = replace(fonecli,'-','') where LojaOrigem = @Loja And NF = @NF And Serie = @Serie
--      update nfcapa set fonecli = replace(fonecli,' ','') where LojaOrigem = @Loja And NF = @NF And Serie = @Serie
--	update nfcapa set fonecli = replace(fonecli,'.','') where LojaOrigem = @Loja And NF = @NF And Serie = @Serie
--	update nfcapa set fonecli = replace(fonecli,'(','') where LojaOrigem = @Loja And NF = @NF And Serie = @Serie
--	update nfcapa set fonecli = replace(fonecli,')','') where LojaOrigem = @Loja And NF = @NF And Serie = @Serie
--      update nfcapa set cepcli = ' ' where LojaOrigem = @Loja And NF = @NF And Serie = @Serie And len(cepcli)<7
	print 'OK 4'
	
	Update nfitens set 
	       CSTICMS = 60 
	  from nfitens, produtoloja 
	 where referencia = pr_referencia 
	   and pr_substituicaoTributaria = 'S' 
	   and LojaOrigem = @Loja 
	   AND Serie = @Serie 
	   AND NF = @NF 
	   and @tiponota <> 'S'
	
	print ('Update nfitens set CSTICMS = 60')

	print 'OK 5'
	
	Update nfitens set 
	       CSTICMS = 20 
	  from nfitens, produtoloja 
	 where referencia = pr_referencia 
	   and pr_substituicaoTributaria = 'N' 
	   and Pr_codigoreducaoicms > 0 
	   and LojaOrigem = @Loja 
	   AND Serie = @Serie 
	   AND NF = @NF
	   and @tiponota <> 'S'
	print ('Update nfitens set CSTICMS = 20')

	print 'OK 6'
	
	Update nfitens set 
	       CSTICMS = 00 
	  from nfitens, produtoloja 
	 where referencia = pr_referencia 
	   and pr_substituicaoTributaria = 'N' 
	   and Pr_codigoreducaoicms = 0 
	   and LojaOrigem = @Loja 
	   AND Serie = @Serie 
	   AND NF = @NF
	   and @tiponota <> 'S'
	   
	select @IDDEST = '1'

	if @Tiponota NOT IN ('E') 
		BEGIN

	IF @UFCliente = 'SP'
	   BEGIN
			IF @pessoa = 'F' or @pessoa = 'U' or @Pessoa = 'J' or @pessoa = 'O' and @Tiponota NOT IN ('S','E') 
				Begin
					
					Update nfitens set 
					       CFOP = 5102 
					  from nfitens, produtoloja 
					 where referencia = pr_referencia 
			           and pr_substituicaoTributaria = 'N' 
			           and LojaOrigem = @Loja 
			           AND Serie = @Serie 
			           AND NF = @NF
			           and @tiponota <> 'S'
			           print ('Update nfitens set CFOP = 5102')
			           
				end
			IF @pessoa = 'F' or @pessoa = 'U' or @pessoa = 'J' or @pessoa = 'O' and @Tiponota NOT IN ('S','E') 
				Begin
					
					Update nfitens set 
						   CFOP = 5405 
					  from nfitens, produtoloja 
					 where referencia = pr_referencia 
					   and pr_substituicaoTributaria = 'S' 
					   and LojaOrigem = @Loja 
					   AND Serie = @Serie 
					   AND NF = @NF
					   and @tiponota <> 'S'
					print ('Update nfitens set CFOP = 5405')
					
				end
		  --print @tiponot
		END

	IF @UFCliente <> 'SP'
		BEGIN
			set @IDDEST = '2'
			IF @pessoa = 'F' or @pessoa = 'U' or @Pessoa = 'J' or @pessoa = 'O' and @Tiponota NOT IN ('S','E') 
				Begin
					Update nfitens set 
					       CFOP = 6404 
					  from nfitens, produtoloja 
					 where referencia = pr_referencia 
					   and pr_substituicaoTributaria = 'S' 
					   and LojaOrigem = @Loja 
					   AND Serie = @Serie 
					   AND NF = @NF  
					   and @tiponota <> 'S'
					print ('Update nfitens set CFOP = 6404')
				end 
				
			IF @pessoa = 'F' or @pessoa = 'U' and @Tiponota NOT IN ('S','E') 
				Begin
					Update nfitens set 
					       CFOP = 6108 
					  from nfitens, produtoloja 
					 where referencia = pr_referencia 
					   and pr_substituicaoTributaria = 'N' 
					   and LojaOrigem = @Loja 
					   AND Serie = @Serie 
					   AND NF = @NF  
					   and @tiponota <> 'S'
					print ('Update nfitens set CFOP = 6108')
				end
				
			IF @Pessoa = 'J' or @pessoa = 'O' and @Tiponota NOT IN ('S','E') 
				Begin
					Update nfitens set 
					       CFOP = 6102 
					  from nfitens, produtoloja 
					 where referencia = pr_referencia 
					   and pr_substituicaoTributaria = 'N' 
					   and LojaOrigem = @Loja 
					   AND Serie = @Serie 
					   AND NF = @NF
					   and @tiponota <> 'S'
					print ('Update nfitens set CFOP = 6102')
				end
		END

	END

	IF rtrim(ltrim(@tiponota)) = 'T'
		Begin
			set @IDDEST = '1'
			Update nfitens set 
			       CFOP = 5409 
			  from nfitens, produtoloja 
			 where referencia = pr_referencia 
			   and pr_substituicaoTributaria = 'S' 
			   and LojaOrigem = @Loja 
			   AND Serie = @Serie 
			   AND NF = @NF
			print ('Update nfitens set CFOP = 5409 transferencia ST')

			Update nfitens set 
			       CFOP = 5152 
			  from nfitens, produtoloja 
			 where referencia = pr_referencia 
			   and pr_substituicaoTributaria = 'N' 
			   and LojaOrigem = @Loja 
			   AND Serie = @Serie 
			   AND NF = @NF
		end
	
			
		--update NFItens set 
		--       CFOP = (select codoper 
		--                 from NFCapa 
		--                where LojaOrigem = @Loja 
		--                  AND Serie = @Serie 
		--                  AND NF = @NF )	
		--  from NFItens 
		-- where LojaOrigem = @Loja 
		--   AND Serie = @Serie 
		--   AND NF = @NF
		
	Update nfcapa set codoper = (select top 1 CFOP from nfitens where LojaOrigem = @Loja AND Serie = @Serie AND NF = @NF) 
	from nfcapa where LojaOrigem = @Loja AND Serie = @Serie AND NF = @NF			
				
	print 'NF'
	print @CondPagtoNF
	
	If @CondPagtoNF = 1
	   Begin
		Select @CondPagto = 0
	   End
	   
	If @CondPagtoNF = 3
	   Begin
		Select @CondPagto = 2
	   End
	   
	If @CondPagtoNF between 4 and 199 
	   Begin
		Select @CondPagto = 1
	   End
	   
	If @CondPagtoNF = 2 or @CondPagtoNF >= 200 
           Begin		
                Select @CondPagto = 2
           End

	select @Operacao = (select top 1 cn_descricaooperacao 
	                      from codigooperacaonovo, NFCapa 
	                     where codoper = cn_codigooperacaonovo 
	                       and LojaOrigem = @Loja 
	                       AND Serie = @Serie 
	                       AND NF = @NF)
	
	if LTrim(Rtrim(@Operacao)) = ''	   
	Begin
		Select @Operacao = 'Venda.'
	End
	  
	/*
	FINNFE
	1 – NF-e normal
	2 – NF-e complementar
	3 – NF-e de ajuste
	4 – Devolução de mercadoria
	*/

	SET @finNFe = '1'
	if  @Tiponota <> 'E' 
	select @entradaSaida =  '1'

	if @cfop = '5202' or @cfop = '5411' or @cfop = '5553' or @cfop = '5909'  or @cfop = '6202' or @cfop = '6411' or @cfop = '6913' 
	begin
		select @entradaSaida = '1'
		select @finNFe = '4'
	end
	
	if @cfop = '1202' or @cfop = '1411' or @cfop = '2202' 
	begin
		select @entradaSaida = '0'
		select @finNFe = '4'
	end	
	
	if @cfop = '5918'  
	begin
		select @entradaSaida = '1'
		select @finNFe = '4'
	end	


	set @IE = (select top 1 ce_inscricaoEstadual 
	             from FIN_Cliente, NFCapa 
	            where cliente = CE_CodigoCliente 
	              and NF = @NF 
	              and serie = @Serie 
	              and LOJAORIGEM = @Loja)

	if @Pessoa = 'F' or @Pessoa = 'U' 
	begin
		set @pessoa = '9'
		set @IE = ''
	end 
	
	if @Pessoa = 'J' or @Pessoa = 'O' 
	begin
		set @pessoa = '1'	
		
		if @IE = 'ISENTO'
		begin
			set @pessoa = '9'	
			set @IE = ''	
		end 
		
	end 

	-- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- 
	
	Select @SQL = 'INSERT INTO NFe_ide (eLoja,eNF,eSerie,cUF,cNF,natOp,indPag,mod,serie,nNF,dEmi,dSaiEnt,hSaiEnt,
	tpNF,cMunFG,tpImp,tpEmis,cDV,tpAmb,finNFe,procEmi,verProc,dhCont,xJust,IDDEST,INDFINAL,INDPRES,refNFe) Select LojaOrigem AS eLoja,nf AS eNF,
	Serie as eSerie,'+''''+ LTrim(RTrim(@UFLoja))+'''' +' AS cUF,'+ LTrim(Rtrim(@NF)) +' As cNF,
	' + '''' + LTrim(RTrim(@Operacao)) + '''' + ' as natop,
	'+ @CondPagto +' As indPag,'+ '''55''' +' AS mod,'+'''1'''+' As serie,
	' + ''''+ LTrim(RTrim(@NroNF_NFe))+'''' +' AS nNF,dataemi As dEmi,DataEmi As dSaiEnt,
	Hora as hSaiEnt,' + '' + @entradaSaida + '' + ' As tpNF,LO_CodigoMunicipio As cMunFG,' + '''1''' + ' As tpImp,
	' + '''1''' + ' As tpEmis,'+ ''' ''' +' As cDV,' +'''2'''+ ' As tpAmb,' + '''' + @finNFe + '''' + ' As finNFe,
	' + '''3''' + ' As procEmi,'+ '''2.0.0''' +' As verProc,getdate() As dhCont,
	' + '''Erro no envio da Nota Fiscal Eletronica devido a problemas com Sefaz''' + ' As xJust, 
	''' + @IDDEST + ''' as IDDEST,''1'' as INDFINAL,''1'' as INDPRES, ChaveNFeDevolucao
	FROM NFCapa (NOLOCK), Loja (NOLOCK) 
	WHERE LojaOrigem = '+''''+ LTrim(Rtrim(@Loja)) +''''+' AND Serie = '+''''+ @Serie + '''' +
	' AND NF = '+ LTrim(Rtrim(@NF)) +' AND LojaOrigem = LO_Loja collate sql_latin1_general_cp1_ci_as'

	Print (@SQL)
	Exec (@SQL)
	
--select * from NFe_ide where eNF = '2049'


	-- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- 

	IF rtrim(ltrim(@tiponota)) = 'T'
		Select @SQL = 'INSERT INTO NFE_controle (eLoja,eNF,eSerie,danfe_IMPRESSORA,danfe_RETORNARESP,
		email_DESTINATARIO,email_ASSUNTO,email_MENSAGEM,email_EMAILEMITENTE,email_NOMEEMITENTE,email_ANEXOPDF,
		email_ANEXOXML,email_ANEXOPROTOCOLO,email_anexoadicional,email_COMPACTADO,email_RETORNARESP) 
		Select LojaOrigem AS eLoja,nf AS eNF,Serie as eSerie,CTS_DanfeImpressora AS danfe_IMPRESSORA,''3'' as danfe_RETORNARESP,
		'''' as email_DESTINATARIO,'''' as email_ASSUNTO,'''' AS email_MENSAGEM,
		''nfesaida@demeo.com.br'' email_EMAILEMITENTE,LO_NomeFantasia AS email_NOMEEMITENTE,''SIM'' as email_ANEXOPDF,
		''SIM'' as email_ANEXOXML,''SIM'' as email_ANEXOPROTOCOLO, ''NAO'' as email_anexoadicional,''NAO'' as email_COMPACTADO, ''1'' email_RETORNARESP
		FROM ControleSistema, NFCapa (NOLOCK), Loja (NOLOCK) 
		WHERE LojaOrigem = LO_loja and LojaOrigem = '+''''+ LTrim(Rtrim(@Loja)) +''''+' AND Serie = '+''''+ @Serie + '''' +
		' AND NF = '+ LTrim(Rtrim(@NF)) +' AND LojaOrigem = LO_Loja collate sql_latin1_general_cp1_ci_as'
	ELSE
		Select @SQL = 'INSERT INTO NFE_controle (eLoja,eNF,eSerie,danfe_IMPRESSORA,danfe_RETORNARESP,
		email_DESTINATARIO,email_ASSUNTO,email_MENSAGEM,email_EMAILEMITENTE,email_NOMEEMITENTE,email_ANEXOPDF,
		email_ANEXOXML,email_ANEXOPROTOCOLO,email_anexoadicional,email_COMPACTADO,email_RETORNARESP) 
		Select LojaOrigem AS eLoja,nf AS eNF,Serie as eSerie,CTS_DanfeImpressora AS danfe_IMPRESSORA,''3'' as danfe_RETORNARESP,
		ce_email as email_DESTINATARIO,''Nota Fiscal Eletrônica ' + LTrim(Rtrim(@NF)) + ' - '' + LO_NomeFantasia as email_ASSUNTO,''Olá '' + ltrim(rtrim(CE_Razao)) + '' 
		Você está recebendo uma cópia da DANFE e o arquivo XML'' AS email_MENSAGEM,
		''nfesaida@demeo.com.br'' email_EMAILEMITENTE,LO_NomeFantasia AS email_NOMEEMITENTE,''SIM'' as email_ANEXOPDF,
		''SIM'' as email_ANEXOXML,''SIM'' as email_ANEXOPROTOCOLO, ''NAO'' as email_anexoadicional,''NAO'' as email_COMPACTADO, ''1'' email_RETORNARESP
		FROM ControleSistema, NFCapa (NOLOCK), fin_Cliente, Loja (NOLOCK) 
		WHERE LojaOrigem = LO_loja and cliente = CE_CodigoCliente and LojaOrigem = '+''''+ LTrim(Rtrim(@Loja)) +''''+' AND Serie = '+''''+ @Serie + '''' +
		' AND NF = '+ LTrim(Rtrim(@NF)) +' AND LojaOrigem = LO_Loja collate sql_latin1_general_cp1_ci_as'

	Print (@SQL)
	Exec (@SQL)

	-- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- 
	
	
	

	Select @SQL = 'INSERT INTO NFe_emit(eLoja,eNF,eSerie,CNPJ,xNome,xFant,xLgr,nro,xCpl,xBairro,cMun,xMun,UF,
	CEP,cPais,xPais,fone,IE,IEST,IM,CNAE,CRT) SELECT LojaOrigem as eLoja,NF as eNF,Serie as eSerie,
	LO_CGC As CNPJ,LO_razao As xNome,LO_NomeFantasia As xFant,
	Lo_Endereco As xLgr,Lo_numero As nro,'''' As xCpl,LO_Bairro As xBairro,
	LO_CodigoMunicipio As cMun,LO_Municipio As xMun,LO_UF As UF,LO_CEP As CEP, 
	'+ '''1058''' +' As cPais, '+'''Brasil'''+' As xPais,LO_DDD + LO_Telefone As fone,
	LO_InscricaoEstadual As IE,'+''' '''+' As IEST,'+''' '''+' As IM,'+''' '''+' As CNAE, 
	'+'''3'''+' As CRT
	FROM Loja (NOLOCK), NFCapa (NOLOCK) WHERE LojaOrigem = LO_loja And 
	LojaOrigem = '+''''+ LTrim(Rtrim(@Loja)) +''''+' AND Serie = '+ '''' + @Serie + '''' +
	' AND NF = '+ LTrim(Rtrim(@NF))

	Print (@SQL)
	Exec (@SQL)

	IF rtrim(ltrim(@tiponota)) = 'T'
		Select @SQL = 'INSERT INTO NFe_dest (eLoja,eNF,eSerie,CNPJ,xNome,xLgr,nro,xCpl,xBairro,cMun,xMun,UF,CEP,cPais,
		xPais,fone,IE,ISUF,email,INDIEDEST) SELECT ' + '''' + LTrim(Rtrim(@Loja)) + '''' + ' as eLoja,' + '''' + LTrim(Rtrim(@NF)) + '''' + ' as eNF, ''NE'' as eSerie,
		(Case When len(lo_CGC) = 14 Then lo_cgc else substring(lo_cgc, 2, 14) end) as CNPJ,
		lo_razao As xNome, lo_endereco As xLgr, lo_numero As nro,'''' As xCpl,
		lo_bairro As xBairro, lo_codigomunicipio As cMun, lo_municipio As xMun, lo_uf As UF,
		lo_cep as CEP,
		''1058'' As cPais,'+'''Brasil'''+' AS xPais,lo_telefone As fone,
		lo_inscricaoEstadual as IE,
		'''' As ISUF,LO_emailoja as Email, ''' + '9' +  ''' as INDIEDEST
		FROM loja (nolock)
		WHERE lo_loja = '+''''+ @ClienteT +''''
	else
	--IF rtrim(ltrim(@tiponota)) = 'E'
	--	Select @SQL = 'INSERT INTO NFe_dest (eLoja,eNF,eSerie,CNPJ,xNome,xLgr,nro,xCpl,xBairro,cMun,xMun,UF,CEP,cPais,
	--	xPais,fone,IE,ISUF,email,INDIEDEST) SELECT ' + '''' + LTrim(Rtrim(@Loja)) + '''' + ' as eLoja,' + '''' + LTrim(Rtrim(@NF)) + '''' + ' as eNF, ''NE'' as eSerie,
	--	(Case When len(lo_CGC) = 14 Then lo_cgc else substring(lo_cgc, 2, 14) end) as CNPJ,
	--	lo_razao As xNome, lo_endereco As xLgr, lo_numero As nro,'''' As xCpl,
	--	lo_bairro As xBairro, lo_codigomunicipio As cMun, lo_municipio As xMun, lo_uf As UF,
	--	lo_cep as CEP,
	--	''1058'' As cPais,'+'''Brasil'''+' AS xPais,lo_telefone As fone,
	--	lo_inscricaoEstadual as IE,
	--	'''' As ISUF,LO_emailoja as Email, ''' + '9' +  ''' as INDIEDEST
	--	FROM loja (nolock)
	--	WHERE lo_loja = '+''''+ @Loja +''''
	--ELSE 
		Select @SQL = 'INSERT INTO NFe_dest (eLoja,eNF,eSerie,CNPJ,CPF,xNome,xLgr,nro,xCpl,xBairro,cMun,xMun,UF,CEP,cPais,
		xPais,fone,IE,ISUF,email, INDIEDEST)SELECT LojaOrigem as eLoja,NF as eNF,Serie as eSerie,
		(Case When len(CE_CGC) = 14 Then CE_CGC else '+''' '''+' end),
		(Case When len(CE_CGC) = 11 Then CE_CGC else '+''' '''+' end),
		CE_Razao As xNome,CE_Endereco As xLgr,CE_numero As nro,CE_Complemento As xCpl,
		CE_bairro As xBairro,CE_CodigoMunicipio As cMun,CE_Municipio As xMun,CE_Estado As UF,
		'+''''+ LTrim(Rtrim(@CEPCliente)) +''''+' as CEP,
		' + '''1058''' + ' As cPais,'+'''Brasil'''+' AS xPais,CE_telefone As fone,
		''' + @IE + ''' as IE,
		CE_InscricaoEstadualSuframa As ISUF,CE_email as Email, ''' + @pessoa +  ''' as INDIEDEST
		FROM NFCapa (NOLOCK),fin_Cliente (nolock)
		WHERE cliente = CE_CodigoCliente AND
		LojaOrigem = '+''''+ LTrim(Rtrim(@Loja)) +''''+
		' AND Serie = '+'''' + @Serie + ''''+' AND NF = '+ LTrim(Rtrim(@NF));		
	


	--Print @SQL-- select ce_cgc,* from fin_cliente where ce_codigocliente = 60046
	Print (@SQL)
	Exec (@SQL)

	-- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- 

	--select * from nfe_estrutura
	Select @SQL = 'INSERT INTO NFe_prod (eLoja,eNF,eSerie,H_nItem,I_cProd,I_cEAN,I_xProd,I_NCM,I_EXTIPI,I_CFOP,
	I_uCom,I_qCom,I_vUnCom,I_vProd,I_cEANTrib,I_uTrib,I_qTrib,I_vUnTrib,I_vFrete,I_vSeg,I_vDesc,I_vOutro,
	I_indTot,N_origICMS,N_CSTICMS,N_modBCICMS,N_vBCICMS,N_pRedBCICMS,N_pICMS,N_vICMS,N_modBCST,N_pMVAST,
	N_pRedBCST,N_vBCST,N_pICMSST,N_vICMSST,O_cIEnq,O_CNPJProd,O_cSelo,O_qSelo,O_cEnq,O_CSTIPI,
	O_vBCIPI,O_qUnid,O_vUnid,O_pIPI,O_vIPI,O_CSTIPINT,P_vBCII,P_vDespAdu,P_vII,P_vIOF,Q_CSTPIS,
	Q_vBCPIS,Q_pPIS,Q_qBCProdPIS,Q_vAliqProdPIS,Q_vPIS,R_vBCPISST,R_pPISST,R_qBCProdPISST,
	R_vAliqProdPISST,R_vPISST,S_CSTCOFINS,S_vBCCOFINS,S_pCOFINS,S_qBCProdCOFINS,S_vAliqProdCOFINS,
	S_vCOFINS,T_vBCCOFINSST,T_pCOFINSST,T_qBCProdCOFINSST,T_vAliqProdCOFINSST,T_vCOFINSST,
	U_vBCISSQN,U_vAliqISSQN,U_vISSQN,U_cMunFGISSQN,U_cListServ,U_cSitTrib,V_infAdProd) 
	SELECT LojaOrigem as eLoja,NF as eNF,Serie as eSerie,ITEM As H_nItem,Referencia As I_cProd,
	'+''' '''+' As I_cEAN,PR_Descricao As I_xProd,PR_ClasseFiscal As I_NCM,'+''' '''+' As I_EXTIPI,
	CFOP As I_CFOP,PR_Unidade As I_uCom,QTDE As I_qCom,VLUnit As I_vUnCom,
	VLTotItem As I_vProd,'+''' '''+' As I_cEANTrib,PR_UNIDADE AS I_uTrib,QTDE aS I_qTrib,
	VLUnit as I_vUnTrib,

	(Case When Item = '+  LTrim(Rtrim(convert(char(20),@Item))) +' Then ((((vltotitem - desconto) * '+ LTrim(Rtrim(convert(char(20),@PercFrete))) +') / 100) + '+ LTrim(Rtrim(convert(char(20),@DiferencaFrete))) +') 
	else (((vltotitem - desconto) * '+ LTrim(Rtrim(convert(char(20),@PercFrete))) +') / 100) end),

	'+'''0'''+' as I_vSeg,desconto as I_vDesc,'+'''0'''+' as I_vOutro, 
	'+'''1'''+' I_indTot,'+ '''0''' +' as N_origICMS,CSTICMS as N_CSTICMS,'+ '''2''' +' as N_modBCICMS,
	baseicms as N_vBCICMS,PR_codigoReducaoICMS as N_pRedBCICMS,ICMSAplicado as N_pICMS,
	ValorICMS as N_vICMS,'+'''0'''+' as N_modBCST,'+'''0'''+' as N_pMVAST,'+'''0'''+' as N_pRedBCST,
	'+'''0'''+' as N_vBCST,'+'''0'''+' as N_pICMSST,'+'''0'''+' as N_vICMSST,
	'+''' '''+' as O_cIEnq,'+''' '''+' as O_CNPJProd,'+''' '''+' as O_cSelo,'+''' '''+' as O_qSelo,
	'+'''999'''+' as O_cEnq,'+'''50'''+' as O_CSTIPI, baseIPI as O_vBCIPI, qtde as O_qUnid,
	vlUnit as O_vUnid, aliqIPI as O_pIPI, vlIpi as O_vIPI,'+''' '''+' as O_CSTIPINT,
	'+'''0'''+' as P_vBCII,'+'''0'''+' as P_vDespAdu,'+'''0'''+' as P_vII,
	'+'''0'''+' as P_vIOF,'+'''01'''+' as Q_CSTPIS,

	(Case When Item = '+  LTrim(Rtrim(convert(char(20),@Item))) +' Then ( (vltotitem - desconto) + (((vltotitem - desconto) * 
	'+ LTrim(Rtrim(convert(char(20),@PercFrete))) +') / 100) + '+ LTrim(Rtrim(convert(char(20),@DiferencaFrete))) +') 
	else ((vltotitem - desconto) + (((vltotitem - desconto) * '+ LTrim(Rtrim(convert(char(20),@PercFrete))) +') / 100)) end), 

	'+'''1.65'''+' as Q_pPIS,'+'''0'''+' as Q_qBCProdPIS,'+'''0'''+' as Q_vAliqProdPIS,

	(Case When Item = '+  LTrim(Rtrim(convert(char(20),@Item))) +'
	Then ((((vltotitem - desconto) + (((vltotitem - desconto) * '+ LTrim(Rtrim(convert(char(20),@PercFrete))) +' ) /100) + 
	'+ LTrim(Rtrim(convert(char(20),@DiferencaFrete))) +' ) * 1.65)/100)
	else ((((vltotitem - desconto) + (((vltotitem - desconto) * '+ LTrim(Rtrim(convert(char(20),@PercFrete))) +' ) /100)) * 1.65)/100) end) as Q_vPIS,

	'+'''0'''+' as R_vBCPISST,'+'''0'''+' as R_pPISST,'+'''0'''+' as R_qBCProdPISST,
	'+'''0'''+' as R_vAliqProdPISST,'+'''0'''+' as R_vPISST,'+'''01'''+' as S_CSTCOFINS,

	(Case When Item = '+  LTrim(Rtrim(convert(char(20),@Item))) +' Then ( (vltotitem - desconto) + (((vltotitem - desconto) * 
	'+ LTrim(Rtrim(convert(char(20),@PercFrete))) +') / 100) + '+ LTrim(Rtrim(convert(char(20),@DiferencaFrete))) +') 
	else ((vltotitem - desconto) + (((vltotitem - desconto) * '+ LTrim(Rtrim(convert(char(20),@PercFrete))) +') / 100)) end),  

	'+'''7.60'''+' as S_pCOFINS,'+'''0'''+' as S_qBCProdCOFINS,'+'''0'''+' as S_vAliqProdCOFINS,

	(Case When Item = '+  LTrim(Rtrim(convert(char(20),@Item))) +'
	Then ((((vltotitem - desconto) + (((vltotitem - desconto) * '+ LTrim(Rtrim(convert(char(20),@PercFrete))) +' ) /100) + 
	'+ LTrim(Rtrim(convert(char(20),@DiferencaFrete))) +' ) * 7.60)/100)
	else ((((vltotitem - desconto) + (((vltotitem - desconto) * '+ LTrim(Rtrim(convert(char(20),@PercFrete))) +' ) /100)) * 7.60)/100) end),

	'+'''0'''+' as T_vBCCOFINSST,'+'''0'''+' as T_pCOFINSST,
	'+'''0'''+' as T_qBCProdCOFINSST,'+'''0'''+' as T_vAliqProdCOFINSST,
	'+'''0'''+' as T_vCOFINSST,'+'''0'''+' as U_vBCISSQN,'+'''0'''+' as U_vAliqISSQN,
	'+'''0'''+' as U_vISSQN,'+''' '''+' as U_cMunFGISSQN,'+''' '''+' as U_cListServ,
	'+''' '''+' as U_cSitTrib,'+''' '''+' as V_infAdProd 
	FROM produtoloja (NOLOCK), NFItens (NOLOCK) 
	WHERE PR_Referencia = Referencia AND LojaOrigem = '+''''+ LTrim(Rtrim(@Loja)) +''''+ 
	' AND Serie = '+'''' + @Serie + ''''+' AND NF = '+ LTrim(Rtrim(@NF)) +' Order by H_nItem'

	--select * from nfe_prod whe

	Print @SQL 
	Exec (@SQL)

	-- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- 

	Select @SQL = 'Insert into NFe_total (eLoja,eNF,eSerie,vBCICMS,vICMS,vBCST,vST,vProd,vFrete,vSeg,vDesc,vII,vIPI,
	vCOFINS,vOutro,vNF,vServ,vBCISSQ,vISS,vPIS,vCOFINSISSQ,vRetPIS,vRetCOFINS,vRetCSLL,vBCIRRF,
	vIRRF,vBCRetPrev,vRetPrev,vVICMSDESON)Select LojaOrigem as eLoja,NF as eNF,Serie as eSerie,

	(Case When baseicms is null Then 0 else baseicms end), 

	VlrICMS AS vICMS,0 as vBCST,0 as vST,
	vlrmercadoria as vProd,Fretecobr as vFrete,'+''' 0''' + ' as vSeg,Desconto as vDesc,
	'+ '''0''' +' as vII,totalipi as vIPI,(((Totalnota-totalipi) * 7.60)/100) as vCOFINS,0 as vOutro,
	TotalNota as vNF,'+ '''0''' +' as vServ,'+ '''0''' +' as vBCISSQ,'+ '''0''' +' as vISS,
	(((Totalnota - totalipi) * 1.65)/100) as vPIS,'+'''0'''+' as vCOFINSISSQ,'+ '''0''' +' as vRetPIS,
	'+ '''0''' +' as vRetCOFINS,'+ '''0''' +' as vRetCSLL,'+ '''0''' +' as vBCIRRF,
	'+ '''0''' +' as vIRRF,'+ '''0''' +' as vBCRetPrev,'+ '''0''' +' as vRetPrev, ''0'' as vVICMSDESON from NFCapa(Nolock) 
	Where LojaOrigem = '+''''+ LTrim(Rtrim(@Loja)) +''''+' AND Serie = '+'''' + @Serie + ''''+
	' AND NF = '+ LTrim(Rtrim(@NF))

	Print @SQL -- baseicms as vBCICMS,
	Exec (@SQL)
	
	--select * from nfe_prod where enf = 3796
	--select * from NFItens where nf = 3796
	
	-- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- 

	Select @SQL = 'Insert into NFe_transp (eLoja,eNF,eSerie,modFrete,CNPJ,CPF,xNome,IE,xEnder,xMun,UF,vServ,vBCRet,pICMSRet,
	vICMSRet,CFOP,cMunFG,placa,UFveic,RNTC,qVol,esq,marca,nVol,pesoL,pesoB,nLacres)
	Select LojaOrigem as eLoja,NF as eNF,Serie as eSerie,TipoFrete as modFrete,'+''' '''+' As CNPJ,
	'+''' '''+' as CPF,'+''' '''+' as xNome,'+''' '''+' as IE,'+''' '''+' as xEnder,
	'+''' '''+' as xMun,'+''' '''+' as UF,'+ '''0'''+' as vServ,'+ '''0''' +' as vBCRet,
	'+ '''0''' +' as pICMSRet,'+ '''0''' +' as vICMSRet,'+''' '''+' as CFOP,'+''' '''+' as cMunFG,
	'+''' '''+' as placa,'+''' '''+' as UFveic,'+''' '''+' as RNTC,
	volume as qVol,'+'''VOLUME(S)'''+' as esq,'+''' '''+' as marca,
	'+ '''0''' +' as nVol,pesolq as pesoL,pesobr as pesoB,
	'+ '''0''' +' as nLacres FROM Loja(NOLOCK), NFCapa (NOLOCK)
	Where lojaOrigem = LO_loja And LojaOrigem = '+''''+ LTrim(Rtrim(@Loja)) +''''+' 
	AND Serie = '+'''' + @Serie + ''''+' AND NF = '+ LTrim(Rtrim(@NF))

	Print @SQL
	Exec (@SQL)

	-- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- 

	declare @descricao varchar(max)
	--declare @sequencia int
	--declare @sequenciaMaxima int

	SET @Carimbo = ''
	Declare Temp_Carimbo insensitive cursor for
			Select rtrim(LTRIM(CNF_Carimbo))
			  from CarimboNotaFiscal 
			 where CNF_Loja = @Loja 
			   and cnf_serie = @Serie 
			   and CNF_NF = @NF
			 order by CNF_TipoCarimbo desc, CNF_Sequencia 
	Open Temp_Carimbo
	Fetch Next From Temp_Carimbo Into @Descricao
	While @@Fetch_Status = 0  
		Begin

		set @Carimbo = @Carimbo + @Descricao + '  -  '
			Fetch Next From Temp_Carimbo Into @Descricao
		end
	close Temp_Carimbo
	Deallocate Temp_Carimbo

	--set @Carimbo = left(@Carimbo,len(@Carimbo)-2)

	Select @SQL = 'insert into NFe_infAdic (eLoja,eNF,eSerie,infAdFisco,infCpl,xCampoCont,
	xTextoCont,xCampoFisco,xTextoFisco,nProc,indProc) Select LojaOrigem as eLoja,
	NF as eNF,Serie as eSerie,'+''' '''+' as infAdFisco,''PEDIDO: '''+' + RTrim(LTrim(Convert(VarChar(10),numeroped)))+ 
	'+''', VENDEDOR: '''+' + RTrim(LTrim(Convert(VarChar(10),Vendedor)))+'+''', COND PAGTO: '''+' + 
	(Case When (RTrim(LTrim(cp_condicao))) is Null Then '+''' '''+' else cp_condicao end) + '+'''  -  ' + @Carimbo + '''' + ''+' as infcpl,
	'+'''E-MAIL'''+' as xCampoCont, Upper(LO_EmaiLoja) as xTextoCont,
	'+''' '''+' as xCampoFisco,'+''' '''+' as xTextoFisco,
	'+''' '''+' as nProc,'+''' '''+' as indProc from nfCapa(nolock),condicaopagamento(nolock),Loja(nolock)
	where cp_codigo = condpag and cp_id = 1 AND LojaOrigem = LO_Loja AND LojaOrigem = '+''''+ LTrim(Rtrim(@Loja)) +''''+' 
	AND Serie = '+'''' + @Serie + ''''+' AND NF = '+ LTrim(Rtrim(@NF))

	Print @SQL
	Exec (@SQL)
	
END


