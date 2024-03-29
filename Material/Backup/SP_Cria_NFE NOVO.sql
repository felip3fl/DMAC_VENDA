USE [Demeo]
GO
/****** Object:  StoredProcedure [dbo].[SP_Cria_NFe]    Script Date: 06/05/2014 15:25:55 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO



-- exec SP_Cria_NFe '497',89,'NE'



ALTER                                            PROCEDURE [dbo].[SP_Cria_NFe]

	@Loja		Char(5),
	@NF		Numeric,
	@Serie		Char(2)

AS

	DECLARE		@SQL        	char(4000),
			@CondPagto	Char(2),
			@CondPagtoNF	Char(2),
			@Parcelas       Char(2),
			@NroNF_NFe	Char(10),
			@Referencia	Char(7),
			@UFCliente	Char(2),
                        @NomeServidor   char(40),
                        @Cliente        numeric,
			@Pessoa         char(1),
			@TipoEmissao    Char(1),
			@QtdeVolume     numeric,
			@TotalFrete     numeric,
			@PercFrete	float,
			@DiferencaFrete float,
			@Item		numeric,
			@Situacao       char(1)
                 
BEGIN

	Select @SQL = 'EXEC SP_Busca_Cliente_NFe' + '''' + LTrim(RTrim(@Loja)) + ''''+','+ LTrim(RTrim(@NF))+','+
		'''' + LTrim(RTrim(@Serie)) + ''''

	exec (@SQL)

	Select @CondPagtoNF = (Select CondPag from NFcapa where LojaOrigem = @Loja And NF = @NF And Serie = @Serie)
	Select @Parcelas = (Select CP_parcelas from CondicaoPagamento Where CP_Codigo = @CondPagtoNF)
	--Update ControleSup set CS_NumeroNFe = (CS_NumeroNFe + 1)
	Select @NroNF_NFe = @NF
	Select @UFCliente = (Select UFCliente from NFCapa where lojaorigem = @Loja and NF = @Nf and Serie = @serie)
	Select @Pessoa = (Select pessoacli from NFCapa where lojaorigem = @Loja and NF = @Nf and Serie = @serie)
	Select @QtdeVolume = (Select sum(qtde) from nfItens where LojaOrigem = @Loja And NF = @NF And Serie = @Serie)
        Select @TotalFrete = (Select fretecobr from NFCapa where lojaorigem = @Loja and NF = @Nf and Serie = @serie)
        Select @PercFrete = (Select ((fretecobr * 100)/ vlrmercadoria) from NFCapa where lojaorigem = @Loja 
		and NF = @Nf and Serie = @serie)
	select @DiferencaFrete = (select ( @TotalFrete - (sum((VlUnit2 * @PercFrete) / 100))) from NFitens
		where lojaorigem = @Loja and NF = @Nf and Serie = @serie)
	Select @Item = (select top 1 Item from nfitens where lojaorigem = @Loja and NF = @Nf and Serie = @serie order by Item)
        update nfcapa set fonecli = replace(fonecli,'-','') where LojaOrigem = @Loja And NF = @NF And Serie = @Serie
        update nfcapa set fonecli = replace(fonecli,' ','') where LojaOrigem = @Loja And NF = @NF And Serie = @Serie
	update nfcapa set fonecli = replace(fonecli,'.','') where LojaOrigem = @Loja And NF = @NF And Serie = @Serie
	update nfcapa set fonecli = replace(fonecli,'(','') where LojaOrigem = @Loja And NF = @NF And Serie = @Serie
	update nfcapa set fonecli = replace(fonecli,')','') where LojaOrigem = @Loja And NF = @NF And Serie = @Serie
        update nfcapa set cepcli = replicate('0',8 - len(cepcli)) + cepcli where LojaOrigem = @Loja And NF = @NF And Serie = @Serie
	--update nfcapa set cepcli = ' ' where LojaOrigem = @Loja And NF = @NF And Serie = @Serie And len(cepcli)<7


	Update nfitens set CSTICMS = 60 from nfitens, produto where referencia = pr_referencia 
		and pr_substituicaoTributaria = 'S' and LojaOrigem = @Loja AND Serie = @Serie AND NF = @NF

	Update nfitens set CSTICMS = 20 from nfitens, produto where referencia = pr_referencia 
		and pr_substituicaoTributaria = 'N' and Pr_codigoreducaoicms > 0 and LojaOrigem = @Loja 
		AND Serie = @Serie AND NF = @NF

	Update nfitens set CSTICMS = 00 from nfitens, produto where referencia = pr_referencia 
		and pr_substituicaoTributaria = 'N' and Pr_codigoreducaoicms = 0 and LojaOrigem = @Loja 
		AND Serie = @Serie AND NF = @NF


	IF @UFCliente = 'SP'
	  BEGIN
	   IF @pessoa = '2' or @Pessoa = '1' or @pessoa = '3'
	      Update nfitens set CFOP = 5102 from nfitens, produto where referencia = pr_referencia 
		     and pr_substituicaoTributaria = 'N' and LojaOrigem = @Loja AND Serie = @Serie AND NF = @NF
	   IF @pessoa = '2' or @Pessoa = '1' or @pessoa = '3'
              Update nfitens set CFOP = 5405 from nfitens, produto where referencia = pr_referencia 
		     and pr_substituicaoTributaria = 'S' and LojaOrigem = @Loja AND Serie = @Serie AND NF = @NF
	END

	IF @UFCliente <> 'SP'
          BEGIN
	   IF @Pessoa = '2' or @Pessoa = '1' or @pessoa = '3' 
	     Update nfitens set CFOP = 6404 from nfitens, produto where referencia = pr_referencia 
		    and pr_substituicaoTributaria = 'S' and LojaOrigem = @Loja AND Serie = @Serie AND NF = @NF  
	   IF @Pessoa = '2'
	     Update nfitens set CFOP = 6108 from nfitens, produto where referencia = pr_referencia 
		    and pr_substituicaoTributaria = 'N' and LojaOrigem = @Loja AND Serie = @Serie AND NF = @NF  
	   IF @Pessoa = '1' or @pessoa = '3' 
	     Update nfitens set CFOP = 6102 from nfitens, produto where referencia = pr_referencia 
		    and pr_substituicaoTributaria = 'N' and LojaOrigem = @Loja AND Serie = @Serie AND NF = @NF
	END


	If @CondPagtoNF = 1
	   Begin
		Select @CondPagto = 0
	   End

	If @CondPagtoNF > 3 
	   Begin
		Select @CondPagto = 1
	   End
	If @CondPagtoNF = 2 or @CondPagtoNF = 3 
	   Begin
		Select @CondPagto = 2
	   End  
	

        If LTrim(RTrim(@Loja)) in ('184','181','28','271','85','316','364','77','353','354','497','535','646','874','943','48')    
            Select @Situacao = 'W'
        else
            Select @Situacao = 'A'
        
    
	
	Select @SQL = ' '
/*
	Select @SQL = 'INSERT INTO NFe_ide (eLoja,eNF,eSerie,cUF,cNF,natOp,indPag,mod,serie,nNF,dEmi,dSaiEnt,hSaiEnt,
			tpNF,cMunFG,tpImp,tpEmis,cDV,tpAmb,finNFe,procEmi,verProc,dhCont,xJust) Select LojaOrigem AS eLoja,nf AS eNF,
			Serie as eSerie,LO_CodigoUF AS cUF,'+''''+ LTrim(RTrim(@NroNF_NFe))+'''' +' As cNF,
			(Case When TipoNota = '+ '''E''' + ' Then ' + '''DEVOLUÇÃO'''+' else '+'''VENDA'''+' end),
			'+ @CondPagto +' As indPag,'+ '''55''' +' AS mod,'+'''1'''+' As serie,
			' + ''''+ LTrim(RTrim(@NroNF_NFe))+'''' +' AS nNF,DataEmi As dEmi,DataEmi As dSaiEnt,
			Hora as hSaiEnt,' + '''1''' + ' As tpNF,LO_CodigoMunicipio As cMunFG,' + '''1''' + ' As tpImp,
			' + '''1''' + ' As tpEmis,'+ ''' ''' +' As cDV,' +'''1'''+ ' As tpAmb,' +'''1''' + ' As finNFe,
			' + '''3''' + ' As procEmi,'+ '''2.0.0''' +' As verProc,getdate() As dhCont,
			' + '''Erro no envio da Nota Fiscal Eletronica devido a problemas com Sefaz''' + ' As xJust 
			FROM NFCapa (NOLOCK), Loja (NOLOCK) 
			WHERE LojaOrigem = '+''''+ LTrim(Rtrim(@Loja)) +''''+' AND Serie = '+''''+ @Serie + '''' +
			' AND NF = '+ LTrim(Rtrim(@NF)) +' AND LojaOrigem = LO_Loja collate sql_latin1_general_cp1_ci_as'


	-- Print @SQL
	 Exec (@SQL)*/

	Select @SQL = 'INSERT INTO NFe_emit(eLoja,eNF,eSerie,CNPJ,xNome,xFant,xLgr,nro,xCpl,xBairro,cMun,xMun,UF,
			CEP,cPais,xPais,fone,IE,IEST,IM,CNAE,CRT) SELECT LojaOrigem as eLoja,NF as eNF,Serie as eSerie,
			substring(LO_CGC,2,len(LO_CGC)) As CNPJ,LO_NomeLoja As xNome,LO_NomeFantasia As xFant,
			Lo_EnderecoNFe As xLgr,Lo_EnderecoNroNFe As nro,LO_ComplementoNFe As xCpl,LO_Bairro As xBairro,
			LO_CodigoMunicipio As cMun,LO_Municipio As xMun,LO_UF As UF,'+'''0'''+'+ Convert(VarChar(7),LO_CEP) As CEP, 
		        '+ '''1058''' +' As cPais, '+'''Brasil'''+' As xPais,LO_Telefone As fone,
			LO_InscricaoEstadual As IE,'+''' '''+' As IEST,'+''' '''+' As IM,'+''' '''+' As CNAE, 
			'+'''3'''+' As CRT
			FROM Loja (NOLOCK), NFCapa (NOLOCK) WHERE LojaOrigem = LO_loja And 
			LojaOrigem = '+''''+ LTrim(Rtrim(@Loja)) +''''+' AND Serie = '+ '''' + @Serie + '''' +
			' AND NF = '+ LTrim(Rtrim(@NF))

                   
	--Print ('teste 1  ' + @SQL)
	Exec (@SQL)

	Select @SQL = 'INSERT INTO NFe_dest (eLoja,eNF,eSerie,CNPJ,CPF,xNome,xLgr,nro,xCpl,xBairro,cMun,xMun,UF,CEP,cPais,
			xPais,fone,IE,ISUF,email)SELECT LojaOrigem as eLoja,NF as eNF,Serie as eSerie,
			(Case When PessoaCli <> '+ '''2''' + ' Then CGCCli else '+''' '''+' end),
			(Case When PessoaCli = '+ '''2''' +' Then CGCCli else '+''' '''+' end),
			NomCli As xNome,NFe_Endereco As xLgr,NFe_NumeroEnd As nro,NFe_Complemento As xCpl,
			BairroCli As xBairro,NFe_CodMunicipio As cMun,MunicipioCli As xMun,UFCliente As UF,
			CEPCli as CEP,
			' + '''1058''' + ' As cPais,'+'''Brasil'''+' AS xPais,FoneCli As fone,
			(Case When PessoaCli <> ' + '''2''' +' Then InscriCli else '+''' '''+' end),
			InscriSufCli As ISUF,Carimbo3 as Email
			FROM NFCapa (NOLOCK),NFe_cliente (nolock) 
			WHERE lojaorigem = NFe_loja AND NF = NFe_NF AND Serie = NFe_Serie AND cliente = NFe_Codigo AND
                        LojaOrigem = '+''''+ LTrim(Rtrim(@Loja)) +''''+
			' AND Serie = '+'''' + @Serie + ''''+' AND NF = '+ LTrim(Rtrim(@NF))
	--Print @SQL
	 Exec (@SQL)

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

			(Case When Item = '+  LTrim(Rtrim(convert(char(20),@Item))) +' Then (((VlUnit2 * '+ LTrim(Rtrim(convert(char(20),@PercFrete))) +') / 100) + '+ LTrim(Rtrim(convert(char(20),@DiferencaFrete))) +') 
			else ((VlUnit2 * '+ LTrim(Rtrim(convert(char(20),@PercFrete))) +') / 100) end),

			'+'''0'''+' as I_vSeg,(vlTotItem-VlUnit2) as I_vDesc,'+'''0'''+' as I_vOutro, 
			'+'''1'''+' I_indTot,'+ '''0''' +' as N_origICMS,CSTICMS as N_CSTICMS,'+ '''2''' +' as N_modBCICMS,
			BaseICMS as N_vBCICMS,PR_codigoReducaoICMS as N_pRedBCICMS,ICMPDV as N_pICMS,
			ValorICMS as N_vICMS,'+'''0'''+' as N_modBCST,ICMSMargem as N_pMVAST,'+'''0'''+' as N_pRedBCST,
			'+'''0'''+' as N_vBCST,'+'''0'''+' as N_pICMSST,'+'''0'''+' as N_vICMSST,
			'+''' '''+' as O_cIEnq,'+''' '''+' as O_CNPJProd,'+''' '''+' as O_cSelo,'+''' '''+' as O_qSelo,
			'+''' '''+' as O_cEnq,'+''' '''+' as O_CSTIPI,'+'''0'''+' as O_vBCIPI,'+ '''0''' +' as O_qUnid,
			'+'''0'''+' as O_vUnid,'+'''0'''+' as O_pIPI,'+'''0'''+' as O_vIPI,'+''' '''+' as O_CSTIPINT,
			'+'''0'''+' as P_vBCII,'+'''0'''+' as P_vDespAdu,'+'''0'''+' as P_vII,
			'+'''0'''+' as P_vIOF,'+'''01'''+' as Q_CSTPIS,

                        (Case When Item = '+  LTrim(Rtrim(convert(char(20),@Item))) +' Then ( VlUnit2 + ((VlUnit2 * 
			'+ LTrim(Rtrim(convert(char(20),@PercFrete))) +') / 100) + '+ LTrim(Rtrim(convert(char(20),@DiferencaFrete))) +') 
			else (vlunit2 + ((VlUnit2 * '+ LTrim(Rtrim(convert(char(20),@PercFrete))) +') / 100)) end),        

                        '+'''1.65'''+' as Q_pPIS,
			'+'''0'''+' as Q_qBCProdPIS,'+'''0'''+' as Q_vAliqProdPIS,

                        (Case When Item = '+  LTrim(Rtrim(convert(char(20),@Item))) +'
                        Then (((vlunit2 + ((vlunit2 * '+ LTrim(Rtrim(convert(char(20),@PercFrete))) +' ) /100) + 
                        '+ LTrim(Rtrim(convert(char(20),@DiferencaFrete))) +' ) * 1.65)/100)
                        else (((vlunit2 + ((vlunit2 * '+ LTrim(Rtrim(convert(char(20),@PercFrete))) +' ) /100)) * 1.65)/100) end),

			'+'''0'''+' as R_vBCPISST,'+'''0'''+' as R_pPISST,'+'''0'''+' as R_qBCProdPISST,
			'+'''0'''+' as R_vAliqProdPISST,'+'''0'''+' as R_vPISST,'+'''01'''+' as S_CSTCOFINS,
			
                        (Case When Item = '+  LTrim(Rtrim(convert(char(20),@Item))) +' Then ( VlUnit2 + ((VlUnit2 * 
                        '+ LTrim(Rtrim(convert(char(20),@PercFrete))) +') / 100) + '+ LTrim(Rtrim(convert(char(20),@DiferencaFrete))) +') 
			else (vlunit2 + ((VlUnit2 * '+ LTrim(Rtrim(convert(char(20),@PercFrete))) +') / 100)) end),  

                        '+'''7.60'''+' as S_pCOFINS,'+'''0'''+' as S_qBCProdCOFINS,
			'+'''0'''+' as S_vAliqProdCOFINS,

                        (Case When Item = '+  LTrim(Rtrim(convert(char(20),@Item))) +'
                        Then (((vlunit2 + ((vlunit2 * '+ LTrim(Rtrim(convert(char(20),@PercFrete))) +' ) /100) + 
                        '+ LTrim(Rtrim(convert(char(20),@DiferencaFrete))) +' ) * 7.60)/100)
                        else (((vlunit2 + ((vlunit2 * '+ LTrim(Rtrim(convert(char(20),@PercFrete))) +' ) /100)) * 7.60)/100) end),
			
                        '+'''0'''+' as T_vBCCOFINSST,'+'''0'''+' as T_pCOFINSST,
			'+'''0'''+' as T_qBCProdCOFINSST,'+'''0'''+' as T_vAliqProdCOFINSST,
			'+'''0'''+' as T_vCOFINSST,'+'''0'''+' as U_vBCISSQN,'+'''0'''+' as U_vAliqISSQN,
			'+'''0'''+' as U_vISSQN,'+''' '''+' as U_cMunFGISSQN,'+''' '''+' as U_cListServ,
			'+''' '''+' as U_cSitTrib,'+''' '''+' as V_infAdProd 
			FROM Produto (NOLOCK), NFItens (NOLOCK) 
			WHERE PR_Referencia = Referencia AND LojaOrigem = '+''''+ LTrim(Rtrim(@Loja)) +''''+ 
			' AND Serie = '+'''' + @Serie + ''''+' AND NF = '+ LTrim(Rtrim(@NF)) +' Order by H_nItem'


 
	--Print @SQL 
	 Exec (@SQL)

	 Select @SQL = 'Insert into NFe_total (eLoja,eNF,eSerie,vBCICMS,vICMS,vBCST,vST,vProd,vFrete,vSeg,vDesc,vII,vIPI,
			vCOFINS,vOutro,vNF,vServ,vBCISSQ,vISS,vPIS,vCOFINSISSQ,vRetPIS,vRetCOFINS,vRetCSLL,vBCIRRF,
			vIRRF,vBCRetPrev,vRetPrev)Select LojaOrigem as eLoja,NF as eNF,Serie as eSerie,
			BaseICMS as vBCICMS,VlrICMS AS vICMS,0 as vBCST,0 as vST,
			Subtotal as vProd,ValFrete as vFrete,'+''' 0''' + ' as vSeg,Desconto as vDesc,
			'+ '''0''' +' as vII,'+ '''0''' +' as vIPI,((Totalnota * 7.60)/100) as vCOFINS,0 as vOutro,
			TotalNota as vNF,'+ '''0''' +' as vServ,'+ '''0''' +' as vBCISSQ,'+ '''0''' +' as vISS,
			((Totalnota * 1.65)/100) as vPIS,'+'''0'''+' as vCOFINSISSQ,'+ '''0''' +' as vRetPIS,
			'+ '''0''' +' as vRetCOFINS,'+ '''0''' +' as vRetCSLL,'+ '''0''' +' as vBCIRRF,
			'+ '''0''' +' as vIRRF,'+ '''0''' +' as vBCRetPrev,'+ '''0''' +' as vRetPrev from NFCapa(Nolock) 
			Where LojaOrigem = '+''''+ LTrim(Rtrim(@Loja)) +''''+' AND Serie = '+'''' + @Serie + ''''+
			' AND NF = '+ LTrim(Rtrim(@NF))
	 --Print @SQL
	 Exec (@SQL)

	Select @SQL =  'insert into NFe_cobr (eLoja,eNF,eSerie,nFat,vOrig,vDesc,vLiq,nDup,dVend,vDup)
			Select dp_loja as eLoja,dp_notaFiscal as eNF,dp_serie as eSerie,'+''' '''+' as nFat,
			'+ '''0''' +' as vOrig,'+ '''0''' +' as vDesc,'+ '''0''' +' as vLiq, 
			RTrim(LTrim(Convert(VarChar(10),dp_loja)))+ Convert(VarChar(10),dp_notafiscal)
			+Convert(VarChar(2),dp_serie)+Convert(VarChar(2),dp_sequencia) as nDup,
			dp_dataVencimento as dVend,dp_valorduplicata as vDup From nfCapa(NOLOCK),duplicata(NOLOCK)
			Where dp_notafiscal = nf And dp_serie = serie And dp_loja = Lojaorigem 
			And dp_loja = '+''''+ LTrim(Rtrim(@Loja)) +''''+' And dp_Serie = '+'''' + @Serie + ''''+' 
			And dp_NotaFiscal = '+ LTrim(Rtrim(@NF))
	--Print @SQL
	 Exec (@SQL)

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
	--Print @SQL
	Exec (@SQL)

	Select @SQL = 'insert into NFe_infAdic (eLoja,eNF,eSerie,infAdFisco,infCpl,xCampoCont,
			xTextoCont,xCampoFisco,xTextoFisco,nProc,indProc) Select LojaOrigem as eLoja,
			NF as eNF,Serie as eSerie,'+''' '''+' as infAdFisco, 
			(Case When (Rtrim(Ltrim(Carimbo1))) is Null Then '+''' '''+' else carimbo1 end) + 
			(Case When (RTrim(LTrim(Carimbo2))) is Null Then '+''' '''+' else carimbo2 end) +
			'+'''   ,PEDIDO:'''+' + RTrim(LTrim(Convert(VarChar(10),numeroped)))+ 
			'+'''   ,VENDEDOR:'''+' + RTrim(LTrim(Convert(VarChar(10),Vendedor)))+'+'''   ,COND PAGTO:'''+' + 
			(Case When (RTrim(LTrim(cp_condicao))) is Null Then '+''' '''+' else cp_condicao end) as infcpl,
			'+'''E-mail'''+' as xCampoCont,LO_EmaiLoja as xTextoCont,
			'+''' '''+' as xCampoFisco,'+''' '''+' as xTextoFisco,
			'+''' '''+' as nProc,'+''' '''+' as indProc from nfCapa(nolock),condicaopagamento(nolock),Loja(nolock)
			where cp_codigo = condpag AND LojaOrigem = LO_Loja AND LojaOrigem = '+''''+ LTrim(Rtrim(@Loja)) +''''+' 
			AND Serie = '+'''' + @Serie + ''''+' AND NF = '+ LTrim(Rtrim(@NF))

          -- Com seguro premiado!!!!!!!
	/*Select @SQL = 'insert into NFe_infAdic (eLoja,eNF,eSerie,infAdFisco,infCpl,xCampoCont,
			xTextoCont,xCampoFisco,xTextoFisco,nProc,indProc) Select LojaOrigem as eLoja,
			NF as eNF,Serie as eSerie,'+''' '''+' as infAdFisco, 
			(Case When (Rtrim(Ltrim(Carimbo1))) is Null Then '+''' '''+' else carimbo1 end) + 
			(Case When (RTrim(LTrim(Carimbo2))) is Null Then '+''' '''+' else carimbo2 end) +
			'+''', PEDIDO:'''+' + RTrim(LTrim(Convert(VarChar(10),numeroped)))+ 
			'+''', VENDEDOR:'''+' + RTrim(LTrim(Convert(VarChar(10),Vendedor)))+'+''', COND PAGTO:'''+' + 
			(Case When (RTrim(LTrim(cp_condicao))) is Null Then '+''' '''+' else cp_condicao end) +
                        (Case When (Rtrim(LTrim(GarantiaEstendida))) = '+'''S'''+' Then 
                        '+''''+', GARANTIA ESTENDIDA: R$ '+''''+'+ RTrim(LTrim(Convert(VarChar(10),totalgarantia))) else '+''' '''+' end)  +
			(Case When seguropremiado > 0 Then 
                        '+''''+', SEGURO PREMIADO: R$ '+''''+'+ RTrim(LTrim(Convert(VarChar(10),seguropremiado))) else '+''' '''+' end)
			as infcpl,
			'+'''E-mail'''+' as xCampoCont,LO_EmaiLoja as xTextoCont,
			'+''' '''+' as xCampoFisco,'+''' '''+' as xTextoFisco,
			'+''' '''+' as nProc,'+''' '''+' as indProc from nfCapa(nolock),condicaopagamento(nolock),Loja(nolock)
			where cp_codigo = condpag AND LojaOrigem = LO_Loja AND LojaOrigem = '+''''+ LTrim(Rtrim(@Loja)) +''''+' 
			AND Serie = '+'''' + @Serie + ''''+' AND NF = '+ LTrim(Rtrim(@NF))*/ 


	 --Print ('teste 2 - ' + @SQL)
	 Exec (@SQL)

        Select @SQL = 'INSERT INTO NFe_ide (eLoja,eNF,eSerie,cUF,cNF,natOp,indPag,mod,serie,nNF,dEmi,dSaiEnt,hSaiEnt,
			tpNF,cMunFG,tpImp,tpEmis,cDV,tpAmb,finNFe,procEmi,verProc,dhCont,xJust,Situacao) Select LojaOrigem AS eLoja,nf AS eNF,
			Serie as eSerie,LO_CodigoUF AS cUF,'+''''+ LTrim(RTrim(@NroNF_NFe))+'''' +' As cNF,
			cn_descricaoOperacao AS natop,
			'+ @CondPagto +' As indPag,'+ '''55''' +' AS mod,'+'''1'''+' As serie,
			' + ''''+ LTrim(RTrim(@NroNF_NFe))+'''' +' AS nNF,DataEmi As dEmi,DataEmi As dSaiEnt,
			Hora as hSaiEnt,' + '''1''' + ' As tpNF,LO_CodigoMunicipio As cMunFG,' + '''1''' + ' As tpImp,
			' + '''1''' + ' As tpEmis,'+ ''' ''' +' As cDV,' +'''1'''+ ' As tpAmb,' +'''1''' + ' As finNFe,
			' + '''3''' + ' As procEmi,'+ '''2.0.0''' +' As verProc,getdate() As dhCont,
			' + '''Erro no envio da Nota Fiscal Eletronica devido a problemas com Sefaz''' + ' As xJust,
                        ' + ''''+ LTrim(RTrim(@Situacao))+'''' +' As Situacao 
			FROM NFCapa (NOLOCK), Loja (NOLOCK), codigooperacaonovo (NOLOCK)
			WHERE LojaOrigem = '+''''+ LTrim(Rtrim(@Loja)) +''''+' AND Serie = '+''''+ @Serie + '''' +
			' AND NF = '+ LTrim(Rtrim(@NF)) +' AND LojaOrigem = LO_Loja collate sql_latin1_general_cp1_ci_as and cn_codigoOperacaoNovo = cfoaux'
			--select * from codigoOperacaoNovo

	 Print ('teste 1 - ' + @SQL)
	 Exec (@SQL)

-- Exec SP_Cria_NFe '181',2034,'NE'
--exec SP_Cria_NFe '181', 2025, 'NE'
--PR_ICMSSaida
--select * from nfe_estrutura
--select * from codigooperacaonovo
END
 


























