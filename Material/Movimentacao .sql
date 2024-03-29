USE [Demeo]
GO
/****** Object:  StoredProcedure [dbo].[MovimentacaoMercadoria]    Script Date: 16/05/2014 12:04:59 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO



ALTER   Procedure [dbo].[MovimentacaoMercadoria]
			@DataInicial 		Char (10),
                        @DataFinal		Char (10),
                        @Loja			Char (05),
			@Referencia		Char (07)
As 

Declare		@LojaVenda		Char(5),
		@Serie			VarChar(2),
		@SerieAlm		VarChar(2)

Begin

	Create Table #TempMovMerc (
		DataMov		DateTime	NULL,
		Descricao	VarChar (60)    NULL,
                Entrada         Int		NULL,
		Saida		Int		NULL,
		CodigoOperacao	SmallInt	NULL,
		EntradaSaida	Char (1)	NULL,
		Loja		Char(5)		NOT NULL
	)


	If @Loja = 'CONSO'
	  Begin

		Declare curLoja Insensitive Cursor For
		Select	LO_Loja
		From	Loja
		Where	LO_Loja <> ('CONSO')

	  End


	Else
	  Begin

		Declare curLoja Insensitive Cursor For
		Select	@Loja

	  End


	Open curLoja

	Fetch Next From curLoja Into
		@Loja


	While @@Fetch_Status = 0
	  Begin

		If @Loja = 'CD'
		  Begin
			Select	@LojaVenda = 'CD',
				@Serie = '%',
				@SerieAlm = ''
		  End
		Else
		  Begin
			If @Loja = '353'
			  Begin
				Select	@LojaVenda = @Loja,
					@Serie = '%',
					@SerieAlm = 'S2'
			  End

			Else
			  Begin
				Select	@LojaVenda = @Loja,
					@Serie = '%',
					@SerieAlm = ''
			  End
		  End

		Insert Into #TempMovMerc (
			DataMov,
			Descricao,
	                Entrada,
			Saida,
			EntradaSaida,
			CodigoOperacao,
			Loja
		)
		Select 	VC_DataEmissao,
			CF_Descricao,
			Sum(Case CF_EntradaSaida
				When 'E' Then VI_Quantidade
				Else 0
			    End),
			Sum(Case CF_EntradaSaida
				When 'S' Then VI_Quantidade
				Else 0
	    		    End),
			CF_EntradaSaida,
			VC_CodigoOperacao,
			@Loja
	   	From	CapaNFVenda,
			ItemNFVenda,
			CodigoOperacao
		Where 	VC_NotaFiscal = VI_NotaFiscal and
			VC_Serie = VI_Serie and
			VC_LojaOrigem = VI_LojaOrigem and
			VC_CodigoOperacao = CF_CodigoOperacao and
			VC_TipoNota <> 'C' and
			VC_DataEmissao between @DataInicial and @DataFinal and
			VI_LojaOrigem = @LojaVenda and 
			VI_Serie like @Serie and
			VI_Serie <> @SerieAlm and
			VI_Referencia = @Referencia and 
			VC_LojaVenda not in ('CMC','CMCS','CMCE', 'MC85E', 'MC85S', 'MC85')
		Group By VC_DataEmissao,
			CF_EntradaSaida,
			CF_Descricao,
			VC_CodigoOperacao


		Insert Into #TempMovMerc (
			DataMov,
			Descricao,
	                Entrada,
			Saida,
			EntradaSaida,
			CodigoOperacao,
			Loja
		)
		Select 	VC_DataEmissao,
			CF_Descricao,
			Sum(VI_Quantidade),
			0,
			'E',
			522,
			@Loja
	   	From	CapaNFVenda,
			ItemNFVenda,
			CodigoOperacao
		Where 	VC_NotaFiscal = VI_NotaFiscal and
			VC_Serie = VI_Serie and
			VC_LojaOrigem = VI_LojaOrigem and
			VC_TipoNota <> 'C' and 
			VC_CodigoOperacao = CF_CodigoOperacao and
			VC_CodigoOperacao in (122, 522) and
			VC_DataEmissao between @DataInicial and @DataFinal and
			VC_LojaDestino = @Loja and 
			VI_Referencia = @Referencia and
			VC_LojaDestino not in ('CMC','CMCE','CMCS', 'MC85E', 'MC85S', 'MC85')
		Group By VC_DataEmissao,
			CF_Descricao


                 If @Loja='CD'

                    Begin                   

-------------------------------------Romaneios do CD para CMC,CMCS,CMCE
			Insert Into #TempMovMerc (
				DataMov,
				Descricao,
		                Entrada,
				Saida,
				EntradaSaida,
				CodigoOperacao,
				Loja
			)
			Select 	RO_DataSaida,
				'Romaneio Transf.',
				0,
				Sum(RO_QuantidadeEnviada),
				'S',
				0,
				@Loja
		   	From	Romaneio
			Where 	RO_LojaOrigem = 'CD' and
				RO_Tipo = 'T' and
				RO_LojaDestino in ('CMC','CMCS','CMCE', 'MC85E', 'MC85S', 'MC85') and
				RO_NumeroRomaneio > 0 and
				RO_Referencia = @Referencia and
				RO_DataSaida between @DataInicial and @DataFinal
			Group By RO_DataSaida

---------------------------------------------------------------------------------------------

-------------------------------------Romaneios do CMC para CD
			Insert Into #TempMovMerc (
				DataMov,
				Descricao,
		                Entrada,
				Saida,
				EntradaSaida,
				CodigoOperacao,
				Loja
			)
			Select 	RO_DataSaida,
				'Romaneio Transf.',
				Sum(RO_QuantidadeEnviada),
				0,
				'E',
				0,
				@Loja
		   	From	Romaneio
			Where 	RO_LojaDestino = 'CD' and
				RO_Tipo = 'T' and
				RO_LojaOrigem in ('CMC','CMCS','CMCE', 'MC85E', 'MC85S', 'MC85') and
				RO_NumeroRomaneio > 0 and
				RO_Referencia = @Referencia and
				RO_DataSaida between @DataInicial and @DataFinal
			Group By RO_DataSaida

---------------------------------------------------------------------------------------------

			Insert Into #TempMovMerc (
				DataMov,
				Descricao,
		                Entrada,
				Saida,
				EntradaSaida,
				CodigoOperacao,
				Loja
			)
			Select 	RO_DataSaida,
				'Romaneios Abertos',
				0,
				Sum(RO_QuantidadeEnviada),
				'S',
				0,
				@Loja
		   	From	Romaneio
			Where 	RO_LojaOrigem = 'CD' and
				RO_Situacao = 'A' and
				RO_NumeroRomaneio > 0 and
				RO_Referencia = @Referencia and
				RO_DataSaida between @DataInicial and @DataFinal
			Group By RO_DataSaida


			Insert Into #TempMovMerc (
				DataMov,
				Descricao,
		                Entrada,
				Saida,
				EntradaSaida,
				CodigoOperacao,
				Loja
			)
			Select 	RO_DataSolicitacao,
				'Romaneios Abertos',
				0,
				Sum(RO_QuantidadePedida),
				'S',
				0,
				@Loja
		   	From	Romaneio
			Where 	RO_LojaOrigem = 'CD' and
				RO_Situacao = 'A' and
				RO_NumeroRomaneio = 0 and
				RO_Referencia = @Referencia and
				RO_DataSolicitacao between @DataInicial and @DataFinal
			Group By RO_DataSolicitacao

                    End


----------------------------  Trata Movimentação entre CMC, CMCS e CMCE --------------------
/*
		 IF @Loja = 'CMC' or @Loja = 'CMCS' or @Loja = 'CMCE' or @Loja = 'MC85E' or @Loja = 'MC85S' or @Loja = 'MC85'
		   Begin
			Insert Into #TempMovMerc (
				DataMov,
				Descricao,
		                Entrada,
				Saida,
				EntradaSaida,
				CodigoOperacao,
				Loja
			)
			Select 	RO_DataSaida,
				'Romaneio Transf.',
				0,
				Sum(RO_QuantidadeEnviada),
				'S',
				522,
				@Loja
		   	From	Romaneio
			Where 	RO_LojaOrigem = @Loja and
				RO_LojaDestino in ('CMC','CMCS','CMCE','CD', 'MC85E', 'MC85S', 'MC85') and
				RO_Situacao = 'P' and
				RO_Referencia = @Referencia and
				RO_DataSaida between @DataInicial and @DataFinal
			Group By RO_DataSaida


			Insert Into #TempMovMerc (
				DataMov,
				Descricao,
		                Entrada,
				Saida,
				EntradaSaida,
				CodigoOperacao,
				Loja
			)
			Select 	RO_DataSaida,
				'Romaneio Transf.',
				Sum(RO_QuantidadeEnviada),
				0,
				'S',
				522,
				@Loja
		   	From	Romaneio
			Where 	RO_LojaOrigem in ('CMC','CMCS','CMCE','CD', 'MC85E', 'MC85S', 'MC85') and
				RO_LojaDestino = @Loja and
				RO_Situacao = 'P' and
				RO_Referencia = @Referencia and
				RO_DataSaida between @DataInicial and @DataFinal
			Group By RO_DataSaida

	 	 End
*/
------------------ Fim do tratamento da Movimentação entre CMC, CMCS e CMCE ----------------

----------------------------  Trata Movimentação Do CMC  --------------------  

		If @Loja='CMC'
		  Begin
			Insert Into #TempMovMerc (
				DataMov,
				Descricao,
		                Entrada,
				Saida,
				EntradaSaida,
				CodigoOperacao,
				Loja
			)
			Select 	Vc_DataEmissao,
				'Saida Para Conserto',
				0,
				Sum(Vi_Quantidade),
				'S',
--				599,
				VC_CodigoOperacao,
				'CMC'
		   	From	CapaNfvenda,ItemNfVenda
			Where   Vc_NotaFiscal=Vi_NotaFiscal and
                                Vc_Serie=Vi_Serie and
                                Vc_DataEmissao=Vi_DataEmissao and
                         	Vc_LojaOrigem = 'CMC' and
				Vc_TipoNota <> 'C' and
				Vc_DataEmissao between @DataInicial and @DataFinal and
                                Vi_Referencia = @Referencia 
	             Group By   Vc_DataEmissao, VC_CodigoOperacao


			Insert Into #TempMovMerc (
				DataMov,
				Descricao,
		                Entrada,
				Saida,
				EntradaSaida,
				CodigoOperacao,
				Loja
			)
			Select 	VC_DataEmissao,
				CF_Descricao,
				Sum(VI_Quantidade),
				0,
				'E',
				522,
				@Loja
		   	From	CapaNFVenda,
				ItemNFVenda,
				CodigoOperacao
			Where 	VC_NotaFiscal = VI_NotaFiscal and
				VC_Serie = VI_Serie and
				VC_LojaOrigem = VI_LojaOrigem and
				VC_TipoNota <> 'C' and
				VC_CodigoOperacao = CF_CodigoOperacao and
				VC_CodigoOperacao in (122, 522) and
				VC_DataEmissao between @DataInicial and @DataFinal and
				VC_LojaDestino = @Loja and 
				VI_Referencia = @Referencia
		       Group By VC_DataEmissao,
				CF_Descricao

------------------------------- Romaneios para o CMC ----------------------------------
  		       Insert Into #TempMovMerc (
				DataMov,
				Descricao,
		                Entrada,
				Saida,
				EntradaSaida,
				CodigoOperacao,
				Loja
			)
			Select 	RO_DataSaida,
				'Romaneio Transf.',
				Sum(RO_QuantidadeEnviada),
				0,
				'S',
				522,
				@Loja
		   	From	Romaneio
			Where 	RO_LojaOrigem in ('CMC','CMCS','CMCE','CD', 'MC85E', 'MC85S', 'MC85') and
				RO_LojaDestino = @Loja and
				RO_Situacao = 'A' and
				RO_Referencia = @Referencia and
				RO_DataSaida between @DataInicial and @DataFinal
			Group By RO_DataSaida
----------------------------------------------------------------------------------------

		  End

		If @Loja='CMCS'
		  Begin
			Insert Into #TempMovMerc (
				DataMov,
				Descricao,
		                Entrada,
				Saida,
				EntradaSaida,
				CodigoOperacao,
				Loja
			)
			Select 	Vc_DataEmissao,
				'Mercadoria em Conserto',
				Sum(Vi_Quantidade),
				0,
				'E',
				599,
				'CMCS'
		   	From	CapaNfvenda,ItemNfVenda
			Where   Vc_NotaFiscal=Vi_NotaFiscal and
                                Vc_Serie=Vi_Serie and
                                Vc_DataEmissao=Vi_DataEmissao and
                         	Vc_LojaOrigem = 'CMC' and
				Vc_TipoNota <> 'C' and
				Vc_DataEmissao between @DataInicial and @DataFinal and
                                Vi_Referencia = @Referencia 
	             Group By   Vc_DataEmissao


                         Insert Into #TempMovMerc (
   		                     DataMov,
			             Descricao,
		                     Entrada,
				     Saida,
				     EntradaSaida,
				     CodigoOperacao,
				     Loja
			)
			Select 	CC_DataEntrada,
                		'Retorno de Conserto',
				0, 
                                Sum(Ci_Quantidade), 
				'S', 
				199,
				'CMCS'
		   	From	CapaNfCompra,ItemNfCompra
			Where   Cc_NotaFiscal=Ci_NotaFiscal and
                                Cc_Serie=Ci_Serie and
                             	Cc_Loja IN ('CMC','CMCS','CMCE') and
				CC_Situacao <> 'C' and
				Cc_DataEntrada between @DataInicial and @DataFinal and
                                Ci_Referencia = @Referencia 
	             Group By   Cc_DataEntrada

		  End


		If @Loja='CMCE'
		  Begin
                         Insert Into #TempMovMerc (
   		                     DataMov,
			             Descricao,
		                     Entrada,
				     Saida,
				     EntradaSaida,
				     CodigoOperacao,
				     Loja
			)
			Select 	CC_DataEntrada,
                		'Retorno de Conserto',
				Sum(Ci_Quantidade), 
                                0, 
				'E', 
				199,
				'CMCE'
		   	From	CapaNfCompra,ItemNfCompra
			Where   Cc_NotaFiscal=Ci_NotaFiscal and
                                Cc_Serie=Ci_Serie and
                              	Cc_Loja IN ('CMC','CMCS','CMCE') and
				CC_Situacao <> 'C' and
				Cc_DataEntrada between @DataInicial and @DataFinal and
                                Ci_Referencia = @Referencia 
	             Group By   Cc_DataEntrada

			Insert Into #TempMovMerc (
				DataMov,
				Descricao,
		                Entrada,
				Saida,
				EntradaSaida,
				CodigoOperacao,
				Loja
			)
			Select 	Vc_DataEmissao,
				'Transferencia',
				0,
				Sum(Vi_Quantidade),
				'S',
				522,
				'CMCE'
		   	From	CapaNfvenda,ItemNfVenda
			Where   Vc_NotaFiscal=Vi_NotaFiscal and
                                Vc_Serie=Vi_Serie and
                                Vc_DataEmissao=Vi_DataEmissao and
                         	Vc_LojaOrigem IN ('CMC','CMCS','CMCE')  and
				Vc_TipoNota = 'T' and
				Vc_DataEmissao between @DataInicial and @DataFinal and
                                Vi_Referencia = @Referencia 
	             Group By   Vc_DataEmissao

		  End


----------------------------  Trata Movimentação Do MC85  --------------------

		If @Loja='MC85'
		  Begin
			Insert Into #TempMovMerc (
				DataMov,
				Descricao,
		                Entrada,
				Saida,
				EntradaSaida,
				CodigoOperacao,
				Loja
			)
			Select 	Vc_DataEmissao,
				'Saida Para Conserto',
                        --        case when Vc_CodigoOperacao = 599 then 'Saida Para Conserto' else 'Outra Saida Não Esp.' end,
				0,
				Sum(Vi_Quantidade),
				'S',
				599,
				'MC85'
		   	From	CapaNfvenda,ItemNfVenda
			Where   Vc_NotaFiscal=Vi_NotaFiscal and
                                Vc_Serie=Vi_Serie and
                                Vc_DataEmissao=Vi_DataEmissao and
--				Vc_CodigoOperacao = 599 and
				Vc_CodigoOperacao in(599,5949) and
                         	Vc_LojaOrigem = 'MC85' and
				Vc_TipoNota <> 'C' and
				Vc_DataEmissao between @DataInicial and @DataFinal and
                                Vi_Referencia = @Referencia 
	             Group By   Vc_DataEmissao

---------------------------------------------------------------------------------------------------------
			Insert Into #TempMovMerc (
				DataMov,
				Descricao,
		                Entrada,
				Saida,
				EntradaSaida,
				CodigoOperacao,
				Loja
			)
			Select 	Vc_DataEmissao,
				'Saida Para Conserto',
                      --          case when Vc_CodigoOperacao = 699 then 'Saida Para Conserto' else 'Outra Saida Não Esp.' end,
				0,
				Sum(Vi_Quantidade),
				'S',
				699,
				'MC85'
		   	From	CapaNfvenda,ItemNfVenda
			Where   Vc_NotaFiscal=Vi_NotaFiscal and
                                Vc_Serie=Vi_Serie and
                                Vc_DataEmissao=Vi_DataEmissao and
	--			Vc_CodigoOperacao = 699 and
                                Vc_CodigoOperacao in(699,6949) and
                         	Vc_LojaOrigem = 'MC85' and
				Vc_TipoNota <> 'C' and
				Vc_DataEmissao between @DataInicial and @DataFinal and
                                Vi_Referencia = @Referencia 
	             Group By   Vc_DataEmissao

---------------------------------------------------------------------------------------------------------

			Insert Into #TempMovMerc (
				DataMov,
				Descricao,
		                Entrada,
				Saida,
				EntradaSaida,
				CodigoOperacao,
				Loja
			)
			Select 	VC_DataEmissao,
				CF_Descricao,
				Sum(VI_Quantidade),
				0,
				'E',
				522,
				@Loja
		   	From	CapaNFVenda,
				ItemNFVenda,
				CodigoOperacao
			Where 	VC_NotaFiscal = VI_NotaFiscal and
				VC_Serie = VI_Serie and
				VC_LojaOrigem = VI_LojaOrigem and
				VC_TipoNota <> 'C' and
				VC_CodigoOperacao = CF_CodigoOperacao and
				VC_CodigoOperacao in (122, 522) and
				VC_DataEmissao between @DataInicial and @DataFinal and
				VC_LojaDestino = @Loja and 
				VI_Referencia = @Referencia
		       Group By VC_DataEmissao,
				CF_Descricao


			Insert Into #TempMovMerc (
				DataMov,
				Descricao,
		                Entrada,
				Saida,
				EntradaSaida,
				CodigoOperacao,
				Loja
			)
			Select 	Vc_DataEmissao,
				'Transferencia',
				0,
				Sum(Vi_Quantidade),
				'S',
				522,
				'MC85'
		   	From	CapaNfvenda,ItemNfVenda
			Where   Vc_NotaFiscal=Vi_NotaFiscal and
                                Vc_Serie=Vi_Serie and
                                Vc_DataEmissao=Vi_DataEmissao and
                         	Vc_LojaOrigem = 'MC85'  and
				Vc_TipoNota = 'T' and
				Vc_Lojaorigem = Vi_lojaorigem and
				Vc_DataEmissao between @DataInicial and @DataFinal and
                                Vi_Referencia = @Referencia 
	             Group By   Vc_DataEmissao


		  End

		If @Loja='MC85S'
		  Begin
			Insert Into #TempMovMerc (
				DataMov,
				Descricao,
		                Entrada,
				Saida,
				EntradaSaida,
				CodigoOperacao,
				Loja
			)
			Select 	Vc_DataEmissao,
				'Mercadoria em Conserto',
                               -- case when Vc_CodigoOperacao = 599 then 'Mercadoria em Conserto' else 'Outra Saida Não Esp.' end,
				Sum(Vi_Quantidade),
				0,
				'E',
				599,
				'MC85S'
		   	From	CapaNfvenda,ItemNfVenda
			Where   Vc_NotaFiscal=Vi_NotaFiscal and
                                Vc_Serie=Vi_Serie and
                                Vc_DataEmissao=Vi_DataEmissao and
                         	Vc_LojaOrigem = 'MC85' and
--				Vc_CodigoOperacao = 599 and
				Vc_CodigoOperacao in(599,5949) and
				Vc_TipoNota <> 'C' and
				Vc_DataEmissao between @DataInicial and @DataFinal and
                                Vi_Referencia = @Referencia 
	             Group By   Vc_DataEmissao

--------------------------------------------------------------------------------------------------------------------
			Insert Into #TempMovMerc (
				DataMov,
				Descricao,
		                Entrada,
				Saida,
				EntradaSaida,
				CodigoOperacao,
				Loja
			)
			Select 	Vc_DataEmissao,
				'Mercadoria em Conserto',
                               -- case when Vc_CodigoOperacao = 699 then 'Mercadoria em Conserto' else 'Outra Saida Não Esp.' end,
				Sum(Vi_Quantidade),
				0,
				'E',
				699,
				'MC85S'
		   	From	CapaNfvenda,ItemNfVenda
			Where   Vc_NotaFiscal=Vi_NotaFiscal and
                                Vc_Serie=Vi_Serie and
                                Vc_DataEmissao=Vi_DataEmissao and
                         	Vc_LojaOrigem = 'MC85' and
--				Vc_CodigoOperacao = 699 and
                                Vc_CodigoOperacao in(699,6949) and
				Vc_TipoNota <> 'C' and
				Vc_DataEmissao between @DataInicial and @DataFinal and
                                Vi_Referencia = @Referencia 
	             Group By   Vc_DataEmissao

--------------------------------------------------------------------------------------------------------------------

                         Insert Into #TempMovMerc (
   		                     DataMov,
			             Descricao,
		                     Entrada,
				     Saida,
				     EntradaSaida,
				     CodigoOperacao,
				     Loja
			)
			Select 	CC_DataEntrada,
                		'Retorno de Conserto',
				0, 
                                Sum(Ci_Quantidade), 
				'S', 
				199,
				'MC85S'
		   	From	CapaNfCompra,ItemNfCompra
			Where   Cc_NotaFiscal=Ci_NotaFiscal and
                                Cc_Serie=Ci_Serie and
                              	Cc_Loja = 'MC85E' and
				Cc_Fornecedor = Ci_fornecedor and
				CC_Situacao <> 'C' and 
				Cc_DataEntrada between @DataInicial and @DataFinal and
                                Ci_Referencia = @Referencia 
	             Group By   Cc_DataEntrada

		  End


		If @Loja='MC85E'
		  Begin
                         Insert Into #TempMovMerc (
   		                     DataMov,
			             Descricao,
		                     Entrada,
				     Saida,
				     EntradaSaida,
				     CodigoOperacao,
				     Loja
			)
			Select 	CC_DataEntrada,
                		'Retorno de Conserto',
				Sum(Ci_Quantidade), 
                                0, 
				'E', 
				199,
				'MC85E'
		   	From	CapaNfCompra,ItemNfCompra
			Where   Cc_NotaFiscal=Ci_NotaFiscal and
                                Cc_Serie=Ci_Serie and
                              	Cc_Loja = 'MC85E' and
				CC_Situacao <> 'C' and
				CC_fornecedor = Ci_fornecedor and
				Cc_DataEntrada between @DataInicial and @DataFinal and
                                Ci_Referencia = @Referencia 
	             Group By   Cc_DataEntrada


			Insert Into #TempMovMerc (
				DataMov,
				Descricao,
		                Entrada,
				Saida,
				EntradaSaida,
				CodigoOperacao,
				Loja
			)
			Select 	Vc_DataEmissao,
				'Transferencia',
				0,
				Sum(Vi_Quantidade),
				'S',
				522,
				'MC85E'
		   	From	CapaNfvenda,ItemNfVenda
			Where   Vc_NotaFiscal=Vi_NotaFiscal and
                                Vc_Serie=Vi_Serie and
                                Vc_DataEmissao=Vi_DataEmissao and
                         	Vc_LojaOrigem = 'MC85E' and
				Vc_LojaOrigem = Vi_LojaOrigem and
				Vc_TipoNota = 'T' and
				Vc_DataEmissao between @DataInicial and @DataFinal and
                                Vi_Referencia = @Referencia 
	             Group By   Vc_DataEmissao

		  End

---------------------------------------------------------------------------------------------
/*
			Insert Into #TempMovMerc (
				DataMov,
				Descricao,
		                Entrada,
				Saida,
				EntradaSaida,
				CodigoOperacao,
				Loja
			)
			Select 	Vc_DataEmissao,
				'Estoque Em Conserto',
				Sum(Vi_Quantidade),
				0,
				'E',
				199,
				'CMC'
		   	From	CapaNfvenda,ItemNfVenda
			Where   Vc_NotaFiscal=Vi_NotaFiscal and
                                Vc_Serie=Vi_Serie and
                                Vc_DataEmissao=Vi_DataEmissao and
                         	Vc_LojaOrigem = 'CMC' and
				Vc_TipoNota <> 'C' and
				Vc_DataEmissao between @DataInicial and @DataFinal and
                                Vi_Referencia = @Referencia 
	             Group By   Vc_DataEmissao


                         Insert Into #TempMovMerc (
				     DataMov,
			             Descricao,
		                     Entrada,
				     Saida,
				     EntradaSaida,
				     CodigoOperacao,
				     Loja
			)
			Select 	CC_DataEntrada,
                		'Retorno de Conserto',
				0, 
                                Sum(Ci_Quantidade), 
				'S', 
				199,
				'CMC'
		   	From	CapaNfCompra,ItemNfCompra
			Where   Cc_NotaFiscal=Ci_NotaFiscal and
                                Cc_Serie=Ci_Serie and
                              	Cc_Loja = 'CMC' and
				CC_Situacao <> 'C' and
				Cc_DataEntrada between @DataInicial and @DataFinal and
                                Ci_Referencia = @Referencia 
	             Group By   Cc_DataEntrada

		  End

       
              If @Loja='CD'
                 Begin
                         Insert Into #TempMovMerc (
   		                     DataMov,
			             Descricao,
		                     Entrada,
				     Saida,
				     EntradaSaida,
				     CodigoOperacao,
				     Loja
			)
			Select 	CC_DataEntrada,
                		'Retorno de Conserto',
				Sum(Ci_Quantidade), 
                                0, 
				'E', 
				199,
				'CD'
		   	From	CapaNfCompra,ItemNfCompra
			Where   Cc_NotaFiscal=Ci_NotaFiscal and
                                Cc_Serie=Ci_Serie and
                              	Cc_Loja = 'CMC' and
				CC_Situacao <> 'C' and
				Cc_DataEntrada between @DataInicial and @DataFinal and
                                Ci_Referencia = @Referencia 
	             Group By   Cc_DataEntrada

		  End
*/    

----------------------------  Final da Movimentacao do CMC -------------------

            If @Loja <> 'CMC' And @Loja <> 'CMCE' And @Loja <> 'CMCS' And @Loja <> 'MC85' And @Loja <> 'MC85E' And @Loja <> 'MC85S'
               Begin     
		Insert Into #TempMovMerc (
			DataMov,
			Descricao,
	                Entrada,
			Saida,
			EntradaSaida,
			CodigoOperacao,
			Loja
		)
		Select 	CC_DataEntrada,
			(Case CC_NaturezaOperacao When 4 Then CF_Descricao + ' NME' ELSE CF_Descricao End),
			Sum(Case CF_EntradaSaida
				When 'E' Then (Case CC_NaturezaOperacao When 4 Then (Case NO_Estoque When 'N' THEN 0 ELSE CI_Quantidade End) 
				ELSE CI_Quantidade End) 
				Else 0
			    End),
			Sum(Case CF_EntradaSaida
				When 'S' Then CI_Quantidade
				Else 0
	    		    End),
			CF_EntradaSaida,
			CF_CodigoOperacao,
			@Loja
	   	From	CapaNFCompra,
			ItemNFCompra,
			CodigoOperacao,
			NaturezaOperacao
		Where 	CC_NotaFiscal = CI_NotaFiscal and
			CC_Serie = CI_Serie and
			CC_Fornecedor = CI_Fornecedor and
			CC_CodigoOperacao = CF_CodigoOperacaoAux and
			NO_CodigoNatureza = CC_NaturezaOperacao and 
			NO_CodigoOperacao = CF_CodigoOperacaoAux and
			CC_Situacao In ('L', 'T') and
			CC_DataEntrada between @DataInicial and @DataFinal and	  		CC_Loja = @Loja and
			CI_Referencia = @Referencia
		Group By CC_DataEntrada,CC_NaturezaOperacao,
			CF_EntradaSaida,
			CF_Descricao,
			CF_CodigoOperacao
               End

		Insert Into #TempMovMerc (
			DataMov,
			Descricao,
	                Entrada,
			Saida,
			EntradaSaida,
			CodigoOperacao,
			Loja
		)
		Select 	AJ_Data,
			(Case  When (AJ_CodigoMotivo) = 99 then 'Inventario'
                               Else  (Case 
			             When Sign(AJ_Quantidade) >= 0 Then 'Ajuste de Entrada'
			             Else 'Ajuste de Sa¡da'
                               End)
			 End
			),
			Sum(Case 
				When Sign(AJ_Quantidade) >= 0 Then AJ_Quantidade
				Else 0
			    End),
			Sum(Case 
				When Sign(AJ_Quantidade) < 0 Then Abs(AJ_Quantidade)
				Else 0
	    		    End),
                        (Case 
				When Sign(AJ_Quantidade) >= 0 Then 'E'
				Else 'S'
			 End), 
                       	0,
			@Loja
	   	From	Ajuste
		Where 	AJ_Data between @DataInicial and @DataFinal and
			AJ_Loja = @Loja and
			AJ_Referencia = @Referencia and 
	                AJ_Alteracao <> 'L' 
		Group By AJ_Data,
			Sign(AJ_Quantidade),Aj_CodigoMotivo

           
		Insert Into #TempMovMerc (
			DataMov,
			Descricao,
	                Entrada,
			Saida,
			EntradaSaida,
			CodigoOperacao,
			Loja
		)
		Select 	AJ_Data,
			'Invent rio',
			AJ_Quantidade,
			0,
			'E',
			0,
			@Loja
	   	From	Ajuste
		Where 	AJ_Data between @DataInicial and @DataFinal and
			AJ_Loja = @Loja and
			AJ_Referencia = @Referencia and
			AJ_CodigoMotivo = 91
		Group By AJ_Data,
			AJ_Quantidade

		Fetch Next From curLoja Into
			@Loja



	End


	Close curLoja
	Deallocate curLoja

	Select 	DataMov,
		Descricao,
                	Sum(Entrada) as Entrada,
		Sum(Saida) as Saida,
		CodigoOperacao,
		Loja
	From 	#TempMovMerc 
	Group By 	Loja,
		DataMov,
		Descricao,
		CodigoOperacao,
        		Entrada,Saida
	Order By 	Loja,
		DataMov,
		Entrada,Saida


End





/*

  Exec movimentacaomercadoria '2012/06/01','2012/06/11','cd','6870004'
  Select * from codigooperacao where cf_codigooperacao='678'


*/



