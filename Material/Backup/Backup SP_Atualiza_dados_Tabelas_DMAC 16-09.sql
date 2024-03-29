USE [DMAC]
GO
/****** Object:  StoredProcedure [dbo].[SP_Atualiza_dados_Tabelas_DMAC]    Script Date: 15/10/2014 15:18:12 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO


ALTER  procedure [dbo].[SP_Atualiza_dados_Tabelas_DMAC]


   as 
    Declare @Promocao as char(1),
            @Produto as char(1), 
            --@ListaPreco as char(1),
            @Estoque as char(1),
            @Outras as char(1),
            @ProdutoLoja as char(1),
            --@PromocaoLoja as char(1),
            @Sugestao as char(1),
            @Ajuste as char(1),
			@Transferencia as char(1),
			@sql as varchar(500),

			@nf as char(6),
			@serie as char(2),
			@lojaorigem as char(3),
			@dataemissao as char(10),
			@observacao as varchar(180),
			@item as numeric,
			@referencia as char(7),
			@quantidade as char(3),
			@lojaDestino as char(3)
            
     select @Promocao      = (Select CTA_Promocao from glb_controletarefas)
     select @Produto       = (Select CTA_Produto from glb_controletarefas) 
     --select @ListaPreco    = (Select CTA_ListaPreco from glb_controletarefas) 
     select @Estoque       = (Select CTA_Estoque from glb_controletarefas)
     select @Outras        = (Select CTA_Outras from glb_controletarefas)
     select @ProdutoLoja   = (Select CTA_Produtoloja from glb_controletarefas)
     --select @PromocaoLoja  = (Select CTA_PromocaoLoja from glb_controletarefas)
     select @Sugestao      = (Select CTA_Sugestao from glb_controletarefas) 
     select @Ajuste        = (Select CTA_Ajuste from glb_controletarefas)
	 select @Transferencia = (Select CTA_Transferencia from glb_controletarefas)

	 IF @Transferencia = 'S'
        Begin
			/*
			select * from glb_controletarefas
			update GLB_ControleTarefas set CTA_Transferencia = 'S'
			exec SP_Atualiza_dados_Tabelas_DMAC
			*/

			IF object_id('capanfvendaSicronizacao') IS NOT NULL 
			BEGIN
				drop table capanfvendaSicronizacao
			END

			select * into capanfvendaSicronizacao from capanfvenda as dmac where not exists 
			(select * from [demeoserver].[demeo].[dbo].capanfvenda as demeo where
			demeo.vc_notafiscal = dmac.vc_notafiscal 
			and demeo.vc_serie = dmac.vc_serie 
			and demeo.vc_dataemissao = dmac.vc_dataemissao
			and demeo.vc_lojaorigem = dmac.VC_LojaOrigem
			and demeo.vc_observacao = dmac.vc_observacao
			and demeo.vc_observacao <> '0')
			and dmac.vc_tiponota = 'T'
			and dmac.vc_dataemissao >= DATEADD(mm,-1, getdate())
			and dmac.vc_lojadestino = '28'
			and dmac.VC_Observacao <> '0'

			while (select count(vc_notafiscal) from capanfvendaSicronizacao) <> 0
			Begin

				select top 1 @nf = VC_NotaFiscal, 
				@serie = VC_Serie, 
				@dataemissao = CONVERT (date,VC_DataEmissao, 111), 
				@lojaorigem = VC_LojaOrigem,
				@observacao = VC_Observacao,
				@lojaDestino = VC_LojaDestino
				from capanfvendaSicronizacao order by VC_NotaFiscal

				select @sql = 'update [DEMEOSERVER].[Demeo].[dbo].capanfvenda 
				set vc_observacao = ''' + @observacao + '''
				where vc_notafiscal = ''' + @nf + '''
				and vc_serie = ''' + @serie + '''
				and vc_dataemissao = ''' + @dataemissao + '''
				and vc_lojaorigem = ''' + @lojaorigem + ''''

				--print (@sql)
				execute (@sql)

				select @item = 1

				while (select COUNT(VI_NumeroItem) from itemnfvenda where VI_NotaFiscal = @nf 
				and VI_LojaOrigem = @lojaorigem 
				and VI_DataEmissao = @dataemissao
				and VI_Serie = @serie and VI_NumeroItem = @item) <> 0
				Begin

					select @referencia = VI_Referencia, 
					@quantidade = VI_Quantidade 
					from itemnfvenda where VI_NotaFiscal = @nf 
					and VI_LojaOrigem = @lojaorigem 
					and VI_DataEmissao = @dataemissao
					and VI_Serie = @serie and VI_NumeroItem = @item

					select @sql = 'Update [DEMEOSERVER].[Demeo].[dbo].Estoque 
					set es_estoque = es_estoque + ' + @quantidade + ', 
					es_transito=es_transito - ' + @quantidade + ' 
					where es_referencia = ''' + @referencia + ''' 
					and es_loja = ''' + @lojaDestino + ''''

					--print (@sql)
					execute (@sql)
				
					select @item = @item + 1

				end

				delete capanfvendaSicronizacao 
				where VC_NotaFiscal = @nf
				and VC_Serie = @serie
				and VC_LojaOrigem = @lojaorigem
				and VC_DataEmissao = @dataemissao

			end

			select @sql = ''
			update GLB_ControleTarefas set CTA_Transferencia = 'N'

        End		

     If @Produto = 'S'
        Begin

         DROP Table Produto
         select * into Produto from [DemeoServer].[Demeo].[dbo].Produto
 		 update produto set pr_cst= '060' where pr_substituicaotributaria='S'
         update produto set pr_cst= '020' where pr_codigoreducaoicms > 0 and pr_substituicaotributaria='N'
         update produto set pr_cst= '000' where pr_codigoreducaoicms = 0 and pr_substituicaotributaria='N'
         update produto set pr_IndicePreco= '2' where pr_codigofornecedor = 680
         update produto set pr_IndicePreco= '1' where pr_codigofornecedor <> 680
         update produto set pr_IndicePreco= '3' where pr_classe = 'P' 

		 
		 DROP Table LinhaProduto
         select * into LinhaProduto from [DemeoServer].[Demeo].[dbo].LinhaProduto

		 exec [DEMEOSERVER].[DEMEO].[DBO].SP_Atualiza_Produto_Barras_Loja 'DMAC'
         
		 -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- 

		 exec [DEMEOSERVER].[DEMEO].[DBO].SP_Atualiza_Produto_Barras_Loja '28'

         delete [dmac28].[dmac_loja].[dbo].produtoLoja from [dmac28].[dmac_loja].[dbo].produtoLoja as l
         where not exists (select PR_Referencia from produto as r where l.pr_referencia=r.PR_Referencia)
           
          update [dmac28].[dmac_loja].[dbo].produtoLoja set PR_PrecoVenda1=r.PR_PrecoVenda1
          from [dmac28].[dmac_loja].[dbo].produtoLoja as l,produto as r
          where r.PR_Referencia=l.pr_referencia and l.PR_PrecoVenda1<>r.PR_PrecoVenda1
         
	
		  insert into [dmac28].[dmac_loja].[dbo].produtoLoja
		  (pr_Referencia , pr_CodigoFornecedor , pr_Descricao , pr_Classe , 
		  pr_Bloqueio , pr_LinhaProduto , pr_ClasseFiscal , pr_Unidade , 
		  pr_ICMSSaida , pr_CodigoReducaoICMS , pr_CustoMedio1 , 
		  pr_PrecoVenda1 , pr_PaginaListaPreco , pr_Peso, pr_Comprador , 
		  pr_Situacao , pr_SubstituicaoTributaria , pr_IcmPdv, 
		  pr_HoraManutencao , pr_CodigoProdutoNoFornecedor , 
		  pr_IcmsSaidaIva, pr_IcmsPdvSaidaIva , pr_ICMSEntrada , 
		  pr_IcmPdvEntrada , pr_ST , pr_GarantiaEstendida , 
		  pr_GarantiaFabricante , pr_IndicePreco,PR_precoVendaLiquido1,PR_custoMedioLiquido1) 
		  select pr_Referencia , pr_CodigoFornecedor , pr_Descricao , pr_Classe , pr_Bloqueio , 
		  pr_LinhaProduto , pr_ClasseFiscal , pr_Unidade , pr_ICMSSaida , pr_CodigoReducaoICMS , 
		  pr_CustoMedio1 , pr_PrecoVenda1 , pr_PaginaListaPreco , pr_Peso, pr_Comprador ,
		  pr_Situacao , pr_SubstituicaoTributaria , pr_IcmPdv, pr_HoraManutencao , 
		  pr_CodigoProdutoNoFornecedor , pr_IcmsSaidaIva, pr_IcmsPdvSaidaIva , 
		  pr_ICMSEntrada , pr_IcmPdvEntrada , pr_CST , pr_GarantiaEstendida , 
		  pr_GarantiaFabricante , pr_IndicePreco,PR_precoVendaLiquido1,PR_custoMedioLiquido1
		  from produto as r where  not exists 
		 (select * from [dmac28].[dmac_loja].[dbo].produtoLoja as l 
		  where l.PR_Referencia=r.pr_referencia) 

		 insert into [dmac28].[dmac_loja].[dbo].LinhaProduto select * from LinhaProduto as r where not exists 
		 (select * from [dmac28].[dmac_loja].[dbo].LinhaProduto as l 
		  where l.LPR_Linha =r.LPR_Linha) 
		 	
         -- Update GLB_ControleTarefas set CTA_Produto = 'N'

		 -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- 

		 exec [DEMEOSERVER].[DEMEO].[DBO].SP_Atualiza_Produto_Barras_Loja 'DESENV'

         delete Desenv_DMAC_Loja..produtoLoja from Desenv_DMAC_Loja..produtoLoja as l
         where not exists (select PR_Referencia from produto as r where l.pr_referencia=r.PR_Referencia)
           
          update Desenv_DMAC_Loja..produtoLoja set PR_PrecoVenda1=r.PR_PrecoVenda1
          from Desenv_DMAC_Loja..produtoLoja as l,produto as r
          where r.PR_Referencia=l.pr_referencia and l.PR_PrecoVenda1<>r.PR_PrecoVenda1
         
	
		  insert into Desenv_DMAC_Loja..produtoLoja
		  (pr_Referencia , pr_CodigoFornecedor , pr_Descricao , pr_Classe , 
		  pr_Bloqueio , pr_LinhaProduto , pr_ClasseFiscal , pr_Unidade , 
		  pr_ICMSSaida , pr_CodigoReducaoICMS , pr_CustoMedio1 , 
		  pr_PrecoVenda1 , pr_PaginaListaPreco , pr_Peso, pr_Comprador , 
		  pr_Situacao , pr_SubstituicaoTributaria , pr_IcmPdv, 
		  pr_HoraManutencao , pr_CodigoProdutoNoFornecedor , 
		  pr_IcmsSaidaIva, pr_IcmsPdvSaidaIva , pr_ICMSEntrada , 
		  pr_IcmPdvEntrada , pr_ST , pr_GarantiaEstendida , 
		  pr_GarantiaFabricante , pr_IndicePreco,PR_precoVendaLiquido1,PR_custoMedioLiquido1) 
		  select pr_Referencia , pr_CodigoFornecedor , pr_Descricao , pr_Classe , pr_Bloqueio , 
		  pr_LinhaProduto , pr_ClasseFiscal , pr_Unidade , pr_ICMSSaida , pr_CodigoReducaoICMS , 
		  pr_CustoMedio1 , pr_PrecoVenda1 , pr_PaginaListaPreco , pr_Peso, pr_Comprador ,
		  pr_Situacao , pr_SubstituicaoTributaria , pr_IcmPdv, pr_HoraManutencao , 
		  pr_CodigoProdutoNoFornecedor , pr_IcmsSaidaIva, pr_IcmsPdvSaidaIva , 
		  pr_ICMSEntrada , pr_IcmPdvEntrada , pr_CST , pr_GarantiaEstendida , 
		  pr_GarantiaFabricante , pr_IndicePreco,PR_precoVendaLiquido1,PR_custoMedioLiquido1
		  from produto as r where  not exists 
		 (select * from Desenv_DMAC_Loja..produtoLoja as l 
		  where l.PR_Referencia=r.pr_referencia) 

		 insert into Desenv_DMAC_Loja..LinhaProduto select * from LinhaProduto as r where not exists 
		 (select * from Desenv_DMAC_Loja..LinhaProduto as l 
		  where l.LPR_Linha =r.LPR_Linha) 

          Update GLB_ControleTarefas set CTA_Produto = 'N'

        End  
		
     If @Promocao = 'S'
        Begin
          truncate Table Promocao
          Insert into Promocao Select * from [DemeoServer].[Demeo].[dbo].Promocao
          Execute SP_VDA_Atualiza_Promocao_Loja
          Update GLB_ControleTarefas set CTA_Promocao = 'N'

        End		
		 
     If @Estoque = 'S'
        Begin
         -- truncate Table Estoque
          Insert into Estoque Select * from [DemeoServer].[Demeo].[dbo].estoque as D
          where  not exists(select es_referencia from estoque as c where
                d.es_loja=c.ES_Loja and d.es_referencia=c.es_referencia) 
              
           delete Estoque FROM ESTOQUE AS M 
           where not exists (select * from [DemeoServer].[Demeo].[dbo].Estoque as d 
           where M.es_referencia=d.es_Referencia)
 

		  
          Update ESTOQUE SET ES_Estoque=D.ES_ESTOQUE FROM ESTOQUE as m,[DemeoServer].[Demeo].[dbo].estoque AS D
          where m.ES_Loja=d.ES_Loja and m.ES_Referencia=d.es_referencia and m.ES_Estoque<>d.ES_Estoque and m.es_loja<>'28'
    
    	  delete Desenv_DMAC_Loja..EstoqueLoja 
	      from Desenv_DMAC_Loja..EstoqueLoja as l
          where not exists (select ES_Referencia from Estoque as r 
          where l.el_referencia=r.es_Referencia and ES_Loja='271')
		  
          insert into Desenv_DMAC_Loja..estoqueloja(EL_Loja,EL_Referencia,
                      EL_CodigoFornecedor,EL_Estoque,EL_EstoqueAnterior)       
          select es_loja,es_referencia,pr_codigofornecedor,0,0 from estoque ,produto 
		  where PR_Referencia=ES_Referencia and ES_Loja='271' and not exists 
		 (select * from Desenv_DMAC_Loja..EstoqueLoja as l 
		  where EL_Referencia=ES_referencia and ES_Loja='271') 

		  -- -- 

		   delete [dmac28].[dmac_loja].[dbo].EstoqueLoja 
	      from [dmac28].[dmac_loja].[dbo].EstoqueLoja as l
          where not exists (select ES_Referencia from Estoque as r 
          where l.el_referencia=r.es_Referencia and ES_Loja='28')
		  
          insert into [dmac28].[dmac_loja].[dbo].estoqueloja(EL_Loja,EL_Referencia,
                      EL_CodigoFornecedor,EL_Estoque,EL_EstoqueAnterior)       
          select es_loja,es_referencia,pr_codigofornecedor,0,0 from estoque ,produto 
		  where PR_Referencia=ES_Referencia and ES_Loja='28' and not exists 
		 (select * from [dmac28].[dmac_loja].[dbo].EstoqueLoja as l 
		  where EL_Referencia=ES_referencia and ES_Loja='28') 

          Update GLB_ControleTarefas set CTA_Estoque = 'N'


        End        
        If @Sugestao = 'S'
           Begin

			IF object_id('SugestaoTransferenciaTEMP') IS NOT NULL 
			BEGIN
				drop table SugestaoTransferenciaTEMP
			END

			select * into SugestaoTransferenciaTEMP from SugestaoTransferencia as dmac where not exists 
			(select * from [demeoserver].[demeo].[dbo].SugestaoTransferencia as demeo where
			demeo.ST_LojaOrigem = dmac.ST_LojaOrigem 
			and demeo.ST_LojaDestino = dmac.ST_LojaDestino 
			and demeo.ST_DataSugestao = dmac.ST_DataSugestao
			and demeo.ST_Referencia = dmac.ST_Referencia
			and demeo.ST_Situacao = dmac.ST_Situacao
			AND demeo.ST_NumeroSugestao = dmac.ST_NumeroSugestao)
			and dmac.ST_DataSugestao >= DATEADD(mm,-1, getdate())

			declare @ST_Situacao char(1),
			@ST_EnviaLoja as char(1),
			@ST_TipoSugestao as char(1)

		    while (select count(ST_NumeroSugestao) from SugestaoTransferenciaTEMP) <> 0
				Begin

					select top 1 @nf = VC_NotaFiscal, 
					@serie = VC_Serie, 
					@dataemissao = CONVERT (date,VC_DataEmissao, 111), 
					@lojaorigem = VC_LojaOrigem,
					@observacao = VC_Observacao,
					@lojaDestino = VC_LojaDestino
					from SugestaoTransferenciaTEMP order by VC_NotaFiscal

					select @sql = 'update [DEMEOSERVER].[Demeo].[dbo].capanfvenda 
					set vc_observacao = ''' + @observacao + '''
					where vc_notafiscal = ''' + @nf + '''
					and vc_serie = ''' + @serie + '''
					and vc_dataemissao = ''' + @dataemissao + '''
					and vc_lojaorigem = ''' + @lojaorigem + ''''

					--print (@sql)
					execute (@sql)

					select @item = 1

					while (select COUNT(VI_NumeroItem) from itemnfvenda where VI_NotaFiscal = @nf 
					and VI_LojaOrigem = @lojaorigem 
					and VI_DataEmissao = @dataemissao
					and VI_Serie = @serie and VI_NumeroItem = @item) <> 0
					Begin

						select @referencia = VI_Referencia, 
						@quantidade = VI_Quantidade 
						from itemnfvenda where VI_NotaFiscal = @nf 
						and VI_LojaOrigem = @lojaorigem 
						and VI_DataEmissao = @dataemissao
						and VI_Serie = @serie and VI_NumeroItem = @item

						select @sql = 'Update [DEMEOSERVER].[Demeo].[dbo].Estoque 
						set es_estoque = es_estoque + ' + @quantidade + ', 
						es_transito=es_transito - ' + @quantidade + ' 
						where es_referencia = ''' + @referencia + ''' 
						and es_loja = ''' + @lojaDestino + ''''

						--print (@sql)
						execute (@sql)
				
						select @item = @item + 1

					end

					delete capanfvendaSicronizacao 
					where VC_NotaFiscal = @nf
					and VC_Serie = @serie
					and VC_LojaOrigem = @lojaorigem
					and VC_DataEmissao = @dataemissao

				end

              truncate Table SugestaoTransferencia
              Insert into SugestaoTransferencia Select * from [DemeoServer].[Demeo].[dbo].SugestaoTransferencia
              Update GLB_ControleTarefas set CTA_Sugestao = 'N'
           End
        If @Ajuste = 'S'     
           Begin
              truncate Table Ajuste
              Insert into ajuste Select * from [DemeoServer].[Demeo].[dbo].ajuste 
                     where aj_situacao='A' and aj_data >'2014/09/01' and AJ_Loja = '28'
              Update GLB_ControleTarefas set CTA_Ajuste = 'N'       
           End  
  
    -- select * from [dmac28].[dmac_loja].[dbo].produtoLoja where pr_referencia='1780513'
   
/*
select * from ajuste where aj_data > '2014/09/01' and aj_situacao='A'

select * from [dmac28].[dmac_loja].[dbo].estoqueloja
select es_loja,es_referencia,count(*) from [DemeoServer].[Demeo].[dbo].estoque
group by es_loja,es_referencia having count(*) >1 



select * from glb_controletarefas
update GLB_ControleTarefas set CTA_Produto='s'
update GLB_ControleTarefas set CTA_Promocao='S'
update GLB_ControleTarefas set CTA_estoque='N'
update GLB_ControleTarefas set CTA_ProdutoLoja='N'
update GLB_ControleTarefas set CTA_PromocaoLoja='N' 
update GLB_ControleTarefas set CTA_sugestao='s'
update GLB_ControleTarefas set CTA_ajuste='s'
  
select pr_IndicePreco,* from produto
update produto set pr_cst= '060' where pr_substituicaotributaria='S'
update produto set pr_cst= '020' where pr_codigoreducaoicms > 0 and pr_substituicaotributaria='N'
update produto set pr_cst= '000' where pr_codigoreducaoicms = 0 and pr_substituicaotributaria='N'
update produto set pr_IndicePreco= '2' where pr_codigofornecedor = 680
update produto set pr_IndicePreco= '1' where pr_codigofornecedor <> 680
update produto set pr_IndicePreco= '3' where pr_classe = 'P'
pr_IndicePreco
where PR_Referencia = '1420019'
drop table TempPromocaoLoja
select * from Promocao
Select * from glb_Controletarefas
update glb_Controletarefas set cta_Promocao='S'
delete promocao where PM_NumeroPromocao <> 3391
insert into glb_Controletarefas(CTA_Produto,CTA_ListaPreco,CTA_Promocao,CTA_Estoque,CTA_Outras) 
values ('N','N','N','N','N')

-- select * from [dmac28].[dmac_loja].[dbo].nfitens where dataemi='2014/08/13'
-- update estoque set es_estoque=el_estoque from estoque,[dmac28].[dmac_loja].[dbo].estoqueloja
   where es_referencia=el_referencia and es_loja='28' and el_loja=es_loja 
   and el_estoque<>es_estoque 

select es_loja,es_referencia,es_estoque,el_estoque from [dmac28].[dmac_loja].[dbo].estoqueloja,
estoque where el_referencia=es_referencia and es_estoque<>el_estoque and es_loja=el_loja




select es_loja,es_referencia,es_estoque,el_estoque from [dmac28].[dmac_loja].[dbo].estoqueloja,


AND ES_ESTOQUE=0

UPDATE ESTOQUE SET es_estoque = 0 WHERE ES_REFERENCIA IN (select es_referencia from [dmac28].[dmac_loja].[dbo].estoqueloja,[DemeoServer].[Demeo].[dbo].estoque where el_referencia=es_referencia and es_estoque<>el_estoqueanterior and es_loja=el_loja AND ES_ESTOQUE=0)
AND ES_LOJA = '28'

UPDATE [dmac28].[dmac_loja].[dbo].ESTOQUELOJA SET el_estoque = 0 WHERE El_REFERENCIA IN (select es_referencia from [dmac28].[dmac_loja].[dbo].estoqueloja,[DemeoServer].[Demeo].[dbo].estoque where el_referencia=es_referencia and es_estoque<>el_estoqueanterior and es_loja=el_loja AND ES_ESTOQUE=0)
AND El_LOJA = '28'



update [dmac28].[dmac_loja].[dbo].produtoloja set PR_precoVendaLiquido1=d.PR_precoVendaLiquido1
,PR_custoMedioLiquido1=d.PR_custoMedioLiquido1 from [dmac28].[dmac_loja].[dbo].produtoloja as l,produto as d
where l.pr_referencia=d.pr_referencia

select * into tranferenciaEntrada281707 select vi_referencia,sum(quantidade)
from itemnfvenda where vi_lojadestino='28' and vi_tiponota='T' 


select es_referencia from [dmac28].[dmac_loja].[dbo].estoqueloja,
[DemeoServer].[Demeo].[dbo].estoque 
where el_referencia=es_referencia and es_estoque<>el_estoqueanterior and es_loja=el_loja


select referencia,qtde from nfitens,estoque where referencia=es_referencia and es_loja='28'
and es_estoqueanterior2 <> 0 and dataemi='2014/xx/xx'


 
select * from estoque where el_estoqueanterior2 >0

==================================================================================================================
---------------------------------------------Acertando Estoque Dmac/loja------------------------------------------ 

update 

update [dmac28].[dmac_loja].[dbo].estoqueloja set el_estoqueanterior=el_estoque

select * from nfitens where dataemi='2014/09/16' and tiponota in('V','T','E')

select * from  [dmac28].[dmac_loja].[dbo].estoqueloja where el_estoque<>el_estoqueanterior

update estoque set es_estoqueanterior=el_estoqueanterior
from estoque,[dmac28].[dmac_loja].[dbo].estoqueloja where el_referencia=es_referencia and es_loja=el_loja



select es_loja,es_referencia,es_estoque,el_estoque from [dmac28].[dmac_loja].[dbo].estoqueloja,
estoque where el_referencia=es_referencia and es_estoque<>el_estoque and es_loja=el_loja


select es_referencia,es_estoque,el_estoqueanterior from [dmac28].[dmac_loja].[dbo].estoqueloja,
[DemeoServer].[Demeo].[dbo].estoque 
where el_referencia=es_referencia and es_estoque<>el_estoqueanterior and es_loja=el_loja

select es_referencia from [DemeoServer].[Demeo].[dbo].estoque where es_transito < 0 and es_loja = '28'


 update estoque set es_estoqueanterior2=9999

 update estoque set es_estoqueanterior2=d.es_estoque 
 from estoque as m,[DemeoServer].[Demeo].[dbo].estoque as d 
 where d.es_referencia=m.es_referencia and 
 d.es_loja=m.es_loja and d.es_loja='28' and m.es_estoqueanterior<>d.es_estoque
 
 
    drop table TmpSomaRef
	Create Table TmpSomaRef (
		Referencia	Char(7) COLLATE SQL_Latin1_General_CP1_CI_AS	Not Null,
		Qtde		Float)
insert into TmpSomaRef select referencia,sum(qtde) from nfitens 
where dataemi='2014/09/18' and tiponota in('T','V') group by referencia

 
 
 update [dmac28].[dmac_loja].[dbo].estoqueloja set el_estoque=es_estoqueanterior2
from [dmac28].[dmac_loja].[dbo].estoqueloja,estoque where el_referencia=es_referencia
and es_loja=el_loja and es_estoqueanterior2 <> 9999 and es_loja='28'

update estoque set es_estoque=es_estoqueanterior2 
where es_estoqueanterior2 <> 9999 and es_loja='28'
 
UPDATE ESTOQUE SET ES_ESTOQUE=(ES_ESTOQUE-qtde) 
from estoque,TmpSomaRef WHERE ES_REFERENCIA =referencia and es_loja='28'

update [dmac28].[dmac_loja].[dbo].estoqueloja set el_estoque=(EL_ESTOQUE-qtde)
from [dmac28].[dmac_loja].[dbo].estoqueloja,TmpSomaRef WHERE EL_REFERENCIA =referencia


----------------------------------------------------------------------------------------------------------------
 ---DEVOLUÇÂO-----
 -----------------
 
 
  drop table TmpSomaRef
	Create Table TmpSomaRef (
		Referencia	Char(7) COLLATE SQL_Latin1_General_CP1_CI_AS	Not Null,
		Qtde		Float)
insert into TmpSomaRef select referencia,sum(qtde) from nfitens 
where dataemi='2014/09/18' and tiponota in('E') group by referencia
 
UPDATE ESTOQUE SET ES_ESTOQUE=(ES_ESTOQUE+qtde) 
from estoque,TmpSomaRef WHERE ES_REFERENCIA =referencia and es_loja='28'

update [dmac28].[dmac_loja].[dbo].estoqueloja set el_estoque=(EL_ESTOQUE+qtde)
from [dmac28].[dmac_loja].[dbo].estoqueloja,TmpSomaRef WHERE EL_REFERENCIA =referencia

----------------------------------------------------------------------------------------------------------------




select * from [dmac28].[dmac_loja].[dbo].nfitens where dataemi='2014/09/10'  and tiponota in('v','t')
and referencia in('0600217','0979368','1583837','1584005','1584034','1780122','1780513','2370013')


==================================================================================================================
*/

--exec SP_Atualiza_dados_Tabelas_DMAC
/*
28      0618086       3           4
28      1583927       2           1
*/