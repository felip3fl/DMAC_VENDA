USE [DMAC]
GO
/****** Object:  Trigger [dbo].[SP_Atualiza_dados_Tabelas_DMAC]    Script Date: 22/08/2014 09:16:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO


ALTER       Trigger [dbo].[SP_Atualiza_dados_Tabelas_DMAC]
   on [dbo].[GLB_ControleTarefas]
for Update 
   as 
    Declare @Promocao as char(1),
            @Produto as char(1), 
            @ListaPreco as char(1),
            @Estoque as char(1),
            @Outras as char(1),
            @ProdutoLoja as char(1),
            @PromocaoLoja as char(1),
            @Sugestao as char(1)
            
     select @Promocao     = (Select Inserted.CTA_Promocao from Inserted)
     select @Produto      = (Select Inserted.CTA_Produto from Inserted) 
     select @ListaPreco   = (Select Inserted.CTA_ListaPreco from Inserted) 
     select @Estoque      = (Select Inserted.CTA_Estoque from Inserted)
     select @Outras       = (Select Inserted.CTA_Outras from Inserted)
     select @ProdutoLoja  = (Select Inserted.CTA_Produtoloja from Inserted)
     select @PromocaoLoja = (Select Inserted.CTA_PromocaoLoja from Inserted)
     select @Sugestao     = (Select Inserted.CTA_Sugestao from Inserted) 
     
     If @Promocao = 'S'
        Begin
          truncate Table Promocao
          Insert into Promocao Select * from [DemeoServer].[Demeo].[dbo].Promocao
          update  [dmac28].[dmac_loja].[dbo].glb_controletarefas set CTA_Promocao='T'
          --Update GLB_ControleTarefas set CTA_Promocao = 'S'

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
		  pr_GarantiaFabricante , pr_IndicePreco) 
		  select pr_Referencia , pr_CodigoFornecedor , pr_Descricao , pr_Classe , pr_Bloqueio , 
		  pr_LinhaProduto , pr_ClasseFiscal , pr_Unidade , pr_ICMSSaida , pr_CodigoReducaoICMS , 
		  pr_CustoMedio1 , pr_PrecoVenda1 , pr_PaginaListaPreco , pr_Peso, pr_Comprador ,
		  pr_Situacao , pr_SubstituicaoTributaria , pr_IcmPdv, pr_HoraManutencao , 
		  pr_CodigoProdutoNoFornecedor , pr_IcmsSaidaIva, pr_IcmsPdvSaidaIva , 
		  pr_ICMSEntrada , pr_IcmPdvEntrada , pr_CST , pr_GarantiaEstendida , 
		  pr_GarantiaFabricante , pr_IndicePreco from produto as r where  not exists 
		 (select * from [dmac28].[dmac_loja].[dbo].produtoLoja as l 
		  where l.PR_Referencia=r.pr_referencia) 

		  DROP TABLE ProdutoBarras
		  select * into produtoBarras from [DemeoServer].[Demeo].[dbo].ProdutoBarras

		  insert into [dmac28].[dmac_loja].[dbo].produtoBarras 
		  select * from produtoBarras as DMAC where not exists 
		  (select PRB_CodigoBarras from [dmac28].[dmac_loja].[dbo].produtoBarras as loja 
		  where loja.PRB_CodigoBarras = dmac.PRB_CodigoBarras)

          Update GLB_ControleTarefas set CTA_Produto = 'N'

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
              truncate Table SugestaoTransferencia
              Insert into SugestaoTransferencia Select * from [DemeoServer].[Demeo].[dbo].SugestaoTransferencia
           End
  
  -- select * from [dmac28].[dmac_loja].[dbo].produtoLoja where pr_referencia='1780513'
   
/*
select * from [dmac28].[dmac_loja].[dbo].estoqueloja
select es_loja,es_referencia,count(*) from [DemeoServer].[Demeo].[dbo].estoque
group by es_loja,es_referencia having count(*) >1 
select * from GLB_ControleTarefas

update GLB_ControleTarefas set CTA_Produto='S'
update GLB_ControleTarefas set CTA_Promocao='s'
update GLB_ControleTarefas set CTA_estoque='S'
update GLB_ControleTarefas set CTA_ProdutoLoja='S'
update GLB_ControleTarefas set CTA_PromocaoLoja='S'   
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
estoque where el_referencia=es_referencia and es_estoque<>el_estoque and es_loja=el_loja


*/







