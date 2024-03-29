

ALTER PROCEDURE [dbo].[SP_Atualiza_Modalidade_Venda_Pedido]
                 @Pedido numeric,
                 @Tipo   char(02),
                 @Codigo char(03),
                 @TipoCondicao char(02) 
As
	Declare  @ToalItens                float
	Declare  @Parcelas                 float
begin
   
	update nfitens set vlunit=ROUND((VLUNIT * cp_coeficiente),2),vltotitem=ROUND(((VLUNIT * QTDE) * cp_coeficiente),2)
	From produtoloja, CondicaoPagamento, nfitens
	where PR_IndicePreco=CP_ID and CP_Tipo=@Tipo and CP_Codigo=@Codigo
	and PR_Referencia = REFERENCIA and NUMEROPED =@Pedido

	update nfitens set DESCRAT = CP_DESCONTO
	From produtoloja, CondicaoPagamento, nfitens
	where PR_IndicePreco=CP_ID and CP_Tipo=@Tipo and CP_Codigo=@Codigo
	and PR_Referencia = REFERENCIA and NUMEROPED =@Pedido
 
	Select @ToalItens =(select SUM(vltotitem)from nfitens where NUMEROPED = @Pedido) 

	update nfcapa set vlrmercadoria=@ToalItens,totalnota=@ToalItens, CONDPAG = @TipoCondicao,
	ModalidadeVenda = (Case When @tipo = 'AV' Then 'A Vista' 
	when @tipo = 'CC' then 'Cartão' 
	when @tipo = 'FA' then 'Faturado' 
	when @tipo = 'FI' then 'Financiado' 
	end), Parcelas = CP_Parcelas
	From  nfcapa, CondicaoPagamento where NUMEROPED =@Pedido 
	and CP_Tipo=@Tipo and CP_Codigo=@Codigo and CP_ID = 1

		  

end 

/*
select top 1 * from CondicaoPagamento where cp_codigo = 100
Exec SP_Atualiza_Modalidade_Venda_Pedido 498,'FA','12'

*/
 