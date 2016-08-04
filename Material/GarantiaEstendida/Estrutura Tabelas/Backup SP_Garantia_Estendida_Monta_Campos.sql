GO
/****** Object:  StoredProcedure [dbo].[SP_Garantia_Estendida_Monta_Campos]    Script Date: 21/11/2013 14:49:02 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO


create PROCEDURE [dbo].[SP_Garantia_Estendida_Monta_Campos]
	@numeroPedido numeric           
As 
Begin

	select itens.referencia, itens.qtdeGarantia, itens.item, itens.planoGarantia, 
	itens.lojaOrigem, itens.certificadoInicio,  itens.dataemi as dataEmissao,
	cast(prod.pr_garantiaFabricante/30 as integer) as garantiaFabricante, 
	prod.pr_descricao, fornec.fo_nomeFantasia as Marca, itens.VLUNIT, 
	faixa.fpg_premioLiquido as CustoDaSegurandora, 
	itens.ge_iof as IOF, itens.ge_valorCustoSeguradora as PremioTotal, itens.ge_premioLiquido as PremioLiquido,
	itens.certificadoInicio, itens.certificadoFim
	from nfitens itens, produto prod, faixapremioge faixa, fornecedor fornec, nfcapa capa
	where itens.numeroPed = @numeroPedido and itens.garantiaEstendida = 'S' and 
	itens.certificadoInicio is not null and
	prod.pr_referencia = itens.referencia and 
	itens.VLUNIT between faixa.fpg_faixainicial and faixa.fpg_faixaFinal and 
	faixa.fpg_plano = itens.planoGarantia and
	fornec.fo_codigoFornecedor = prod.pr_codigoFornecedor and
	capa.numeroPed = @numeroPedido and capa.garantiaEstendida = 'S' and capa.tipoNota = 'V'
	order by item

end
