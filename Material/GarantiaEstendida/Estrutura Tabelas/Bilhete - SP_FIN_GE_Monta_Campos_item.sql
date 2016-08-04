create PROCEDURE [dbo].[SP_FIN_GE_Monta_Campos_item]
	@numeroPedido numeric           
As 
Begin
	 
	IF object_id('temp_GE_Itens') IS NOT NULL 
	BEGIN
		drop table temp_GE_Itens
	END

	SELECT 
	convert(varchar(20),
	REPLICATE ( '0' ,4 - LEN(controle.ct_codigoEstipulanteGE)) + 
	RTRIM(controle.ct_codigoEstipulanteGE) + 
	'01' + REPLICATE ( '0' ,6 - LEN(loja.LO_Loja)) + 
	RTRIM(loja.LO_Loja) + REPLICATE ( '0' ,8 - LEN(itens.CertificadoInicio)) + 
	RTRIM(itens.CertificadoInicio))
	as NumeroDoBilhete,

	itens.dataemi as DataEmissaoSeguro,
	itens.dataemi as DataCompraBem,
	itens.dataemi as PerildoGarantiaInicio,
	DATEADD (mm, cast(prod.pr_garantiaFabricante/30 as integer), itens.dataemi) as PerildoGarantiaFIM,
	DATEADD (dd, 01, DATEADD (mm, cast(prod.pr_garantiaFabricante/30 as integer), itens.dataemi)) as PerildoVigenciaInicio,
	DATEADD (mm, itens.planoGarantia, itens.dataemi) as PerildoVigenciaFIM,

	'R$ ' + convert(char(15),replace(convert(decimal(10,2), itens.VLUNIT),'.',',')) as limiteMaximoIndenizacao,

	prod.pr_descricao as ProdutoSegurado, 
	fornec.fo_nomeFantasia as Marca, 
	itens.referencia as Modelo, 
	
	'R$ ' + convert(char(15),replace(convert(decimal(10,2), itens.VLUNIT),'.',',')) as ValorProduto, 
	'R$ ' + convert(char(15),replace(convert(decimal(10,2), itens.ge_premioLiquido),'.',',')) as PremioLiquido, 
	'R$ ' + convert(char(15),replace(convert(decimal(10,2), itens.ge_iof),'.',',')) as IOF, 
	'R$ ' + convert(char(15),replace(convert(decimal(10,2), itens.ge_valorCustoSeguradora),'.',',')) as PremioTotal, 

	itens.certificadoInicio as CertificadoInicio, 
	itens.certificadoFim as CertificadoFim
	
	into temp_GE_Itens
	from nfitens itens, 
	produto prod, 
	faixapremioge faixa, 
	fornecedor fornec, 
	nfcapa capa,
	controle controle,
	lojas loja

	where itens.numeroPed = @numeroPedido  
	and itens.garantiaEstendida = 'S' 
	and itens.certificadoInicio is not null 
	and prod.pr_referencia = itens.referencia  
	and itens.VLUNIT between faixa.fpg_faixainicial 
	and faixa.fpg_faixaFinal  
	and faixa.fpg_plano = itens.planoGarantia 
	and fornec.fo_codigoFornecedor = prod.pr_codigoFornecedor 
	and capa.numeroPed = @numeroPedido  
	and capa.nf = itens.nf
	and capa.garantiaEstendida = 'S' 
	and capa.tipoNota = 'V'
	and capa.LOJAORIGEM = loja.lo_loja
	order by CertificadoInicio

	declare @contador int

	update temp_GE_Itens set CertificadoFim = '' where certificadoInicio = certificadofim
	set @contador = (select count(CertificadoInicio) from temp_GE_Itens where CertificadoFim not in (CertificadoInicio,''))

	while @contador <> 0
		Begin
		declare @certificadoInicio as char(12)
		declare @certificadoFIM as char(12)
		declare @NumeroCertificado as char(12)

		set @certificadoInicio = (select top 1 CertificadoInicio from temp_GE_Itens where CertificadoFim not in (CertificadoInicio,''))
		set @certificadoFIM = (select top 1 CertificadoFIM from temp_GE_Itens where CertificadoFim not in (CertificadoInicio,''))

		while @certificadoInicio < @certificadoFIM
		Begin
			set @certificadoInicio = @certificadoInicio + 1
			insert into temp_GE_Itens select top 1 SUBSTRING(NumeroDoBilhete,1,12) + 
			REPLICATE ( '0' ,8 - LEN(@CertificadoInicio)) + RTRIM(@CertificadoInicio), 
			DataEmissaoSeguro, 
			DataCompraBem, 
			PerildoGarantiaInicio, 
			PerildoGarantiaFIM, 
			PerildoVigenciaInicio, 
			PerildoVigenciaFIM, 
			limiteMaximoIndenizacao, 
			ProdutoSegurado, 
			Marca, 
			Modelo, 
			ValorProduto, 
			PremioLiquido, 
			IOF,
			PremioTotal, 
			@CertificadoInicio, 
			'' 
			from temp_GE_Itens 
			where CertificadoFim not in (CertificadoInicio,'')
			print @certificadoInicio
		end 

		update temp_GE_Itens 
		set CertificadoFim = '' 
		where certificadoInicio <> @certificadoFIM 
		and CertificadoFIM = @certificadoFIM
		
		set @contador = (select count(CertificadoInicio) 
		from temp_GE_Itens 
		where CertificadoFim not in (CertificadoInicio,''))
		
	end

end

/*
SP_FIN_GE_Monta_Campos_item 3755
*/