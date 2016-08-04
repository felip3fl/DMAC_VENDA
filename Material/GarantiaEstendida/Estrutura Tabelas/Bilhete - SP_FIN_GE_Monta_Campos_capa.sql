CREATE PROCEDURE [dbo].[SP_FIN_GE_Monta_Campos_capa]
	@numeroPedido numeric           
As 
Begin

	SET LANGUAGE Português

	declare @CNPJLoja char(14)

	select @CNPJLoja = SUBSTRING(lo_cgc,len(lo_cgc)-13,14) 
	from lojas, nfcapa
	where numeroPed = @numeroPedido 
	and lojaorigem = lo_loja

	select 
	'SERRO ADM. CORR. SEG. LTDA' as Corretor,
	'059626.1.02.8522-6' as RegistroNaSusep,
	'Reparo' as CoberturaContratada,
	lo_razao as RepresentanteDeSeguros,
	SUBSTRING(@CNPJLoja,1,2) + '.'
	+ SUBSTRING(@CNPJLoja,3,3) + '.'
	+ SUBSTRING(@CNPJLoja,6,3) + '/'
	+ SUBSTRING(@CNPJLoja,9,4) + '-'
	+ SUBSTRING(@CNPJLoja,13,2) as CNPJRepresentante,
	'0800 198915' as CanaisDeAtendimento,
	dataemi as DataEmissaoSeguro,
	
	ce_razao as Segurado,
	SUBSTRING(ce_cgc,1,3) + '.'
	+ SUBSTRING(ce_cgc,4,3) + '.'
	+ SUBSTRING(ce_cgc,7,3) + '-'
	+ SUBSTRING(ce_cgc,10,2) as CPF,
	ce_cgc as terste,
	ce_telefone as Telefone,
	ce_endereco as Endereco, 
	ce_bairro as Bairro, 
	ce_numero as Numero,
	ce_complemento as Complemento, 
	ce_municipio as Cidade, 
	ce_cep as CEP,
	ce_estado as UF,

	lo_municipio as EmissaoCidade,
	'dia ' + convert(varchar(10),day(dataemi)) + ' de ' + 
	convert(varchar(10),DATENAME(MONTH, dataemi)) + ' de ' + 
	convert(varchar(10),year(dataemi)) as EmissaoData

	from controle, 
	nfcapa, 
	cliente,
	lojas
	where numeroPed = @numeroPedido 
	and cliente = ce_codigoCliente
	and lojaorigem = lo_loja
	and tiponota = 'V'

end

/*
SP_FIN_GE_Monta_Campos_capa 3756
*/

