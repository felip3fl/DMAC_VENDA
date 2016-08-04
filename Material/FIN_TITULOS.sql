	select nf,serie,lojaorigem, * from nfcapa where dataemi >= '2015/06/08' and condpag > '3' order by lojaorigem
	
	select nf,serie,lojaorigem, * from nfcapa,fin_titulos where dataemi >= '2015/06/08' and condpag > '3' 
	and tit_documento = rtrim(ltrim(str(nf))) and tit_serie = serie and tit_loja = lojaorigem
	order by lojaorigem

	select nf,serie,lojaorigem, * from nfcapa where dataemi >= '2015/06/08' and condpag > '3' 
	and not exists (select * from fin_titulos where tit_documento = rtrim(ltrim(str(nf))) and tit_serie = serie and tit_loja = lojaorigem)
	
	sp_help nfcapa
	select * from nfcapa where dataemi > '2015/06/06' and condpag > 3
		select * from condicaopagamento where cp_codigo = 4
	select * from fin_titulos where tit_documento = '2664' and tit_serie = 'NE' and tit_loja = '181'
	select * from Duplicata where dp_notafiscal = '2638'


exec SP_FIN_Cria_titulos_Receber_Faturado '2015/06/09','28','7086','NE'

delete fin_titulos where tit_documento = '7086' and tit_serie = 'NE' and tit_loja = '28'

select top  2 * from fin_titulos where tit_documento = '7086' and tit_serie = 'NE' and tit_loja = '28'
select * from demeo..Duplicata where DP_NotaFiscal = '7086' and DP_Serie = 'NE' AND DP_Loja = '28'
select * from Duplicata where DP_NotaFiscal = '7086' and DP_Serie = 'NE' AND DP_Loja = '28'

SELECT * FROM demeo..Duplicata where DP_DataEmissao > '2015/06/08'

insert into Duplicata(DP_Loja,DP_NotaFiscal,DP_Serie,DP_Sequencia,DP_CodigoCliente,DP_DataEmissao,
DP_Vendedor,DP_Banco,DP_DocumentoBancario,DP_ValorDuplicata,DP_DataVencimento,DP_NotaCredito,
DP_Abatimento,DP_Desconto,DP_Despesas,DP_Juros,DP_ValorPago,DP_DataPagamento,DP_DataBaixa,
DP_DataCartorio,DP_Historico,DP_TipoPagamento,DP_Agrupamento,DP_Situacao) 
select TIT_Loja, TIT_Documento, TIT_Serie, TIT_Parcela, TIT_EmitenteCedente, 
TIT_DataEmissao, TIT_Vendedor, TIT_Banco, TIT_DocumentoBanco, TIT_ValorTitulo, TIT_Vencimento, '', TIT_Abatimento,
0, 0,TIT_JurosMora, TIT_ValorPago, null, TIT_DataBaixa, TIT_DataCartorio, null, null, TIT_AgrupaTitulo, TIT_Situacao 
from fin_titulos
where tit_documento = '7086' and tit_serie = 'NE' and tit_loja = '28'

UPDATE fin_titulos set TIT_ValorPago = 0 where TIT_ValorPago is null
UPDATE AP set AP_Bordero = 0 where AP_Bordero is null

--insert into Duplicata select * from demeo..Duplicata


Select AP_NumeroAP, AP_DataDocumento,AP_NumeroDocumento,AP_ValorTotalDocumento,
AP_Vencimento,AP_NumeroBanco, AP_InfComplementares, AP_Valor, FO_CGC, CC_Serie, 
CC_ValorTotalNota, CC_ValorCalculado, LO_Loja, LO_OrdemLoja, AP_NumeroParcela  
from AP, Fornecedor, CapaNFCompra, Loja  Where AP_CodigoFornecedor=FO_CodigoFornecedor 
And AP_NumeroDocumento= CC_NotaFiscal And LO_Loja = CC_Loja And AP_Usuario=LO_Loja 
And AP_CodigoFornecedor=CC_Fornecedor And AP_DataDocumento between '06/08/2015' And '06/11/2015' 
And ap_usuario<>'800' and CC_SITUACAO IN ('L','T') Order By AP_NumeroDocumento


select * from fin_titulos where TIT_TipoDocumento = '02' and TIT_DataEmissao > '2015/06/07'

SELECT * FROM ap WHERE AP_NumeroDocumento = '1352' AND AP_CodigoFornecedor = '62'
--delete ap WHERE AP_DataDocumento > '2015/06/07'
--DELETE ap WHERE AP_NumeroDocumento = '1352' AND AP_CodigoFornecedor = '62'

SELECT * FROM capanfcompra WHERE CC_DataEntrada > '2015/06/07' AND CC_TipoEntrada = '1'
exec CriaAp '27083','25','S1','CD'

/*
SELECT * FROM capanfcompra WHERE CC_DataEntrada > '2015/06/07' AND CC_TipoEntrada = '1'
SELECT * FROM capanfcompra WHERE CC_DataEntrada > '2015/06/07' AND CC_TipoEntrada = '1' AND NOT EXISTS (select * from AP WHERE ap_numeroDocumento = cc_notafiscal and ap_codigoFornecedor = cc_fornecedor) ORDER BY CC_SERIE, CC_Fornecedor

declare @nf as varchar(10)
declare @serie as varchar(10)
declare @fornecedor as varchar(10)
declare @loja as varchar(10)
while (SELECT count(*) FROM capanfcompra WHERE CC_DataEntrada > '2015/06/07' AND CC_TipoEntrada = '1' AND NOT EXISTS (select * from AP WHERE ap_numeroDocumento = cc_notafiscal and ap_codigoFornecedor = cc_fornecedor) ) > 0
begin
	SELECT top 1 @nf = cc_notafiscal, @serie = cc_serie, @fornecedor = cc_fornecedor, @loja = cc_loja FROM capanfcompra WHERE CC_DataEntrada > '2015/06/07' AND CC_TipoEntrada = '1' AND NOT EXISTS (select * from AP WHERE ap_numeroDocumento = cc_notafiscal and ap_codigoFornecedor = cc_fornecedor)
	exec CriaAp @nf,@fornecedor,@serie,@loja
end



/*

SELECT VENCIMENTOSFORNECEDOR.* FROM capanfcompra,VENCIMENTOSFORNECEDOR WHERE CC_DataEntrada > '2015/06/07' and 
CC_NotaFiscal = VF_NotaFiscal and VF_Serie = CC_Serie and VF_Fornecedor = CC_Fornecedor
SELECT * FROM VENCIMENTOSFORNECEDOR WHERE VF_DataVencimento > '2015/06/06'

SELECT * FROM tipopedidocompra

select * from vencimento

SELECT * FROM CodigoOperacaoNovo WHERE CN_CodigoOperacaoNovo = '1949'