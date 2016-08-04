select * from fin_cliente where ce_cgc = '70721769349'


insert into svdmac.dmac.dbo.ClienteFidelidade
select CE_CodigoFidelidade as CEF_CodigoFidelidade,
CE_CodigoCliente as CEF_CodigoCliente,
cE_cgc as CEF_CNPJCPF,
(select top 1 cts_loja from ControleSistema) as CEF_LojaCadastro,
CE_Vendedor as CEF_VendedorCadastro,
CE_Razao as CEF_Razao,
CE_Endereco as CEF_Endereco,
ce_numero as CEF_Numero,
CE_Bairro as CEF_Bairro,
CE_Municipio as CEF_Municipio,
CE_Estado as CEF_Estado,
CE_CEP as CEF_CEP,
CE_Telefone as CEF_Telefone,
CE_EMail as CEF_EMail,
CE_TipoPessoa as CEF_TipoPessoa,
CE_DataCadastro as CEF_DataCadastro,
CE_Situacao CEF_Situacao,
CE_PontosFidelidade AS CEF_PontosFidelidade,
'' as CEF_DataUltimaCompra,
'' as CEF_DataUltimoResgate,
0 as CEF_CompraAcumulada,
0 as CEF_PontoResgatado,
0 as CEF_AutoPrint,
0 as CEF_ValorUltimoCupom
 from fin_cliente
 where ce_cgc = '70721769349'


 exec svdmac.dmac.dbo.SP_Ponto_Clientes_Fidelidade


