use DMAC_Loja

update svdmac.dmac.dbo.loja set lo_nomeServidor = '[DMAC181].[DMAC_Loja].[dbo].', lo_dmac = 'S' where lo_loja = '181'
select * into loja from svdmac.dmac.dbo.loja
select * from loja

--select * from estoqueloja
--truncate table EstoqueLoja

insert into estoqueloja(el_loja, el_referencia, el_codigoFornecedor, el_estoque, el_estoqueAnterior) 
select el_loja, el_referencia, el_codigoFornecedor, el_estoque, el_estoqueAnterior from [demeoserver].[loja181].[dbo].estoqueLoja where el_loja = '181'

-- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 

update controlesistema 
set cts_loja = ct_loja, 
cts_NUMERONF = ct_seqNota, 
cts_NUMEROPedido = 100, 
cts_NUMERONE = ct_numeronfe, 
cts_certificado = ct_certificado ,
CTS_SequenciaCliente = ct_seqCliente,
CTS_Serienota = 'NE',
CTS_DescontoVendedor = 3
from [demeoserver].[loja181].[dbo].Controle

--select ct_loja,ct_seqnota,ct_numped,ct_numeronfe,ct_certificado,* from [demeoserver].[loja181].[dbo].Controle
--select * from ControleSistema

- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 

update ParametroCaixa set par_loja = '181'

--select * from ParametroCaixa

- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 

truncate table vende
insert into Vende(VE_Codigo, 
VE_TotalVenda, 
VE_MargemVenda, 
VE_Nome, ve_senha) 
select loja.VE_Codigo, 
loja.VE_TotalVenda, 
loja.VE_MargemVenda, 
loja.VE_Nome, ''
from [demeoserver].[loja181].[dbo].vende as loja,
[demeoserver].[loja181].[dbo].usuario as usuario
where loja.ve_codigo not in ('999','888') 
and loja.ve_codigo = usuario.US_Codigo

--select * from Vende
--select * from [demeoserver].[loja181].[dbo].vende
--select * from [demeoserver].[loja181].[dbo].usuario

/*
update Vende set ve_senha = '1970' where ve_codigo = '291'
select * from Vende
*/

- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 

--select * from [demeoserver].[loja181].[dbo].usuario

truncate table UsuarioCaixa
insert into UsuarioCaixa(USU_Codigo,USU_Nome,USU_TipoUsuario,USU_Senha,USU_Situacao) select us_codigo, us_usuario, 'O' as tipoUsuario, '', 'A' as Situacao from [demeoserver].[loja181].[dbo].usuario where us_codigo not in ('888','999')

update usuariocaixa set usu_tipoUsuario = 'S', USU_situacao = 'S' where usu_nome in ('gerente')
delete UsuarioCaixa where usu_nome not in ('Caixa','Gerente')
insert into UsuarioCaixa values (99, 'fecger', 'F', '123456', 'F')
insert into UsuarioCaixa values (98, 'dmactrator', 'S', 'root#root', 'S')
insert into UsuarioCaixa values (97, 'dmactrator', 'O', 'root#root', 'A')

--select * from UsuarioCaixa
--delete UsuarioCaixa where usu_codigo = 3

/*
delete UsuarioCaixa where usu_nome not in ('Caixa','Gerente','dmactrator','fecger')

update UsuarioCaixa set usu_senha = '' where usu_nome = 'Gerente'
update UsuarioCaixa set usu_nome = 'Elias' where usu_nome = 'Gerente'

update vende set ve_nome = 'Elias', ve_senha = 'lili' where ve_nome = 'Gerente'

update UsuarioCaixa set usu_senha = '' where usu_nome = 'Caixa'
update UsuarioCaixa set usu_nome = 'Flavia' where usu_nome = 'Caixa'

update ControleSistema set cts_senhaLiberacao = 'lili', cts_senhaDesconto = 'lili'
*/

- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 

--truncate table fin_cliente
insert into fin_cliente 
(CE_CodigoCliente,CE_CGC,CE_InscricaoEstadual,CE_Razao,CE_Endereco,CE_Bairro,CE_Municipio,CE_Estado,CE_CEP,CE_Telefone,
CE_Fax,CE_EMail,CE_TipoPessoa,CE_Praca,CE_PagamentoCarteira,CE_EnderecoCobranca,CE_BairroCobranca,CE_MunicipioCobranca,
CE_EstadoCobranca,CE_CEPCobranca,CE_LimiteCredito,CE_DataLimiteCredito,CE_MaiorCompra,CE_DataMaiorCompra,CE_UltimaCompra,
CE_DataUltimaCompra,CE_UltimoPagamento,CE_DataUltimoPagamento,CE_MaiorAtraso,CE_QuantidadeCompras,CE_JurosCartorio,CE_DataCadastro,
CE_DataCancelamento,CE_Alteracao,CE_Situacao,CE_HoraManutencao,CE_Numero) 
select CE_CodigoCliente,CE_CGC,CE_InscricaoEstadual,CE_Razao,CE_Endereco,CE_Bairro,CE_Municipio,CE_Estado,CE_CEP,
CE_Telefone,CE_Fax,CE_EMail,CE_TipoPessoa,CE_Praca,CE_PagamentoCarteira,CE_EnderecoCobranca,CE_BairroCobranca,CE_MunicipioCobranca,
CE_EstadoCobranca,CE_CEPCobranca,CE_LimiteCredito,CE_DataLimiteCredito,CE_MaiorCompra,CE_DataMaiorCompra,CE_UltimaCompra,
CE_DataUltimaCompra,CE_UltimoPagamento,CE_DataUltimoPagamento,CE_MaiorAtraso,CE_QuantidadeCompras,CE_JurosCartorio,
CE_DataCadastro,CE_DataCancelamento,CE_Alteracao,CE_Situacao,CE_HoraManutencao,CE_Numero
from [demeoserver].[loja181].[dbo].cliente 

--select count(CE_CodigoCliente) from FIN_Cliente

--select top 100 * from FIN_Cliente

- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 

exec SP_Atualiza_Produto_Loja '181'
exec SP_Atualiza_Produto_Barras_Loja '181'

- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 

/*
truncate table nfcapa
truncate table nfitens
truncate table EMKTLoja
truncate table Agenda
truncate table MovimentoCaixa


*/