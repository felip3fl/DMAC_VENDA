Select LO_Razao, * from lojas where lo_loja = '333'
Select LO_NomeLoja,LO_DDD,lo_Fax, * from loja where lo_loja = '333'

SELECT * FROM CONTROLESISTEMA
Select CTS_Loja,Loja.* from loja,Controlesistema where lo_loja= CTS_Loja

/*

alter table loja add LO_Razao char(30)
update loja set LO_Razao = 'CONTINENTAL FERRAMENTAS LTDA' where lo_loja = '333'

alter table loja add lo_numero char(10)
update loja set lo_numero = '333' where lo_loja = '333'

update loja set LO_DDD = '11',lo_Fax = '00000000' where lo_loja = '333'

select * from vende where ve_nome = 'gerente'

select * into fornecedor from [svcentralfer].[dmac].[dbo].fornecedor

update produtoloja set PR_GarantiaEstendida = 'N', PR_GarantiaFabricante = 0, PR_custoMedioLiquido1 = 0, PR_precoVendaLiquido1 = 0, PR_IndicePreco = 1

select CP_TipoCondicao,* from CondicaoPagamento
alter table condicaoPagamento add CP_id int
alter table condicaoPagamento add CP_TipoCondicao nvarchar(4)
update CondicaoPagamento set cp_id = 1
update CondicaoPagamento set CP_TipoCondicao = cp_codigo

SP_Atualiza_Modalidade_Venda_Pedido

*/

SELECT * FROM NFCAPA





select LO_TipoTransferencia, * from loja

alter table loja add LO_TipoTransferencia char(1)
update loja set LO_TipoTransferencia = 'T'

select * from NfCapa where numeroped = 82254


SELECT * FROM UsuarioCaixa

insert into UsuarioCaixa values('99',	'fecger',                   	'F',	'123456',	'F')