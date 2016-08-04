
select * from nfcapa where nf = 2786 and serie = 'NE'
select * from MovimentoCaixaTEMP where MC_Documento = 2626 and MC_Serie = 'NE'

DROP TABLE nfcapaTEMP
DROP TABLE nfitensTEMP
DROP TABLE CarimboNotaFiscalTEMP
DROP TABLE MovimentoCaixaTEMP

select * into nfcapaTEMP from nfcapa where nf = 2626 and serie = 'NE'
select * into nfitensTEMP from nfitens where  nf = 2626 and serie = 'NE'
select * into CarimboNotaFiscalTEMP from CarimboNotaFiscal where  cnf_nf = 2626 and cnf_serie = 'NE'
select * into MovimentoCaixaTEMP from MovimentoCaixa where MC_Documento = 2626 and MC_Serie = 'NE'


select CTS_NumeroPedido, * from controlesistema
update controlesistema set CTS_NumeroPedido = CTS_NumeroPedido + 1

select CTS_NumeroNE, * from controlesistema
update controlesistema set CTS_NumeroNE = CTS_NumeroNE + 1

SELECT top 10 * FROM ControleCaixa where CTR_Operador <> 99 order by CTR_DataInicial desc

update nfcapaTEMP set nf = '2787', DataEmi = '2015/07/16',DataProcesso = '2015/07/16',NUMEROPED='10806'--, TipoNota = 'R'--,  NotaFiscal = 0, CodigoOperacao = '2949', CodigoOperacaonovo = '2949'
update nfitensTEMP set nf = '2787', DataEmi = '2015/07/16',DataProcesso = '2015/07/16',NUMEROPED='10806'--, TipoNota = 'R'--, DataEmissao = '2015/06/30', NotaFiscal = 0, cfop = '2949'

update CarimboNotaFiscalTEMP set cnf_nf = '2787', CNF_data = '2015/07/16',CNF_DataProcesso = '2015/07/15',CNF_NumeroPed='10806'
update MovimentoCaixaTEMP set MC_Documento = '2787', MC_Data = '2015/07/15',MC_DataProcesso = '2015/07/15',MC_Pedido='10806', MC_Protocolo = 161


select nf, * from nfcapaTEMP
select nf,* from nfitensTEMP
select * from CarimboNotaFiscalTEMP

insert into nfcapa select * from nfcapaTEMP
insert into NFItens select * from nfitensTEMP
insert into CarimboNotaFiscal select * from CarimboNotaFiscalTEMP
insert into MovimentoCaixa(MC_NumeroECF,MC_CodigoOperador,MC_Loja,MC_Data,MC_Grupo,MC_SubGrupo,MC_Documento,MC_Serie,MC_Valor,MC_Banco,MC_Agencia,MC_ContaCorrente,MC_NumeroCheque,MC_BomPara,MC_Parcelas,MC_Remessa,MC_SituacaoEnvio,MC_ControleAVR,MC_DataBaixaAVR,MC_Protocolo,MC_NroCaixa,MC_GrupoAuxiliar,MC_Situacao,MC_Pedido,MC_DataProcesso,MC_TipoNota,MC_SequenciaTEF) select MC_NumeroECF,MC_CodigoOperador,MC_Loja,MC_Data,MC_Grupo,MC_SubGrupo,MC_Documento,MC_Serie,MC_Valor,MC_Banco,MC_Agencia,MC_ContaCorrente,MC_NumeroCheque,MC_BomPara,MC_Parcelas,MC_Remessa,MC_SituacaoEnvio,MC_ControleAVR,MC_DataBaixaAVR,MC_Protocolo,MC_NroCaixa,MC_GrupoAuxiliar,MC_Situacao,MC_Pedido,MC_DataProcesso,MC_TipoNota,MC_SequenciaTEF from MovimentoCaixaTEMP

SELECT * FROM CarimboNotaFiscal WHERE cnf_nf = 2830
UPDATE CarimboNotaFiscal SET cnf_nf =2783 WHERE cnf_nf = 2830 AND CNF_NumeroPed = 10791
UPDATE nfitens SET NF =2783 WHERE NF = 2830 AND numeroped = 10791

exec sp_vda_cria_nfe '181','2783','NE',''

DROP TABLE nfcapaTEMP
DROP TABLE nfitensTEMP
DROP TABLE CarimboNotaFiscalTEMP
DROP TABLE MovimentoCaixaTEMP

/*
select * from capanfvenda2015
select * from ITEMNFVENDA2015

update capanfvenda2015 set VC_Serie = 'S3', VC_TipoNota = 'R', VC_DataEmissao = '2015/06/30', VC_NotaFiscal = 0, VC_CodigoOperacao = '2949', VC_CodigoOperacaonovo = '2949'
update itemnfvenda2015 set vi_Serie = 'S3', vi_TipoNota = 'R', vi_DataEmissao = '2015/06/30', vi_NotaFiscal = 0, Vi_cfop = '2949'




*/

