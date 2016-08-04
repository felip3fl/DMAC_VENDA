select cliente,UFCLIENTE,* from nfcapa where dataemi > = '2014/07/17' and tiponota IN 'V' and SERIE IN ('ne','CF') and LOJAORIGEM = '28'

select vc_ufcliente,* from capanfvenda where vc_lojaorigem = '28' and vc_dataemissao > = '2014/07/17' and vc_serie IN ('ne','CF') and VC_Tiponota in ('V','T','c')
update capanfvenda set vc_cgclojadestino = 0 where vc_lojaorigem = '28' and vc_dataemissao > = '2014/07/17' and vc_serie IN ('ne','CF') and VC_Tiponota in ('V','T') and vc_cgclojadestino is null
/*

update nfcapa 
set NOMCLI = CE_Razao,
FONECLI = CE_Telefone,
CGCCLI = CE_CGC,
INSCRICLI = CE_InscricaoEstadual,
ENDCLI = CE_Endereco,
UFCLIENTE = CE_Estado,
MUNICIPIOCLI = CE_Municipio,
BAIRROCLI = CE_Bairro,
CEPCLI = CE_CEP, 
codmun = '',
CompleResidencia = '',
NroResidencia = CE_Numero
from nfcapa,fin_cliente_28_TEMP where 
dataemi > = '2014/07/17' 
and tiponota in ('V','C')
and SERIE IN ('NE','CF')
and LOJAORIGEM = '28'
and cliente = CE_CodigoCliente

*/

/*

update capanfvenda 
set vc_nomecliente = NOMCLI,
vc_telefonecliente = FONECLI,
vc_cgccliente = CGCCLI,
vc_inscEstCliente = INSCRICLI,
vc_enderecoCliente = ENDCLI,
vc_ufcliente =  UFCLIENTE,
vc_municipioCliente = MUNICIPIOCLI,
vc_bairroCliente = BAIRROCLI,
vC_cepcliente = CEPCLI, 
vc_codigoMunicipio = 0,
vc_compleResidencia =  CompleResidencia,
vc_nroResidencia = NroResidencia
from nfcapa, capanfvenda where 
dataemi  > '2015/06/07' 
and LOJAORIGEM = VC_LojaOrigem
and nf = vc_notafiscal
and serie = vc_serie
and dataemi = vc_dataemissao
and VC_NomeCliente IS NULL

*/

/*
select * from loja

update nfcapa 
set NOMCLI = lo_nomeloja,
FONECLI = lo_telefone,
CGCCLI = lo_cgc,
INSCRICLI = lo_inscricaoEstadual,
ENDCLI = lo_endereco,
UFCLIENTE = lo_uf,
MUNICIPIOCLI = lo_municipio,
BAIRROCLI = lo_bairro,
CEPCLI = lo_cep, 
codmun = lo_codigoMunicipio,
CompleResidencia = '',
NroResidencia = lo_endereconronfe
from nfcapa,loja where 
dataemi > = '2014/07/17' 
and tiponota IN ('T','C')
and SERIE IN ('NE','CF')
and LOJAORIGEM = '28'
and cliente = lo_loja
and lo_gruporegiao <  80

*/

--UPDATE capanfvenda 