select  * from capanfvenda WHERE vc_lojaorigem = '28' and vc_cliente = '999999' and vc_dataemissao > '2014/08/01' and vc_tiponota = 'v'
select * from nfcapa where DATAEMI > '2014/08/01' and LOJAORIGEM = '28' AND CPFNFP <> '' and cliente = '999999' and tiponota = 'V'

/*
update capanfvenda set vc_cgccliente = cgccli, vc_cgclojadestino = cgccli 
from nfcapa, capanfvenda
where vc_cgccliente = '000000000000000'
and DATAEMI >= '2014/08/01' 
and LOJAORIGEM = '28' 
and CPFNFP <> '' 
and cliente = '999999' 
and tiponota = 'V'
and vc_lojaorigem = '28' 
and vc_cliente = '999999' 
and vc_dataemissao >= '2014/09/01' 
and vc_tiponota = 'v'
and vc_notafiscal = nf
and vc_lojaorigem = lojaorigem
and vc_serie = serie
and vc_dataemissao = dataemi


UPDATE capanfvenda set vc_cgccliente = '000000000000000', 
vc_cgclojadestino = '000000000000000' 
WHERE vc_lojaorigem = '28' 
and vc_cliente = '999999' 
and vc_dataemissao >= '2014/08/01' 
and vc_tiponota = 'v' 
and vc_cgclojadestino = ''

*/
