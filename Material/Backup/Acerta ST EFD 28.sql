select * from itemnfvenda 
where vi_lojaorigem = '28' 
and vi_serie = 'CF'
and vi_cst = '060'
and vi_dataEmissao >= '2014/08/01'
and vi_dataEmissao <= '2014/08/30'
AND VI_TIPONOTA = 'V'




select * from capanfvenda where vc_notafiscal = 463 and vc_lojaorigem = '28' and vc_serie = 'CF'
select * from itemnfvenda where vi_notafiscal = 463 and vi_lojaorigem = '28' and vi_serie = 'CF'


select * from itemnfvenda 
where vi_lojaorigem = '28' 
and vi_serie = 'CF'
and vi_cst = '060'
and vi_dataEmissao >= '2014/08/01'
and vi_dataEmissao <= '2014/08/30'
AND VI_TIPONOTA = 'V'

/*

update itemnfvenda 
set vi_aliquotaICMS = 0
where vi_lojaorigem = '28' 
and vi_serie = 'CF'
and vi_cst = '060'
and vi_dataEmissao >= '2014/08/01'
and vi_dataEmissao <= '2014/08/30'
AND VI_TIPONOTA = 'V'

*/