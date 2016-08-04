select VI_NOTAFISCAL, vi_serie, vi_dataemissao from itemnfvenda where vi_lojaorigem = '28' and vi_dataemissao > '2014/07/01'  and Vi_tiponota = ('V') and vi_desconto < 0 
select vc_notafiscal, VC_serie, VC_VALORMERCADORIAS, vc_desconto, VC_TOTALNOTA,((vc_desconto / VC_VALORMERCADORIAS) * 100) as porcentagem from capanfvenda, itemnfvenda where vC_lojaorigem = '28' and vC_dataemissao > '2014/07/01' and vc_serie = vi_serie and vc_notafiscal = vi_notafiscal and vi_lojaorigem = '28' and vi_lojaorigem = vc_lojaorigem and vi_dataemissao > '2014/07/01'  and Vi_tiponota = ('V') and vi_desconto < 0 

7,6271186440678

select vi_desconto,* from itemnfvenda where vi_SERIE = 'CF' and vi_notafiscal = '67'


update itemNFVENDA SET vi_desconto = desconto from nfitens where nf = vi_notafiscal and serie = vi_serie and dataemi = vi_dataemissao and vi_referencia = referencia and vi_desconto < 0 and vi_dataemissao between '2014/07/18' and '2014/07/23'

select * from capanfvenda where vc_SERIE = 'CF' and vc_notafiscal = '67'

select  * from nfcapa where serie = 'CF' and NF = '67'

197 = 2370


select 