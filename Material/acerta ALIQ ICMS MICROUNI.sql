select * from cfosaidcad as c,nfsaidacad as m  where c.valicm > 0 and c.alqicm = 0 and c.numord = m.numord and m.dtemis >= '2014/10/01' and m.filial = '02'

select * from cfosaidcad as c,nfsaidacad as m  where c.valicm > 0 and c.alqicm = 0 and c.numord = m.numord and m.dtemis >= '2014/10/01' and m.filial = '02' 
and c.valicm2 = 0 and c.valicm3 = 0 and c.valicm4 = 0  and c.valicm5 = 0  

select top 5 * from nfsaidacad

select * from cfosaidcad where numnota = '002667' and serie = 'NE' 



begin transaction

update cfosaidcad set alqicm = vi_aliquotaICMS 
FROM demeo..cfopAcertoTEMP , cfosaidcad as c, nfsaidacad as m WHERE 
vi_notafiscal = c.numnota collate sql_latin1_general_cp1_ci_as
and vi_serie = c.serie collate sql_latin1_general_cp1_ci_as
and vi_cfop = c.cfo collate sql_latin1_general_cp1_ci_as 
and c.valicm > 0 
and c.alqicm = 0 
and c.numord = m.numord 
and m.dtemis >= '2014/10/01' 
and m.filial = '02'

commit transaction
rollback

select * from demeo..cfopAcertoTEMP

--update cfosaidcad set alqicm = 18 where  numnota = '002667' and serie = 'NE' and cfo = '5152'

*/


select * from cfosaidcad where numnota = '002667' and serie = 'NE' and cfo = '5152'


select vi_aliquotaICMS, * from demeo..itemnfvenda where vi_notafiscal = 2667 and vi_lojaorigem = '28' and vi_serie = 'ne'

select vi_aliquotaICMS, * from demeo..itemnfvenda where vi_notafiscal in (002169,002172,002173,002175) and vi_lojaorigem = '28' and vi_serie = 'ne'

select DISTINCT vi_notafiscal,vi_cfop,vi_aliquotaICMS from demeo..itemnfvenda where 
and vi_lojaorigem = '28' and vi_serie = 'ne' and vi_aliquotaICMS > 0

select * from demeo..cfopAcertoTEMP
DROP TABLE demeo..cfopAcertoTEMP

select vi_notafiscal,vi_serie,vi_cfop,vi_aliquotaICMS into demeo..cfopAcertoTEMP from demeo..itemnfvenda

INSERT INTO demeo..cfopAcertoTEMP select DISTINCT vi_notafiscal,vi_serie,vi_cfop,vi_aliquotaICMS from demeo..itemnfvenda where vi_dataemissao >= '2014/10/01'
and vi_lojaorigem = '28' and vi_serie = 'CF' and vi_aliquotaICMS > 0 and vi_aliquotaICMS <> 8.8 and vi_cfop <> 5409 and vi_cfop <> ''


select vi_notafiscal,vi_cfop,vi_aliquotaICMS, pr_icmssaida, pr_referencia from demeo..itemnfvenda, demeo..produto where vi_dataemissao >= '2014/10/01'
and vi_lojaorigem = '28' and vi_serie = 'ne' and vi_aliquotaICMS > 0 and pr_referencia = vi_referencia and vi_aliquotaICMS <> 8.8 and vi_cfop <> 5409 and vi_cfop = ''


