
select * from movimentocaixa where mc_data = '2014/09/06' and mc_documento = 241041 and mc_valor = 35.81

--update movimentocaixa set mc_valor = 35.81 where mc_documento = 241041 and mc_data = '2014/09/06' and mc_grupo = 11003

select nf,* from nfcapa where dataemi = '2014/09/06' and tiponota = 'V'

select * from movimentocaixa where mc_data = '2014/09/06' and mc_documento = 241039

select * from nfcapa where dataemi = '2014/09/06' and tiponota = 'V'
SELECT * FROM movimentocaixa where mc_data = '2014/09/06' and mc_documento in (select nf from nfcapa where dataemi = '2014/09/06' and tiponota = 'V') and mc_grupo = 20102 order by mc_documento, mc_grupo

SELECT mc_valor, totalnota, * FROM movimentocaixa, nfcapa where mc_data = '2014/09/06' 
and mc_documento = nf and dataemi = '2014/09/06' and tiponota = 'v'
and mc_documento in (select nf from nfcapa where dataemi = '2014/09/06' and tiponota = 'V') 
and mc_grupo = 20102 order by mc_documento, mc_grupo

update movimentocaixa set mc_valor = 254 where mc_documento = 241044 and mc_grupo = 20102 
--289,81

Select * from Modalidade where MO_OrdemApresentacao<>0 and MO_Grupo not in (30105,20104,20105,20106,20203,20204) order by MO_OrdemApresentacao

select totalnota,* from nfcapa where nf = 241041