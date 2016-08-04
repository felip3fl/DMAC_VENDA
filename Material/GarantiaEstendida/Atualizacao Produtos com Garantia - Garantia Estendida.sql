/*
select top 100 * from produtoComGarantia

update produto set 
pr_garantiaEstendida = d.pr_garantiaEstendida,
pr_garantiaFabricante = d.pr_garantiaFabricante 
from produto as l, produtoComGarantia as d 
where l.pr_referencia = d.pr_referencia 
collate sql_latin1_general_cp1_ci_as

drop table produtoComGarantia

select top 100 * from produto where pr_garantiaEstendida = 'S'

*/


/*

update [demeoServer].[loja181].[dbo].produto set 
pr_garantiaEstendida = d.pr_garantiaEstendida,
pr_garantiaFabricante = d.pr_garantiaFabricante 
from [demeoServer].[loja181].[dbo].produto as l, produtoComGarantia as d 
where l.pr_referencia = d.pr_referencia 
collate sql_latin1_general_cp1_ci_as

*/