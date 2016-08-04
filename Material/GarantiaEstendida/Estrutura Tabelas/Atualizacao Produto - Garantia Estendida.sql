alter table produto add 
pr_GarantiaEstendida varchar(1),
pr_GarantiaFabricante int

ALTER TABLE produto ADD DEFAULT 'N' FOR pr_garantiaEstendida
update produto set pr_garantiaEstendida = 'N', pr_garantiaFabricante = 0 where pr_garantiaEstendida is null

--sp_help produto