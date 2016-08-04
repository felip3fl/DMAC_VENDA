select substring(CE_CGC,2,14),* from fin_cliente where substring(CE_CGC,1,1) = '0' and len(RTRIM(CE_CGC)) = 15
update fin_cliente set cE_cgc = substring(CE_CGC,2,14) where substring(CE_CGC,1,1) = '0' and len(RTRIM(CE_CGC)) = 15
select * from fin_cliente where CE_CodigoCliente =  '55555'
select * from fin_cliente where CE_Razao = 'CASSARINO DECORACOES DE CINEMAS LTDA'

select replace(CE_InscricaoEstadual,'.',''), * from fin_cliente where CE_InscricaoEstadual like '%.%'
update fin_cliente set CE_InscricaoEstadual = replace(CE_InscricaoEstadual,'.','') where CE_InscricaoEstadual like '%.%'

select CE_Telefone,replace(CE_Telefone,'-',''),* from fin_cliente where CE_Telefone like '%-%'
update fin_cliente set CE_Telefone = replace(CE_Telefone,'-','') where CE_Telefone like '%-%'
update fin_cliente set CE_Telefone = replace(CE_Telefone,' ','') where CE_Telefone like '% %'

update fin_cliente set CE_Celular = replace(CE_Celular,'-','') where CE_Celular like '%-%'
update fin_cliente set CE_Celular = replace(CE_Celular,' ','') where CE_Celular like '% %'

select * from FIN_Municipio where Mun_Nome = 'OSASCO'
select * from FIN_Cliente where CE_CodigoMunicipio is null

select Mun_Codigo, Mun_Nome, CE_Municipio from FIN_Municipio, FIN_Cliente where CE_CodigoMunicipio is null and CE_Municipio = Mun_Nome
update FIN_Cliente set CE_CodigoMunicipio = Mun_Codigo from FIN_Municipio, FIN_Cliente where CE_CodigoMunicipio is null and rtrim(CE_Municipio) = Mun_Nome
update FIN_Cliente set CE_CodigoMunicipio = '3550308'  where CE_Municipio = 'SAO PAULO'
