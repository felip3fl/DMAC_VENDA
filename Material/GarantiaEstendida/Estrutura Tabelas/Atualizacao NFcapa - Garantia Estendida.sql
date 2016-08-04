alter table nfcapa add 
GarantiaEstendida char(1), 
TotalGarantia float

alter table nfcapa ADD DEFAULT 'N' FOR GarantiaEstendida
alter table nfcapa ADD DEFAULT 0 FOR TotalGarantia
update nfcapa set garantiaEstendida = 'N' where garantiaEstendida IS NULL


alter table nfcapa add 
vendedorGarantia numeric
alter table nfcapa ADD DEFAULT 0 FOR vendedorGarantia
update nfcapa set vendedorGarantia = 0 where vendedorGarantia IS NULL

--sp_help nfcapa


alter table nfcapa add 
SeguroPremiado float, 
CertificadoSorte numeric, 
numeroSorte numeric, 
sp_premioLiquido float, 
sp_IOF float, 
sp_valorRemuneracao float, 
sp_percentualRemuneracao float, 
sp_valorRepasse float