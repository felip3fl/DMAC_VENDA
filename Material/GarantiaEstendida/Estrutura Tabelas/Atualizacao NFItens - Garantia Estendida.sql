alter table nfitens add 
GarantiaEstendida char(1),
PlanoGarantia int,
CoeficientePlano float,
QtdeGarantia int,
ValorGarantia float,
CertificadoInicio char(12),
CertificadoFim char(12),
ge_premioLiquido float,
ge_IOF float,
ge_dataInicioVigencia datetime,
ge_dataFinalVigencia datetime, 
ge_valorCustoSeguradora float


alter table GarantiaEstendida ADD DEFAULT 'N' FOR garantiaEstendida
alter table PlanoGarantia ADD DEFAULT 0 FOR garantiaEstendida
alter table CoeficientePlano ADD DEFAULT 0 FOR garantiaEstendida
alter table QtdeGarantia ADD DEFAULT 0 FOR garantiaEstendida
alter table ValorGarantia ADD DEFAULT 0 FOR garantiaEstendida
alter table CertificadoInicio ADD DEFAULT '' FOR garantiaEstendida
alter table CertificadoFim ADD DEFAULT '' FOR garantiaEstendida
alter table ge_premioLiquido ADD DEFAULT 0 FOR garantiaEstendida
alter table ge_IOF ADD DEFAULT 0 FOR garantiaEstendida
alter table ge_dataInicioVigencia ADD DEFAULT 0 FOR garantiaEstendida
alter table ge_dataFinalVigencia ADD DEFAULT 0 FOR garantiaEstendida
alter table ge_valorCustoSeguradora ADD DEFAULT 0 FOR garantiaEstendida

update nfitens set garantiaEstendida = 'N' where garantiaEstendida IS NULL



--sp_help nfitens