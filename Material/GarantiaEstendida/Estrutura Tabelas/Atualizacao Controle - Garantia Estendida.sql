alter table controle add
CT_Certificado numeric(9), 
CT_sequenciaArquivoGE int,
CT_IdentificacaoRamo char(4),
CT_codigoEstipulanteGE char(4),
CT_Apolice char(30)

alter table controle add 
CT_VlrSegPremiado float,
CT_PremioControle float,
CT_Cobertura float,
CT_CoberturaPorExtenso char(20),
CT_NumeroSorte int
CT_CoeficientePremio float,

update controle set 
ct_certificado = 0,
CT_sequenciaArquivoGE = 0,
CT_IdentificacaoRamo = '0095',
CT_codigoEstipulanteGE = '0205', 
ct_apolice = '06530.2013.01.0095.000091'

update controle set 
ct_vlrSegPremiado = 0,
CT_CoeficientePremio = 1.0378,
ct_numeroSorte = 0
ct_premioControle = '0', 
ct_cobertura = 0,
ct_coberturaPorExtenso = 'zero'

alter table controle add
CT_codigoInternoProdutoGE char(2),
CT_codigoEstipulanteGE char(4),

update controle set 
CT_codigoEstipulanteGE = '0205',
CT_codigoInternoProdutoGE = '01'

--sp_help controle
--select * from controle