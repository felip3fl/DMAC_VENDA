select nf,* from nfcapa where dataemi = '2015/03/20' and tiponota = 'S'
select nf,* from nfitens where dataemi = '2015/03/20' and tiponota = 'S'
select * from carimbonotafiscal where cnf_numeroped = 4292

'REMESSA DE TROCA REF NF 25720 E NF 25753   TRANSPORTADORA MINUANO - FRETE POR CONTA DO DESTINO'

update controlesistema set CTS_NumeroNE = CTS_NumeroNE + 1
update controlesistema set CTS_NumeroPedido = CTS_NumeroPedido + 1
select CTS_NumeroNE from controlesistema
select CTS_NumeroPedido from controlesistema

insert into nfcapa (NUMEROPED,DATAEMI,VENDEDOR,VLRMERCADORIA,DESCONTO,SUBTOTAL,LOJAORIGEM,TIPONOTA,CONDPAG,AV,CLIENTE,CODOPER,DATAPAG,PGENTRA,LOJAT,QTDITEM,PEDCLI,TM,PESOBR,PESOLQ,VALFRETE,FRETECOBR,OUTRALOJA,OUTROVEND,NF,TOTALNOTA,NATOPERACAO,DATAPED,BASEICMS,ALIQICMS,VLRICMS,SERIE,HORA,TOTALIPI,ECF,NUMEROSF,NOMCLI,FONECLI,CGCCLI,INSCRICLI,ENDCLI,UFCLIENTE,MUNICIPIOCLI,BAIRROCLI,CEPCLI,PESSOACLI,REGIAOCLI,CFOAUX,AnexoAUx,PAGINANF,ECFNF,Carimbo1,Carimbo2,Carimbo3,Carimbo4,Carimbo5,CustoMedioLiquido,VendaLiquida,MargemContribuicao,ValorTotalCodigoZero,TotalNotaAlternativa,ValorMercadoriaAlternativa,SituacaoEnvio,VendedorLojaVenda,LojaVenda,NotaCredito,NfDevolucao,SerieDevolucao,EmiteDataSaida,CancelarNota,HoraManutencao,DataProcessamento,SituacaoFec,ObsSituacaoFec,CodMunicipioCli,EnderecoNFeCli,EnderecoNroNFeCli,ComplementoNFeCli,InscriSufCli,BaseICMSST,ValorICMSST,ValorCOFINS,ValorOutros,SenhaDesconto,Volume,TipoFrete,ParcelasTEF,AutorizacaoTEF,GarantiaEstendida,TotalGarantia,NroResidencia,CompleResidencia,SeguroPremiado,CertificadoSorte,numeroSorte,sp_premioLiquido,sp_IOF,sp_valorRemuneracao,sp_percentualRemuneracao,sp_valorRepasse,codmun,ChaveNFe,SituacaoProcesso,DataProcesso,Parcelas,NroCaixa,Protocolo,TipoTransporte,Criticaprocesso,CPFNFP,vendedorGarantia,ModalidadeVenda,LiberaBloqueio) 
select '4387','2015/03/23',VENDEDOR,VLRMERCADORIA,DESCONTO,SUBTOTAL,LOJAORIGEM,TIPONOTA,CONDPAG,AV,CLIENTE,CODOPER,DATAPAG,PGENTRA,LOJAT,QTDITEM,PEDCLI,TM,PESOBR,PESOLQ,VALFRETE,FRETECOBR,OUTRALOJA,OUTROVEND,'1208',TOTALNOTA,NATOPERACAO,DATAPED,BASEICMS,ALIQICMS,VLRICMS,SERIE,HORA,TOTALIPI,ECF,NUMEROSF,NOMCLI,FONECLI,CGCCLI,INSCRICLI,ENDCLI,UFCLIENTE,MUNICIPIOCLI,BAIRROCLI,CEPCLI,PESSOACLI,REGIAOCLI,CFOAUX,AnexoAUx,PAGINANF,ECFNF,Carimbo1,Carimbo2,Carimbo3,Carimbo4,Carimbo5,CustoMedioLiquido,VendaLiquida,MargemContribuicao,ValorTotalCodigoZero,TotalNotaAlternativa,ValorMercadoriaAlternativa,SituacaoEnvio,VendedorLojaVenda,LojaVenda,NotaCredito,NfDevolucao,SerieDevolucao,EmiteDataSaida,CancelarNota,HoraManutencao,DataProcessamento,SituacaoFec,ObsSituacaoFec,CodMunicipioCli,EnderecoNFeCli,EnderecoNroNFeCli,ComplementoNFeCli,InscriSufCli,BaseICMSST,ValorICMSST,ValorCOFINS,ValorOutros,SenhaDesconto,Volume,TipoFrete,ParcelasTEF,AutorizacaoTEF,GarantiaEstendida,TotalGarantia,NroResidencia,CompleResidencia,SeguroPremiado,CertificadoSorte,numeroSorte,sp_premioLiquido,sp_IOF,sp_valorRemuneracao,sp_percentualRemuneracao,sp_valorRepasse,codmun,ChaveNFe,SituacaoProcesso,DataProcesso,Parcelas,NroCaixa,Protocolo,TipoTransporte,Criticaprocesso,CPFNFP,vendedorGarantia,ModalidadeVenda,LiberaBloqueio from nfcapa 
where dataemi = '2015/03/20' and tiponota = 'S'

insert into nfitens (NUMEROPED,DATAEMI,REFERENCIA,QTDE,VLUNIT,VLUNIT2,VLTOTITEM,DESCRAT,ICMS,ITEM,VLIPI,DESCONTO,PLISTA,COMISSAO,VALORICMS,BCOMIS,CSPROD,LINHA,SECAO,VBUNIT,ICMPDV,CODBARRA,NF,SERIE,LOJAORIGEM,CLIENTE,VENDEDOR,ALIQIPI,TIPONOTA,REDUCAOICMS,BASEICMS,TIPOMOVIMENTACAO,DETALHEIMPRESSAO,SerieProd1,SerieProd2,CustoMedioLiquido,VendaLiquida,MargemContribuicao,EncargosVendaLiquida,EncargosCustoMedioLiquido,PrecoUnitAlternativa,ValorMercadoriaAlternativa,ReferenciaAlternativa,SituacaoEnvio,DescricaoAlternativa,Tributacao,IcmsMargem,PisCofins,DeducoesVendas,EncargosFinanceiros,EstoqueAntes,EstoqueDepois,CFOP,CSTICMS,GarantiaEstendida,PlanoGarantia,CoeficientePlano,QtdeGarantia,ValorGarantia,CertificadoInicio,CertificadoFim,ge_premioLiquido,ge_IOF,ge_dataInicioVigencia,ge_dataFinalVigencia,ge_valorCustoSeguradora,ge_seqCancelamento,ge_dataCancelamento,SituacaoProcesso,dataprocesso,ICMSAplicado,Parcelas) 
select '4387','2015/03/23',REFERENCIA,QTDE,VLUNIT,VLUNIT2,VLTOTITEM,DESCRAT,ICMS,ITEM,VLIPI,DESCONTO,PLISTA,COMISSAO,VALORICMS,BCOMIS,CSPROD,LINHA,SECAO,VBUNIT,ICMPDV,CODBARRA,'1208',SERIE,LOJAORIGEM,CLIENTE,VENDEDOR,ALIQIPI,TIPONOTA,REDUCAOICMS,BASEICMS,TIPOMOVIMENTACAO,DETALHEIMPRESSAO,SerieProd1,SerieProd2,CustoMedioLiquido,VendaLiquida,MargemContribuicao,EncargosVendaLiquida,EncargosCustoMedioLiquido,PrecoUnitAlternativa,ValorMercadoriaAlternativa,ReferenciaAlternativa,SituacaoEnvio,DescricaoAlternativa,Tributacao,IcmsMargem,PisCofins,DeducoesVendas,EncargosFinanceiros,EstoqueAntes,EstoqueDepois,CFOP,CSTICMS,GarantiaEstendida,PlanoGarantia,CoeficientePlano,QtdeGarantia,ValorGarantia,CertificadoInicio,CertificadoFim,ge_premioLiquido,ge_IOF,ge_dataInicioVigencia,ge_dataFinalVigencia,ge_valorCustoSeguradora,ge_seqCancelamento,ge_dataCancelamento,SituacaoProcesso,dataprocesso,ICMSAplicado,Parcelas 
from nfitens 
where dataemi = '2015/03/20' and tiponota = 'S'

update NFItens set CSTICMS = '20' where dataemi = '2015/03/23' and numeroped = '4387'

SELECT * FROM NFItens WHERE numeroped = '4387'
SELECT nf,* FROM NFCAPA WHERE tiponota = 'S' and dataemi = '2015/03/23'
SELECT nf,* FROM nfitens WHERE tiponota = 'S' and dataemi = '2015/03/23'
--update nfcapa set totalnota = 753.1, vlrmercadoria = 753.1, subtotal = 753.1 WHERE tiponota = 'S' and dataemi = '2015/03/23' 
--update nfitens set vltotitem = 616,vlunit2 = 616 WHERE tiponota = 'S' and dataemi = '2015/03/23' and REFERENCIA = '7330070'

exec SP_vda_cria_nfe '353','1208','NE',''


select * from NFItens WHERE DATAEMI = '2015/03/23' AND ITEM > 1