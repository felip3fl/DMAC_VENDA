select codoper,cfoaux, * from NFCapa where DATAEMI = '2015/01/14' and lojaorigem in ('28','353')
--update NFCapa set cfoaux = codoper where DATAEMI = '2015/01/14' and lojaorigem in ('28','353')
select  tipomovimentacao, * from nfitens where DATAEMI = '2015/01/14' and lojaorigem in ('28','353')
--update nfitens set tipomovimentacao = '11' where DATAEMI = '2015/01/14' and lojaorigem in ('28','353') and tiponota = 'V'
--update nfitens set tipomovimentacao = '12' where DATAEMI = '2015/01/14' and lojaorigem in ('28','353') and tiponota = 'T'

update nfitens set descrat = 0, vlipi = 0, comissao = 0, bcomis = 0, csprod = 0, linha = 0, secao = 0, vbunit = 0, icmpdv = 0,
codbarra = '', cliente = 0, vendedor = 0, reducaoICMS = 0, serieProd1 = '', serieProd2 = '', custoMedioLiquido = 0,
vendaLiquida = 0, margemContribuicao = 0, encargosVendaLiquida = 0, encargosCustoMedioLiquido = 0, PrecoUnitAlternativa = 0,
ValorMercadoriaAlternativa = 0, referenciaAlternativa = 0, situacaoEnvio = 'A', descricaoAlternativa = '', tributacao = '',
icmsMargem = 0, piscofins = 0, deducoesVendas = 0, encargosFinanceiros = 0, estoqueAntes = 0, estoqueDepois = 0, parcelas = 0
 where lojaorigem = '353' and dataemi = '2015/01/15' and vlipi is null
 
 update lojat = 
 
 update nfcapa set natoperacao = 0, datapag = '1900-01-01', pedcli = 0, aliqicms = 0, totalipi = 0, 
 numerosf = 0, pessoacli = 0, regiaocli = 0, anexoAUX = '', carimbo1 = '', carimbo2 = '',carimbo3 = '',carimbo4 = '', carimbo5 = '',
 custoMedioLiquido = 0, vendaLiquida = 0, margemContribuicao = 0, valorTotalCodigoZero = 0, 
 valorMercadoriaaLTERNATIVA = 0, situacaoEnvio = 'A', notacredito = 0, nfdevolucao = 0, serieDevolucao = '', emiteDataSaida = '', cancelarNota = 'N', 
 dataProcessamento = '2015/01/14', situacaofec = 'A', OBSSITUACAOFEC = '', codMunicipioCli = '', endereconfecli = '', inscrisufcli = '', baseICMSST = 0, 
 valorICMSST = 0, valorOutros = 0, senhadesconto = '', parcelasTEF = 0, autorizacaoTEF = '', ecf = '1', paginanf = 1
  where lojaorigem = '353' and dataemi = '2015/01/15' and av is null