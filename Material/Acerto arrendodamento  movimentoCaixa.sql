select * from NFCapa where TIPONOTA = 'T' and DATAEMI = '2015/04/06'
select * from MovimentoCaixa where MC_Data = '2015/04/06' and MC_TipoNota = 'TA'

select subtotal,TOTALNOTA,MC_Valor, NF,* from MovimentoCaixa, nfcapa where MC_Data = '2015/04/06' and MC_TipoNota = 'TA' and MC_Documento = NF 
and TIPONOTA = 'T' and DATAEMI = '2015/04/06' and MC_Valor <> SUBTOTAL ORDER by MC_Documento



UPDATE MOVIMENTOCAIXA SET MC_VALOR = TOTALNOTA from MovimentoCaixa, nfcapa where MC_Data = '2015/04/06' and MC_TipoNota = 'TA' and MC_Documento = NF 
and TIPONOTA = 'T' and DATAEMI = '2015/04/06' and MC_Valor <> TOTALNOTA 

update MovimentoCaixa 