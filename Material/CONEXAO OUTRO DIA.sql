--Select * from NFCapa as L where L.LojaOrigem = (select top 1 cts_loja from ControleSistema) and L.dataemi = '2015/06/29' and tiponota not in('PA','PD','TA')
--                          and situacaoProcesso = 'A' and L.NF is not null and  not exists
--                         (select * from svdmac.dmac.dbo.NFCapa as C where C.NF = L.NF and C.Serie = L.Serie 
--                          and c.lojaorigem=l.lojaorigem)


update NFCapa set DataProcesso = '2015/06/30' from NFCapa as L where L.LojaOrigem = (select top 1 cts_loja from ControleSistema) and L.dataemi = '2015/06/29' and tiponota not in('PA','PD','TA')
                          and situacaoProcesso = 'A' and L.NF is not null and  not exists
                         (select * from svdmac.dmac.dbo.NFCapa as C where C.NF = L.NF and C.Serie = L.Serie 
                          and c.lojaorigem=l.lojaorigem)



 --Select * from NFItens as L where L.LojaOrigem = (select top 1 cts_loja from ControleSistema) and L.dataemi ='2015/06/29' and tiponota not in('PA','PD','PD','TA')
 --                         and situacaoProcesso =  'A'  and L.NF is not null and  not exists
 --                        (select * from svdmac.dmac.dbo.NFItens as C where C.NF = L.NF and C.Serie = L.Serie 
 --                         and c.lojaorigem=l.lojaorigem)

						  
update NFItens set DataProcesso = '2015/06/30' from NFItens as L where L.LojaOrigem = (select top 1 cts_loja from ControleSistema) and L.dataemi ='2015/06/29' and tiponota not in('PA','PD','PD','TA')
                          and situacaoProcesso =  'A'  and L.NF is not null and  not exists
                         (select * from svdmac.dmac.dbo.NFItens as C where C.NF = L.NF and C.Serie = L.Serie 
                          and c.lojaorigem=l.lojaorigem)

exec SP_Atualiza_Processos_Venda_Central