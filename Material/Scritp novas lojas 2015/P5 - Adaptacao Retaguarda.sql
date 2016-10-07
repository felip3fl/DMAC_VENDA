use DMAC_LOJA

-- DMAC

exec svdmac.dmac.dbo.SP_Atualiza_Produto_Loja '181'
exec svdmac.dmac.dbo.SP_Atualiza_Produto_Barras_Loja '181'

delete svdmac.dmac.dbo.nfcapa where lojaorigem = ''
delete svdmac.dmac.dbo.nfitens where lojaorigem = ''
delete svdmac.dmac.dbo.capanfvenda where lojaorigem = ''
delete svdmac.dmac.dbo.itemnfvenda where lojaorigem = ''
delete svdmac.dmac.dbo.movimentocaixa where mC_loja = ''


/*
ALTERAR A PROCEDURE
SP_VDA_Conexao_DEMEOSERVER

Adicionar +1 loja na linha:
set @lojas = '(''28'',''353'',''354'')'

*/