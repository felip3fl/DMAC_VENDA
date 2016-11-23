
	exec SP_EST_Movimentacao_Estoque_Anterior_6 'CD','2016/05/01','2016/05/30',0,0
	exec SP_EST_Movimentacao_Estoque_Anterior_5 'CD','2016/06/01','2016/06/30',1,0
	exec SP_EST_Movimentacao_Estoque_Anterior_4 'CD','2016/07/01','2016/07/30',1,0
	exec SP_EST_Movimentacao_Estoque_Anterior_3 'CD','2016/08/01','2016/08/30',1,0
	exec SP_EST_Movimentacao_Estoque_Anterior_2 'CD','2016/09/01','2016/09/30',1,0
	exec SP_EST_Movimentacao_Estoque_Anterior_1 'CD','2016/10/01','2016/10/30',1,0
	exec SP_EST_Movimentacao_Estoque_certo 'CMCE',1,0
	EXEC SP_ATUALIZA_TABELA_ESTOQUE_CDM ''