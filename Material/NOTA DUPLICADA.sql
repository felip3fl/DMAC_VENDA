select * from nfe_nflojas where nfl_nronfe = '2276' group by NFL_Sequencia,NFL_Descricao,NFL_Dados,NFL_Loja,NFL_NroNFE,NFL_DataEmissao order by NFL_Sequencia
select * into temp_nfe_nflojas from nfe_nflojas where nfl_nronfe = '2276' group by NFL_Sequencia,NFL_Descricao,NFL_Dados,NFL_Loja,NFL_NroNFE,NFL_DataEmissao order by NFL_Sequencia
delete nfe_nflojas where nfl_nronfe = '2276'
insert into NFE_NFLojas select * from temp_nfe_nflojas
drop table temp_nfe_nflojas


--sp_help nfe_nflojas