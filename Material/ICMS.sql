select pr_substituicaoTributaria, pr_codigoReducaoICMS, PR_ICMSSAIDA, PR_ICMPDV, * from produto where pr_referencia = '0600365'
update produto set pr_codigoReducaoICMS = 0 where  pr_referencia = '0600365'



exec sp_tv '181','0617943','B'

select pr_substituicaotributaria,pr_codigoReducaoICMS, * from produto where pr_referencia = '7318546'

--update produto set pr_codigoReducaoICMS = 0 where pr_referencia = '7318548'




update produtoloja set pr_st = '060' where PR_SubstituicaoTributaria = 's' and pr_st is null
update produtoloja set pr_st = '020' where PR_CodigoReducaoICMS > 0 and pr_st is null
update produtoloja set pr_st = '000' where PR_CodigoReducaoICMS = 0 and PR_SubstituicaoTributaria = 'N' and pr_st is null