
SELECT  * FROM NFITENS WHERE NF = '12254' and SERIE = 'NE'
SELECT SUBTOTAL, VLRMERCADORIA, TOTALNOTA,  * FROM NFCAPA WHERE NF = '1183' and SERIE = 'NE'
select round(SUM(VLTOTITEM),2) FROM NFITENS WHERE NF = '12254' and SERIE = 'NE' and lojaorigem = 'CD'

/*
update NFITENS set VLUNIT = 29.98 WHERE NF = 12254  and SERIE = 'NE' and lojaorigem = 'CD'
update NFITENS set VLTOTITEM = round((VLUNIT * qtde),2) WHERE NF = 12254 and SERIE = 'NE' and lojaorigem = 'CD'
update NFITENS set VLUNIT2 = round((VLUNIT * qtde),2) WHERE NF = 12254 and SERIE = 'NE' and lojaorigem = 'CD'
*/

/*
UPDATE NFCAPA SET SUBTOTAL = (select round(SUM(VLTOTITEM),2) FROM NFITENS WHERE NF = '12254' and SERIE = 'NE' and lojaorigem = 'CD'), 
VLRMERCADORIA= (select round(SUM(VLTOTITEM),2) FROM NFITENS WHERE NF = '12254' and SERIE = 'NE' and lojaorigem = 'CD'), 
TOTALNOTA = (select round(SUM(VLTOTITEM),2) FROM NFITENS WHERE NF = '12254' and SERIE = 'NE' and lojaorigem = 'CD') 
WHERE NF = '12254' and SERIE = 'NE' and lojaorigem = 'CD'

UPDATE NFCAPA SET SUBTOTAL = 599.82, VLRMERCADORIA= 599.82, TOTALNOTA = 599.82 WHERE NF = '12254'  and SERIE = 'NE' and lojaorigem = 'CD'

*/
 


 update 


 