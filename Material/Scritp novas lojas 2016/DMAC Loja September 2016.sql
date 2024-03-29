USE [DMAC_LOJA]
GO
/****** Object:  UserDefinedFunction [dbo].[SP_GLB_Valida_Email]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

CREATE FUNCTION [dbo].[SP_GLB_Valida_Email] (
 @EmailAddr varchar(255) -- Email address to check
)   RETURNS BIT -- 1 if @EmailAddr is a valid email address
/*
* Checks an text string to be sure it's a valid e-mail address.
* Returns 1 when it is, otherwise 0.
* Example:
SELECT CASE WHEN 1=dbo.udf_Txt_IsEmail('anovick@NovickSoftware.com')
    THEN 'Is an e-mail address' ELSE 'Not an e-mail address' END
*
* Test:
print case when 1=dbo.udf_txt_isEmail('anovick@novicksoftware.com')
       then 'Passes' else 'Fails' end + ' test for good addr'
print case when 0=dbo.udf_txt_isEmail('@novicksoftware.com')
       then 'Passes' else 'Fails' end + ' test for no user'
print case when 0=dbo.udf_txt_isEmail('anovick@n.com')
       then 'Passes' else 'Fails' end + ' test for 1 char domain'
print case when 1=dbo.udf_txt_isEmail('anovick@no.com')
       then 'Passes' else 'Fails' end + ' test for 2 char domain'
print case when 0=dbo.udf_txt_isEmail('anovick@.com')
       then 'Passes' else 'Fails' end + ' test for no domain'
print case when 0=dbo.udf_txt_isEmail('anov ick@novicksoftware.com')
       then 'Passes' else 'Fails' end + ' test for space in name'
print case when 0=dbo.udf_txt_isEmail('ano#vick@novicksoftware.com')
       then 'Passes' else 'Fails' end + ' test for # in user'
print case when 0=dbo.udf_txt_isEmail('anovick@novick*software.com')
       then 'Passes' else 'Fails' end + ' test for * asterisk in domain'
****************************************************************/
AS BEGIN
DECLARE @AlphabetPlus VARCHAR(255)
      , @Max INT -- Length of the address
      , @Pos INT -- Position in @EmailAddr
      , @OK BIT  -- Is @EmailAddr OK
-- Check basic conditions
IF @EmailAddr IS NULL 
   OR NOT @EmailAddr LIKE '_%@__%.__%' 
   OR CHARINDEX(' ',LTRIM(RTRIM(@EmailAddr))) > 0
       RETURN(0)
SELECT @AlphabetPlus = 'abcdefghijklmnopqrstuvwxyz01234567890_-.@'
     , @Max = LEN(@EmailAddr)
     , @Pos = 0
     , @OK = 1
WHILE @Pos < @Max AND @OK = 1 BEGIN
    SET @Pos = @Pos + 1
    IF NOT @AlphabetPlus LIKE '%' 
                             + SUBSTRING(@EmailAddr, @Pos, 1) 
                             + '%' 
        SET @OK = 0
END -- WHILE
RETURN @OK
END

GO
/****** Object:  UserDefinedFunction [dbo].[SP_GLB_VALIDA_IE]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

CREATE FUNCTION [dbo].[SP_GLB_VALIDA_IE] 
(	
	-- Add the parameters for the function here
	@uf varchar(2),
	@ie varchar(18)
)
RETURNS INT 
BEGIN
	DECLARE @counter INT;
	DECLARE @b INT;
	DECLARE @soma INT;	
	DECLARE @dig INT;
	
	--auxiliares
	DECLARE @p INT;
	DECLARE @d INT;
	DECLARE @i VARCHAR(18);
	DECLARE @die VARCHAR(13);
	
	IF LEN(@ie) > 0 AND LEN(@uf) = 2
	BEGIN		
		--retira caracteres especiais da IE
		SET @ie = REPLACE(REPLACE(REPLACE(LTRIM(RTRIM(@ie)), '.', ''), '-', ''), '/', '');		
		
		-- verifica IE para o estado AC
		IF @uf = 'AC' 
		BEGIN
			--verifica tamanho da IE
			IF LEN(@ie) <> 13 OR SUBSTRING(@ie, 1, 2) <> '01'
				RETURN(0);
			
			--calcula primeiro digito verificador
			SET @counter = 1;
			SET @b = 4;
			SET @soma = 0;			
			WHILE @counter < 12
			BEGIN			  
			  SET @soma = @soma + (SUBSTRING(@ie, @counter, 1) * @b);			  
			  SET @counter = @counter + 1;
			  SET @b = @b -1;
			  IF @b = 1
				SET @b = 9;
			END
			SET @dig = 11 - (@soma % 11);
			IF @dig >= 10
				SET @dig = 0;
			IF @dig <> SUBSTRING(@ie, 12, 1)
				RETURN (0);
			
			--calcula segundo digito verificador
			SET @counter = 1;
			SET @b = 5;
			SET @soma = 0;
			WHILE @counter < 13
			BEGIN			  
			  SET @soma = @soma + (SUBSTRING(@ie, @counter, 1) * @b);
			  SET @counter = @counter + 1;
			  SET @b = @b -1;
			  IF @b = 1
				SET @b = 9;
			END
			SET @dig = 11 - (@soma % 11);
			IF @dig >= 10
				SET @dig = 0;
			IF @dig <> SUBSTRING(@ie, 13, 1)
				RETURN (0);							
			
		END
		--AC
		
		-- verifica IE para o estado AL
		ELSE IF @uf = 'AL' 
		BEGIN
			--verifica tamanho da IE
			IF LEN(@ie) <> 9 OR SUBSTRING(@ie, 1, 2) <> '24'
				RETURN(0);
			
			--calcula primeiro digito verificador
			SET @counter = 1;
			SET @b = 9;
			SET @soma = 0;			
			WHILE @counter < 9
			BEGIN			  
			  SET @soma = @soma + (SUBSTRING(@ie, @counter, 1) * @b);			  
			  SET @counter = @counter + 1;
			  SET @b = @b -1;			  
			END
			SET @soma = @soma * 10;
			SET @dig = @soma - FLOOR(@soma / 11) * 11;
			IF @dig = 10
				SET @dig = 0;
			IF @dig <> SUBSTRING(@ie, 9, 1)
				RETURN (0);			
		END
		--AL
		
		-- verifica IE para o estado AM
		ELSE IF @uf = 'AM' 
		BEGIN
			--verifica tamanho da IE
			IF LEN(@ie) <> 9
				RETURN(0);
			
			--calcula primeiro digito verificador
			SET @counter = 1;
			SET @b = 9;
			SET @soma = 0;			
			WHILE @counter < 9
			BEGIN			  
			  SET @soma = @soma + (SUBSTRING(@ie, @counter, 1) * @b);			  
			  SET @counter = @counter + 1;
			  SET @b = @b -1;			  
			END
			IF @soma < 11
				SET @dig = 11 - @soma;
			ELSE
				IF @soma % 11 <= 1
					SET @dig = 0;
				ELSE
					SET @dig = 11 - (@soma % 11);			
			IF @dig <> SUBSTRING(@ie, 9, 1)
				RETURN (0);			
		END
		--AM
		
		-- verifica IE para o estado AP
		ELSE IF @uf = 'AP' 
		BEGIN
			--verifica tamanho da IE
			IF LEN(@ie) <> 9
				RETURN(0);			
			
			--calcula primeiro digito verificador
			SET @p = 0;
			SET @d = 0;
			SET @i = SUBSTRING(@ie, 1, 8);			
			IF @i >= 3000001 AND @i <= 3017000
			BEGIN
				SET @p = 5;
				SET @d = 0;
			END
			ELSE IF @i >= 3017001 AND @i <= 3019022
			BEGIN
				SET @p = 9;
				SET @d = 1;
			END
			SET @counter = 1;
			SET @b = 9;
			SET @soma = @p;			
			WHILE @counter < 9
			BEGIN			  
			  SET @soma = @soma + (SUBSTRING(@i, @counter, 1) * @b);			  
			  SET @counter = @counter + 1;
			  SET @b = @b -1;			  
			END
			SET @dig = 11 - (@soma % 11);
			IF @dig = 10
				SET @dig = 0;
			ELSE IF @dig = 11
				SET @dig = @d;
			IF @dig <> SUBSTRING(@ie, 9, 1)
				RETURN (0);			
		END
		--AP
		
		-- verifica IE para o estado BA
		ELSE IF @uf = 'BA' 
		BEGIN
			--verifica tamanho da IE
			IF LEN(@ie) <> 8
				RETURN(0);
			
			--calcula segundo digito verificador
			SET @counter = 1;
			SET @b = 7;
			SET @soma = 0;			
			WHILE @counter < 7
			BEGIN			  
			  SET @soma = @soma + (SUBSTRING(@ie, @counter, 1) * @b);			  
			  SET @counter = @counter + 1;
			  SET @b = @b -1;			  
			END
			IF SUBSTRING(@ie, 1, 1) IN('0','1','2','3','4','5','8')
				SET @dig = 10 - (@soma % 10);
			ELSE
			BEGIN
				SET @dig = 11 - (@soma % 11);
				IF @dig <= 1
					SET @dig = 0;
			END
			IF @dig <> SUBSTRING(@ie, 8, 1)
				RETURN (0);
			
			--calcula primeiro digito verificador
			SET @counter = 1;
			SET @b = 8;
			SET @soma = 0;
			WHILE @counter < 9
			BEGIN			  
			  IF @counter <> 7
			  BEGIN
				SET @soma = @soma + (SUBSTRING(@ie, @counter, 1) * @b);			  
				SET @b = @b -1;
			  END
			  SET @counter = @counter + 1;			  
			END
			IF SUBSTRING(@ie, 1, 1) IN('0','1','2','3','4','5','8')
				SET @dig = 10 - (@soma % 10);
			ELSE
			BEGIN
				SET @dig = 11 - (@soma % 11);
				IF @dig <= 1
					SET @dig = 0;
			END
			IF @dig <> SUBSTRING(@ie, 7, 1)
				RETURN (0);		
		END
		--BA
		
		-- verifica IE para o estado CE
		ELSE IF @uf = 'CE' 
		BEGIN
			--verifica tamanho da IE
			IF LEN(@ie) > 9
				RETURN(0);			
			SET @die = @ie;
			IF(LEN(@ie) < 9)
			BEGIN
				WHILE LEN(@die) <= 8
				BEGIN
					SET @die = '0' + @die;
				END
			END
			--calcula primeiro digito verificador			
			SET @counter = 1;
			SET @b = 9;
			SET @soma = 0;			
			WHILE @counter < 9
			BEGIN			  
			  SET @soma = @soma + (SUBSTRING(@die, @counter, 1) * @b);			  
			  SET @counter = @counter + 1;
			  SET @b = @b -1;			  
			END			
			SET @dig = 11 - (@soma % 11);
			IF @dig >= 10
				SET @dig = 0;
			IF @dig <> SUBSTRING(@ie, 9, 1)
				RETURN (0);			
		END
		--CE
		
		-- verifica IE para o estado DF
		ELSE IF @uf = 'DF' 
		BEGIN
			--verifica tamanho da IE
			IF LEN(@ie) <> 13 OR SUBSTRING(@ie, 1, 2) <> '07'
				RETURN(0);
			
			--calcula primeiro digito verificador
			SET @counter = 1;
			SET @b = 4;
			SET @soma = 0;			
			WHILE @counter < 12
			BEGIN			  
			  SET @soma = @soma + (SUBSTRING(@ie, @counter, 1) * @b);			  
			  SET @counter = @counter + 1;
			  SET @b = @b -1;
			  IF @b = 1
				SET @b = 9;
			END
			SET @dig = 11 - (@soma % 11);
			IF @dig >= 10
				SET @dig = 0;
			IF @dig <> SUBSTRING(@ie, 12, 1)
				RETURN (0);
			
			--calcula segundo digito verificador
			SET @counter = 1;
			SET @b = 5;
			SET @soma = 0;
			WHILE @counter < 13
			BEGIN			  
			  SET @soma = @soma + (SUBSTRING(@ie, @counter, 1) * @b);
			  SET @counter = @counter + 1;
			  SET @b = @b -1;
			  IF @b = 1
				SET @b = 9;
			END
			SET @dig = 11 - (@soma % 11);
			IF @dig >= 10
				SET @dig = 0;
			IF @dig <> SUBSTRING(@ie, 13, 1)
				RETURN (0);							
			
		END
		--DF
		
		-- verifica IE para o estado ES
		ELSE IF @uf = 'ES' 
		BEGIN
			--verifica tamanho da IE
			IF LEN(@ie) <> 9
				RETURN(0);
			
			--calcula primeiro digito verificador
			SET @counter = 1;
			SET @b = 9;
			SET @soma = 0;			
			WHILE @counter < 9
			BEGIN			  
			  SET @soma = @soma + (SUBSTRING(@ie, @counter, 1) * @b);			  
			  SET @counter = @counter + 1;
			  SET @b = @b -1;			  
			END
			SET @dig = @soma % 11;			
			IF @dig < 2
				SET @dig = 0;
			ELSE				
				SET @dig = 11 - @dig;			
			IF @dig <> SUBSTRING(@ie, 9, 1)
				RETURN (0);			
		END
		--ES
		
		-- verifica IE para o estado GO
		ELSE IF @uf = 'GO' 
		BEGIN
			--verifica tamanho da IE
			IF LEN(@ie) <> 9 OR SUBSTRING(@ie, 1, 2) NOT IN ('10', '11', '15')
				RETURN(0);
			
			--calcula primeiro digito verificador
			SET @counter = 1;
			SET @b = 9;
			SET @soma = 0;			
			WHILE @counter < 9
			BEGIN			  
			  SET @soma = @soma + (SUBSTRING(@ie, @counter, 1) * @b);			  
			  SET @counter = @counter + 1;
			  SET @b = @b -1;
			END
			SET @dig = @soma % 11;			
			IF(@dig = 1)
			BEGIN
				IF SUBSTRING(@ie, 1, 8) >= 10103105 AND SUBSTRING(@ie, 1, 8) <= 10119997
					SET @dig = 1;
				ELSE
					SET @dig = 0
			END			
			ELSE IF @dig > 1
				SET @dig = 11 - @dig;
			IF @dig <> SUBSTRING(@ie, 9, 1)
				RETURN (0);			
		END
		--GO
		
		-- verifica IE para o estado MA
		ELSE IF @uf = 'MA'
		BEGIN			
			--verifica tamanho da IE
			IF LEN(@ie) <> 9 OR SUBSTRING(@ie, 1, 2) <> '12'
				RETURN(0);
			
			--calcula primeiro digito verificador
			SET @counter = 1;
			SET @b = 9;
			SET @soma = 0;			
			WHILE @counter < 9
			BEGIN			  
			  SET @soma = @soma + (SUBSTRING(@ie, @counter, 1) * @b);			  
			  SET @counter = @counter + 1;
			  SET @b = @b -1;			  
			END
			SET @dig = @soma % 11;
			IF @dig <= 1
				SET @dig = 0;
			ELSE
				SET @dig = 11 - @dig;
			IF @dig <> SUBSTRING(@ie, 9, 1)
				RETURN (0);									
		END
		--MA
		
		-- verifica IE para o estado MG
		ELSE IF @uf = 'MG' 
		BEGIN
			IF SUBSTRING(@ie, 1, 2) = 'PR' OR SUBSTRING(@ie, 1, 5) = 'ISENT'
				RETURN(1);
			--verifica tamanho da IE
			IF LEN(@ie) <> 13
				RETURN(0);
			
			SET @die = SUBSTRING(@ie, 1, 3) + '0' + SUBSTRING(@ie, 4, 11);
			--calcula primeiro digito verificador
			SET @counter = 1;
			SET @b = 1;
			SET @soma = 0;			
			WHILE @counter < 13
			BEGIN
			  IF (SUBSTRING(@die, @counter, 1) * @b) >= 10
				SET @soma = @soma + ((SUBSTRING(@die, @counter, 1) * @b) - 9);
			  ELSE
				SET @soma = @soma + (SUBSTRING(@die, @counter, 1) * @b);
			  SET @counter = @counter + 1;
			  SET @b = @b +1;
			  IF @b = 3
				SET @b = 1;
			END
			SET @dig = ((FLOOR(@soma/10)+1) * 10) - @soma;
			IF @dig <> SUBSTRING(@ie, 12, 1)
				RETURN (0);
			
			--calcula segundo digito verificador
			SET @counter = 1;
			SET @b = 3;
			SET @soma = 0;
			WHILE @counter < 13
			BEGIN			  
			  SET @soma = @soma + (SUBSTRING(@ie, @counter, 1) * @b);
			  SET @counter = @counter + 1;
			  SET @b = @b -1;
			  IF @b = 1
				SET @b = 11;
			END
			SET @dig = 11 - (@soma % 11);
			IF @dig >= 10
				SET @dig = 0;
			IF @dig <> SUBSTRING(@ie, 13, 1)
				RETURN (0);							
			
		END
		--MG
		
		-- verifica IE para o estado MT
		ELSE IF @uf = 'MT' 
		BEGIN
			--verifica tamanho da IE
			IF LEN(@ie) < 9
				RETURN(0);
							
			SET @die = @ie;
			IF(LEN(@ie) < 11)
			BEGIN
				WHILE LEN(@die) <= 11
				BEGIN
					SET @die = '0' + @die;
				END
			END
			--calcula primeiro digito verificador
			SET @counter = 1;
			SET @b = 3;
			SET @soma = 0;			
			WHILE @counter < 11
			BEGIN			  
			  SET @soma = @soma + (SUBSTRING(@die, @counter, 1) * @b);			  
			  SET @counter = @counter + 1;
			  SET @b = @b -1;
			  IF @b = 1
				SET @b = 9;
			END
			SET @dig = @soma % 11;
			IF @dig <= 1
				SET @dig = 0;
			ELSE
				SET @dig = 11 - @dig;
			IF @dig <> SUBSTRING(@ie, 11, 1)
				RETURN (0);			
		END
		--MT
		
		-- verifica IE para o estado MS
		ELSE IF @uf = 'MS'
		BEGIN
			--verifica tamanho da IE
			IF LEN(@ie) <> 9 OR SUBSTRING(@ie, 1, 2) <> '28'
				RETURN(0);
			
			--calcula primeiro digito verificador
			SET @counter = 1;
			SET @b = 9;
			SET @soma = 0;			
			WHILE @counter < 9
			BEGIN			  
			  SET @soma = @soma + (SUBSTRING(@ie, @counter, 1) * @b);			  
			  SET @counter = @counter + 1;
			  SET @b = @b -1;			  
			END
			SET @dig = @soma % 11;
			IF @dig <= 1
				SET @dig = 0;
			ELSE
				SET @dig = 11 - @dig;
			IF @dig <> SUBSTRING(@ie, 9, 1)
				RETURN (0);			
		END
		--MS
		
		-- verifica IE para o estado PA
		ELSE IF @uf = 'PA'
		BEGIN
			--verifica tamanho da IE
			IF LEN(@ie) <> 9 OR SUBSTRING(@ie, 1, 2) <> '15'
				RETURN(0);
			
			--calcula primeiro digito verificador
			SET @counter = 1;
			SET @b = 9;
			SET @soma = 0;			
			WHILE @counter < 9
			BEGIN			  
			  SET @soma = @soma + (SUBSTRING(@ie, @counter, 1) * @b);			  
			  SET @counter = @counter + 1;
			  SET @b = @b -1;			  
			END
			SET @dig = @soma % 11;
			IF @dig <= 1
				SET @dig = 0;
			ELSE
				SET @dig = 11 - @dig;
			IF @dig <> SUBSTRING(@ie, 9, 1)
				RETURN (0);			
		END
		--PA
		
		-- verifica IE para o estado PA
		/*ELSE IF @uf = 'PA'
		BEGIN
			--verifica tamanho da IE
			IF LEN(@ie) <> 9 OR SUBSTRING(@ie, 1, 2) <> '15'
				RETURN(0);
			
			--calcula primeiro digito verificador
			SET @counter = 1;
			SET @b = 9;
			SET @soma = 0;			
			WHILE @counter < 9
			BEGIN			  
			  SET @soma = @soma + (SUBSTRING(@ie, @counter, 1) * @b);			  
			  SET @counter = @counter + 1;
			  SET @b = @b -1;			  
			END
			SET @dig = @soma % 11;
			IF @dig <= 1
				SET @dig = 0;
			ELSE
				SET @dig = 11 - @dig;
			IF @dig <> SUBSTRING(@ie, 9, 1)
				RETURN (0);			
		END*/
		--PA
		
		-- verifica IE para o estado PB
		ELSE IF @uf = 'PB'
		BEGIN
			--verifica tamanho da IE
			IF LEN(@ie) <> 9
				RETURN(0);
			
			--calcula primeiro digito verificador
			SET @counter = 1;
			SET @b = 9;
			SET @soma = 0;			
			WHILE @counter < 9
			BEGIN			  
			  SET @soma = @soma + (SUBSTRING(@ie, @counter, 1) * @b);			  
			  SET @counter = @counter + 1;
			  SET @b = @b -1;			  
			END
			SET @dig = @soma % 11;
			IF @dig <= 1
				SET @dig = 0;
			ELSE
				SET @dig = 11 - @dig;
			IF @dig <> SUBSTRING(@ie, 9, 1)
				RETURN (0);			
		END
		--PB
		
		-- verifica IE para o estado PR
		ELSE IF @uf = 'PR' 
		BEGIN
			--verifica tamanho da IE
			IF LEN(@ie) <> 10
				RETURN(0);
			
			--calcula primeiro digito verificador
			SET @counter = 1;
			SET @b = 3;
			SET @soma = 0;			
			WHILE @counter < 9
			BEGIN			  
			  SET @soma = @soma + (SUBSTRING(@ie, @counter, 1) * @b);			  
			  SET @counter = @counter + 1;
			  SET @b = @b -1;
			  IF @b = 1
				SET @b = 7;
			END
			SET @dig = 11 - (@soma % 11);
			IF @dig >= 10
				SET @dig = 0;
			IF @dig <> SUBSTRING(@ie, 9, 1)
				RETURN (0);
			
			--calcula segundo digito verificador
			SET @counter = 1;
			SET @b = 4;
			SET @soma = 0;
			WHILE @counter < 10
			BEGIN			  
			  SET @soma = @soma + (SUBSTRING(@ie, @counter, 1) * @b);
			  SET @counter = @counter + 1;
			  SET @b = @b -1;
			  IF @b = 1
				SET @b = 7;
			END
			SET @dig = 11 - (@soma % 11);
			IF @dig >= 10
				SET @dig = 0;
			IF @dig <> SUBSTRING(@ie, 10, 1)
				RETURN (0);							
			
		END
		--PR
		
		-- verifica IE para o estado PE
		ELSE IF @uf = 'PE' 
		BEGIN
			--verifica tamanho da IE
			IF LEN(@ie) <> 9
				RETURN(0);
			
			--calcula primeiro digito verificador
			SET @counter = 1;
			SET @b = 8;
			SET @soma = 0;			
			WHILE @counter < 8
			BEGIN			  
			  SET @soma = @soma + (SUBSTRING(@ie, @counter, 1) * @b);			  
			  SET @counter = @counter + 1;
			  SET @b = @b -1;			  
			END
			SET @dig = 11 - (@soma % 11);
			IF @dig >= 10
				SET @dig = 0;
			IF @dig <> SUBSTRING(@ie, 8, 1)
				RETURN (0);
			
			--calcula segundo digito verificador
			SET @counter = 1;
			SET @b = 9;
			SET @soma = 0;
			WHILE @counter < 9
			BEGIN			  
			  SET @soma = @soma + (SUBSTRING(@ie, @counter, 1) * @b);
			  SET @counter = @counter + 1;
			  SET @b = @b -1;			  
			END
			SET @dig = 11 - (@soma % 11);
			IF @dig >= 10
				SET @dig = 0;
			IF @dig <> SUBSTRING(@ie, 9, 1)
				RETURN (0);							
			
		END
		--PE
		
		-- verifica IE para o estado PI
		ELSE IF @uf = 'PI'
		BEGIN
			--verifica tamanho da IE
			IF LEN(@ie) <> 9
				RETURN(0);
			
			--calcula primeiro digito verificador
			SET @counter = 1;
			SET @b = 9;
			SET @soma = 0;			
			WHILE @counter < 9
			BEGIN			  
			  SET @soma = @soma + (SUBSTRING(@ie, @counter, 1) * @b);			  
			  SET @counter = @counter + 1;
			  SET @b = @b -1;			  
			END
			SET @dig = @soma % 11;
			IF @dig <= 1
				SET @dig = 0;
			ELSE
				SET @dig = 11 - @dig;
			IF @dig <> SUBSTRING(@ie, 9, 1)
				RETURN (0);			
		END
		--PI
		
		-- verifica IE para o estado RJ
		ELSE IF @uf = 'RJ'
		BEGIN
			--verifica tamanho da IE
			IF LEN(@ie) <> 8
				RETURN(0);
			
			--calcula primeiro digito verificador
			SET @counter = 1;
			SET @b = 2;
			SET @soma = 0;			
			WHILE @counter < 8
			BEGIN			  
			  SET @soma = @soma + (SUBSTRING(@ie, @counter, 1) * @b);			  
			  SET @counter = @counter + 1;
			  SET @b = @b -1;
			  IF @b = 1
				SET @b = 7;			  
			END
			SET @dig = @soma % 11;
			IF @dig <= 1
				SET @dig = 0;
			ELSE
				SET @dig = 11 - @dig;
			IF @dig <> SUBSTRING(@ie, 8, 1)
				RETURN (0);			
		END
		--RJ
		
		-- verifica IE para o estado RN
		ELSE IF @uf = 'RN'
		BEGIN
			--verifica tamanho da IE
			IF LEN(@ie) <> 9
				RETURN(0);
			
			--calcula primeiro digito verificador
			SET @counter = 1;
			SET @b = 9;
			SET @soma = 0;			
			WHILE @counter < 9
			BEGIN			  
			  SET @soma = @soma + (SUBSTRING(@ie, @counter, 1) * @b);			  
			  SET @counter = @counter + 1;
			  SET @b = @b -1;			  
			END
			SET @soma = @soma * 10;
			SET @dig = @soma % 11;
			IF @dig = 10
				SET @dig = 0;			
			IF @dig <> SUBSTRING(@ie, 9, 1)
				RETURN (0);			
		END
		--RN
		
		-- verifica IE para o estado RS
		ELSE IF @uf = 'RS'
		BEGIN
			--verifica tamanho da IE
			IF LEN(@ie) <> 10 OR SUBSTRING(@ie, 1, 3) > '467'
				RETURN(0);
			
			--calcula primeiro digito verificador
			SET @counter = 1;
			SET @b = 2;
			SET @soma = 0;			
			WHILE @counter < 10
			BEGIN			  
			  SET @soma = @soma + (SUBSTRING(@ie, @counter, 1) * @b);			  
			  SET @counter = @counter + 1;
			  SET @b = @b -1;
			  IF @b = 1
				SET @b = 9;			  
			END
			SET @dig = 11 - (@soma % 11);
			IF @dig >= 10
				SET @dig = 0;			
			IF @dig <> SUBSTRING(@ie, 10, 1)
				RETURN (0);			
		END
		--RS
		
		-- verifica IE para o estado RO
		ELSE IF @uf = 'RO'
		BEGIN
			--verifica tamanho da IE
			IF LEN(@ie) <> 14
				RETURN(0);
			
			--calcula primeiro digito verificador
			SET @counter = 1;
			SET @b = 6;
			SET @soma = 0;			
			WHILE @counter < 14
			BEGIN			  
			  SET @soma = @soma + (SUBSTRING(@ie, @counter, 1) * @b);			  
			  SET @counter = @counter + 1;
			  SET @b = @b -1;
			  IF @b = 1
				SET @b = 9;			  
			END
			SET @dig = 11 - (@soma % 11);
			IF @dig >= 10
				SET @dig = @dig - 10;			
			IF @dig <> SUBSTRING(@ie, 14, 1)
				RETURN (0);			
		END
		--RO
		
		-- verifica IE para o estado RR
		ELSE IF @uf = 'RR'
		BEGIN
			--verifica tamanho da IE
			IF LEN(@ie) <> 9 OR SUBSTRING(@ie, 1, 2) <> '24'
				RETURN(0);
			
			--calcula primeiro digito verificador
			SET @counter = 1;
			SET @b = 1;
			SET @soma = 0;			
			WHILE @counter < 9
			BEGIN			  
			  SET @soma = @soma + (SUBSTRING(@ie, @counter, 1) * @b);			  
			  SET @counter = @counter + 1;
			  SET @b = @b + 1;
			END
			
			SET @dig = @soma % 9;			
			IF @dig <> SUBSTRING(@ie, 9, 1)
				RETURN (0);			
		END
		--RR
		
		-- verifica IE para o estado SC
		ELSE IF @uf = 'SC' 
		BEGIN
			--verifica tamanho da IE
			IF LEN(@ie) <> 9
				RETURN(0);
			
			--calcula primeiro digito verificador
			SET @counter = 1;
			SET @b = 9;
			SET @soma = 0;			
			WHILE @counter < 9
			BEGIN			  
			  SET @soma = @soma + (SUBSTRING(@ie, @counter, 1) * @b);			  
			  SET @counter = @counter + 1;
			  SET @b = @b -1;			  
			END
			SET @dig = @soma % 11;
			IF @dig <= 1
				SET @dig = 0;
			ELSE				
				SET @dig = 11 - @dig;
			IF @dig <> SUBSTRING(@ie, 9, 1)
				RETURN (0);			
		END
		--SC
		
		-- verifica IE para o estado SP
		ELSE IF @uf = 'SP' 
		BEGIN
			IF SUBSTRING(@ie, 1, 1) = 'P'
			BEGIN				
				SET @die = SUBSTRING(@ie, 2, 9);				
				SET @soma = (SUBSTRING(@die, 1, 1) * 1) + (SUBSTRING(@die, 2, 1) * 3) + (SUBSTRING(@die, 3, 1) * 4) + (SUBSTRING(@die, 4, 1) * 5) + (SUBSTRING(@die, 5, 1) * 6) + (SUBSTRING(@die, 6, 1) * 7) + (SUBSTRING(@die, 7, 1) * 8) + (SUBSTRING(@die, 8, 1) * 10);
				SET @dig = @soma % 11;
				IF @dig >= 10
					SET @dig = 0;				
				IF @dig <> SUBSTRING(@ie, 10, 1)
					RETURN (0);	
			END
			ELSE
			BEGIN
				IF LEN(@ie) < 12
					RETURN(0);
				SET @soma = (SUBSTRING(@ie, 1, 1) * 1) + (SUBSTRING(@ie, 2, 1) * 3) + (SUBSTRING(@ie, 3, 1) * 4) + (SUBSTRING(@ie, 4, 1) * 5) + (SUBSTRING(@ie, 5, 1) * 6) + (SUBSTRING(@ie, 6, 1) * 7) + (SUBSTRING(@ie, 7, 1) * 8) + (SUBSTRING(@ie, 8, 1) * 10);
				SET @dig = @soma % 11;
				IF @dig >= 10
					SET @dig = 0;
				IF @dig <> SUBSTRING(@ie, 9, 1)
					RETURN (0);
					
				SET @counter = 1;
				SET @b = 3;
				SET @soma = 0;			
				WHILE @counter < 12
				BEGIN			  
				  SET @soma = @soma + (SUBSTRING(@ie, @counter, 1) * @b);			  
				  SET @counter = @counter + 1;
				  SET @b = @b -1;
				  IF @b = 1
					SET @b = 10;			  
				END
				SET @dig = @soma % 11;
				IF @dig >= 10
					SET @dig = 0;
				IF @dig <> SUBSTRING(@ie, 12, 1)
					RETURN (0);
			END
		END
		--SP
		
		-- verifica IE para o estado SE
		ELSE IF @uf = 'SE' 
		BEGIN
			--verifica tamanho da IE
			IF LEN(@ie) <> 9
				RETURN(0);
			
			--calcula primeiro digito verificador
			SET @counter = 1;
			SET @b = 9;
			SET @soma = 0;			
			WHILE @counter < 9
			BEGIN			  
			  SET @soma = @soma + (SUBSTRING(@ie, @counter, 1) * @b);			  
			  SET @counter = @counter + 1;
			  SET @b = @b -1;			  
			END
			SET @dig = 11 - (@soma % 11);
			IF @dig >= 10
				SET @dig = 0;			
			IF @dig <> SUBSTRING(@ie, 9, 1)
				RETURN (0);			
		END
		--SE
		
		-- verifica IE para o estado TO
		ELSE IF @uf = 'TO' 
		BEGIN
			--verifica tamanho da IE
			IF LEN(@ie) <> 11 OR SUBSTRING(@ie, 3, 2) NOT IN('01','02','03','99')
				RETURN(0);
			
			--calcula primeiro digito verificador
			SET @counter = 1;
			SET @b = 9;
			SET @soma = 0;			
			WHILE @counter < 11
			BEGIN			  
			  IF @counter NOT IN (3, 4)
			  BEGIN
				  SET @soma = @soma + (SUBSTRING(@ie, @counter, 1) * @b);				  
				  SET @b = @b -1;			  
			  END
			  SET @counter = @counter + 1;
			END
			SET @dig = @soma % 11;
			IF @dig < 2
				SET @dig = 0;
			ELSE
				SET @dig = 11 - @dig;
			IF @dig <> SUBSTRING(@ie, 11, 1)
				RETURN (0);			
		END
		--TO
		
		
		
		
		
		ELSE
			RETURN(0);
	END			
	ELSE	
		RETURN(0);
	
	--retorna true	
	RETURN(1);
END


GO
/****** Object:  UserDefinedFunction [dbo].[TiraLetras]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO


create FUNCTION [dbo].[TiraLetras]
 (
 @Resultado VARCHAR(8000)
 )
 RETURNS VARCHAR(8000)
 AS
 BEGIN

    DECLARE @CharInvalido SMALLINT
    SET @CharInvalido = PATINDEX('%[^0-9]%', @Resultado)
    WHILE @CharInvalido > 0
    BEGIN
       SET @Resultado = STUFF(@Resultado, @CharInvalido, 1, '')
       SET @CharInvalido = PATINDEX('%[^0-9]%', @Resultado)
    END
    SET @Resultado = @Resultado
    RETURN @Resultado
	
 END
 
 
GO
/****** Object:  Table [dbo].[Agenda]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Agenda](
	[AGE_Data] [datetime] NULL,
	[AGE_Loja] [char](5) NULL,
	[AGE_Vendedor] [int] NULL,
	[AGE_Sequencia] [decimal](18, 0) IDENTITY(1,1) NOT FOR REPLICATION NOT NULL,
	[AGE_Assunto] [char](50) NULL
) ON [PRIMARY]

GO
/****** Object:  Table [dbo].[AjusteProcessado]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[AjusteProcessado](
	[AP_CodigoAjuste] [int] NOT NULL,
	[AP_Data] [datetime] NOT NULL,
	[AP_Usuario] [char](20) NULL,
PRIMARY KEY CLUSTERED 
(
	[AP_CodigoAjuste] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]

GO
/****** Object:  Table [dbo].[barcoldContagem]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[barcoldContagem](
	[BAR_Local] [int] NULL,
	[BAR_Dupla] [int] NULL,
	[BAR_Contagem] [int] NULL,
	[BAR_digitador] [int] NULL,
	[BAR_CodigoBarras] [char](30) NULL,
	[BAR_Referencia] [char](7) NULL,
	[BAR_Quantidade] [int] NULL,
	[BAR_Sequencia] [numeric](18, 0) NOT NULL,
	[BAR_Controle] [int] NULL,
	[BAR_Situacao] [char](1) NULL
) ON [PRIMARY]

GO
/****** Object:  Table [dbo].[CarimboNotaFiscal]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[CarimboNotaFiscal](
	[CNF_NumeroPed] [decimal](10, 0) NULL,
	[CNF_Loja] [char](5) NULL,
	[CNF_Serie] [char](3) NULL,
	[CNF_NF] [decimal](10, 0) NULL,
	[CNF_Sequencia] [decimal](10, 0) NULL,
	[CNF_Carimbo] [char](150) NULL,
	[CNF_TipoCarimbo] [char](1) NULL,
	[CNF_DetalheImpressao] [char](1) NULL,
	[CNF_Data] [datetime] NULL,
	[CNF_SituacaoProcesso] [char](1) NULL,
	[CNF_DataProcesso] [datetime] NULL
) ON [PRIMARY]

GO
/****** Object:  Table [dbo].[CarimbosEspeciais]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[CarimbosEspeciais](
	[CE_Referencia] [nvarchar](7) NULL,
	[CE_Linha1] [nvarchar](60) NULL,
	[CE_Linha2] [nvarchar](60) NULL,
	[CE_Linha3] [nvarchar](60) NULL,
	[CE_Linha4] [nvarchar](60) NULL,
	[CE_Linha5] [nvarchar](60) NULL,
	[CE_Linha6] [nvarchar](60) NULL,
	[CE_Linha7] [nvarchar](60) NULL,
	[CE_Linha8] [nvarchar](60) NULL,
	[CE_Linha9] [nvarchar](60) NULL,
	[CE_Linha10] [nvarchar](30) NULL,
	[CE_Linha11] [nvarchar](30) NULL,
	[CE_Linha12] [nvarchar](90) NULL
) ON [PRIMARY]

GO
/****** Object:  Table [dbo].[CertificadoInicio]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[CertificadoInicio](
	[DataEmissaoSeguro] [datetime] NULL,
	[DataCompraBem] [datetime] NULL,
	[PerildoGarantiaInicio] [datetime] NULL,
	[PerildoGarantiaFIM] [datetime] NULL,
	[PerildoVigenciaInicio] [datetime] NULL,
	[PerildoVigenciaFIM] [datetime] NULL,
	[limiteMaximoIndenizacao] [varchar](5) NOT NULL,
	[ProdutoSegurado] [char](50) NOT NULL,
	[Marca] [varchar](15) NOT NULL,
	[Modelo] [nvarchar](7) NOT NULL,
	[ValorDoProdutoSegurado] [money] NULL,
	[PremioLiquido] [float] NULL,
	[IOF] [float] NULL,
	[PremioTotal] [float] NULL,
	[CertificadoInicio] [char](12) NULL,
	[CertificadoFim] [char](12) NULL
) ON [PRIMARY]

GO
/****** Object:  Table [dbo].[CFOPEntradaSaida]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[CFOPEntradaSaida](
	[CFO_Codigo] [char](4) NOT NULL,
	[CFO_DescricaoOperacao] [char](40) NULL,
	[CFO_EntradaSaida] [char](3) NULL,
	[CFO_DentroForaEstado] [char](1) NULL,
	[CFO_Tributado] [char](1) NULL,
	[CFO_SubstituicaoTributaria] [char](1) NULL,
	[CFO_UF] [char](2) NULL
) ON [PRIMARY]

GO
/****** Object:  Table [dbo].[clientefichafinanceira]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[clientefichafinanceira](
	[OID] [int] NOT NULL,
	[CODIGO] [varchar](10) NOT NULL,
	[RAZAOSOCIAL] [varchar](254) NULL,
	[CNPJ] [varchar](18) NULL,
	[SITUACAO] [varchar](9) NULL,
	[PGTOCARTORIO] [decimal](18, 2) NULL,
	[INSCESTADUAL] [varchar](40) NULL,
	[RUA] [varchar](254) NULL,
	[NUMERO] [varchar](254) NULL,
	[COMPLEMENTO] [varchar](254) NULL,
	[BAIRRO] [varchar](254) NULL,
	[CIDADE] [varchar](254) NULL,
	[ESTADO] [varchar](2) NULL,
	[CEP] [varchar](9) NULL,
	[TELEFONE] [varchar](254) NULL,
	[FAX] [varchar](254) NULL,
	[DTCADASTRO] [datetime] NULL,
	[PESSOA] [varchar](1) NULL,
	[LIMITECREDITO] [decimal](18, 2) NULL,
	[DTLIMITE] [datetime] NULL,
	[JUROS] [decimal](18, 2) NULL,
	[VALORABERTO] [decimal](18, 2) NULL,
	[QTDEABERTO] [int] NULL,
	[QTDECOMPRA] [int] NULL,
	[DTULTCOMPRA] [datetime] NULL,
	[SALDOCOMPRA] [decimal](18, 2) NULL,
	[MAIORCOMPRA] [decimal](18, 2) NULL,
	[DTMAIORCOMPRA] [datetime] NULL,
	[ULTIMOPAGTO] [decimal](18, 2) NULL,
	[DTULTPAGTO] [datetime] NULL,
	[QTDEPAGTOATRASO] [int] NULL,
	[DUPLATRASO] [int] NULL,
	[MAIORATRASO] [int] NULL,
	[VALORULTCOMPRA] [decimal](18, 0) NULL
) ON [PRIMARY]

GO
/****** Object:  Table [dbo].[CodigoOperacao]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[CodigoOperacao](
	[CF_CodigoOperacao] [smallint] NOT NULL,
	[CF_CodigoOperacaoAux] [smallint] NOT NULL,
	[CF_Descricao] [varchar](60) NOT NULL,
	[CF_EntradaSaida] [char](1) NOT NULL,
	[CF_Transferencia] [char](1) NOT NULL,
	[CF_SimplesRemessa] [char](1) NOT NULL,
	[CF_Interestadual] [char](1) NOT NULL,
	[CF_Importacao] [char](1) NOT NULL,
	[CF_Devolucao] [char](1) NOT NULL,
	[CF_CodigoTributo] [char](2) NOT NULL,
	[CF_TipoCodigo] [varchar](3) NULL,
	[CF_CodigoOperacaoNovo] [smallint] NOT NULL,
 CONSTRAINT [PKCF_CodigoOperacao] PRIMARY KEY CLUSTERED 
(
	[CF_CodigoOperacao] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 90) ON [PRIMARY]
) ON [PRIMARY]

GO
/****** Object:  Table [dbo].[CODIGOOPERACAONOVO]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[CODIGOOPERACAONOVO](
	[CN_CodigoOperacaoNovo] [int] NOT NULL,
	[CN_CodigoOperacaoAntigo] [int] NOT NULL,
	[CN_DescricaoOperacao] [varchar](40) NULL
) ON [PRIMARY]

GO
/****** Object:  Table [dbo].[COMPLEMENTOVENDA]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[COMPLEMENTOVENDA](
	[COV_NumeroPedido] [decimal](18, 0) NOT NULL,
	[COV_CodigoComplemento] [decimal](18, 0) NOT NULL,
	[COV_SequenciaComplemento] [decimal](18, 0) NOT NULL,
	[COV_ValorComplemento] [char](50) NULL
) ON [PRIMARY]

GO
/****** Object:  Table [dbo].[comprador]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[comprador](
	[CO_CodigoComprador] [smallint] NOT NULL,
	[CO_Nome] [varchar](25) NOT NULL,
	[CO_Assinatura] [char](20) NULL,
	[CO_Email] [char](35) NULL,
	[CO_PercentualCompras] [float] NULL,
	[CO_PercentualVenda] [float] NULL,
	[CO_PercentualCobertura] [float] NULL,
	[CO_PrazoMedioDias] [int] NULL,
	[CO_PercentualSemGiro90] [float] NULL,
	[CO_PercentualBonificaçao] [float] NULL,
	[CO_Margem] [float] NULL,
	[CO_PercentualForadeLinha] [float] NULL,
	[CO_PercentualEstoqueCMC] [float] NULL
) ON [PRIMARY]

GO
/****** Object:  Table [dbo].[condicaoPagamento]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[condicaoPagamento](
	[CP_ID] [int] NOT NULL,
	[CP_Tipo] [nvarchar](2) NULL,
	[CP_TipoCondicao] [nvarchar](2) NULL,
	[CP_Codigo] [int] NOT NULL,
	[CP_Condicao] [nvarchar](30) NULL,
	[CP_Parcelas] [int] NULL,
	[CP_Coeficiente] [money] NULL,
	[CP_TipoDocumento] [int] NULL,
	[CP_IntervaloParcelas] [nvarchar](30) NULL,
	[CP_Desconto] [money] NULL,
PRIMARY KEY CLUSTERED 
(
	[CP_ID] ASC,
	[CP_Codigo] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]

GO
/****** Object:  Table [dbo].[CondicaoPagto]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[CondicaoPagto](
	[CP_CodigoCondicao] [smallint] NOT NULL,
	[CP_Descricao] [varchar](50) NOT NULL,
	[CP_QuantidadeParcelas] [smallint] NOT NULL,
	[CP_TipoCondicao] [char](2) NOT NULL,
	[CP_Parcelas] [varchar](25) NOT NULL,
	[CP_VendaCompra] [char](1) NOT NULL,
	[CP_Intervalo] [smallint] NULL,
 CONSTRAINT [PKCP_CondPagto] PRIMARY KEY CLUSTERED 
(
	[CP_CodigoCondicao] ASC,
	[CP_VendaCompra] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 90) ON [PRIMARY]
) ON [PRIMARY]

GO
/****** Object:  Table [dbo].[contagem]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[contagem](
	[CC_Local] [int] NOT NULL,
	[CC_Referencia] [varchar](7) NOT NULL,
	[CC_Quantidade1] [int] NULL,
	[CC_Quantidade2] [int] NULL,
	[CC_Quantidade3] [int] NULL,
	[CC_Quantidade4] [int] NULL,
	[CC_Quantidade5] [int] NULL,
	[CC_QuantidadeOK] [int] NULL,
	[CC_Situacao] [varchar](1) NOT NULL,
	[CC_CodigoBarras] [varchar](15) NULL
) ON [PRIMARY]

GO
/****** Object:  Table [dbo].[ControleCaixa]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[ControleCaixa](
	[CTR_Operador] [decimal](18, 0) NULL,
	[CTR_Supervisor] [decimal](18, 0) NULL,
	[CTR_DataInicial] [datetime] NULL,
	[CTR_DataFinal] [datetime] NULL,
	[CTR_SaldoAnterior] [float] NULL,
	[CTR_SaldoFinal] [float] NULL,
	[CTR_NroSangriaReforco] [decimal](18, 0) NULL,
	[CTR_SituacaoCaixa] [char](1) NULL,
	[CTR_Protocolo] [decimal](18, 0) IDENTITY(1,1) NOT FOR REPLICATION NOT NULL,
	[CTR_NumeroCaixa] [decimal](18, 0) NOT NULL,
	[CTR_ProtocoloAnterior] [decimal](18, 0) NULL
) ON [PRIMARY]

GO
/****** Object:  Table [dbo].[controleinv]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[controleinv](
	[CI_Loja] [varchar](5) NOT NULL,
	[CI_DataInventario] [datetime] NOT NULL,
	[CI_TotalLocais] [int] NULL,
	[CI_Situacao] [varchar](1) NOT NULL,
	[CI_AtualizaEstoqueLocal] [char](1) NULL,
	[CI_SenhaEspecial] [char](15) NULL
) ON [PRIMARY]

GO
/****** Object:  Table [dbo].[ControleSerie]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[ControleSerie](
	[CS_NroCaixa] [int] NOT NULL,
	[CS_SerieCF] [varchar](3) NOT NULL,
	[CS_Serie] [varchar](3) NOT NULL,
	[CS_DACFE] [varchar](50) NULL,
 CONSTRAINT [PK_ControleSerie] PRIMARY KEY CLUSTERED 
(
	[CS_NroCaixa] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]

GO
/****** Object:  Table [dbo].[ControleSistema]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[ControleSistema](
	[CTS_Loja] [char](5) NULL,
	[CTS_NumeroNF] [decimal](18, 0) NULL,
	[CTS_Numero00] [decimal](18, 0) NULL,
	[CTS_NumeroNE] [decimal](18, 0) NULL,
	[CTS_NumeroPedido] [decimal](18, 0) NULL,
	[CTS_NumeroNCredito] [decimal](18, 0) NULL,
	[CTS_DescontoVendedor] [money] NULL,
	[CTS_SenhaDesconto] [char](6) NULL,
	[CTS_ValidadeCotacao] [int] NULL,
	[CTS_SequenciaCliente] [float] NULL,
	[CTS_ServidorRetaguarda] [char](50) NULL,
	[CTS_SerieNota] [varchar](2) NULL,
	[CTS_SerieTransferencia] [varchar](50) NULL,
	[CTS_CaminhoWeb1] [char](100) NULL,
	[CTS_CaminhoWeb2] [char](100) NULL,
	[CTS_CaminhoWeb3] [char](100) NULL,
	[CTS_CaminhoWeb4] [char](100) NULL,
	[CTS_SenhaLiberacao] [nvarchar](12) NULL,
	[CTS_LogoPedido] [char](100) NULL,
	[CTS_MensagemECF] [char](40) NULL,
	[CTS_QtdeViaRomaneio] [int] NULL,
	[CTS_QtdeViaMovimento] [int] NULL,
	[CTS_EmiteCupom] [char](1) NULL,
	[CTS_CaminhoNFe] [char](100) NULL,
	[CTS_CodigoClienteFaturado] [numeric](18, 0) NULL,
	[CTS_EmiteCodigoZero] [char](1) NULL,
	[CTS_LiberaPOS] [char](1) NULL,
	[CTS_LimiteFormaPagamento] [numeric](18, 0) NULL,
	[CTS_Apolice] [char](30) NULL,
	[CTS_Certificado] [numeric](5, 0) NULL,
	[CTS_codigoInternoProdutoGE] [char](2) NULL,
	[CTS_codigoEstipulanteGE] [char](4) NULL,
	[CTS_DataAjuste] [datetime] NULL,
	[CTS_MostraMgSimulador] [char](1) NULL,
	[cts_ServidorAtualizacao] [char](100) NULL,
	[CTS_DataEstoque] [datetime] NULL,
	[CTS_LiberaBloqueio] [char](1) NULL,
	[CTS_DanfeImpressora] [varchar](100) NULL,
	[CTS_LiberaBloqueioPreco] [char](1) NULL,
	[cts_CaminhoBanner] [varchar](300) NULL,
	[CTS_EmiteSAT] [char](1) NULL,
	[CTS_TipoEmpresa] [char](2) NULL
) ON [PRIMARY]

GO
/****** Object:  Table [dbo].[Crediario]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Crediario](
	[CRE_CodigoCrediario] [decimal](18, 0) NOT NULL,
	[CRE_NroCampos] [decimal](18, 0) NULL,
	[CRE_Parcelas1] [char](15) NULL,
	[CRE_Coeficiente1] [float] NULL,
	[CRE_Parcelas2] [char](15) NULL,
	[CRE_Coeficiente2] [float] NULL,
	[CRE_Parcelas3] [char](15) NULL,
	[CRE_Coeficiente3] [float] NULL,
	[CRE_Parcelas4] [char](15) NULL,
	[CRE_Coeficiente4] [float] NULL,
	[CRE_Parcelas5] [char](15) NULL,
	[CRE_Coeficiente] [float] NULL,
	[CRE_Parcelas6] [char](15) NULL,
	[CRE_Coeficiente6] [float] NULL,
	[CRE_Parcelas7] [char](15) NULL,
	[CRE_Coeficiente7] [float] NULL,
	[CRE_Parcelas8] [char](15) NULL,
	[CRE_Coeficiente8] [float] NULL,
	[CRE_Parcelas9] [char](15) NULL,
	[CRE_Coeficiente9] [float] NULL,
	[CRE_Parcelas10] [char](15) NULL,
	[CRE_Coeficiente10] [float] NULL,
	[CRE_Parcelas11] [char](15) NULL,
	[CRE_Coeficiente11] [float] NULL,
	[CRE_Parcelas12] [char](15) NULL,
	[CRE_Coeficiente12] [float] NULL,
	[CRE_Parcelas13] [char](15) NULL,
	[CRE_Coeficiente13] [float] NULL,
	[CRE_Parcelas14] [char](15) NULL,
	[CRE_Coeficiente14] [float] NULL,
	[CRE_Parcelas15] [char](15) NULL,
	[CRE_Coeficiente15] [float] NULL,
	[CRE_Parecelas16] [char](15) NULL,
	[CRE_Coeficiente16] [float] NULL,
	[CRE_Parcelas17] [char](15) NULL,
	[CRE_Coeficiente17] [float] NULL,
	[CRE_Parcelas18] [char](15) NULL,
	[CRE_Coeficiente18] [float] NULL,
	[CRE_Parcelas19] [char](15) NULL,
	[CRE_Coeficiente19] [float] NULL,
	[CRE_Parcelas20] [char](15) NULL,
	[CRE_Coeficiente20] [float] NULL,
	[CRE_Parcelas21] [char](15) NULL,
	[CRE_Coeficiente21] [float] NULL,
	[CRE_Parcelas22] [char](15) NULL,
	[CRE_Coeficiente22] [float] NULL,
	[CRE_Parcelas23] [char](15) NULL,
	[CRE_Coeficiente23] [float] NULL,
	[CRE_Parcelas24] [char](15) NULL,
	[CRE_Coeficiente24] [float] NULL
) ON [PRIMARY]

GO
/****** Object:  Table [dbo].[DivergenciaEstoque]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[DivergenciaEstoque](
	[DE_Loja] [varchar](5) NOT NULL,
	[DE_descricao] [varchar](100) NULL,
	[DE_Referencia] [char](7) NOT NULL,
	[DE_EstoqueLoja] [int] NOT NULL,
	[DE_EstoqueCentral] [int] NOT NULL,
	[DE_Sequencia] [decimal](10, 0) IDENTITY(1,1) NOT FOR REPLICATION NOT NULL
) ON [PRIMARY]

GO
/****** Object:  Table [dbo].[dupla]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[dupla](
	[CD_CodigoDupla] [int] NOT NULL,
	[CD_NomeDupla] [varchar](20) NOT NULL,
	[CD_TipoDupla] [int] NULL
) ON [PRIMARY]

GO
/****** Object:  Table [dbo].[Duplicata]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Duplicata](
	[DP_Loja] [char](5) NOT NULL,
	[DP_NotaFiscal] [int] NOT NULL,
	[DP_Serie] [char](2) NOT NULL,
	[DP_Sequencia] [tinyint] NOT NULL,
	[DP_CodigoCliente] [int] NOT NULL,
	[DP_DataEmissao] [datetime] NOT NULL,
	[DP_Vendedor] [smallint] NOT NULL,
	[DP_Banco] [smallint] NOT NULL,
	[DP_DocumentoBancario] [varchar](10) NULL,
	[DP_ValorDuplicata] [float] NOT NULL,
	[DP_DataVencimento] [datetime] NOT NULL,
	[DP_NotaCredito] [smallint] NULL,
	[DP_Abatimento] [float] NOT NULL,
	[DP_Desconto] [float] NOT NULL,
	[DP_Despesas] [float] NOT NULL,
	[DP_Juros] [float] NOT NULL,
	[DP_ValorPago] [float] NOT NULL,
	[DP_DataPagamento] [datetime] NULL,
	[DP_DataBaixa] [datetime] NULL,
	[DP_DataCartorio] [datetime] NULL,
	[DP_Historico] [varchar](250) NULL,
	[DP_TipoPagamento] [char](2) NULL,
	[DP_Agrupamento] [int] NULL,
	[DP_Situacao] [char](1) NOT NULL,
 CONSTRAINT [PKDP_Duplicata] PRIMARY KEY CLUSTERED 
(
	[DP_Loja] ASC,
	[DP_NotaFiscal] ASC,
	[DP_Serie] ASC,
	[DP_Sequencia] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 90) ON [PRIMARY]
) ON [PRIMARY]

GO
/****** Object:  Table [dbo].[EMKTLoja]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[EMKTLoja](
	[MKT_Loja] [char](5) NOT NULL,
	[MKT_Vendedor] [int] NOT NULL,
	[MKT_Email] [char](30) NOT NULL,
	[MKT_DataCadastro] [datetime] NULL,
	[MKT_RamoAtividade] [int] NULL,
	[MKT_Nome] [char](30) NULL,
	[MKT_CEP] [char](8) NULL,
	[MKT_Situacao] [char](1) NULL
) ON [PRIMARY]

GO
/****** Object:  Table [dbo].[EstoqueInv]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[EstoqueInv](
	[EI_Referencia] [varchar](7) NOT NULL,
	[EI_Quantidade] [int] NULL,
	[EI_QuantidadeVezes] [int] NULL
) ON [PRIMARY]

GO
/****** Object:  Table [dbo].[EstoqueLoja]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[EstoqueLoja](
	[EL_Loja] [char](7) NOT NULL,
	[EL_Referencia] [char](7) NOT NULL,
	[EL_CodigoFornecedor] [int] NULL,
	[EL_Estoque] [int] NULL,
	[EL_EstoqueAnterior] [int] NULL,
	[EL_NaoComercializado] [int] NULL,
	[EL_NaoComercializadoCONSO] [int] NULL,
PRIMARY KEY CLUSTERED 
(
	[EL_Referencia] ASC,
	[EL_Loja] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]

GO
/****** Object:  Table [dbo].[FaixaPremioGE]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[FaixaPremioGE](
	[FPG_Plano] [numeric](18, 0) NOT NULL,
	[FPG_CodigoFaixa] [numeric](18, 0) NOT NULL,
	[FPG_FaixaInicial] [float] NULL,
	[FPG_FaixaFinal] [float] NULL,
	[FPG_PremioLiquido] [float] NULL,
	[FPG_Remuneracao] [float] NULL,
	[FPG_PISCOFINS] [float] NULL,
	[FPG_IOF] [float] NULL,
	[FPG_Premio] [float] NULL
) ON [PRIMARY]

GO
/****** Object:  Table [dbo].[FIN_CEP]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[FIN_CEP](
	[CEP] [varchar](50) NOT NULL,
	[LOGRADOURO] [varchar](151) NULL,
	[BAIRRO] [varchar](150) NULL,
	[BAIRROA] [varchar](50) NULL,
	[MUNICIPIO] [varchar](50) NULL,
	[UF] [varchar](50) NULL,
 CONSTRAINT [pk_cep] PRIMARY KEY CLUSTERED 
(
	[CEP] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]

GO
/****** Object:  Table [dbo].[FIN_Cliente]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[FIN_Cliente](
	[CE_Loja] [varchar](5) NOT NULL,
	[CE_CodigoCliente] [int] NOT NULL,
	[CE_CGC] [varchar](20) NOT NULL,
	[CE_InscricaoEstadual] [varchar](15) NOT NULL,
	[CE_InscricaoEstadualSuframa] [char](20) NULL,
	[CE_Razao] [varchar](60) NOT NULL,
	[CE_Endereco] [varchar](60) NOT NULL,
	[CE_Bairro] [varchar](40) NOT NULL,
	[CE_Municipio] [varchar](40) NOT NULL,
	[CE_Estado] [char](2) NOT NULL,
	[CE_CEP] [char](8) NOT NULL,
	[CE_Telefone] [varchar](15) NULL,
	[CE_Fax] [varchar](15) NULL,
	[CE_EMail] [varchar](60) NULL,
	[CE_TipoPessoa] [char](1) NOT NULL,
	[CE_Praca] [smallint] NOT NULL,
	[CE_PagamentoCarteira] [char](1) NOT NULL,
	[CE_EnderecoCobranca] [varchar](60) NULL,
	[CE_BairroCobranca] [varchar](60) NULL,
	[CE_MunicipioCobranca] [varchar](40) NULL,
	[CE_EstadoCobranca] [char](2) NULL,
	[CE_CEPCobranca] [char](8) NULL,
	[CE_DataCadastro] [datetime] NOT NULL,
	[CE_DataCancelamento] [datetime] NULL,
	[CE_Alteracao] [char](1) NULL,
	[CE_Situacao] [smallint] NOT NULL,
	[CE_HoraManutencao] [datetime] NULL,
	[CE_Numero] [varchar](10) NULL,
	[CE_Complemento] [varchar](15) NULL,
	[CE_CodigoMunicipio] [char](7) NULL,
	[CE_NumeroCobranca] [varchar](10) NULL,
	[CE_Celular] [varchar](15) NULL,
	[CE_ramoatividade] [decimal](10, 0) NULL,
	[CE_DataNasc] [datetime] NULL,
	[CE_ComplCobranca] [varchar](30) NULL,
	[CE_segmento] [int] NULL,
	[CE_TipoCliente] [char](1) NULL,
	[CE_ClienteFidelidade] [char](18) NULL,
	[CE_Vendedor] [int] NULL,
	[CE_LimiteCredito] [float] NULL,
	[CE_dataLimiteCredito] [date] NULL,
	[CE_MaiorCompra] [float] NULL,
	[CE_DataMaiorCompra] [date] NULL,
	[CE_UltimaCompra] [float] NULL,
	[CE_dataUltimaCompra] [date] NULL,
	[CE_ultimoPagamento] [float] NULL,
	[CE_DataUltimoPagamento] [date] NULL,
	[CE_maiorAtraso] [int] NULL,
	[CE_quantidadeCompras] [int] NULL,
	[CE_JurosCartorio] [float] NULL,
	[CE_Mun_Codigo] [char](7) NULL,
 CONSTRAINT [PK_fin_cliente_2] PRIMARY KEY CLUSTERED 
(
	[CE_Loja] ASC,
	[CE_CodigoCliente] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]

GO
/****** Object:  Table [dbo].[FIN_Estado]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[FIN_Estado](
	[UF_Estado] [nvarchar](2) NOT NULL,
	[UF_Nome] [nvarchar](25) NULL,
	[UF_Regiao] [int] NULL,
	[UF_ICMSInterno] [float] NULL,
	[UF_ICMSInterEstadual] [float] NULL,
	[UF_ICMSInterImport] [float] NULL,
	[UF_ICMSDifal] [float] NULL,
	[UF_ICMSDifalImportado] [float] NULL,
	[UF_FECP] [float] NULL,
	[UF_Participacao] [float] NULL,
PRIMARY KEY CLUSTERED 
(
	[UF_Estado] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]

GO
/****** Object:  Table [dbo].[FIN_FaixaPremioGE]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[FIN_FaixaPremioGE](
	[FPG_Plano] [numeric](18, 0) NOT NULL,
	[FPG_CodigoFaixa] [numeric](18, 0) NOT NULL,
	[FPG_FaixaInicial] [float] NULL,
	[FPG_FaixaFinal] [float] NULL,
	[FPG_PremioLiquido] [float] NULL,
	[FPG_Remuneracao] [float] NULL,
	[FPG_PISCOFINS] [float] NULL,
	[FPG_IOF] [float] NULL,
	[FPG_Premio] [float] NULL
) ON [PRIMARY]

GO
/****** Object:  Table [dbo].[FIN_Municipio]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[FIN_Municipio](
	[Mun_Nome] [char](60) NULL,
	[Mun_Codigo] [char](7) NULL,
	[Mun_UF] [char](2) NULL
) ON [PRIMARY]

GO
/****** Object:  Table [dbo].[FIN_RamoAtividade]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[FIN_RamoAtividade](
	[RMO_Codigo] [decimal](18, 0) NOT NULL,
	[RMO_Pessoa] [char](1) NOT NULL,
	[RMO_DescricaoRamo] [char](50) NOT NULL
) ON [PRIMARY]

GO
/****** Object:  Table [dbo].[FIN_Segmento]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[FIN_Segmento](
	[SEG_RamoAtividade] [decimal](18, 0) NOT NULL,
	[SEG_CodigoSegmento] [decimal](18, 0) NOT NULL,
	[SEG_Descricao] [char](50) NOT NULL,
	[SEG_Situacao] [char](1) NOT NULL
) ON [PRIMARY]

GO
/****** Object:  Table [dbo].[FIN_SituacaoCliente]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[FIN_SituacaoCliente](
	[SI_CodigoSituacao] [smallint] NOT NULL,
	[SI_Descricao] [varchar](30) NOT NULL
) ON [PRIMARY]

GO
/****** Object:  Table [dbo].[Fornecedor]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Fornecedor](
	[FO_CodigoFornecedor] [smallint] NOT NULL,
	[FO_RazaoSocial] [varchar](40) NOT NULL,
	[FO_NomeFantasia] [varchar](15) NOT NULL,
	[FO_TipoFornecedor] [char](1) NOT NULL,
	[FO_PessoaContato] [varchar](200) NULL,
	[FO_Endereco] [varchar](40) NOT NULL,
	[FO_Municipio] [varchar](15) NOT NULL,
	[FO_Estado] [char](2) NOT NULL,
	[FO_Cep] [char](8) NOT NULL,
	[FO_Telefone] [varchar](15) NOT NULL,
	[FO_Fax] [varchar](15) NOT NULL,
	[FO_InscricaoEstadual] [varchar](15) NOT NULL,
	[FO_CGC] [varchar](15) NOT NULL,
	[FO_FrequenciaVisita] [char](1) NULL,
	[FO_DiaVisita] [char](2) NULL,
	[FO_CarenciaEntrega] [smallint] NOT NULL,
	[FO_PedidoMinimo] [float] NOT NULL,
	[FO_VerbaPropaganda] [float] NOT NULL,
	[FO_ValorFrete] [float] NOT NULL,
	[FO_PercentualFrete] [float] NOT NULL,
	[FO_PercentualVendor] [float] NOT NULL,
	[FO_CIFFOB] [tinyint] NOT NULL,
	[FO_Observacao] [text] NULL,
	[FO_CFODefault] [smallint] NULL,
	[FO_MargemSubstituicaoTrib] [float] NULL,
	[FO_IPISobreFrete] [char](1) NULL,
	[FO_EMail] [varchar](50) NULL,
	[FO_DataCadastro] [datetime] NULL,
	[FO_BloqueioFornecedor] [char](1) NULL,
	[FO_MotivoBloqueio] [char](250) NULL,
	[FO_WebEDI] [char](1) NULL,
	[FO_eMaileCommerce] [varchar](50) NULL,
	[FO_CodigoFornecedorMU] [varchar](6) NULL,
	[FO_HoraManutencao] [datetime] NULL,
	[Fo_SequenciaProduto] [smallint] NULL,
	[FO_CaminhoFTP] [char](50) NULL,
	[FO_LeadTime] [int] NULL,
	[FO_LeadTimeAutomatico] [char](1) NULL,
	[FO_ToleranciaPedidoPendente] [int] NULL,
	[FO_Cobertura] [int] NULL,
	[FO_RevisaoCompras] [datetime] NULL,
	[FO_ControlaFornecedor] [int] NULL,
	[FO_FornecedorRecebimento] [smallint] NULL,
	[FO_PrioridadeCompras] [char](1) NULL,
	[FO_PrioridadeComprasMes] [char](1) NULL,
	[FO_Ranking] [int] NULL,
	[FO_CodigoMunicipio] [numeric](18, 0) NULL,
	[FO_bairro] [varchar](40) NULL,
	[FO_numero] [varchar](10) NULL,
	[FO_Complemento] [varchar](30) NULL,
	[FO_fornecedoraPagar] [int] NULL,
PRIMARY KEY CLUSTERED 
(
	[FO_CodigoFornecedor] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]

GO
/****** Object:  Table [dbo].[GarantiaEstendida]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[GarantiaEstendida](
	[NF] [numeric](9, 0) NOT NULL,
	[SERIE] [nvarchar](6) NOT NULL,
	[DATAEMI] [datetime] NOT NULL,
	[REFERENCIA] [nvarchar](14) NOT NULL,
	[LOJAORIGEM] [nvarchar](10) NOT NULL,
	[GarantiaEstendida] [char](1) NOT NULL,
	[PlanoGarantia] [int] NOT NULL,
	[CoeficientePlano] [real] NOT NULL,
	[QtdeGarantia] [int] NOT NULL,
	[ValorGarantia] [real] NOT NULL,
	[CertificadoInicio] [char](12) NOT NULL,
	[CertificadoFim] [char](12) NOT NULL,
	[ge_premioLiquido] [real] NOT NULL,
	[ge_IOF] [real] NOT NULL,
	[ge_dataInicioVigencia] [datetime] NOT NULL,
	[ge_dataFinalVigencia] [datetime] NOT NULL,
	[ge_valorCustoSeguradora] [float] NOT NULL,
	[ge_seqCancelamento] [int] NULL,
	[ge_dataCancelamento] [datetime] NULL,
 CONSTRAINT [PK_GarantiaEstendida] PRIMARY KEY CLUSTERED 
(
	[NF] ASC,
	[SERIE] ASC,
	[LOJAORIGEM] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]

GO
/****** Object:  Table [dbo].[ICMSInterEstadual]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[ICMSInterEstadual](
	[IE_Codigo] [char](6) NOT NULL,
	[IE_ICMSDestino] [float] NULL,
	[IE_ICMSAplicado] [float] NULL,
	[IE_CFOP] [int] NULL,
	[IE_BasedeReducao] [float] NULL,
	[IE_CST] [char](3) NULL
) ON [PRIMARY]

GO
/****** Object:  Table [dbo].[LembreMe]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[LembreMe](
	[LEM_Loja] [char](5) NULL,
	[LEM_Vendedor] [char](5) NULL,
	[LEM_Referencia] [char](7) NULL,
	[LEM_Data] [datetime] NULL,
	[LEM_Observacao] [char](135) NULL,
	[LEM_Situacao] [char](1) NULL,
	[LEM_Sequencia] [int] IDENTITY(1,1) NOT FOR REPLICATION NOT NULL
) ON [PRIMARY]

GO
/****** Object:  Table [dbo].[LinhaProduto]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[LinhaProduto](
	[LPR_Linha] [char](8) NOT NULL,
	[LPR_Descricao] [char](35) NULL,
	[LPR_Comissao] [float] NULL,
	[LPR_Margem] [float] NULL,
	[LPR_Ordem] [numeric](18, 0) NOT NULL,
	[LPR_TipoRegistro] [tinyint] NULL
) ON [PRIMARY]

GO
/****** Object:  Table [dbo].[locais]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[locais](
	[CL_CodigoLocal] [int] NOT NULL,
	[CL_NomeLocal] [varchar](20) NOT NULL,
	[CL_EmiteFolha] [int] NULL,
	[CL_Contagem] [int] NULL,
	[CL_Situacao] [varchar](2) NOT NULL,
	[CL_Dupla1] [int] NULL,
	[CL_Dupla2] [int] NULL,
	[CL_Dupla3] [int] NULL,
	[CL_Dupla4] [int] NULL,
	[CL_Dupla5] [int] NULL,
	[CL_Digitador1] [int] NULL,
	[CL_Digitador2] [int] NULL,
	[CL_Digitador3] [int] NULL,
	[CL_Digitador4] [int] NULL,
	[CL_Digitador5] [int] NULL
) ON [PRIMARY]

GO
/****** Object:  Table [dbo].[Loja]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Loja](
	[LO_Loja] [char](5) NOT NULL,
	[LO_Empresa] [smallint] NOT NULL,
	[LO_Endereco] [varchar](40) NOT NULL,
	[LO_Numero] [numeric](18, 0) NULL,
	[LO_Bairro] [varchar](15) NOT NULL,
	[LO_Municipio] [varchar](15) NOT NULL,
	[LO_UF] [char](2) NOT NULL,
	[LO_Cep] [char](8) NOT NULL,
	[LO_CGC] [varchar](15) NOT NULL,
	[LO_InscricaoEstadual] [varchar](15) NOT NULL,
	[LO_Telefone] [varchar](20) NOT NULL,
	[LO_Fax] [varchar](20) NOT NULL,
	[LO_Localizacao] [char](30) NOT NULL,
	[LO_Registro] [varchar](15) NOT NULL,
	[LO_UltimaConexao] [datetime] NOT NULL,
	[LO_SituacaoConexao] [char](1) NOT NULL,
	[LO_TelComunicacao] [varchar](15) NOT NULL,
	[LO_UltimoProcesso] [datetime] NOT NULL,
	[LO_ArquivosProcessar] [char](1) NOT NULL,
	[LO_SituacaoProcesso] [char](1) NOT NULL,
	[LO_SituacaoCaixa] [char](1) NOT NULL,
	[LO_Situacao] [char](1) NULL,
	[LO_Gerente] [varchar](15) NULL,
	[LO_CriaArquivos] [char](1) NULL,
	[LO_ProximoInventario] [datetime] NULL,
	[LO_UltimoInventario] [datetime] NULL,
	[LO_Razao] [char](30) NULL,
	[LO_NomeLogotipo] [varchar](60) NULL,
	[LO_OrdemLoja] [int] NULL,
	[LO_PosicaoLoja] [int] NULL,
	[LO_Regiao] [int] NULL,
	[LO_SituacaoEnvio] [char](1) NULL,
	[LO_ProcessaNota] [char](2) NULL,
	[LO_ComprasOnLine] [char](2) NOT NULL,
	[LO_ProcessaAuditorEstoque] [char](1) NULL,
	[LO_ProcessaEstoque] [char](1) NULL,
	[LO_MostraEstoque] [char](1) NULL,
	[LO_PerInvInicial] [datetime] NULL,
	[LO_PerInvFinal] [datetime] NULL,
	[LO_IpLoja] [char](15) NULL,
	[LO_NomeServidor] [char](31) NULL,
	[LO_BancoOnLine] [int] NULL,
	[LO_LinkedServer] [char](1) NULL,
	[LO_OrdemLojaMU] [char](2) NULL,
	[LO_CriaEstoque] [char](1) NULL,
	[LO_NroFuncionario] [int] NULL,
	[LO_MetragemLoja] [float] NULL,
	[LO_CodigoUF] [numeric](18, 0) NULL,
	[LO_CodigoMunicipio] [numeric](18, 0) NULL,
	[LO_NomeFantasia] [char](20) NULL,
	[LO_CodigoPais] [char](4) NULL,
	[LO_EnderecoNFe] [char](60) NULL,
	[LO_EnderecoNroNFe] [int] NULL,
	[LO_ComplementoNFe] [char](60) NULL,
	[LO_OrdemDistribuicao] [int] NULL,
	[LO_EmaiLoja] [char](50) NULL,
	[LO_EnviaProdutoLoja] [char](2) NULL,
	[LO_Conexao] [char](1) NULL,
	[LO_GrupoRegiao] [char](2) NULL,
	[LO_NroVendedor] [int] NULL,
	[LO_LogoPedido] [char](30) NULL,
	[LO_InscricaoMunicipal] [char](15) NULL,
	[LO_DDD] [char](2) NULL,
	[LO_TipoTransferencia] [char](2) NULL,
	[LO_DMAC] [char](1) NULL,
	[Lo_site] [varchar](50) NULL,
	[Lo_Televendas] [varchar](20) NULL,
	[lo_nomeLoja] [varchar](30) NULL,
PRIMARY KEY CLUSTERED 
(
	[LO_Loja] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]

GO
/****** Object:  Table [dbo].[mensagens]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[mensagens](
	[ME_loja] [char](5) NULL,
	[ME_Nota] [int] NULL,
	[ME_Serie] [char](3) NULL,
	[ME_Assunto] [varchar](max) NULL,
	[ME_Mensagem] [varchar](max) NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]

GO
/****** Object:  Table [dbo].[modalidade]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[modalidade](
	[MO_Grupo] [int] NOT NULL,
	[MO_Descricao] [nvarchar](60) NULL,
	[MO_AtualizaPDV] [nvarchar](2) NULL,
	[MO_OrdemApresentacao] [varchar](6) NULL,
	[MO_TipoPag] [char](2) NULL,
	[MO_Bandeira] [char](3) NULL,
 CONSTRAINT [PK_modalidade] PRIMARY KEY CLUSTERED 
(
	[MO_Grupo] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]

GO
/****** Object:  Table [dbo].[MovimentoCaixa]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[MovimentoCaixa](
	[MC_NumeroECF] [int] NULL,
	[MC_CodigoOperador] [int] NULL,
	[MC_Loja] [nvarchar](10) NULL,
	[MC_Data] [datetime] NULL,
	[MC_Grupo] [int] NULL,
	[MC_SubGrupo] [char](10) NULL,
	[MC_Documento] [int] NULL,
	[MC_Serie] [nvarchar](3) NULL,
	[MC_Valor] [money] NULL,
	[MC_Banco] [int] NULL,
	[MC_Agencia] [nvarchar](30) NULL,
	[MC_ContaCorrente] [int] NULL,
	[MC_NumeroCheque] [int] NULL,
	[MC_BomPara] [datetime] NULL,
	[MC_Parcelas] [int] NULL,
	[MC_Remessa] [int] NULL,
	[MC_SituacaoEnvio] [nvarchar](2) NULL,
	[MC_ControleAVR] [char](1) NULL,
	[MC_DataBaixaAVR] [datetime] NULL,
	[MC_Protocolo] [decimal](18, 0) NULL,
	[MC_NroCaixa] [decimal](18, 0) NULL,
	[MC_GrupoAuxiliar] [int] NULL,
	[MC_Situacao] [char](2) NULL,
	[MC_Pedido] [char](8) NULL,
	[MC_DataProcesso] [datetime] NULL,
	[MC_TipoNota] [char](2) NULL,
	[MC_Sequencia] [numeric](18, 0) IDENTITY(1,1) NOT FOR REPLICATION NOT NULL,
	[MC_SequenciaTEF] [int] NULL
) ON [PRIMARY]

GO
/****** Object:  Table [dbo].[News]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[News](
	[NWS_Codigo] [numeric](18, 0) IDENTITY(1,1) NOT NULL,
	[NWS_Usuario] [nchar](10) NULL,
	[NWS_Assunto] [char](200) NULL,
	[NWS_Mensagem] [varchar](8000) NULL,
	[NWS_Data] [datetime] NULL,
	[NWS_Lido] [char](1) NULL,
 CONSTRAINT [PK_News] PRIMARY KEY CLUSTERED 
(
	[NWS_Codigo] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]

GO
/****** Object:  Table [dbo].[NFCapa]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[NFCapa](
	[NUMEROPED] [numeric](18, 0) NOT NULL,
	[DATAEMI] [datetime] NULL,
	[VENDEDOR] [numeric](18, 0) NULL,
	[VLRMERCADORIA] [money] NULL,
	[DESCONTO] [money] NULL,
	[SUBTOTAL] [money] NULL,
	[LOJAORIGEM] [nvarchar](5) NULL,
	[TIPONOTA] [nvarchar](2) NULL,
	[CONDPAG] [nvarchar](4) NULL,
	[AV] [numeric](18, 0) NULL,
	[CLIENTE] [numeric](20, 0) NULL,
	[CODOPER] [numeric](18, 0) NULL,
	[DATAPAG] [datetime] NULL,
	[PGENTRA] [money] NULL,
	[LOJAT] [varchar](5) NULL,
	[QTDITEM] [numeric](18, 0) NULL,
	[PEDCLI] [numeric](18, 0) NULL,
	[TM] [numeric](18, 0) NULL,
	[PESOBR] [numeric](18, 0) NULL,
	[PESOLQ] [numeric](18, 0) NULL,
	[VALFRETE] [money] NULL,
	[FRETECOBR] [money] NULL,
	[OUTRALOJA] [char](5) NULL,
	[OUTROVEND] [numeric](18, 0) NULL,
	[NF] [numeric](18, 0) NULL,
	[TOTALNOTA] [money] NULL,
	[NATOPERACAO] [int] NULL,
	[DATAPED] [datetime] NULL,
	[BASEICMS] [money] NULL,
	[ALIQICMS] [money] NULL,
	[VLRICMS] [money] NULL,
	[SERIE] [varchar](3) NULL,
	[HORA] [datetime] NULL,
	[TOTALIPI] [money] NULL,
	[ECF] [int] NULL,
	[NUMEROSF] [int] NULL,
	[NOMCLI] [varchar](50) NULL,
	[FONECLI] [varchar](50) NULL,
	[CGCCLI] [varchar](50) NULL,
	[INSCRICLI] [varchar](50) NULL,
	[ENDCLI] [varchar](50) NULL,
	[UFCLIENTE] [varchar](50) NULL,
	[MUNICIPIOCLI] [varchar](50) NULL,
	[BAIRROCLI] [varchar](30) NULL,
	[CEPCLI] [varchar](15) NULL,
	[PESSOACLI] [int] NULL,
	[REGIAOCLI] [int] NULL,
	[CFOAUX] [nvarchar](36) NULL,
	[AnexoAUx] [varchar](20) NULL,
	[PAGINANF] [int] NULL,
	[ECFNF] [int] NULL,
	[Carimbo1] [varchar](132) NULL,
	[Carimbo2] [varchar](132) NULL,
	[Carimbo3] [varchar](132) NULL,
	[Carimbo4] [varchar](132) NULL,
	[Carimbo5] [varchar](132) NULL,
	[CustoMedioLiquido] [money] NULL,
	[VendaLiquida] [money] NULL,
	[MargemContribuicao] [money] NULL,
	[ValorTotalCodigoZero] [money] NULL,
	[TotalNotaAlternativa] [money] NULL,
	[ValorMercadoriaAlternativa] [money] NULL,
	[SituacaoEnvio] [varchar](2) NULL,
	[VendedorLojaVenda] [int] NULL,
	[LojaVenda] [varchar](5) NULL,
	[NotaCredito] [int] NULL,
	[NfDevolucao] [int] NULL,
	[SerieDevolucao] [nvarchar](3) NULL,
	[EmiteDataSaida] [char](1) NULL,
	[CancelarNota] [varchar](1) NULL,
	[HoraManutencao] [datetime] NULL,
	[DataProcessamento] [datetime] NULL,
	[SituacaoFec] [varchar](1) NULL,
	[ObsSituacaoFec] [char](50) NULL,
	[CodMunicipioCli] [char](9) NULL,
	[EnderecoNFeCli] [char](60) NULL,
	[EnderecoNroNFeCli] [char](10) NULL,
	[ComplementoNFeCli] [char](60) NULL,
	[InscriSufCli] [char](15) NULL,
	[BaseICMSST] [float] NULL,
	[ValorICMSST] [float] NULL,
	[ValorCOFINS] [float] NULL,
	[ValorOutros] [float] NULL,
	[SenhaDesconto] [char](8) NULL,
	[Volume] [int] NULL,
	[TipoFrete] [int] NULL,
	[ParcelasTEF] [int] NULL,
	[AutorizacaoTEF] [char](50) NULL,
	[GarantiaEstendida] [char](1) NULL,
	[TotalGarantia] [float] NULL,
	[NroResidencia] [varchar](10) NULL,
	[CompleResidencia] [varchar](30) NULL,
	[SeguroPremiado] [float] NULL,
	[CertificadoSorte] [numeric](18, 0) NULL,
	[numeroSorte] [numeric](18, 0) NULL,
	[sp_premioLiquido] [float] NULL,
	[sp_IOF] [float] NULL,
	[sp_valorRemuneracao] [float] NULL,
	[sp_percentualRemuneracao] [float] NULL,
	[sp_valorRepasse] [float] NULL,
	[codmun] [nvarchar](9) NULL,
	[ChaveNFe] [nchar](44) NULL,
	[SituacaoProcesso] [char](1) NULL,
	[DataProcesso] [datetime] NULL,
	[Parcelas] [int] NULL,
	[NroCaixa] [decimal](18, 0) NULL,
	[Protocolo] [decimal](18, 0) NULL,
	[TipoTransporte] [char](60) NULL,
	[Criticaprocesso] [char](60) NULL,
	[CPFNFP] [char](14) NULL,
	[vendedorGarantia] [numeric](18, 0) NULL,
	[ModalidadeVenda] [char](15) NULL,
	[LiberaBloqueio] [nchar](1) NULL,
	[valorSub] [float] NULL,
	[baseSub] [float] NULL,
	[sub] [float] NULL,
	[ChaveNFeDevolucao] [char](44) NULL,
	[valICMSRemet] [float] NULL,
	[valICMSDest] [float] NULL,
	[valorICMSFECP] [float] NULL,
 CONSTRAINT [PK_NFCapa] PRIMARY KEY CLUSTERED 
(
	[NUMEROPED] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]

GO
/****** Object:  Table [dbo].[NFe_cobr]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[NFe_cobr](
	[eLoja] [char](5) NULL,
	[eNF] [char](10) NULL,
	[eSerie] [char](2) NULL,
	[Situacao] [char](1) NULL,
	[nFat] [char](60) NULL,
	[vOrig] [float] NULL,
	[vDesc] [float] NULL,
	[vLiq] [float] NULL,
	[nDup] [char](60) NULL,
	[dVend] [datetime] NULL,
	[vDup] [float] NULL
) ON [PRIMARY]

GO
/****** Object:  Table [dbo].[NFE_controle]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[NFE_controle](
	[eLoja] [char](5) NULL,
	[eNF] [char](10) NULL,
	[eSerie] [char](2) NULL,
	[Situacao] [char](1) NULL,
	[danfe_IMPRESSORA] [varchar](100) NULL,
	[danfe_RETORNARESP] [char](1) NULL,
	[email_DESTINATARIO] [varchar](60) NULL,
	[email_ASSUNTO] [varchar](60) NULL,
	[email_MENSAGEM] [varchar](120) NULL,
	[email_EMAILEMITENTE] [varchar](60) NULL,
	[email_NOMEEMITENTE] [varchar](60) NULL,
	[email_ANEXOPDF] [char](3) NULL,
	[email_ANEXOXML] [char](3) NULL,
	[email_ANEXOPROTOCOLO] [char](3) NULL,
	[email_anexoadicional] [char](3) NULL,
	[email_COMPACTADO] [char](3) NULL,
	[email_RETORNARESP] [char](1) NULL
) ON [PRIMARY]

GO
/****** Object:  Table [dbo].[NFe_dest]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[NFe_dest](
	[eLoja] [char](5) NULL,
	[eNF] [char](10) NULL,
	[eSerie] [char](2) NULL,
	[Situacao] [char](1) NULL,
	[CNPJ] [char](15) NULL,
	[CPF] [char](15) NULL,
	[xNome] [char](60) NULL,
	[xLgr] [char](60) NULL,
	[Nro] [char](10) NULL,
	[xCpl] [char](60) NULL,
	[xBairro] [char](60) NULL,
	[cMun] [char](7) NULL,
	[xMun] [char](60) NULL,
	[UF] [char](2) NULL,
	[CEP] [char](8) NULL,
	[cPais] [char](4) NULL,
	[xPais] [char](60) NULL,
	[fone] [char](12) NULL,
	[IE] [char](15) NULL,
	[ISUF] [char](15) NULL,
	[email] [char](60) NULL,
	[INDIEDEST] [char](1) NULL
) ON [PRIMARY]

GO
/****** Object:  Table [dbo].[NFE_dup]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[NFE_dup](
	[eLoja] [char](5) NULL,
	[eNF] [char](10) NULL,
	[eSerie] [char](2) NULL,
	[Situacao] [char](1) NULL,
	[nDup] [varchar](60) NULL,
	[dVend] [datetime] NULL,
	[vDup] [float] NULL
) ON [PRIMARY]

GO
/****** Object:  Table [dbo].[NFe_emit]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[NFe_emit](
	[eLoja] [char](5) NULL,
	[eNF] [char](10) NULL,
	[eSerie] [char](2) NULL,
	[Situacao] [char](1) NULL,
	[CNPJ] [char](15) NULL,
	[xNome] [char](60) NULL,
	[xFant] [char](20) NULL,
	[xLgr] [char](60) NULL,
	[nro] [char](60) NULL,
	[xCpl] [char](60) NULL,
	[xBairro] [char](60) NULL,
	[cMun] [char](10) NULL,
	[xMun] [char](60) NULL,
	[UF] [char](10) NULL,
	[CEP] [char](8) NULL,
	[cPais] [char](4) NULL,
	[xPais] [char](60) NULL,
	[fone] [char](12) NULL,
	[IE] [char](15) NULL,
	[IEST] [char](15) NULL,
	[IM] [char](14) NULL,
	[CNAE] [char](7) NULL,
	[CRT] [char](1) NULL
) ON [PRIMARY]

GO
/****** Object:  Table [dbo].[nfe_estrutura]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[nfe_estrutura](
	[ETR_Sequencia] [numeric](18, 0) NOT NULL,
	[ETR_Rotulo] [nvarchar](255) NULL,
	[ETR_Campo] [nvarchar](255) NULL,
	[ETR_Tabela_DE] [nvarchar](255) NULL,
	[ETR_Campo_DE] [nvarchar](255) NULL,
 CONSTRAINT [PK_nfe_estrutura] PRIMARY KEY CLUSTERED 
(
	[ETR_Sequencia] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]

GO
/****** Object:  Table [dbo].[NFE_fat]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[NFE_fat](
	[eLoja] [char](5) NULL,
	[eNF] [char](10) NULL,
	[eSerie] [char](2) NULL,
	[Situacao] [char](1) NULL,
	[nFat] [varchar](60) NULL,
	[vOrig] [float] NULL,
	[vDesc] [float] NULL,
	[vLiq] [float] NULL
) ON [PRIMARY]

GO
/****** Object:  Table [dbo].[NFe_ide]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[NFe_ide](
	[eLoja] [char](5) NULL,
	[eNF] [char](10) NULL,
	[eSerie] [char](2) NULL,
	[Situacao] [char](1) NULL,
	[cUF] [char](2) NULL,
	[cNF] [char](9) NULL,
	[natOp] [char](60) NULL,
	[indPag] [char](1) NULL,
	[mod] [char](2) NULL,
	[serie] [char](3) NULL,
	[nNF] [char](9) NULL,
	[dEmi] [datetime] NULL,
	[dSaiEnt] [datetime] NULL,
	[hSaiEnt] [datetime] NULL,
	[tpNF] [char](1) NULL,
	[cMunFG] [char](7) NULL,
	[tpImp] [char](1) NULL,
	[tpEmis] [char](1) NULL,
	[cDV] [char](1) NULL,
	[tpAmb] [char](1) NULL,
	[finNFe] [char](1) NULL,
	[procEmi] [char](1) NULL,
	[verProc] [char](20) NULL,
	[dhCont] [datetime] NULL,
	[xJust] [char](256) NULL,
	[ChaveAcesso] [char](44) NULL,
	[refNFe] [char](44) NULL,
	[IDDEST] [char](1) NULL,
	[INDFINAL] [char](1) NULL,
	[INDPRES] [char](1) NULL
) ON [PRIMARY]

GO
/****** Object:  Table [dbo].[NFe_infAdic]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[NFe_infAdic](
	[eLoja] [char](5) NULL,
	[eNF] [char](10) NULL,
	[eSerie] [char](2) NULL,
	[Situacao] [char](1) NULL,
	[infAdFisco] [char](256) NULL,
	[infCpl] [char](5000) NULL,
	[xCampoCont] [char](20) NULL,
	[xTextoCont] [char](60) NULL,
	[xCampoFisco] [char](20) NULL,
	[xTextoFisco] [char](60) NULL,
	[nProc] [char](60) NULL,
	[indProc] [char](1) NULL
) ON [PRIMARY]

GO
/****** Object:  Table [dbo].[NFE_NFLojas]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[NFE_NFLojas](
	[NFL_Sequencia] [int] NULL,
	[NFL_Descricao] [char](18) NULL,
	[NFL_Dados] [char](2000) NULL,
	[NFL_Loja] [char](5) NULL,
	[NFL_NroNFE] [numeric](18, 0) NULL,
	[NFL_DataEmissao] [char](10) NULL
) ON [PRIMARY]

GO
/****** Object:  Table [dbo].[NFe_prod]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[NFe_prod](
	[eLoja] [char](5) NULL,
	[eNF] [char](10) NULL,
	[eSerie] [char](2) NULL,
	[Situacao] [char](1) NULL,
	[H_nItem] [int] NULL,
	[I_cProd] [char](60) NULL,
	[I_cEAN] [char](15) NULL,
	[I_xProd] [char](120) NULL,
	[I_NCM] [char](10) NULL,
	[I_EXTIPI] [char](3) NULL,
	[I_CFOP] [char](4) NULL,
	[I_uCom] [char](6) NULL,
	[I_qCom] [float] NULL,
	[I_vUnCom] [float] NULL,
	[I_vProd] [float] NULL,
	[I_cEANTrib] [char](14) NULL,
	[I_uTrib] [char](6) NULL,
	[I_qTrib] [float] NULL,
	[I_vUnTrib] [float] NULL,
	[I_vFrete] [float] NULL,
	[I_vSeg] [float] NULL,
	[I_vDesc] [float] NULL,
	[I_vOutro] [float] NULL,
	[I_indTot] [char](1) NULL,
	[N_origICMS] [char](1) NULL,
	[N_CSTICMS] [char](2) NULL,
	[N_modBCICMS] [char](1) NULL,
	[N_vBCICMS] [float] NULL,
	[N_pRedBCICMS] [float] NULL,
	[N_pICMS] [float] NULL,
	[N_vICMS] [float] NULL,
	[N_modBCST] [char](1) NULL,
	[N_pMVAST] [float] NULL,
	[N_pRedBCST] [float] NULL,
	[N_vBCST] [float] NULL,
	[N_pICMSST] [float] NULL,
	[N_vICMSST] [float] NULL,
	[O_cIEnq] [char](5) NULL,
	[O_CNPJProd] [char](14) NULL,
	[O_cSelo] [char](60) NULL,
	[O_qSelo] [char](12) NULL,
	[O_cEnq] [char](3) NULL,
	[O_CSTIPI] [char](2) NULL,
	[O_vBCIPI] [float] NULL,
	[O_qUnid] [float] NULL,
	[O_vUnid] [float] NULL,
	[O_pIPI] [float] NULL,
	[O_vIPI] [float] NULL,
	[O_CSTIPINT] [char](2) NULL,
	[P_vBCII] [float] NULL,
	[P_vDespAdu] [float] NULL,
	[P_vII] [float] NULL,
	[P_vIOF] [float] NULL,
	[Q_CSTPIS] [char](2) NULL,
	[Q_vBCPIS] [float] NULL,
	[Q_pPIS] [float] NULL,
	[Q_qBCProdPIS] [float] NULL,
	[Q_vAliqProdPIS] [float] NULL,
	[Q_vPIS] [float] NULL,
	[R_vBCPISST] [float] NULL,
	[R_pPISST] [float] NULL,
	[R_qBCProdPISST] [float] NULL,
	[R_vAliqProdPISST] [float] NULL,
	[R_vPISST] [float] NULL,
	[S_CSTCOFINS] [char](2) NULL,
	[S_vBCCOFINS] [float] NULL,
	[S_pCOFINS] [float] NULL,
	[S_qBCProdCOFINS] [float] NULL,
	[S_vAliqProdCOFINS] [float] NULL,
	[S_vCOFINS] [float] NULL,
	[T_vBCCOFINSST] [float] NULL,
	[T_pCOFINSST] [float] NULL,
	[T_qBCProdCOFINSST] [float] NULL,
	[T_vAliqProdCOFINSST] [float] NULL,
	[T_vCOFINSST] [float] NULL,
	[U_vBCISSQN] [float] NULL,
	[U_vAliqISSQN] [float] NULL,
	[U_vISSQN] [float] NULL,
	[U_cMunFGISSQN] [char](7) NULL,
	[U_cListServ] [char](4) NULL,
	[U_cSitTrib] [char](1) NULL,
	[V_infAdProd] [char](500) NULL,
	[W_vBCUFDEST] [float] NULL,
	[W_pFCPUFDEST] [float] NULL,
	[W_pICMSUFDEST] [float] NULL,
	[W_pICMSINTER] [float] NULL,
	[W_pICMSINTERPART] [float] NULL,
	[W_vFCPUFDEST] [float] NULL,
	[W_vICMSUFDEST] [float] NULL,
	[W_vICMSUFREMET] [float] NULL,
	[vVICMSDESON] [float] NULL,
	[vBCUFDEST] [float] NULL,
	[I_CEST] [char](7) NULL,
	[X_orig] [char](1) NULL,
	[X_CSOSN] [char](3) NULL
) ON [PRIMARY]

GO
/****** Object:  Table [dbo].[nfe_total]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[nfe_total](
	[eLoja] [char](5) NULL,
	[eNF] [char](10) NULL,
	[eSerie] [char](2) NULL,
	[Situacao] [char](1) NULL,
	[vBCICMS] [float] NULL,
	[vICMS] [float] NULL,
	[vBCST] [float] NULL,
	[vST] [float] NULL,
	[vProd] [float] NULL,
	[vFrete] [float] NULL,
	[vSeg] [float] NULL,
	[vDesc] [numeric](8, 2) NULL,
	[vII] [float] NULL,
	[vIPI] [float] NULL,
	[vCOFINS] [float] NULL,
	[vOutro] [float] NULL,
	[vNF] [float] NULL,
	[vServ] [float] NULL,
	[vBCISSQ] [float] NULL,
	[vISS] [float] NULL,
	[vPIS] [float] NULL,
	[vCOFINsISSQ] [float] NULL,
	[vRetPIS] [float] NULL,
	[vRetCOFINS] [float] NULL,
	[vRetCSLL] [float] NULL,
	[vBCIRRF] [float] NULL,
	[vIRRF] [float] NULL,
	[vBCRetPrev] [float] NULL,
	[vRetPrev] [float] NULL,
	[vTOTTRIB] [float] NULL,
	[vFCPUFDEST] [float] NULL,
	[vICMSUFDEST] [float] NULL,
	[vICMSUFREMET] [float] NULL,
	[vVICMSDESON] [float] NULL
) ON [PRIMARY]

GO
/****** Object:  Table [dbo].[NFe_transp]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[NFe_transp](
	[eLoja] [char](5) NULL,
	[eNF] [char](10) NULL,
	[eSerie] [char](2) NULL,
	[Situacao] [char](1) NULL,
	[modFrete] [char](1) NULL,
	[CNPJ] [char](14) NULL,
	[CPF] [char](11) NULL,
	[xNome] [char](60) NULL,
	[IE] [char](14) NULL,
	[xEnder] [char](60) NULL,
	[xMun] [char](60) NULL,
	[UF] [char](2) NULL,
	[vServ] [float] NULL,
	[vBCRet] [float] NULL,
	[pICMSRet] [float] NULL,
	[vICMSRet] [float] NULL,
	[CFOP] [char](4) NULL,
	[cMunFG] [char](7) NULL,
	[placa] [char](8) NULL,
	[UFveic] [char](2) NULL,
	[RNTC] [char](20) NULL,
	[qVol] [char](15) NULL,
	[esq] [char](60) NULL,
	[marca] [char](60) NULL,
	[nVol] [char](60) NULL,
	[pesoL] [float] NULL,
	[pesoB] [float] NULL,
	[lacres] [char](60) NULL,
	[nLacres] [char](60) NULL
) ON [PRIMARY]

GO
/****** Object:  Table [dbo].[NFItens]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[NFItens](
	[NUMEROPED] [numeric](18, 0) NOT NULL,
	[DATAEMI] [datetime] NULL,
	[REFERENCIA] [nvarchar](7) NOT NULL,
	[QTDE] [float] NULL,
	[VLUNIT] [money] NULL,
	[VLUNIT2] [money] NULL,
	[VLTOTITEM] [money] NULL,
	[DESCRAT] [money] NULL,
	[ICMS] [money] NOT NULL,
	[ITEM] [int] NOT NULL,
	[VLIPI] [money] NULL,
	[DESCONTO] [money] NULL,
	[PLISTA] [money] NULL,
	[COMISSAO] [money] NULL,
	[VALORICMS] [money] NULL,
	[BCOMIS] [money] NULL,
	[CSPROD] [numeric](18, 0) NULL,
	[LINHA] [numeric](18, 0) NULL,
	[SECAO] [numeric](18, 0) NULL,
	[VBUNIT] [numeric](18, 0) NULL,
	[ICMPDV] [numeric](18, 0) NULL,
	[CODBARRA] [nvarchar](20) NULL,
	[NF] [numeric](18, 0) NULL,
	[SERIE] [nvarchar](3) NULL,
	[LOJAORIGEM] [nvarchar](5) NULL,
	[CLIENTE] [numeric](10, 0) NULL,
	[VENDEDOR] [int] NULL,
	[ALIQIPI] [money] NULL,
	[TIPONOTA] [nvarchar](2) NULL,
	[REDUCAOICMS] [money] NULL,
	[BASEICMS] [money] NULL,
	[TIPOMOVIMENTACAO] [int] NULL,
	[DETALHEIMPRESSAO] [nvarchar](1) NULL,
	[SerieProd1] [nvarchar](50) NULL,
	[SerieProd2] [nvarchar](50) NULL,
	[CustoMedioLiquido] [money] NULL,
	[VendaLiquida] [money] NULL,
	[MargemContribuicao] [real] NULL,
	[EncargosVendaLiquida] [money] NULL,
	[EncargosCustoMedioLiquido] [money] NULL,
	[PrecoUnitAlternativa] [money] NULL,
	[ValorMercadoriaAlternativa] [money] NULL,
	[ReferenciaAlternativa] [nvarchar](8) NULL,
	[SituacaoEnvio] [nvarchar](1) NULL,
	[DescricaoAlternativa] [char](50) NULL,
	[Tributacao] [char](3) NULL,
	[IcmsMargem] [float] NULL,
	[PisCofins] [float] NULL,
	[DeducoesVendas] [float] NULL,
	[EncargosFinanceiros] [float] NULL,
	[EstoqueAntes] [int] NULL,
	[EstoqueDepois] [int] NULL,
	[CFOP] [char](4) NULL,
	[CSTICMS] [char](2) NULL,
	[GarantiaEstendida] [char](1) NULL,
	[PlanoGarantia] [int] NULL,
	[CoeficientePlano] [float] NULL,
	[QtdeGarantia] [int] NULL,
	[ValorGarantia] [float] NULL,
	[CertificadoInicio] [char](12) NULL,
	[CertificadoFim] [char](12) NULL,
	[ge_premioLiquido] [float] NULL,
	[ge_IOF] [float] NULL,
	[ge_dataInicioVigencia] [datetime] NULL,
	[ge_dataFinalVigencia] [datetime] NULL,
	[ge_valorCustoSeguradora] [float] NULL,
	[ge_seqCancelamento] [int] NULL,
	[ge_dataCancelamento] [datetime] NULL,
	[SituacaoProcesso] [char](1) NULL,
	[dataprocesso] [datetime] NULL,
	[ICMSAplicado] [float] NULL,
	[Parcelas] [float] NULL,
	[valorSub] [float] NULL,
	[baseSub] [float] NULL,
	[sub] [float] NULL,
	[baseIPI] [float] NULL,
	[aliqICMSDest] [float] NULL,
	[aliqICMSInter] [float] NULL,
	[ICMSInterpart] [float] NULL,
	[valICMSRemet] [float] NULL,
	[valICMSDest] [float] NULL,
	[valorICMSFECP] [float] NULL,
	[aliqICMSFECP] [float] NULL,
	[Cest] [char](7) NULL,
 CONSTRAINT [PK_NFItens_1] PRIMARY KEY CLUSTERED 
(
	[NUMEROPED] ASC,
	[ITEM] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]

GO
/****** Object:  Table [dbo].[ParametroCaixa]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[ParametroCaixa](
	[PAR_Loja] [char](5) NULL,
	[PAR_NroCaixa] [int] NULL,
	[PAR_NroECF] [int] NULL,
	[PAR_CaixaECF] [char](1) NULL,
	[PAR_CaixaSN] [char](1) NULL,
	[PAR_Caixa00] [char](1) NULL,
	[PAR_CaixaSM] [char](1) NULL,
	[PAR_ECFPedido] [char](1) NULL
) ON [PRIMARY]

GO
/****** Object:  Table [dbo].[produto]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[produto](
	[PR_Referencia] [char](7) NOT NULL,
	[PR_CodigoFornecedor] [smallint] NOT NULL,
	[PR_CodigoBarra] [varchar](20) NULL,
	[PR_Descricao] [varchar](38) NOT NULL,
	[PR_DataCadastro] [datetime] NOT NULL,
	[PR_Linha] [smallint] NOT NULL,
	[PR_Secao] [smallint] NOT NULL,
	[PR_Classe] [char](1) NOT NULL,
	[PR_Bloqueio] [char](1) NOT NULL,
	[PR_ClasseABC] [char](3) NULL,
	[PR_ClasseFiscal] [varchar](10) NULL,
	[PR_Unidade] [varchar](2) NOT NULL,
	[PR_UnidadeDistribuicao] [smallint] NOT NULL,
	[PR_PercentualComissao] [float] NOT NULL,
	[PR_ICMSEntrada] [float] NOT NULL,
	[PR_ICMSSaida] [float] NOT NULL,
	[PR_AliquotaIPI] [float] NOT NULL,
	[PR_CodigoIPI] [tinyint] NOT NULL,
	[PR_CodigoReducaoICMS] [tinyint] NULL,
	[PR_PrecoFornecedor] [float] NOT NULL,
	[PR_DescontoFornecedor] [float] NOT NULL,
	[PR_ItemCondicoesGerais] [smallint] NULL,
	[PR_PercentualFrete] [float] NOT NULL,
	[PR_PercentualEmbalagem] [float] NOT NULL,
	[PR_CustoMedio1] [float] NOT NULL,
	[PR_CustoMedio2] [float] NOT NULL,
	[PR_CustoMedio3] [float] NOT NULL,
	[PR_CustoMedioLiquido1] [float] NOT NULL,
	[PR_CustoMedioLiquido2] [float] NOT NULL,
	[PR_CustoMedioLiquido3] [float] NOT NULL,
	[PR_PrecoCusto1] [float] NOT NULL,
	[PR_PrecoCusto2] [float] NOT NULL,
	[PR_PrecoCusto3] [float] NOT NULL,
	[PR_CustoLiquido1] [float] NOT NULL,
	[PR_CustoLiquido2] [float] NOT NULL,
	[PR_CustoLiquido3] [float] NOT NULL,
	[PR_PrecoEntrada1] [float] NOT NULL,
	[PR_PrecoEntrada2] [float] NOT NULL,
	[PR_PrecoEntrada3] [float] NOT NULL,
	[PR_DataPrecoCusto1] [datetime] NULL,
	[PR_DataPrecoCusto2] [datetime] NULL,
	[PR_DataPrecoCusto3] [datetime] NULL,
	[PR_PrecoVenda1] [float] NOT NULL,
	[PR_PrecoVenda2] [float] NOT NULL,
	[PR_PrecoVenda3] [float] NOT NULL,
	[PR_DataPrecoVenda1] [datetime] NULL,
	[PR_DataPrecoVenda2] [datetime] NULL,
	[PR_DataPrecoVenda3] [datetime] NULL,
	[PR_PrecoVendaObjetivo] [float] NOT NULL,
	[PR_PaginaListaPreco] [smallint] NOT NULL,
	[PR_Peso] [float] NOT NULL,
	[PR_MenorUnidadeCompra] [smallint] NOT NULL,
	[PR_MetodoCompra] [tinyint] NOT NULL,
	[PR_TipoCalculoReposicao] [tinyint] NOT NULL,
	[PR_MetodoDistribuicao] [tinyint] NOT NULL,
	[PR_Residencia] [tinyint] NOT NULL,
	[PR_MargemObjetiva] [float] NOT NULL,
	[PR_MargemPrevista] [float] NOT NULL,
	[PR_Markup] [float] NOT NULL,
	[PR_Comprador] [smallint] NOT NULL,
	[PR_EmiteEtiqueta] [char](1) NULL,
	[PR_Situacao] [char](1) NOT NULL,
	[PR_SubstituicaoTributaria] [char](1) NULL,
	[PR_DeducoesVenda] [float] NULL,
	[PR_IcmPdv] [float] NULL,
	[PR_DescricaoPDV] [varchar](30) NULL,
	[PR_Grupo] [smallint] NULL,
	[PR_Complemento] [varchar](50) NULL,
	[PR_Sazonal] [char](15) NULL,
	[PR_CodigoProdutoNoFornecedor] [varchar](25) NULL,
	[PR_PrecoVendaLiquido1] [float] NULL,
	[PR_PrecoVendaLiquido2] [float] NULL,
	[PR_PrecoVendaLiquido3] [float] NULL,
	[PR_HoraManutencao] [datetime] NULL,
	[PR_PrecoVendaSemIcms1] [float] NULL,
	[PR_PrecoVendaSemIcms2] [float] NULL,
	[PR_PrecoVendaSemIcms3] [float] NULL,
	[Pr_UltimoProduto] [char](1) NULL,
	[Pr_IpiCalculado] [float] NOT NULL,
	[PR_CodigoReducaoIcmsEntrada] [tinyint] NOT NULL,
	[Pr_IcmPdvEntrada] [float] NULL,
	[Pr_ProdutoInternet] [char](1) NULL,
	[PR_CustomediosemIPI] [float] NULL,
	[PR_ValorICMSCustoMediosemIPI] [float] NULL,
	[PR_ValorPisCofins] [float] NULL,
	[PR_DiaMesContagem] [char](6) NULL,
	[c1] [float] NULL,
	[c2] [float] NULL,
	[c3] [float] NULL,
	[c4] [float] NULL,
	[c5] [float] NULL,
	[c6] [float] NULL,
	[PR_LinhaProduto] [char](8) NULL,
	[PR_IVA] [float] NULL,
	[PR_ICMSEntradaIVA] [float] NULL,
	[PR_ICMSPDVEntradaIVA] [float] NULL,
	[PR_CodigoReducaoIcmsEntradaIVA] [tinyint] NULL,
	[PR_CodigoMercadoriaST] [varchar](3) NULL,
	[PR_ICMSSaidaIVA] [float] NULL,
	[PR_ICMSPDVSaidaIVA] [float] NULL,
	[PR_CodigoReducaoIcmsSaidaIVA] [tinyint] NULL,
	[PR_ValorICMSIVA] [float] NULL,
	[PR_MortalidadeProduto] [char](1) NULL,
	[PR_CustoContabil1] [float] NULL,
	[PR_CustoContabil2] [float] NULL,
	[PR_CustoContabil3] [float] NULL,
	[PR_CustoContabilLiquido1] [float] NULL,
	[PR_CustoContabilLiquido2] [float] NULL,
	[PR_CustoContabilLiquido3] [float] NULL,
	[PR_GarantiaFabricante] [numeric](18, 0) NULL,
	[PR_GarantiaEstendida] [char](1) NULL,
	[PR_PrecoUnitarioNFCompra1] [float] NULL,
	[PR_PrecoUnitarioNFCompra2] [float] NULL,
	[PR_PrecoUnitarioNFCompra3] [float] NULL,
	[PR_IndicePreco] [char](3) NULL,
	[PR_CST] [char](3) NULL,
	[PR_PrecoPromocao] [float] NULL,
	[PR_DescricaoFornecedor] [char](100) NULL,
	[PR_Comprimento] [float] NULL,
	[PR_Largura] [float] NULL,
	[PR_Altura] [float] NULL
) ON [PRIMARY]

GO
/****** Object:  Table [dbo].[produtoBarras]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[produtoBarras](
	[PRB_Referencia] [char](7) NULL,
	[PRB_CodigoBarras] [char](15) NOT NULL,
	[PRB_CodigoFornecedor] [numeric](18, 0) NOT NULL,
	[PRB_Embalagem] [numeric](18, 0) NULL,
	[PRB_HoraManutencao] [datetime] NULL,
	[PRB_TipoCodigo] [char](1) NOT NULL,
PRIMARY KEY CLUSTERED 
(
	[PRB_CodigoBarras] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]

GO
/****** Object:  Table [dbo].[ProdutoDescricao]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[ProdutoDescricao](
	[PRO_CODIGO] [float] NULL,
	[PRO_REFERENCIA] [nvarchar](7) NOT NULL,
	[PRO_DESCR_LONGA] [nvarchar](max) NULL,
	[PRO_ITENS_INCLUSOS] [nvarchar](max) NULL,
	[PRO_ESPECIFICACAO_SITE] [nvarchar](max) NULL,
PRIMARY KEY CLUSTERED 
(
	[PRO_REFERENCIA] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]

GO
/****** Object:  Table [dbo].[produtoLoja]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[produtoLoja](
	[pr_Referencia] [char](7) NOT NULL,
	[pr_CodigoFornecedor] [smallint] NOT NULL,
	[pr_Descricao] [varchar](38) NOT NULL,
	[pr_Classe] [char](1) NOT NULL,
	[pr_Bloqueio] [char](1) NOT NULL,
	[pr_LinhaProduto] [char](8) NULL,
	[pr_ClasseFiscal] [varchar](10) NULL,
	[pr_Unidade] [varchar](2) NOT NULL,
	[pr_ICMSSaida] [float] NOT NULL,
	[pr_CodigoReducaoICMS] [tinyint] NULL,
	[pr_CustoMedio1] [float] NOT NULL,
	[pr_PrecoVenda1] [float] NOT NULL,
	[pr_PaginaListaPreco] [smallint] NOT NULL,
	[pr_Peso] [float] NOT NULL,
	[pr_Comprador] [smallint] NOT NULL,
	[pr_Situacao] [char](1) NOT NULL,
	[pr_SubstituicaoTributaria] [char](1) NULL,
	[pr_IcmPdv] [float] NULL,
	[pr_HoraManutencao] [datetime] NULL,
	[pr_CodigoProdutoNoFornecedor] [varchar](25) NULL,
	[pr_IcmsSaidaIva] [float] NULL,
	[pr_IcmsPdvSaidaIva] [float] NULL,
	[pr_ICMSEntrada] [float] NOT NULL,
	[pr_IcmPdvEntrada] [float] NULL,
	[pr_ST] [char](3) NULL,
	[pr_CST] [char](3) NULL,
	[pr_GarantiaEstendida] [char](1) NULL,
	[pr_GarantiaFabricante] [numeric](18, 0) NULL,
	[pr_IndicePreco] [char](3) NULL,
	[PR_precoVendaLiquido1] [float] NULL,
	[PR_custoMedioLiquido1] [float] NOT NULL,
	[PR_PrecoCusto1] [float] NOT NULL,
	[pr_cest] [char](7) NULL,
PRIMARY KEY CLUSTERED 
(
	[pr_Referencia] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]

GO
/****** Object:  Table [dbo].[SAT_NF]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[SAT_NF](
	[snf_Sequencia] [int] IDENTITY(1,1) NOT NULL,
	[snf_Descricao] [char](18) NULL,
	[snf_Sinal] [char](1) NULL,
	[snf_Dados] [varchar](2000) NULL,
	[snf_pedido] [numeric](18, 0) NULL
) ON [PRIMARY]

GO
/****** Object:  Table [dbo].[tempCalculoICMS]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[tempCalculoICMS](
	[BasedeCalculoICMS] [float] NULL,
	[ValorCalculadoICMS] [float] NULL,
	[Tributacao] [char](3) NULL,
	[icmsdestino] [float] NULL
) ON [PRIMARY]

GO
/****** Object:  Table [dbo].[tipoIcms]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[tipoIcms](
	[tpi_codigo] [char](3) NOT NULL,
	[tpi_descricao] [varchar](100) NULL
) ON [PRIMARY]

GO
/****** Object:  Table [dbo].[Usuariocaixa]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Usuariocaixa](
	[USU_Codigo] [numeric](18, 0) NOT NULL,
	[USU_Nome] [char](25) NULL,
	[USU_TipoUsuario] [char](1) NULL,
	[USU_Senha] [char](15) NULL,
	[USU_Situacao] [char](1) NULL
) ON [PRIMARY]

GO
/****** Object:  Table [dbo].[Vende]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Vende](
	[VE_Codigo] [float] NULL,
	[VE_TotalVenda] [money] NULL,
	[VE_MargemVenda] [money] NULL,
	[VE_Nome] [nvarchar](20) NULL,
	[VE_Senha] [char](15) NULL
) ON [PRIMARY]

GO
/****** Object:  Table [dbo].[vende_detalhe]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[vende_detalhe](
	[VDE_CODIGO] [real] NOT NULL,
	[VDE_EMAIL] [varchar](100) NULL,
	[VDE_ASSINATURA] [varchar](max) NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]

GO
ALTER TABLE [dbo].[CodigoOperacao] ADD  DEFAULT ((0)) FOR [CF_CodigoOperacaoNovo]
GO
ALTER TABLE [dbo].[CondicaoPagto] ADD  CONSTRAINT [DF__CondicaoP__CP_In__16451E08]  DEFAULT ((0)) FOR [CP_Intervalo]
GO
ALTER TABLE [dbo].[ControleSistema] ADD  CONSTRAINT [DF_ControleSistema_CTS_LiberaBloqueio]  DEFAULT ('N') FOR [CTS_LiberaBloqueio]
GO
ALTER TABLE [dbo].[Duplicata] ADD  CONSTRAINT [DF__Duplicata__DP_Ab__1D072A30]  DEFAULT ((0)) FOR [DP_Abatimento]
GO
ALTER TABLE [dbo].[Duplicata] ADD  CONSTRAINT [DF__Duplicata__DP_De__1DFB4E69]  DEFAULT ((0)) FOR [DP_Desconto]
GO
ALTER TABLE [dbo].[Duplicata] ADD  CONSTRAINT [DF__Duplicata__DP_De__1EEF72A2]  DEFAULT ((0)) FOR [DP_Despesas]
GO
ALTER TABLE [dbo].[Duplicata] ADD  CONSTRAINT [DF__Duplicata__DP_Ju__1FE396DB]  DEFAULT ((0)) FOR [DP_Juros]
GO
ALTER TABLE [dbo].[Duplicata] ADD  CONSTRAINT [DF__Duplicata__DP_Va__20D7BB14]  DEFAULT ((0)) FOR [DP_ValorPago]
GO
ALTER TABLE [dbo].[Duplicata] ADD  CONSTRAINT [DF__Duplicata__DP_Si__21CBDF4D]  DEFAULT ('A') FOR [DP_Situacao]
GO
ALTER TABLE [dbo].[EstoqueLoja] ADD  CONSTRAINT [CS_EL_NaoComercializado]  DEFAULT ((0)) FOR [EL_NaoComercializado]
GO
ALTER TABLE [dbo].[EstoqueLoja] ADD  CONSTRAINT [CS_EL_NaoComercializadoCONSO]  DEFAULT ((0)) FOR [EL_NaoComercializadoCONSO]
GO
ALTER TABLE [dbo].[FIN_Cliente] ADD  CONSTRAINT [DF_fin_cliente_CE_Loja]  DEFAULT ('999') FOR [CE_Loja]
GO
ALTER TABLE [dbo].[locais] ADD  CONSTRAINT [DF_locais_CL_EmiteFolha]  DEFAULT ((0)) FOR [CL_EmiteFolha]
GO
ALTER TABLE [dbo].[locais] ADD  CONSTRAINT [DF_locais_CL_Contagem]  DEFAULT ((0)) FOR [CL_Contagem]
GO
ALTER TABLE [dbo].[locais] ADD  CONSTRAINT [DF_locais_CL_Situacao]  DEFAULT (' ') FOR [CL_Situacao]
GO
ALTER TABLE [dbo].[locais] ADD  CONSTRAINT [DF_locais_CL_Dupla1]  DEFAULT ((0)) FOR [CL_Dupla1]
GO
ALTER TABLE [dbo].[locais] ADD  CONSTRAINT [DF_locais_CL_Dupla2]  DEFAULT ((0)) FOR [CL_Dupla2]
GO
ALTER TABLE [dbo].[locais] ADD  CONSTRAINT [DF_locais_CL_Dupla3]  DEFAULT ((0)) FOR [CL_Dupla3]
GO
ALTER TABLE [dbo].[locais] ADD  CONSTRAINT [DF_locais_CL_Dupla4]  DEFAULT ((0)) FOR [CL_Dupla4]
GO
ALTER TABLE [dbo].[locais] ADD  CONSTRAINT [DF_locais_CL_Dupla5]  DEFAULT ((0)) FOR [CL_Dupla5]
GO
ALTER TABLE [dbo].[locais] ADD  CONSTRAINT [DF_locais_CL_Digitador1]  DEFAULT ((0)) FOR [CL_Digitador1]
GO
ALTER TABLE [dbo].[locais] ADD  CONSTRAINT [DF_locais_CL_Digitador2]  DEFAULT ((0)) FOR [CL_Digitador2]
GO
ALTER TABLE [dbo].[locais] ADD  CONSTRAINT [DF_locais_CL_Digitador3]  DEFAULT ((0)) FOR [CL_Digitador3]
GO
ALTER TABLE [dbo].[locais] ADD  CONSTRAINT [DF_locais_CL_Digitador4]  DEFAULT ((0)) FOR [CL_Digitador4]
GO
ALTER TABLE [dbo].[locais] ADD  CONSTRAINT [DF_locais_CL_Digitador5]  DEFAULT ((0)) FOR [CL_Digitador5]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_NUMEROPED]  DEFAULT ((0)) FOR [NUMEROPED]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_DATAEMI]  DEFAULT ('') FOR [DATAEMI]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_VENDEDOR]  DEFAULT ((0)) FOR [VENDEDOR]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_VLRMERCADORIA]  DEFAULT ((0)) FOR [VLRMERCADORIA]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_DESCONTO]  DEFAULT ((0)) FOR [DESCONTO]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_SUBTOTAL]  DEFAULT ((0)) FOR [SUBTOTAL]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_LOJAORIGEM]  DEFAULT ('') FOR [LOJAORIGEM]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_TIPONOTA]  DEFAULT ('') FOR [TIPONOTA]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_CONDPAG]  DEFAULT ('01') FOR [CONDPAG]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_AV]  DEFAULT ((0)) FOR [AV]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_CLIENTE]  DEFAULT ((0)) FOR [CLIENTE]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_CODOPER]  DEFAULT ((0)) FOR [CODOPER]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_DATAPAG]  DEFAULT ('') FOR [DATAPAG]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_PGENTRA]  DEFAULT ((0)) FOR [PGENTRA]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_LOJAT]  DEFAULT ('999') FOR [LOJAT]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_QTDITEM]  DEFAULT ((0)) FOR [QTDITEM]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_PEDCLI]  DEFAULT ((0)) FOR [PEDCLI]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_TM]  DEFAULT ((0)) FOR [TM]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_PESOBR]  DEFAULT ((0)) FOR [PESOBR]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_PESOLQ]  DEFAULT ((0)) FOR [PESOLQ]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_VALFRETE]  DEFAULT ((0)) FOR [VALFRETE]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_FRETECOBR]  DEFAULT ((0)) FOR [FRETECOBR]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_OUTRALOJA]  DEFAULT ('') FOR [OUTRALOJA]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_OUTROVEND]  DEFAULT ((0)) FOR [OUTROVEND]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_NF]  DEFAULT ((0)) FOR [NF]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_TOTALNOTA]  DEFAULT ((0)) FOR [TOTALNOTA]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_NATOPERACAO]  DEFAULT ((0)) FOR [NATOPERACAO]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_DATAPED]  DEFAULT ((0)) FOR [DATAPED]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_BASEICMS]  DEFAULT ((0)) FOR [BASEICMS]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_ALIQICMS]  DEFAULT ((0)) FOR [ALIQICMS]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_VLRICMS]  DEFAULT ((0)) FOR [VLRICMS]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_SERIE]  DEFAULT ('') FOR [SERIE]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_HORA]  DEFAULT ('') FOR [HORA]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_TOTALIPI]  DEFAULT ((0)) FOR [TOTALIPI]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_ECF]  DEFAULT ((0)) FOR [ECF]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_NUMEROSF]  DEFAULT ((0)) FOR [NUMEROSF]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_NOMCLI]  DEFAULT ('') FOR [NOMCLI]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_FONECLI]  DEFAULT ('') FOR [FONECLI]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_CGCCLI]  DEFAULT ('') FOR [CGCCLI]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_INSCRICLI]  DEFAULT ('') FOR [INSCRICLI]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_ENDCLI]  DEFAULT ('') FOR [ENDCLI]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_UFCLIENTE]  DEFAULT ('') FOR [UFCLIENTE]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_MUNICIPIOCLI]  DEFAULT ('') FOR [MUNICIPIOCLI]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_BAIRROCLI]  DEFAULT ('') FOR [BAIRROCLI]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_CEPCLI]  DEFAULT ('') FOR [CEPCLI]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_PESSOACLI]  DEFAULT ((0)) FOR [PESSOACLI]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_REGIAOCLI]  DEFAULT ((0)) FOR [REGIAOCLI]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_CFOAUX]  DEFAULT ('') FOR [CFOAUX]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_AnexoAUx]  DEFAULT ('') FOR [AnexoAUx]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_PAGINANF]  DEFAULT ((0)) FOR [PAGINANF]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_ECFNF]  DEFAULT ((2)) FOR [ECFNF]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_Carimbo1]  DEFAULT ('') FOR [Carimbo1]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_Carimbo2]  DEFAULT ('') FOR [Carimbo2]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_Carimbo3]  DEFAULT ('') FOR [Carimbo3]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_Carimbo4]  DEFAULT ('') FOR [Carimbo4]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_Carimbo5]  DEFAULT ('') FOR [Carimbo5]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_CustoMedioLiquido]  DEFAULT ((0)) FOR [CustoMedioLiquido]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_VendaLiquida]  DEFAULT ((0)) FOR [VendaLiquida]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_MargemContribuicao]  DEFAULT ((0)) FOR [MargemContribuicao]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_ValorTotalCodigoZero]  DEFAULT ((0)) FOR [ValorTotalCodigoZero]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_TotalNotaAlternativa]  DEFAULT ((0)) FOR [TotalNotaAlternativa]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_ValorMercadoriaAlternativa]  DEFAULT ((0)) FOR [ValorMercadoriaAlternativa]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_SituacaoEnvio]  DEFAULT ('A') FOR [SituacaoEnvio]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_VendedorLojaVenda]  DEFAULT ((0)) FOR [VendedorLojaVenda]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_LojaVenda]  DEFAULT ('') FOR [LojaVenda]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_NotaCredito]  DEFAULT ((0)) FOR [NotaCredito]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_NfDevolucao]  DEFAULT ((0)) FOR [NfDevolucao]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_SerieDevolucao]  DEFAULT ('') FOR [SerieDevolucao]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_EmiteDataSaida]  DEFAULT ('') FOR [EmiteDataSaida]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_CancelarNota]  DEFAULT ('N') FOR [CancelarNota]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_HoraManutencao]  DEFAULT ('') FOR [HoraManutencao]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_DataProcessamento]  DEFAULT ('') FOR [DataProcessamento]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_SituacaoFec]  DEFAULT ('A') FOR [SituacaoFec]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_ObsSituacaoFec]  DEFAULT ('') FOR [ObsSituacaoFec]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_CodMunicipioCli]  DEFAULT ('') FOR [CodMunicipioCli]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_EnderecoNFeCli]  DEFAULT ('') FOR [EnderecoNFeCli]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_EnderecoNroNFeCli]  DEFAULT ('') FOR [EnderecoNroNFeCli]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_ComplementoNFeCli]  DEFAULT ('') FOR [ComplementoNFeCli]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_InscriSufCli]  DEFAULT ('') FOR [InscriSufCli]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_BaseICMSST]  DEFAULT ((0)) FOR [BaseICMSST]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_ValorICMSST]  DEFAULT ((0)) FOR [ValorICMSST]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_ValorCOFINS]  DEFAULT ((0)) FOR [ValorCOFINS]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_ValorOutros]  DEFAULT ((0)) FOR [ValorOutros]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_SenhaDesconto]  DEFAULT ('') FOR [SenhaDesconto]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_Volume]  DEFAULT ((0)) FOR [Volume]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_TipoFrete]  DEFAULT ((0)) FOR [TipoFrete]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_ParcelasTEF]  DEFAULT ((0)) FOR [ParcelasTEF]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_AutorizacaoTEF]  DEFAULT ('') FOR [AutorizacaoTEF]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_GarantiaEstendida]  DEFAULT ('N') FOR [GarantiaEstendida]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_TotalGarantia]  DEFAULT ((0)) FOR [TotalGarantia]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_NroResidencia]  DEFAULT ('') FOR [NroResidencia]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_CompleResidencia]  DEFAULT ('') FOR [CompleResidencia]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_SeguroPremiado]  DEFAULT ((0)) FOR [SeguroPremiado]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_CertificadoSorte]  DEFAULT ((0)) FOR [CertificadoSorte]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_numeroSorte]  DEFAULT ((0)) FOR [numeroSorte]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_sp_premioLiquido]  DEFAULT ((0)) FOR [sp_premioLiquido]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_sp_IOF]  DEFAULT ((0)) FOR [sp_IOF]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_sp_valorRemuneracao]  DEFAULT ((0)) FOR [sp_valorRemuneracao]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_sp_percentualRemuneracao]  DEFAULT ((0)) FOR [sp_percentualRemuneracao]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_sp_valorRepasse]  DEFAULT ((0)) FOR [sp_valorRepasse]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_codmun]  DEFAULT ('') FOR [codmun]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_ChaveNFe]  DEFAULT ('') FOR [ChaveNFe]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_SituacaoProcesso]  DEFAULT ('A') FOR [SituacaoProcesso]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_DataProcesso]  DEFAULT ('') FOR [DataProcesso]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_Parcelas]  DEFAULT ((0)) FOR [Parcelas]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_NroCaixa]  DEFAULT ((0)) FOR [NroCaixa]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_Protocolo]  DEFAULT ((0)) FOR [Protocolo]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_TipoTransporte]  DEFAULT ('') FOR [TipoTransporte]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_Criticaprocesso]  DEFAULT ('') FOR [Criticaprocesso]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_CPFNFP]  DEFAULT ('') FOR [CPFNFP]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_vendedorGarantia]  DEFAULT ((0)) FOR [vendedorGarantia]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_ModalidadeVenda]  DEFAULT ('') FOR [ModalidadeVenda]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_LiberaBloqueio]  DEFAULT ('N') FOR [LiberaBloqueio]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_valorSub]  DEFAULT ((0)) FOR [valorSub]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_baseSub]  DEFAULT ((0)) FOR [baseSub]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_sub]  DEFAULT ((0)) FOR [sub]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF_NFCapa_ChaveNFeDevolucao]  DEFAULT ('') FOR [ChaveNFeDevolucao]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF__NFCapa__valICMSR__7A076D29]  DEFAULT ((0)) FOR [valICMSRemet]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF__NFCapa__valICMSD__7AFB9162]  DEFAULT ((0)) FOR [valICMSDest]
GO
ALTER TABLE [dbo].[NFCapa] ADD  CONSTRAINT [DF__NFCapa__valorICM__7BEFB59B]  DEFAULT ((0)) FOR [valorICMSFECP]
GO
ALTER TABLE [dbo].[NFe_ide] ADD  CONSTRAINT [DF_NFe_ide_Situacao]  DEFAULT ('W') FOR [Situacao]
GO
ALTER TABLE [dbo].[NFe_ide] ADD  CONSTRAINT [DF_NFe_ide_refNFe]  DEFAULT ('') FOR [refNFe]
GO
ALTER TABLE [dbo].[NFe_prod] ADD  DEFAULT ((0)) FOR [vVICMSDESON]
GO
ALTER TABLE [dbo].[NFe_prod] ADD  DEFAULT ((0)) FOR [vBCUFDEST]
GO
ALTER TABLE [dbo].[NFItens] ADD  CONSTRAINT [DF_NFItens_DATAEMI]  DEFAULT ('') FOR [DATAEMI]
GO
ALTER TABLE [dbo].[NFItens] ADD  CONSTRAINT [DF_NFItens_QTDE]  DEFAULT ((0)) FOR [QTDE]
GO
ALTER TABLE [dbo].[NFItens] ADD  CONSTRAINT [DF_NFItens_VLUNIT]  DEFAULT ((0)) FOR [VLUNIT]
GO
ALTER TABLE [dbo].[NFItens] ADD  CONSTRAINT [DF_NFItens_VLUNIT2]  DEFAULT ((0)) FOR [VLUNIT2]
GO
ALTER TABLE [dbo].[NFItens] ADD  CONSTRAINT [DF_NFItens_VLTOTITEM]  DEFAULT ((0)) FOR [VLTOTITEM]
GO
ALTER TABLE [dbo].[NFItens] ADD  CONSTRAINT [DF_NFItens_DESCRAT]  DEFAULT ((0)) FOR [DESCRAT]
GO
ALTER TABLE [dbo].[NFItens] ADD  CONSTRAINT [DF_NFItens_ICMS]  DEFAULT ((0)) FOR [ICMS]
GO
ALTER TABLE [dbo].[NFItens] ADD  CONSTRAINT [DF_NFItens_ITEM]  DEFAULT ((0)) FOR [ITEM]
GO
ALTER TABLE [dbo].[NFItens] ADD  CONSTRAINT [DF_NFItens_VLIPI]  DEFAULT ((0)) FOR [VLIPI]
GO
ALTER TABLE [dbo].[NFItens] ADD  CONSTRAINT [DF_NFItens_DESCONTO]  DEFAULT ((0)) FOR [DESCONTO]
GO
ALTER TABLE [dbo].[NFItens] ADD  CONSTRAINT [DF_NFItens_PLISTA]  DEFAULT ((0)) FOR [PLISTA]
GO
ALTER TABLE [dbo].[NFItens] ADD  CONSTRAINT [DF_NFItens_COMISSAO]  DEFAULT ((0)) FOR [COMISSAO]
GO
ALTER TABLE [dbo].[NFItens] ADD  CONSTRAINT [DF_NFItens_VALORICMS]  DEFAULT ((0)) FOR [VALORICMS]
GO
ALTER TABLE [dbo].[NFItens] ADD  CONSTRAINT [DF_NFItens_BCOMIS]  DEFAULT ((0)) FOR [BCOMIS]
GO
ALTER TABLE [dbo].[NFItens] ADD  CONSTRAINT [DF_NFItens_CSPROD]  DEFAULT ((0)) FOR [CSPROD]
GO
ALTER TABLE [dbo].[NFItens] ADD  CONSTRAINT [DF_NFItens_LINHA]  DEFAULT ((0)) FOR [LINHA]
GO
ALTER TABLE [dbo].[NFItens] ADD  CONSTRAINT [DF_NFItens_SECAO]  DEFAULT ((0)) FOR [SECAO]
GO
ALTER TABLE [dbo].[NFItens] ADD  CONSTRAINT [DF_NFItens_VBUNIT]  DEFAULT ((0)) FOR [VBUNIT]
GO
ALTER TABLE [dbo].[NFItens] ADD  CONSTRAINT [DF_NFItens_ICMPDV]  DEFAULT ((0)) FOR [ICMPDV]
GO
ALTER TABLE [dbo].[NFItens] ADD  CONSTRAINT [DF_NFItens_CODBARRA]  DEFAULT ('') FOR [CODBARRA]
GO
ALTER TABLE [dbo].[NFItens] ADD  CONSTRAINT [DF_NFItens_NF]  DEFAULT ((0)) FOR [NF]
GO
ALTER TABLE [dbo].[NFItens] ADD  CONSTRAINT [DF_NFItens_SERIE]  DEFAULT ('') FOR [SERIE]
GO
ALTER TABLE [dbo].[NFItens] ADD  CONSTRAINT [DF_NFItens_LOJAORIGEM]  DEFAULT ('') FOR [LOJAORIGEM]
GO
ALTER TABLE [dbo].[NFItens] ADD  CONSTRAINT [DF_NFItens_CLIENTE]  DEFAULT ((0)) FOR [CLIENTE]
GO
ALTER TABLE [dbo].[NFItens] ADD  CONSTRAINT [DF_NFItens_VENDEDOR]  DEFAULT ((0)) FOR [VENDEDOR]
GO
ALTER TABLE [dbo].[NFItens] ADD  CONSTRAINT [DF_NFItens_ALIQIPI]  DEFAULT ((0)) FOR [ALIQIPI]
GO
ALTER TABLE [dbo].[NFItens] ADD  CONSTRAINT [DF_NFItens_TIPONOTA]  DEFAULT ('') FOR [TIPONOTA]
GO
ALTER TABLE [dbo].[NFItens] ADD  CONSTRAINT [DF_NFItens_REDUCAOICMS]  DEFAULT ((0)) FOR [REDUCAOICMS]
GO
ALTER TABLE [dbo].[NFItens] ADD  CONSTRAINT [DF_NFItens_BASEICMS]  DEFAULT ((0)) FOR [BASEICMS]
GO
ALTER TABLE [dbo].[NFItens] ADD  CONSTRAINT [DF_NFItens_TIPOMOVIMENTACAO]  DEFAULT ((0)) FOR [TIPOMOVIMENTACAO]
GO
ALTER TABLE [dbo].[NFItens] ADD  CONSTRAINT [DF_NFItens_DETALHEIMPRESSAO]  DEFAULT ('') FOR [DETALHEIMPRESSAO]
GO
ALTER TABLE [dbo].[NFItens] ADD  CONSTRAINT [DF_NFItens_SerieProd1]  DEFAULT ('') FOR [SerieProd1]
GO
ALTER TABLE [dbo].[NFItens] ADD  CONSTRAINT [DF_NFItens_SerieProd2]  DEFAULT ('') FOR [SerieProd2]
GO
ALTER TABLE [dbo].[NFItens] ADD  CONSTRAINT [DF_NFItens_CustoMedioLiquido]  DEFAULT ((0)) FOR [CustoMedioLiquido]
GO
ALTER TABLE [dbo].[NFItens] ADD  CONSTRAINT [DF_NFItens_VendaLiquida]  DEFAULT ((0)) FOR [VendaLiquida]
GO
ALTER TABLE [dbo].[NFItens] ADD  CONSTRAINT [DF_NFItens_MargemContribuicao]  DEFAULT ((0)) FOR [MargemContribuicao]
GO
ALTER TABLE [dbo].[NFItens] ADD  CONSTRAINT [DF_NFItens_EncargosVendaLiquida]  DEFAULT ((0)) FOR [EncargosVendaLiquida]
GO
ALTER TABLE [dbo].[NFItens] ADD  CONSTRAINT [DF_NFItens_EncargosCustoMedioLiquido]  DEFAULT ((0)) FOR [EncargosCustoMedioLiquido]
GO
ALTER TABLE [dbo].[NFItens] ADD  CONSTRAINT [DF_NFItens_PrecoUnitAlternativa]  DEFAULT ((0)) FOR [PrecoUnitAlternativa]
GO
ALTER TABLE [dbo].[NFItens] ADD  CONSTRAINT [DF_NFItens_ValorMercadoriaAlternativa]  DEFAULT ((0)) FOR [ValorMercadoriaAlternativa]
GO
ALTER TABLE [dbo].[NFItens] ADD  CONSTRAINT [DF_NFItens_ReferenciaAlternativa]  DEFAULT ('') FOR [ReferenciaAlternativa]
GO
ALTER TABLE [dbo].[NFItens] ADD  CONSTRAINT [DF_NFItens_SituacaoEnvio]  DEFAULT ('A') FOR [SituacaoEnvio]
GO
ALTER TABLE [dbo].[NFItens] ADD  CONSTRAINT [DF_NFItens_DescricaoAlternativa]  DEFAULT ('') FOR [DescricaoAlternativa]
GO
ALTER TABLE [dbo].[NFItens] ADD  CONSTRAINT [DF_NFItens_Tributacao]  DEFAULT ('') FOR [Tributacao]
GO
ALTER TABLE [dbo].[NFItens] ADD  CONSTRAINT [DF_NFItens_IcmsMargem]  DEFAULT ((0)) FOR [IcmsMargem]
GO
ALTER TABLE [dbo].[NFItens] ADD  CONSTRAINT [DF_NFItens_PisCofins]  DEFAULT ((0)) FOR [PisCofins]
GO
ALTER TABLE [dbo].[NFItens] ADD  CONSTRAINT [DF_NFItens_DeducoesVendas]  DEFAULT ((0)) FOR [DeducoesVendas]
GO
ALTER TABLE [dbo].[NFItens] ADD  CONSTRAINT [DF_NFItens_EncargosFinanceiros]  DEFAULT ((0)) FOR [EncargosFinanceiros]
GO
ALTER TABLE [dbo].[NFItens] ADD  CONSTRAINT [DF_NFItens_EstoqueAntes]  DEFAULT ((0)) FOR [EstoqueAntes]
GO
ALTER TABLE [dbo].[NFItens] ADD  CONSTRAINT [DF_NFItens_EstoqueDepois]  DEFAULT ((0)) FOR [EstoqueDepois]
GO
ALTER TABLE [dbo].[NFItens] ADD  CONSTRAINT [DF_NFItens_CFOP]  DEFAULT ('') FOR [CFOP]
GO
ALTER TABLE [dbo].[NFItens] ADD  CONSTRAINT [DF_NFItens_CSTICMS]  DEFAULT ('') FOR [CSTICMS]
GO
ALTER TABLE [dbo].[NFItens] ADD  CONSTRAINT [DF_nfitens_GarantiaEstendida]  DEFAULT ('N') FOR [GarantiaEstendida]
GO
ALTER TABLE [dbo].[NFItens] ADD  CONSTRAINT [DF_nfitens_PlanoGarantia]  DEFAULT ((0)) FOR [PlanoGarantia]
GO
ALTER TABLE [dbo].[NFItens] ADD  CONSTRAINT [DF_nfitens_CoeficientePlano]  DEFAULT ((0)) FOR [CoeficientePlano]
GO
ALTER TABLE [dbo].[NFItens] ADD  CONSTRAINT [DF_nfitens_QtdeGarantia]  DEFAULT ((0)) FOR [QtdeGarantia]
GO
ALTER TABLE [dbo].[NFItens] ADD  CONSTRAINT [DF_nfitens_ValorGarantia]  DEFAULT ((0)) FOR [ValorGarantia]
GO
ALTER TABLE [dbo].[NFItens] ADD  CONSTRAINT [DF_nfitens_CertificadoInicio]  DEFAULT ('') FOR [CertificadoInicio]
GO
ALTER TABLE [dbo].[NFItens] ADD  CONSTRAINT [DF_nfitens_CertificadoFim]  DEFAULT ('') FOR [CertificadoFim]
GO
ALTER TABLE [dbo].[NFItens] ADD  CONSTRAINT [DF_nfitens_ge_premioLiquido]  DEFAULT ((0)) FOR [ge_premioLiquido]
GO
ALTER TABLE [dbo].[NFItens] ADD  CONSTRAINT [DF_nfitens_ge_IOF]  DEFAULT ((0)) FOR [ge_IOF]
GO
ALTER TABLE [dbo].[NFItens] ADD  CONSTRAINT [DF_NFItens_ge_dataInicioVigencia]  DEFAULT ('') FOR [ge_dataInicioVigencia]
GO
ALTER TABLE [dbo].[NFItens] ADD  CONSTRAINT [DF_NFItens_ge_dataFinalVigencia]  DEFAULT ('') FOR [ge_dataFinalVigencia]
GO
ALTER TABLE [dbo].[NFItens] ADD  CONSTRAINT [DF_nfitens_ge_valorCustoSeguradora]  DEFAULT ((0)) FOR [ge_valorCustoSeguradora]
GO
ALTER TABLE [dbo].[NFItens] ADD  CONSTRAINT [DF_nfitens_ge_seqCancelamento]  DEFAULT ((0)) FOR [ge_seqCancelamento]
GO
ALTER TABLE [dbo].[NFItens] ADD  CONSTRAINT [DF_NFItens_ge_dataCancelamento]  DEFAULT ('') FOR [ge_dataCancelamento]
GO
ALTER TABLE [dbo].[NFItens] ADD  CONSTRAINT [DF_NFItens_SituacaoProcesso]  DEFAULT ('') FOR [SituacaoProcesso]
GO
ALTER TABLE [dbo].[NFItens] ADD  CONSTRAINT [DF_NFItens_dataprocesso]  DEFAULT ('') FOR [dataprocesso]
GO
ALTER TABLE [dbo].[NFItens] ADD  CONSTRAINT [DF_NFItens_ICMSAplicado]  DEFAULT ((0)) FOR [ICMSAplicado]
GO
ALTER TABLE [dbo].[NFItens] ADD  CONSTRAINT [DF_NFItens_Parcelas]  DEFAULT ((0)) FOR [Parcelas]
GO
ALTER TABLE [dbo].[NFItens] ADD  CONSTRAINT [DF_nfitens_valorSub]  DEFAULT ((0)) FOR [valorSub]
GO
ALTER TABLE [dbo].[NFItens] ADD  CONSTRAINT [DF_nfitens_baseSub]  DEFAULT ((0)) FOR [baseSub]
GO
ALTER TABLE [dbo].[NFItens] ADD  CONSTRAINT [DF_nfitens_sub]  DEFAULT ((0)) FOR [sub]
GO
ALTER TABLE [dbo].[NFItens] ADD  CONSTRAINT [DF_nfitens_baseIPI]  DEFAULT ((0)) FOR [baseIPI]
GO
ALTER TABLE [dbo].[NFItens] ADD  DEFAULT ((0)) FOR [aliqICMSDest]
GO
ALTER TABLE [dbo].[NFItens] ADD  DEFAULT ((0)) FOR [aliqICMSInter]
GO
ALTER TABLE [dbo].[NFItens] ADD  DEFAULT ((0)) FOR [ICMSInterpart]
GO
ALTER TABLE [dbo].[NFItens] ADD  DEFAULT ((0)) FOR [valICMSRemet]
GO
ALTER TABLE [dbo].[NFItens] ADD  DEFAULT ((0)) FOR [valICMSDest]
GO
ALTER TABLE [dbo].[NFItens] ADD  DEFAULT ((0)) FOR [valorICMSFECP]
GO
ALTER TABLE [dbo].[NFItens] ADD  DEFAULT ((0)) FOR [aliqICMSFECP]
GO
/****** Object:  StoredProcedure [dbo].[SP_AcertaCondicaoPagamento]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
/*

select * from CondicaoPagamento where cp_tipo = 'FA' and CP_Codigo <> 85
exec SP_AcertaCondicaoPagamento


*/
create            Procedure [dbo].[SP_AcertaCondicaoPagamento]

As
Declare	@Int	    int,
		@ContLen	int,
		@Condicao    char(60),
		@NovaCondicao char(60),
		@Aux          char(60)
       


 Begin Transaction
       select @NovaCondicao = ''
       select  @int = 4
       while @int < 99
         begin 
           Select @Condicao = (select rtrim(ltrim(CP_Condicao)) from CondicaoPagamento 
                                        where cp_tipo = 'FA' and CP_Codigo <> 85 and CP_Codigo = @Int)

           select @Condicao = REPLACE(@Condicao,'/',' ')
           select @Condicao = REPLACE(@Condicao,'D','')
           select @Condicao = REPLACE(@Condicao,'L','')
           select @ContLen = 1
           select @Aux = ''
		   while @ContLen <  50
		      Begin
		         if substring(Ltrim(Rtrim(@Condicao)),@ContLen,1) <> ''
		           Begin
		              select @Aux = rtrim(Ltrim(@aux)) + '' +  substring(Ltrim(Rtrim(@Condicao)),@ContLen,1)
		           end
		         if substring(Ltrim(Rtrim(@Condicao)),@ContLen,1) = ''
		           begin 
		              if LEN(Ltrim(Rtrim(@aux))) = 2 
		                    Select @Aux = '0' +  (Ltrim(Rtrim(@aux)))
		                
		              Select @NovaCondicao = rtrim(ltrim(convert(char(50),@NovaCondicao))) 
		                                                + convert(char(3),(Ltrim(Rtrim(@aux))))
		                                                
		              select @Aux = ''
		              
				   end
		           select @ContLen = @ContLen + 1
		        end


		    select @Aux = ''
		    		print @NovaCondicao   
             update CondicaoPagamento set cp_intervaloparcelas = RTRIM(ltrim(@NovaCondicao))
                      where cp_tipo = 'FA' and CP_Codigo <> 85 and CP_Codigo = @Int
                
             select @NovaCondicao = ''
            select @Int = @int + 1
		    end 
  

      

   If @@Error <> 0 
      Begin
         Rollback Transaction
      End 
   Else 
      Begin  
         Commit Transaction
      End

GO
/****** Object:  StoredProcedure [dbo].[SP_ALERTA_FATURADA]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

/*

SELECT nf,serie,* FROM NFCAPA WHERE liberaBloqueio = 'S' and dataemi > '2014/11/26'

EXEC SP_ALERTA_DESBLOQUEIO '28','3269','ne'

EXEC SP_ALERTA_DESBLOQUEIO '28','3269','ne'

sele
select  from nfcapa where dataemi >'2016/05/01' and serie = 'NE'  and cliente < 90000
select top 1 * from nfitens where dataemi >'2016/05/01' and serie = 'NE' 
select top 1 * from historicoClientes 
select * into vende from dmac271.dmac_loja.dbo.vende
[SP_ALERTA_FATURADA] '181',3959, 'NE'
*/

CREATE Procedure [dbo].[SP_ALERTA_FATURADA]
		@loja					char(5),
		@nf						varchar(20),
		@serie					char(3)

As

	Declare	@mensagem         as varchar(max),
 			@assunto          as varchar(1000),
			@cliente          as varchar(200),
			@texto1           as varchar(max),
			@valor            as varchar(20),
			@data             as char(12),
			@codvendedor      as varchar(10),
			@vendedor         as char(200),
			@referencia       as varchar(7),
			@descricaoRef     as char(50),
			@qtde             as varchar(20),
			@ValorItens       as varchar(20),
			@codCli           as varchar(20),
			@cgcCli           as varchar(15),
			@Limite           as varchar(20),
			@maiorCompra      as varchar(20),
			@ultimaCompra     as varchar(20),
			@ultimoPagamento  as varchar(20),
			@maiorAtraso      as varchar(20),
			@qtdeCompras      as varchar(20),
			@totalCompras     as varchar(20),
			@duplAberta       as varchar(20),
			@qtdeDuplAberta   as varchar(20),
			@saldoCompras     as varchar(20),
			@duplAtrasado     as varchar(20),
			@duplPagas        as varchar(20),
			@dataLimite       as char(12),
			@dataMaiorCompra  as char(12),
			@dataUltimaCompra as char(12),
			@dataUltimoPagto  as char(12)

			
Begin
	
	select @cliente = NOMCLI,
	       @valor = TOTALNOTA,
		   @data = DATAEMI ,
		   @codCli = Cliente, 
		   @cgcCli = CGCCLI, 
		   @codvendedor = VENDEDOR 
	  from nfcapa 
	 where nf = @nf and SERIE = @serie and LOJAORIGEM = @loja


	select @vendedor = ve_nome 
	  from vende
	 where ve_codigo = @codvendedor 

	select @dataLimite = HIC_DataLimiteCredito,
	       @Limite = HIC_LimiteCredito,
		   @maiorCompra = HIC_MaiorCompra,
		   @dataMaiorCompra = HIC_DataMaiorCompra,
		   @qtdeCompras = HIC_QuantidadeCompras,
		   @ultimoPagamento =  HIC_UltimoPagamento,
		   @maiorAtraso = HIC_MaiorAtraso,
		   @totalCompras = HIC_TotalCompras,
		   @ultimaCompra = HIC_UltimaCompra ,
		   @dataUltimaCompra = HIC_DataUltimaCompra ,
		   @dataUltimoPagto = HIC_DataUltimoPagamento,
		   @saldoCompras = HIC_SaldoCompras,
		   @duplAberta = HIC_DuplAbertas,
		   @duplAtrasado = HIC_DuplAtrasado 
	  from svdmac.dmac.dbo.HistoricoClientes 
	 where HIC_CodigoCliente = @codCli  
	

	set @assunto = 'Venda Faturada - Nota: ' + @nf + ', Loja: ' + @loja + ' Cliente: ' + @cliente
	
	select @texto1 = '**** Informações da Venda ****' + char(10) 
	select @texto1 = @texto1 + 'Nota fiscal: ' + @nf + '/' + @serie + ' - Loja ' + @loja  + char(10)
	select @texto1 = @texto1 + 'Valor Total da Nota: ' + @valor + char(10)
	select @texto1 = @texto1 + 'Data de Emissão: ' + @data  + char(10) 
	select @texto1 = @texto1 + 'Vendedor: ' + @codvendedor + ' - ' + @vendedor + char(10)+ char(10)
	 
	select @texto1 = @texto1 + '**** Informações do Cliente **** ' + char(10) 
	select @texto1 = @texto1 + 'Código: ' + @codCli + char(10)  
	select @texto1 = @texto1 + 'Razão:' + @cliente + char(10) 
	select @texto1 = @texto1 + 'CNPJ/CPF:' + @cgcCli + char(10)+ char(10)

	select @texto1 = @texto1 + '**** Ficha Financeira ****' + char(10)
	select @texto1 = @texto1 + 'Limite de Crédito : R$' + @Limite + char(10)
	select @texto1 = @texto1 + 'Data do Limite de Crédito :' + @dataLimite + char(10)  
	select @texto1 = @texto1 + 'Duplicatas em Aberto: R$' + @duplAberta + char(10)
	select @texto1 = @texto1 + 'Duplicatas em Atraso: ' + @duplAtrasado + char(10)
	select @texto1 = @texto1 + 'Saldo para Compras: R$' + @saldoCompras + char(10)
	select @texto1 = @texto1 + 'Última Compra: R$' + @ultimaCompra + char(10)
	select @texto1 = @texto1 + 'Data da Última Compra:' + @dataUltimaCompra + char(10)
	select @texto1 = @texto1 + 'Maior Compra: R$' + @maiorCompra + char(10)
	select @texto1 = @texto1 + 'Data Maior Compra: ' + @dataMaiorCompra + char(10)
	select @texto1 = @texto1 + 'Quantidade de Compras: ' + @qtdeCompras + char(10) 
	select @texto1 = @texto1 + 'Último Pagamento: R$' + @ultimoPagamento + char(10)
	select @texto1 = @texto1 + 'Data do Último Pagamento: ' +  @dataUltimoPagto + char(10)
	select @texto1 = @texto1 + 'Maior Atraso: ' + @maiorAtraso + char(10)
	select @texto1 = @texto1 + 'Total de Compras: R$' + @totalCompras + char(10)


	print (@assunto)
	print (@texto1) 
	
	insert into mensagens (ME_Loja, ME_Nota, ME_Serie, ME_Assunto, ME_Mensagem)
	values (@loja, @nf, @serie, @assunto, @texto1)

END 
--select lojaOrigem, nf, serie, condpag from nfcapa where numeroped = 4809
-- and CONDPAG >= 3



/*

--USE [DMAC_LOJA]
--GO

--/****** Object:  Table [dbo].[mensagens]    Script Date: 17/05/2016 11:46:18 ******/
--SET ANSI_NULLS ON
--GO

--SET QUOTED_IDENTIFIER ON
--GO

--SET ANSI_PADDING ON
--GO

--CREATE TABLE [dbo].[mensagens](
--	[ME_loja] [char](5) NULL,
--	[ME_Nota] [int] NULL,
--	[ME_Serie] [char](3) NULL,
--	[ME_Assunto] [varchar](max) NULL,
--	[ME_Mensagem] [varchar](max) NULL
--) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]

--GO

--SET ANSI_PADDING OFF
--GO





*/
GO
/****** Object:  StoredProcedure [dbo].[SP_alterar_dividir_modalidade]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
 

CREATE Procedure [dbo].[SP_alterar_dividir_modalidade]
		 @sequencia		int,
		 @grupoNovo		int,
		 @valor			float

As

Begin
	
	IF @valor < (select MC_Valor from MovimentoCaixa where MC_Sequencia = @sequencia ) 
	BEGIN

		INSERT INTO MovimentoCaixa (MC_NumeroECF, MC_CodigoOperador, MC_Loja, MC_Data, MC_Grupo, 
		MC_SubGrupo, MC_Documento, MC_Serie, MC_Valor, MC_Banco, MC_Agencia, MC_ContaCorrente, 
		MC_NumeroCheque, MC_BomPara, MC_Parcelas, MC_Remessa, MC_SituacaoEnvio, MC_ControleAVR, 
		MC_DataBaixaAVR, MC_Protocolo, MC_NroCaixa, MC_GrupoAuxiliar, MC_Situacao, MC_Pedido, 
		MC_DataProcesso, MC_TipoNota, MC_SequenciaTEF)
		SELECT MC_NumeroECF, MC_CodigoOperador, MC_Loja, MC_Data, 
		@grupoNovo, MC_SubGrupo, MC_Documento, MC_Serie, 
		@valor, MC_Banco, MC_Agencia, MC_ContaCorrente, MC_NumeroCheque, 
		MC_BomPara, MC_Parcelas, MC_Remessa, MC_SituacaoEnvio, MC_ControleAVR, 
		MC_DataBaixaAVR, MC_Protocolo, MC_NroCaixa, MC_GrupoAuxiliar, 
		MC_Situacao, MC_Pedido, MC_DataProcesso, MC_TipoNota, 
		MC_SequenciaTEF 
		from MovimentoCaixa 
		where MC_Sequencia = @sequencia 


		update MovimentoCaixa 
		set MC_Valor = mc_valor - @valor 
		where MC_Sequencia = @sequencia 

	END

--exec SP_Delete_NFe '28',612,'NE'
--exec SP_Cria_NFe '181',1915,'NE'

End




GO
/****** Object:  StoredProcedure [dbo].[SP_Atualiza_Cliente_NFCAPA_Local]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
create    Procedure [dbo].[SP_Atualiza_Cliente_NFCAPA_Local]
		@NF					char(6),
		@Serie				Char(3)                    
As

Begin
	Declare @SQL Char(4000)
	Declare @tiponota char(1)

--set @tiponota = '(''V'',''C'')'
--set @SERIE = 'CF'
--set @Loja = '271'
--set @nf = '74'

	--select top 10 * from [DEMEOSERVER].[Desenv_Demeo].[dbo].nfcapa where numeroped = 121
	--select top 10 * from [DEMEOSERVER].[Desenv_Demeo].[dbo].nfcapa where numeroped = 121


	select @tiponota = (select top 1 tiponota from nfcapa where nf = @NF and serie = @Serie)

	-- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- 

	if @tiponota = 'V' OR @tiponota = 'E' or @tiponota = 'S'
	BEGIN

		Select @sql = 'update nfcapa 
		set NOMCLI = CE_Razao,
		FONECLI = CE_Telefone,
		CGCCLI = CE_CGC,
		INSCRICLI = CE_InscricaoEstadual,
		ENDCLI = CE_Endereco,
		UFCLIENTE = CE_Estado,
		MUNICIPIOCLI = CE_Municipio,
		BAIRROCLI = CE_Bairro,
		CEPCLI = CE_CEP, 
		codmun = '''',
		CompleResidencia = '''',
		NroResidencia = CE_Numero
		from nfcapa, fin_cliente where 
		tiponota = ''' + @tiponota + '''
		and SERIE = ''' + @Serie  + '''
		and nf = ''' + @nf + '''
		and cliente = CE_CodigoCliente'

		--print @sql
		Execute (@SQL)

	end 

	-- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- 

	if @tiponota = 'T'
	BEGIN
	
		Select @sql = 'update nfcapa 
		set NOMCLI = LO_Razao,
		FONECLI = lo_telefone,
		CGCCLI = lo_cgc,
		INSCRICLI = lo_inscricaoEstadual,
		ENDCLI = lo_endereco,
		UFCLIENTE = lo_uf,
		MUNICIPIOCLI = lo_municipio,
		BAIRROCLI = lo_bairro,
		CEPCLI = lo_cep, 
		codmun = lo_codigoMunicipio,
		CompleResidencia = '''',
		NroResidencia = lo_endereconronfe
		from nfcapa,loja where 
		tiponota = ''' + @tiponota + '''
		and SERIE = ''' + @Serie  + '''
		and nf = ''' + @nf + '''
		and lojat = lo_loja'

		--print @sql
		Execute (@sql)

	end 

	-- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- 


	If @@ERROR <> 0
	   Begin	
	   	Rollback Transaction		
	   	Return
	   End
	

End
 
GO
/****** Object:  StoredProcedure [dbo].[SP_Atualiza_Modalidade_Venda_Pedido]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO


CREATE PROCEDURE [dbo].[SP_Atualiza_Modalidade_Venda_Pedido]
                 @Pedido numeric,
                 @Tipo   char(02),
                 @Codigo char(03),
                 @TipoCondicao char(02) 
As
	Declare  @ToalItens                float
	Declare  @Parcelas                 float
	Declare  @libera				   char(1)
begin

	set @libera = (select top 1 LiberaBloqueio from NFCapa where NUMEROPED = @Pedido)

	if @libera = 'T'
	begin
		update nfitens set vlunit=ROUND((PrecoUnitAlternativa * cp_coeficiente),2),vltotitem=ROUND(((PrecoUnitAlternativa * QTDE) * cp_coeficiente),2)
		From produtoloja, CondicaoPagamento, nfitens
		where PR_IndicePreco=CP_ID and CP_Tipo=@Tipo and CP_Codigo=@Codigo
		and PR_Referencia = REFERENCIA and NUMEROPED = @Pedido
	end 
	else
	begin
		update nfitens set vlunit=ROUND((PR_PrecoVenda1 * cp_coeficiente),2),vltotitem=ROUND(((PR_PrecoVenda1 * QTDE) * cp_coeficiente),2)
		From produtoloja, CondicaoPagamento, nfitens
		where PR_IndicePreco=CP_ID and CP_Tipo=@Tipo and CP_Codigo=@Codigo
		and PR_Referencia = REFERENCIA and NUMEROPED = @Pedido
	end 
	 

	update nfitens set DESCRAT = CP_DESCONTO
	From produtoloja, CondicaoPagamento, nfitens
	where PR_IndicePreco=CP_ID and CP_Tipo=@Tipo and CP_Codigo=@Codigo
	and PR_Referencia = REFERENCIA and NUMEROPED =@Pedido
 
	Select @ToalItens =(select SUM(vltotitem)from nfitens where NUMEROPED = @Pedido) 

	update nfcapa set vlrmercadoria=@ToalItens,totalnota=@ToalItens, CONDPAG = @TipoCondicao,
	ModalidadeVenda = (Case When @tipo = 'AV' Then 'A Vista' 
	when @tipo = 'CC' then 'Cartão' 
	when @tipo = 'FA' then 'Faturado' 
	when @tipo = 'FI' then 'Financiado' 
	end), Parcelas = CP_Parcelas
	From  nfcapa, CondicaoPagamento where NUMEROPED =@Pedido 
	and CP_Tipo=@Tipo and CP_Codigo=@Codigo and CP_ID = 1

		  

end 



GO
/****** Object:  StoredProcedure [dbo].[SP_Atualiza_Processos_Venda]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO



CREATE           Procedure [dbo].[SP_Atualiza_Processos_Venda]
                          @NumeroPedido integer,
                          @NumeroNF     integer,
                          @Protocolo    integer,
                          @NumeroCaixa integer
As
Declare @data                   char(10),
	    @condPag                nvarchar(8),
        @serie					char(3),
		@loja					varchar(5)
		
SELECT @data = convert(char(10),getdate(),111)
select top 1 @serie = serie, @condPag = CONDPAG from NFCapa where NUMEROPED = @NumeroPedido

 IF EXISTS (select * from sysobjects where name = '#TempAtualiza_Saida_Estoque ' and upper(xtype) = 'U')
        Drop Table #TempAtualiza_Saida_Estoque  

Begin Transaction
    Create Table  #TempAtualiza_Saida_Estoque 
          (TPS_CodigoProduto   Char(16),
           TPS_Quantidade      Numeric)
          
    Insert Into #TempAtualiza_Saida_Estoque (TPS_CodigoProduto,TPS_Quantidade)
           select Referencia ,sum(Qtde)
           from NFItens  Where NumeroPed=@NumeroPedido and TipoNota in ('PA','TA') 
           group by Referencia

      if substring((select top 1 CAST(CODOPER as varchar(10)) from NFCapa where NumeroPed = @NumeroPedido),1,1) = '1'
            Update EstoqueLoja set EL_Estoque =(EL_Estoque + TPS_Quantidade)
                     From EstoqueLoja,#TempAtualiza_Saida_Estoque 
                     Where EL_Referencia = TPS_CodigoProduto collate SQL_latin1_general_cp1_ci_as

      if substring((select top 1 CAST(CODOPER as varchar(10)) from NFCapa where NumeroPed = @NumeroPedido),1,1) <> '1'
            Update EstoqueLoja set EL_Estoque =(EL_Estoque - TPS_Quantidade)
                     From EstoqueLoja,#TempAtualiza_Saida_Estoque 
                     Where EL_Referencia = TPS_CodigoProduto collate SQL_latin1_general_cp1_ci_as           

    Update NfItens set 
    NF = @NumeroNF,
    TipoNota=(select CASE TipoNota 
     WHEN 'PA' THEN 'V' --
	 WHEN 'V' THEN 'V' --
     WHEN 'SA' THEN 'S' --
	 WHEN 'S' THEN 'S' --
     WHEN 'EA' THEN 'E' --
	 WHEN 'E' THEN 'E' --
      WHEN 'TA' THEN 'T' --
	  WHEN 'T' THEN 'T' --
      end), 
    DATAEMI = @data, 
    dataprocesso = @data  
    Where Numeroped = @NumeroPedido
       
    Update NfCapa set 
    NF = @NumeroNF,
    TipoNota=(select CASE TipoNota 
     WHEN 'PA' THEN 'V' --
	 WHEN 'V' THEN 'V' --
     WHEN 'SA' THEN 'S' --
	 WHEN 'S' THEN 'S' --
     WHEN 'EA' THEN 'E' --
	 WHEN 'E' THEN 'E' --
      WHEN 'TA' THEN 'T' --
	  WHEN 'T' THEN 'T' --
      end),
    hora=  convert(varchar(10),getdate(),108),
    Protocolo = @Protocolo,
    NroCaixa =  @NumeroCaixa,
    DATAEMI = @data, 
    dataprocesso = @data
    Where numeroped = @NumeroPedido
      
    Update CarimboNotaFiscal set 
    CNF_NF = @NumeroNF,
    CNF_DataProcesso = @data
    Where CNF_NumeroPed = @NumeroPedido      
    
    Update movimentocaixa set
    MC_Documento = @NumeroNF,
    MC_DataProcesso = @data,
    MC_TipoNota=(select CASE MC_TipoNota 
     WHEN 'PA' THEN 'V' --
	 WHEN 'V' THEN 'V' --
     WHEN 'SA' THEN 'S' --
	 WHEN 'S' THEN 'S' --
     WHEN 'EA' THEN 'E' --
	 WHEN 'E' THEN 'E' --
      WHEN 'TA' THEN 'T' --
	  WHEN 'T' THEN 'T' --
      end)
    where MC_Pedido = @NumeroPedido
    
    EXEC SP_Atualiza_Cliente_NFCAPA_Local @NumeroNF,@serie
	
	select @loja = (select top 1 CTS_Loja from ControleSistema)
	if @condPag > 3
		EXEC Sp_Cria_Duplicatas @loja, @data,@data

   If @@Error <> 0 
      Begin
         Rollback Transaction
      End 
   Else 
      Begin  
         Commit Transaction
      End
      
    
    

 
GO
/****** Object:  StoredProcedure [dbo].[SP_Atualiza_Processos_Venda_Central]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
/*

Atualiza Estoque atraves da Confirmação do ItensVenda

exec SP_Atualiza_Processos_Venda_Central
*/

CREATE            Procedure [dbo].[SP_Atualiza_Processos_Venda_Central]
                      
As
 Declare	@Loja  	Char(5)

 Begin

   Select @Loja=(Select CTS_Loja From ControleSistema)
		--Update dmac..Lojas set LO_Conexao='S' where LO_Loja='271'
		--select LO_Loja,LO_Conexao, * from lojas

   Update [SVDMAC].[DMAC].[dbo].Loja set LO_Conexao='S' where LO_Loja=@loja

 End

GO
/****** Object:  StoredProcedure [dbo].[SP_Busca_Codigo_Cliente_Faturado]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
/*

Atualiza Estoque atraves da Confirmação do ItensVenda

*/

CREATE       Procedure [dbo].[SP_Busca_Codigo_Cliente_Faturado]
                      
As
 Declare	@Loja  	Char(5)

 Begin

   
  -- Update [SVCentralfer].[Dmac].[dbo].fin_controlefinanceiro set CF_CodigoCliente =( CF_CodigoCliente + 1)
  -- Select  CF_CodigoCliente from   [SVCentralfer].[Dmac].[dbo].fin_controlefinanceiro

    Update dmac..fin_controlefinanceiro set CF_CodigoCliente =( CF_CodigoCliente + 1)
    update controlesistema set cts_codigoclientefaturado = r.CF_CodigoCliente from 
           controlesistema, dmac..fin_controlefinanceiro as r
    End



/* 

Exec SP_Busca_Codigo_Cliente_Faturado

*/

GO
/****** Object:  StoredProcedure [dbo].[SP_Cancela_NotaFiscal]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE                  Procedure [dbo].[SP_Cancela_NotaFiscal]
                            @NotaFiscal integer,
                            @Serie      char(03)
                         
As
Declare	@TipoNota	Char(2),
        @Retorno int,
        @Codigo		        Integer,
        @Sequencia              Integer,
        @ValorCampo             char(50)

 IF EXISTS (select * from sysobjects where name = '#TempCancela_NotaFiscal' and upper(xtype) = 'U')
        Drop Table #TempCancela_NotaFiscal  


 Begin Transaction
    Create Table  #TempCancela_NotaFiscal
          (TPS_CodigoProduto   Char(16),
           TPS_Quantidade      Numeric,
           TPS_TipoNota        Char(2))

    Select @TipoNota =(select TipoNota from  NfCapa  Where NF = @NotaFiscal and Serie = @Serie) 
          
    Insert Into #TempCancela_NotaFiscal  (TPS_CodigoProduto,TPS_Quantidade)
           select Referencia ,sum(Qtde)
           from NFItens  Where NF = @NotaFiscal and Serie = @Serie  
           group by Referencia
  
    If @TipoNota = 'V' or @TipoNota = 'T' or @TipoNota = 'S' or @TipoNota = 'TB'
       Update EstoqueLoja set EL_Estoque =(EL_Estoque + TPS_Quantidade)
              From EstoqueLoja,#TempCancela_NotaFiscal
              Where EL_Referencia = TPS_CodigoProduto collate sql_latin1_general_cp1_ci_as

   If @TipoNota = 'E' or @TipoNota = 'R'
       Update EstoqueLoja set EL_Estoque =(EL_Estoque - TPS_Quantidade)
              From EstoqueLoja,#TempCancela_NotaFiscal
              Where EL_Referencia = TPS_CodigoProduto collate sql_latin1_general_cp1_ci_as

   Update NfItens set TipoNota ='C' Where  NF = @NotaFiscal and Serie = @Serie
       
   Update NfCapa set TipoNota ='C',hora=  getdate()
           Where  NF = @NotaFiscal and Serie = @Serie

   Update Movimentocaixa set MC_TipoNota = 'C' where MC_Documento = @NotaFiscal and MC_Serie = ltrim(rtrim(@Serie)) 


   if @Serie = 'NE'
		update nfe_ide set Situacao = 'C' 
		where enf = @NotaFiscal


   If @@Error <> 0 
      Begin
         Rollback Transaction
      End 
   Else 
      Begin  
         Commit Transaction
      End


	  
GO
/****** Object:  StoredProcedure [dbo].[SP_CORRECAO_DIARIA]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO


CREATE    Procedure [dbo].[SP_CORRECAO_DIARIA]

As

Begin
	Declare  @SQL 		       char(5000)                

	declare @data as char(10)
	select @data = convert(char(10),getdate(),111)
	
	--select @data = '2015/12/04'
        
          Select @SQL = 'UPDATE MOVIMENTOCAIXA 
					     SET MC_VALOR = TOTALNOTA 
					     from MovimentoCaixa, nfcapa 
					     where MC_Data = ''' + @data + ''' 
					     and MC_TipoNota in (''TA'',''T'')
					     and MC_Documento = NF 
						 and TIPONOTA = ''T'' 
						 and DATAEMI = ''' + @data + ''' 
						 and MC_Valor <> TOTALNOTA 
'

		  --print (@SQL)
          execute (@SQL)
end

--exec SP_CORRECAO_DIARIA

 



GO
/****** Object:  StoredProcedure [dbo].[Sp_Cria_Duplicatas]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE   Procedure [dbo].[Sp_Cria_Duplicatas]
	@Loja		Char(5),
	@DataInicial	Char(20),
	@DataFinal	Char(20)

As


Declare		@LojaVenda		VarChar(5),
		@Serie			VarChar(2),
		@SerieAlm		VarChar(2),
		@NotaFiscal		Int,
		@SerieLida		Char(2),
		@LojaOrigem		Char(5),
		@Cliente		Int,
		@DataEmissao		DateTime,
		@CodigoVendedor		SmallInt,
		@CondicaoPagamento	SmallInt,
		@DataPagamento		DateTime,
		@PagamentoEntrada	Float,
		@TotalNota		Float,
		@Descricao		VarChar(60),
		@QuantidadeParcelas	SmallInt,
		@TipoCondicao		Char(2),
		@TipoPessoa		Char(1),
		@PagamentoCarteira	Char(1),
		@MaiorCompra		Float,
		@DataUltimaCompra	DateTime,
		@Inicio			SmallInt,
		@DataBase		DateTime,
		@ValorParcela		Float,
		@Sequencia		SmallInt,
		@ParcelaInt		Int,
		@Dias			SmallInt,
		@Situacao	Char(1)


Begin


	Begin Transaction

	If @Loja = 'ALM01'
	  Begin
		Select	@LojaVenda = '353',
			--@Serie = 'S2',
			@SerieAlm = ''
	  End
	Else
	  Begin
		If @Loja = '353'
		  Begin
			Select	@LojaVenda = @Loja,
				--@Serie = '%',
				@SerieAlm = 'S2'
		  End

		Else
		  Begin
			
			Select	@LojaVenda = @Loja,
				--@Serie = '%',
				@SerieAlm = ''
		  End
	  End

	--Select @serieNota	

	--print ('DUPLICATA ETAPA 1')

		Declare curVenda Insensitive Cursor For
		
				Select  NF,
			Serie,
			LojaOrigem,
			Cliente,
			DataEmi,
			Vendedor,
			CondPag,
			DataPag,
			PgEntra,
			TotalNota,
			CP_Parcelas,
			CP_QuantidadeParcelas,
			CP_TipoCondicao,
			CE_TipoPessoa,
			CE_PagamentoCarteira,
			CE_MaiorCompra,
			CE_DataUltimaCompra
		From	NFCapa,
			CondicaoPagto,
			FIN_Cliente,
			CodigoOperacao
		Where	CondPag = CP_CodigoCondicao and
			Cliente = CE_CodigoCliente and
			CondPag > 3 and
			DataEmi between @DataInicial and @DataFinal and
			CF_TipoCodigo = 'V' and 
			CP_VendaCompra = 'V' and
			SituacaoEnvio ='A' and
			TipoNota <> 'C' and 
			LojaOrigem = @Loja and
			Serie = 'NE'



	Open curVenda
	Fetch Next From curVenda Into
		@NotaFiscal,
		@SerieLida,
		@LojaOrigem,
		@Cliente,
		@DataEmissao,
		@CodigoVendedor,
		@CondicaoPagamento,
		@DataPagamento,
		@PagamentoEntrada,
		@TotalNota,
		@Descricao,
		@QuantidadeParcelas,
		@TipoCondicao,
		@TipoPessoa,
		@PagamentoCarteira,
		@MaiorCompra,
		@DataUltimaCompra


	--print ('DUPLICATA')
	While @@FETCH_STATUS = 0
	  Begin

		Delete 	Duplicata 
		Where	DP_Loja = @LojaOrigem and
			DP_NotaFiscal = @NotaFiscal and
			DP_Serie = @SerieLida

		If @@Error <> 0 Goto Desfaz


		Select  @Inicio = 1,
			@Sequencia = 0,
			@DataBase = (Case @TipoCondicao
					When 'DD' Then @DataEmissao
					When 'DL' Then DateAdd(day, 1, @DataEmissao)
					When 'DI' Then @DataPagamento
				     End
				    )

		If @PagamentoEntrada > 0
		  Begin

			Select 	@TotalNota = @TotalNota - @PagamentoEntrada,
				@Sequencia = @Sequencia + 1
			

			Insert Into Duplicata (
				DP_Loja,
				DP_NotaFiscal,
				DP_Serie,
				DP_Sequencia,
				DP_CodigoCliente,
				DP_DataEmissao,
				DP_Vendedor,
				DP_Banco,
				DP_ValorDuplicata,
				DP_DataVencimento,
				DP_Situacao
			)
			Select	@LojaOrigem,
				@NotaFiscal,
				@SerieLida,
				@Sequencia,
				(Case
					When @CondicaoPagamento = 2 Then 999998
					When @CondicaoPagamento = 3 Then 999997
					Else @Cliente
				 End
				),
				@DataEmissao,
				@CodigoVendedor,
				(Case
					When @LojaOrigem = '800' Then 315
					Else 800
				 End
				),
				@PagamentoEntrada,
				@DataEmissao,
				'W'

			If @@Error <> 0 Goto Desfaz

		  End


		Select 	@ParcelaInt = Abs(Ceiling(-(@TotalNota / @QuantidadeParcelas) * 100))

		Select 	@ValorParcela = @ParcelaInt / 100.0

		--print ('WHILE')
		--PRINT @Descricao
		
		While @QuantidadeParcelas > 0
		  Begin
			Select 	@Sequencia = @Sequencia + 1,
				@Dias = Convert(SmallInt, SubString(@Descricao, @Inicio, 3))

			If @QuantidadeParcelas = 1
			  Begin
				Select @ValorParcela = @TotalNota
			  End

			--print ('INSERIR DUPLICATA')
			Insert Into Duplicata (
				DP_Loja,
				DP_NotaFiscal,
				DP_Serie,
				DP_Sequencia,
				DP_CodigoCliente,
				DP_DataEmissao,
				DP_Vendedor,
				DP_Banco,
				DP_ValorDuplicata,
				DP_DataVencimento,
				DP_Situacao
			)
			Select	@LojaOrigem,
				@NotaFiscal,
				@SerieLida,
				@Sequencia,
				(Case
					When @CondicaoPagamento = 2 Then 999998
					When @CondicaoPagamento = 3 Then 999997
					Else @Cliente
				 End
				),
				@DataEmissao,
				@CodigoVendedor,
				(Case
					When @CondicaoPagamento = 2 Then 997
					When @CondicaoPagamento = 3 Then 998
					When @CondicaoPagamento > 3 and @PagamentoCarteira = 'S' and @LojaOrigem <> '800' Then 800
					When @CondicaoPagamento > 3 and @TipoPessoa = 'U' and @LojaOrigem <> '800' Then 802
					When @CondicaoPagamento > 3 and @TipoPessoa = 'A' and @LojaOrigem <> '800' Then 804
					When @CondicaoPagamento > 3 and @LojaOrigem = '800' Then 314
					When @LojaOrigem = '85' then 422
                                        Else 422
				 End
				),
				@ValorParcela,
				DateAdd(day, @Dias, @DataBase),
				'W'

				If @@Error <> 0 Goto Desfaz


			Select 	@QuantidadeParcelas = @QuantidadeParcelas - 1,
				@Inicio = @Inicio + 3,
				@TotalNota = @TotalNota - @ValorParcela

		  End



		Fetch Next From curVenda Into
			@NotaFiscal,
			@SerieLida,
			@LojaOrigem,
			@Cliente,
			@DataEmissao,
			@CodigoVendedor,
			@CondicaoPagamento,
			@DataPagamento,
			@PagamentoEntrada,
			@TotalNota,
			@Descricao,
			@QuantidadeParcelas,
			@TipoCondicao,
			@TipoPessoa,
			@PagamentoCarteira,
			@MaiorCompra,
			@DataUltimaCompra

	  End


	Close curVenda
	Deallocate curVenda


Final:



	Commit Transaction

	Return(0)


Desfaz:

	Close curVenda
	Deallocate curVenda


	Rollback Transaction

	Return(1)


End


GO
/****** Object:  StoredProcedure [dbo].[SP_Delete_NFe]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO



CREATE Procedure [dbo].[SP_Delete_NFe]
		 @loja			char(5),	
		 @notafiscal		int,
		 @serie			char(2)
As

Begin

delete NFe_ide WHERE eloja = @loja and eNF = @notaFiscal and eSerie = @serie
delete NFe_emit WHERE eloja = @loja and eNF = @notaFiscal and eSerie = @serie
delete NFe_dest WHERE eloja = @loja and eNF = @notaFiscal and eSerie = @serie
delete NFe_total WHERE eloja = @loja and eNF = @notaFiscal and eSerie = @serie
delete NFe_prod WHERE eloja = @loja and eNF = @notaFiscal and eSerie = @serie
delete NFe_transp WHERE eloja = @loja and eNF = @notaFiscal and eSerie = @serie
delete NFe_cobr WHERE eloja = @loja and eNF = @notaFiscal and eSerie = @serie
delete NFe_infAdic WHERE eloja = @loja and eNF = @notaFiscal and eSerie = @serie
delete NFe_DUP WHERE eloja = @loja and eNF = @notaFiscal and eSerie = @serie


End

    

 
GO
/****** Object:  StoredProcedure [dbo].[SP_Divergencia_Estoque]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
/*

Atualiza Estoque atraves da Confirmação do ItensVenda

*/

CREATE           Procedure [dbo].[SP_Divergencia_Estoque]
                      
		@TipoDivergencia int,
		@TipoPesquisa	 char(1),
		@Pesquisa        Char(7)

As
 
	Declare	
                @Servidor	 Char(80),
		@Loja 		 Char(5),
		@SQL 		 char(5000),
		@wWhere		 Char(500)
		
 Begin
          

   Select @Servidor =(Select CTS_ServidorRetaguarda From ControleSistema)
   Select @Loja =(Select CTS_Loja From ControleSistema)

   If @TipoDivergencia = 1 
      Begin
	If @TipoPesquisa = 'F'
		Select EL_Referencia,EL_Estoque,ES_Estoque,Pr_descricao, (ES_Estoque - EL_Estoque) as Divergencia 
   	   	   from dmac..Estoque,EstoqueLoja,ProdutoLoja 
   		   where es_loja = LTrim(Rtrim(@Loja)) and es_referencia = El_referencia 
                   and es_referencia = pr_referencia and ES_Estoque <> EL_Estoque and PR_codigofornecedor = RTrim(LTrim(@Pesquisa))
                   Order by EL_Referencia 

        If @TipoPesquisa = 'R'
                Select EL_Referencia,EL_Estoque,ES_Estoque,Pr_descricao, (ES_Estoque - EL_Estoque) as Divergencia 
   	   	   from dmac..Estoque,EstoqueLoja,ProdutoLoja 
   		   where es_loja = LTrim(Rtrim(@Loja)) and es_referencia = El_referencia 
                   and es_referencia = pr_referencia and ES_Estoque <> EL_Estoque and EL_Referencia = RTrim(LTrim(@Pesquisa))
                   Order by EL_Referencia 

	If @TipoPesquisa = 'N'
                Select EL_Referencia,EL_Estoque,ES_Estoque,Pr_descricao, (ES_Estoque - EL_Estoque) as Divergencia 
   	   	   from dmac..Estoque,EstoqueLoja,ProdutoLoja 
   		   where es_loja = LTrim(Rtrim(@Loja)) and es_referencia = El_referencia 
                   and es_referencia = pr_referencia and ES_Estoque <> EL_Estoque 
                   Order by EL_Referencia 
     End



   
    If @TipoDivergencia = 2 
      begin
	If @TipoPesquisa = 'F'
		Select EL_Referencia,EL_Estoque,ES_Estoque,Pr_descricao, (ES_Estoque - EL_Estoque) as Divergencia 
   	   	   from dmac..Estoque,EstoqueLoja,ProdutoLoja 
   		   where es_loja = LTrim(Rtrim(@Loja)) and es_referencia = El_referencia 
                   and es_referencia = pr_referencia and ES_Estoque = EL_Estoque 
                   and EL_Estoque < 0 and PR_codigofornecedor = RTrim(LTrim(@Pesquisa))
                   Order by EL_Referencia 

        If @TipoPesquisa = 'R'
                Select EL_Referencia,EL_Estoque,ES_Estoque,Pr_descricao, (ES_Estoque - EL_Estoque) as Divergencia 
   	   	   from dmac..Estoque,EstoqueLoja,ProdutoLoja 
   		   where es_loja = LTrim(Rtrim(@Loja)) and es_referencia = El_referencia 
                   and es_referencia = pr_referencia and ES_Estoque = EL_Estoque 
                   and EL_Estoque < 0 and EL_Referencia = RTrim(LTrim(@Pesquisa))
                   Order by EL_Referencia 

	If @TipoPesquisa = 'N'
                Select EL_Referencia,EL_Estoque,ES_Estoque,Pr_descricao, (ES_Estoque - EL_Estoque) as Divergencia 
   	   	   from dmac..Estoque,EstoqueLoja,ProdutoLoja 
   		   where es_loja = LTrim(Rtrim(@Loja)) and es_referencia = El_referencia 
                   and es_referencia = pr_referencia and ES_Estoque = EL_Estoque 
                   and EL_Estoque < 0 
                   Order by EL_Referencia 
     end
END

GO
/****** Object:  StoredProcedure [dbo].[SP_EST_Devolucao]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
CREATE    Procedure [dbo].[SP_EST_Devolucao]
                 @NotaFiscal           varchar(09),
				 @serie				   varchar(02)
As

Begin
	Declare  @SQL 		       char(5000)                


        
          Select @SQL = 'UPDATE EstoqueLoja Set EL_Estoque = (EL_Estoque + QTDE) 
                         FROM NFItens, EstoqueLoja 
                         Where EL_Referencia = Referencia and numeroped = ' + @NotaFiscal + 
						 'and serie = ''' + @serie + ''''


          execute (@SQL)

end

GO
/****** Object:  StoredProcedure [dbo].[SP_EST_Outras_Entradas_Estoque_Loja]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
/*
exec SP_EST_Outras_Entradas_Estoque_Loja 130
*/
CREATE     Procedure [dbo].[SP_EST_Outras_Entradas_Estoque_Loja]
		      @PedidoPar	Numeric
AS

Begin 

               
  		       Update EstoqueLoja set EL_Estoque = (EL_Estoque +Qtde)
  		              From EstoqueLoja,NfItens  
		              Where  EL_Loja = Lojaorigem  and
			                 EL_Referencia = Referencia and
			                 NUMEROPED=@PedidoPar
/*			          
  		       Update [SVCENTRALFER].[DMAC].[DBO].Estoque set ES_Estoque = (ES_Estoque +qtde) 
  		              From [SVCENTRALFER].[DMAC].[DBO].Estoque,NfItens 
		              Where  ES_Loja =  Lojaorigem  and
			                 ES_Referencia =  Referencia and
			                 NUMEROPED=@PedidoPar		          
*/			     
	
			  Update NfCapa set TIPONOTA = 'ET', SERIE = 'ET', NF = NUMEROPED where  NUMEROPED=@PedidoPar 
			  Update NfItens set TIPONOTA = 'ET', SERIE = 'ET', NF = NUMEROPED where  NUMEROPED=@PedidoPar       
       
end       
/*
 exec SP_EST_Outras_Entradas_Estoque_Loja
 
 
 select * from nfitens where dataemi = '2012/05/25' and tiponota =  'T'
 
 
*/

GO
/****** Object:  StoredProcedure [dbo].[SP_EST_Transferencia]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
create    Procedure [dbo].[SP_EST_Transferencia]
                 @Pedido           varchar(09)
As

Begin
	Declare  @SQL 		       char(5000)                


        
          Select @SQL = 'UPDATE EstoqueLoja Set EL_Estoque = (EL_Estoque - QTDE) 
                         FROM NFItens, EstoqueLoja 
                         Where EL_Referencia = Referencia and numeroped = ' + @Pedido

          execute (@SQL)

end

GO
/****** Object:  StoredProcedure [dbo].[SP_FIN_Altera_Cliente]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE              PROCEDURE [dbo].[SP_FIN_Altera_Cliente]
			@Codigo int,
                  	@Razao   varchar (40),
			@Cnpj varchar(20),
			@Pessoa char(1),
			@Situacao int,
			@PagamentoCarteira char(1),
			@InscricaoEstadual varchar(15),
			@Cep char(8),
			@Endereco varchar(40),
			@Numero int,
			@Municipio varchar(15),
			@CodigoMunicipio int,
			@Estado char(2),
			@Complemento varchar(15),
			@Bairro varchar(15),
			@Praca int,
			@Telefone varchar(15),
			@Celular varchar(15),
			@Fax varchar(15),
			@DataNascimento datetime,
			@RamoAtividade int,
			@Email varchar(50),
			@Segmento int,
			@EnderecoCobranca varchar(40),
			@NumeroCobranca varchar (5),
			@ComplementoCobranca varchar(15),
			@CepCobranca char(8),
			@BairroCobranca varchar(15),
			@MunicipioCobranca varchar(15),
			@EstadoCobranca char(2),
			@LimiteCredito float

As 

Begin
                  
	update FIN_Cliente set CE_Razao=@Razao, CE_CGC=@Cnpj, CE_TipoPessoa=@Pessoa,
			CE_Situacao=@Situacao,
			CE_PagamentoCarteira=@PagamentoCarteira,
			CE_InscricaoEstadual=@InscricaoEstadual, CE_Cep=@Cep, 
			CE_Endereco=@Endereco, CE_Numero=@Numero, 
			CE_Municipio=@Municipio, 
			CE_CodigoMunicipio=@CodigoMunicipio, CE_Estado=@Estado,
			CE_Complemento=@Complemento, CE_Bairro=@Bairro, 
			CE_Praca=@Praca, CE_Telefone=@Telefone, 
			CE_Celular=@Celular, CE_Fax=@Fax, 
			CE_DataNasc=@DataNascimento, 
			CE_RamoAtividade=@RamoAtividade, CE_EMail=@Email, 
			CE_Segmento=@Segmento,
			CE_EnderecoCobranca=@EnderecoCobranca, 
			CE_NumeroCobranca=@NumeroCobranca,
			CE_ComplCobranca=@ComplementoCobranca, 
			CE_CEPCobranca=@CepCobranca, 
			CE_BairroCobranca=@BairroCobranca,
			CE_MunicipioCobranca=@MunicipioCobranca,
			CE_EstadoCobranca=@EstadoCobranca, 
			CE_LimiteCredito=@LimiteCredito 
			 where CE_CodigoCliente=@Codigo
		 
			
end

/*
select*from fin_cliente where CE_CGC='37021008813'
SP_FIN_Altera_Cliente '935542','marina duarte utysch','37021008813','F', 0,'N', 'Isento','02870150','rua',123,'São Paulo',3550308,'SP','compl','BAIRRO',1,'1111111111','1111111111','1111111111','1991/08/14',17,'email',8, 'rua','123','compl','02870150','BAIRRO','São Paulo','SP',0
*/

GO
/****** Object:  StoredProcedure [dbo].[SP_FIN_Altera_Cliente_Loja]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
create              PROCEDURE [dbo].[SP_FIN_Altera_Cliente_Loja]
			@Codigo int,
                  	@Razao   varchar (40),
			@Cnpj varchar(20),
			@Pessoa char(1),
			@Situacao int,
			@PagamentoCarteira char(1),
			@InscricaoEstadual varchar(15),
			@Cep char(8),
			@Endereco varchar(40),
			@Numero int,
			@Municipio varchar(15),
			@CodigoMunicipio int,
			@Estado char(2),
			@Complemento varchar(15),
			@Bairro varchar(15),
			@Praca int,
			@Telefone varchar(15),
			@Celular varchar(15),
			@Fax varchar(15),
			@DataNascimento datetime,
			@RamoAtividade int,
			@Email varchar(50),
			@Segmento int,
			@EnderecoCobranca varchar(40),
			@NumeroCobranca varchar (5),
			@ComplementoCobranca varchar(15),
			@CepCobranca char(8),
			@BairroCobranca varchar(15),
			@MunicipioCobranca varchar(15)

As 

Begin
                  
	update FIN_Cliente set CE_Razao=@Razao, CE_CGC=@Cnpj, CE_TipoPessoa=@Pessoa,
			CE_Situacao=@Situacao,
			CE_PagamentoCarteira=@PagamentoCarteira,
			CE_InscricaoEstadual=@InscricaoEstadual, CE_Cep=@Cep, 
			CE_Endereco=@Endereco, CE_Numero=@Numero, 
			CE_Municipio=@Municipio, 
			CE_CodigoMunicipio=@CodigoMunicipio, CE_Estado=@Estado,
			CE_Complemento=@Complemento, CE_Bairro=@Bairro, 
			CE_Praca=@Praca, CE_Telefone=@Telefone, 
			CE_Celular=@Celular, CE_Fax=@Fax, 
			CE_DataNasc=@DataNascimento, 
			CE_RamoAtividade=@RamoAtividade, CE_EMail=@Email, 
			CE_Segmento=@Segmento,
			CE_EnderecoCobranca=@EnderecoCobranca, 
			CE_NumeroCobranca=@NumeroCobranca,
			CE_ComplCobranca=@ComplementoCobranca, 
			CE_CEPCobranca=@CepCobranca, 
			CE_BairroCobranca=@BairroCobranca,
			CE_MunicipioCobranca=@MunicipioCobranca
			 where CE_CodigoCliente=@Codigo
		 
			
end

/*
select*from fin_cliente where CE_CGC='37021008813'
SP_FIN_Altera_Cliente '935542','marina duarte utysch','37021008813','F', 0,'N', 'Isento','02870150','rua',123,'São Paulo',3550308,'SP','compl','BAIRRO',1,'1111111111','1111111111','1111111111','1991/08/14',17,'email',8, 'rua','123','compl','02870150','BAIRRO','São Paulo','SP',0
*/

GO
/****** Object:  StoredProcedure [dbo].[SP_FIN_Atualizando_Codigo]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE             PROCEDURE [dbo].[SP_FIN_Atualizando_Codigo]
			@UltimoNumeroCliente int
As 

Begin
                  
	    	update controleSistema set CTS_SequenciaCliente=@UltimoNumeroCliente

end
/*


*/

GO
/****** Object:  StoredProcedure [dbo].[SP_FIN_Calcula_ICMS]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
/*

EXEC SP_FIN_Calcula_ICMS 55555, '0070094', 109,'V', 0
EXEC SP_FIN_Calcula_ICMS 55555, '0070094', 109,'V', 1

*/

create PROCEDURE [dbo].[SP_FIN_Calcula_ICMS]
		@codigoCliente as char(6),
		@referencia as char(7),
		@vlTotalItem as float, -- (valorTotal - desconto)
		@tipoNota as char(1),
		@log as bit
AS
BEGIN

	 print char(10) + '-- -- Inicio SP_FIN_Calcula_ICMS --  --' + char(10)

	 declare @estadoCliente as char(2),
			 @tipoPessoa as char(1),
			 @codigoTipoPessoa as varchar(1),
			 @regiao as varchar(1),
			 @chaveICMS as varchar(10)

		IF object_id('tempCalculoICMS') IS NOT NULL 
		BEGIN
			drop table tempCalculoICMS
		END

	 select @tipoPessoa = ce_Tipopessoa,
	 @estadoCliente = ce_estado 
	 from fin_cliente 
	 where CE_CodigoCliente = @codigoCliente

	 select @regiao = UF_Regiao from fin_Estado where UF_Estado = @estadoCliente

	 select @codigoTipoPessoa = 
	 (select CASE @codigoTipoPessoa 
     WHEN 'F' THEN 2 --FATURADO
	 WHEN 'U' THEN 2 --
	 WHEN 'O' THEN 3 --
	 ELSE 1
	 end)

	 select @tipoNota = 
	 (select CASE @tipoNota 
     WHEN 'P' THEN 'V' --FATURADO
	 end)

	 select @chaveICMS = @regiao + @codigoTipoPessoa

	 if @log = 1
	 BEGIN
		 print 'Código Cliente: ' + @codigoCliente
		 print 'Tipo Cliente: ' + @tipoPessoa
		 print 'Código Tipo Pessoa: ' + @codigoTipoPessoa
		 print 'Estado Cliente: ' + @estadoCliente
		 print 'Região: ' + @regiao
		 print 'Chave @regiao + codigoTipoPessoa: ' + @chaveICMS + char(13)
	 end

	 -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- --

	 declare @substituicaoTributaria as char(2),
			 @codigoReducaoIcms as varchar(1),
			 @icmsSaida as CHAR(2)

	 select @substituicaoTributaria = PR_substituicaotributaria,
	 @codigoReducaoIcms = PR_codigoreducaoicms,
	 @icmsSaida = pr_icmssaida
	 from produtoloja 
	 where pr_referencia = @referencia

	 IF @substituicaoTributaria = 'N' and @codigoReducaoIcms > 0
	 BEGIN
		PRINT 'ICMS20 '
		select @chaveICMS = @regiao + @codigoTipoPessoa +  REPLICATE('0', 2 - LEN(@icmsSaida)) + @icmsSaida + @codigoReducaoIcms + '0'
	 end
	 else
	 begin
		IF @substituicaoTributaria = 'S'
	 		select @chaveICMS = @regiao + @codigoTipoPessoa + '000' + '1';
		ELSE
			select @chaveICMS = @regiao + @codigoTipoPessoa + '000' + @codigoReducaoIcms + '0'
	 end 

	 if @log = 1
	 BEGIN
		 print 'Substituicao Tributaria: ' + @substituicaoTributaria
		 print 'Codigoreducao ICMS: ' + @codigoReducaoIcms
		 print 'ICMS Saida: ' + @icmsSaida
		 print 'Chave: ' + @chaveICMS + char(13)
	 end
	 
	 -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- --	

	 declare @icmsAplicado as float,
			 @Tributacao as char(3),
			 @Cfo as float,
			 @BasedeReducao  as float,
			 @icmsdestino as float,
			 @ValorCalculadoICMS as float,
			 @BasedeCalculoICMS as float
		
	 SELECT @icmsAplicado = IE_ICMSAplicado, --GLB_AliquotaAplicadaICMS
	 @BasedeReducao = IE_BasedeReducao,
	 @icmsdestino = IE_ICMSDestino,			 --wAnexoIten
	 @Tributacao = IE_CST,
	 @BasedeCalculoICMS = 0,
	 @ValorCalculadoICMS = 0
	 from IcmsInterEstadual 
	 where IE_Codigo = @chaveICMS
		 
	 select @ValorCalculadoICMS = (@vlTotalItem * @icmsAplicado) / 100

	 if @BasedeReducao > 0
		select @BasedeCalculoICMS = (@vlTotalItem) - (@vlTotalItem * @BasedeReducao) / 100
	 else
		select @BasedeCalculoICMS = (@vlTotalItem)

     if @log = 1
	 BEGIN
		 print 'Valor ITEM: ' + cast(@vlTotalItem as varchar(20))
		 print 'ICMS Aplicado: ' + cast(@icmsAplicado as varchar(20))
		 print 'Base Reducao: ' + cast(@BasedeReducao as varchar(20))
		 print 'ICMS Destino: ' + cast(@icmsdestino as varchar(20))
		 print 'CST/Tributacao: ' + cast(@Tributacao as varchar(20))
		 print 'Base Calculo ICMS: ' + cast(@BasedeCalculoICMS as varchar(20))
		 print 'Valor Calculado ICMS: ' + cast(@ValorCalculadoICMS as varchar(20))
	 end

	-- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- --	

	if @log = 0
	begin
		--SELECT @BasedeCalculoICMS, @ValorCalculadoICMS, @Tributacao, @icmsdestino into #tempICMS


		CREATE TABLE tempCalculoICMS (
			BasedeCalculoICMS float,
			ValorCalculadoICMS float,
			Tributacao char(3),
			icmsdestino float
		);
		
		insert into tempCalculoICMS values (@BasedeCalculoICMS, @ValorCalculadoICMS, @Tributacao, @icmsdestino)
		select * from tempCalculoICMS
	end 

		--@BasedeCalculoICMS as BaseICMS, @ValorCalculadoICMS as Valoricms, @Tributacao as CST, @icmsdestino AS ICMSAplicado
		--SELECT * FROM #tempICMS

		--if @log = 0
		--RETURN @BasedeCalculoICMS

	print char(10) + '-- -- FIM SP_FIN_Calcula_ICMS --  --' + char(10)

END

/*

EXEC SP_FIN_Calcula_ICMS 55555, '0070094', 109,'V',0

sp_help produtoloja
Select produtoloja.PR_codigoreducaoicms, produtoloja.*, nfitens.* from produtoloja,nfitens 
where nfitens.numeroped = 2612 and pr_referencia = nfitens.referencia 
order by NfItens.Item

SELECT PR_substituicaotributaria,* FROM PRODUTOLOJA WHERE (PR_substituicaoTributaria = 'N' and PR_codigoreducaoicms >0 ) and PR_substituicaotributaria <> 'S'

SELECT * FROM IcmsInterEstadual
SELECT * FROM 
*/

GO
/****** Object:  StoredProcedure [dbo].[SP_FIN_Calcula_ICMS_NFCAPA]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
/*
select * from NFCAPA
SP_FIN_Calcula_ICMS_NFCAPA 2811, '2014/10/13'
EXEC SP_FIN_Calcula_ICMS 55555, '0070094', 109,'V',1

select  BASEICMS, VALORICMS, Tributacao, ICMSAplicado,  * from nfitens where numeroped = 2811
select  BASEICMS, vlricms, * from nfcapa where numeroped = 2811

Select FO_NomeFantasia,C.Cliente,C.BASEICMS,C.vlricms,VE_Codigo, VE_Nome, PR_Descricao, I.Qtde, I.Vlunit,(I.VLUnit * I.Qtde) as VLUnit2,PR_Referencia, PR_ClasseFiscal, C.Desconto as Desconto ,C.VALFRETE,C.TOTALNOTA,PR_Unidade,PR_ICMSSAIDA,PR_ST From ProdutoLoja, NFItens as I, NFCapa as C, Vende,Fornecedor Where PR_Referencia = I.Referencia and VE_Codigo = C.Vendedor and I.NumeroPed = C.NumeroPed and  I.DataEmi = C.DataEmi and PR_CodigoFornecedor=FO_CodigoFornecedor and C.NumeroPed = 2811

*/

--sp_help nfcapa

create PROCEDURE [dbo].[SP_FIN_Calcula_ICMS_NFCAPA]
		@numeroPed as numeric,
		@dataemi as datetime
AS
BEGIN

	 declare @vlTotalItem as float,
			 @tipoNota as char(1),
			 @referencia as char(7),
			 @codigoCliente as char(6),
			 @item as int

	 declare @BasedeCalculoICMS as float,
			 @ValorCalculadoICMS as float,
			 @Tributacao as char(3),
			 @icmsdestino as float,
			 @TotalBasedeCalculoICMS as float,
			 @TotalValorCalculadoICMS as float

	select @tiponota = TIPONOTA, @codigoCliente = cliente 
	from nfcapa 
	where numeroped = @numeroPed and dataemi = @dataemi

	print 'Tipo Nota: ' + cast(@tiponota as varchar(20))
	print 'Numero Ped: ' + cast(@numeroPed as varchar(20))

	select @item = 1

	select @TotalBasedeCalculoICMS = 0
	select @TotalValorCalculadoICMS = 0

	while @item = (select item from nfitens where numeroped = @numeroPed and dataemi = @dataemi and item = @item)
	Begin

		select @vlTotalItem = VLTOTITEM, @referencia = REFERENCIA 
		from nfitens 
		where numeroped = @numeroPed 
		and dataemi = @dataemi 
		and item = @item

		print char(10) + char(10) + 'Total Item: ' + cast(@vlTotalItem as varchar(20))
		print 'Referencia: ' + cast(@referencia as varchar(20))

		EXEC SP_FIN_Calcula_ICMS @codigoCliente, @referencia, @vlTotalItem,@tiponota, 0

		select @BasedeCalculoICMS = BasedeCalculoICMS, 
		@ValorCalculadoICMS = ValorCalculadoICMS, 
		@Tributacao = Tributacao, 
		@icmsdestino = icmsdestino 
		from tempCalculoICMS

		select @TotalBasedeCalculoICMS = @TotalBasedeCalculoICMS + @BasedeCalculoICMS
		select @TotalValorCalculadoICMS = @TotalValorCalculadoICMS + @ValorCalculadoICMS

		print char(10) + char(10) + 'Total Base ICMS: ' + cast(@TotalBasedeCalculoICMS as varchar(20))
		print 'Total Calculo ICMS: ' + cast(@ValorCalculadoICMS as varchar(20))

		update nfitens set BASEICMS = @BasedeCalculoICMS, 
		VALORICMS = @ValorCalculadoICMS, 
		Tributacao = @Tributacao, 
		ICMSAplicado = @icmsdestino
		where numeroped = @numeroPed 
		and dataemi = @dataemi 
		and item = @item

		select @item = @item + 1

	end 

	update nfcapa set BASEICMS = @TotalBasedeCalculoICMS,
	vlricms = @TotalValorCalculadoICMS
	where numeroped = @numeroPed 
	and dataemi = @dataemi

END

/*

EXEC SP_FIN_Calcula_ICMS 55555, '0070094', 109,'V',0

sp_help produtoloja
Select produtoloja.PR_codigoreducaoicms, produtoloja.*, nfitens.* from produtoloja,nfitens 
where nfitens.numeroped = 2612 and pr_referencia = nfitens.referencia 
order by NfItens.Item

SELECT PR_substituicaotributaria,* FROM PRODUTOLOJA WHERE (PR_substituicaoTributaria = 'N' and PR_codigoreducaoicms >0 ) and PR_substituicaotributaria <> 'S'

SELECT * FROM IcmsInterEstadual
SELECT * FROM 
*/

GO
/****** Object:  StoredProcedure [dbo].[SP_FIN_GE_Monta_Campos_capa]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
/*

SP_FIN_GE_Monta_Campos_capa 3747

*/

CREATE PROCEDURE [dbo].[SP_FIN_GE_Monta_Campos_capa]
	@numeroPedido numeric           
As 
Begin

	SET LANGUAGE Português

	declare @CNPJLoja char(14)

	select @CNPJLoja = SUBSTRING(lo_cgc,len(lo_cgc)-13,14) 
	from loja, nfcapa
	where numeroPed = @numeroPedido 
	and lojaorigem = lo_loja

	select 
	'SERRO ADM. CORR. SEG. LTDA' as Corretor,
	'059626.1.02.8522-6' as RegistroNaSusep,
	'Reparo' as CoberturaContratada,
	lo_razao as RepresentanteDeSeguros,
	SUBSTRING(@CNPJLoja,1,2) + '.'
	+ SUBSTRING(@CNPJLoja,3,3) + '.'
	+ SUBSTRING(@CNPJLoja,6,3) + '/'
	+ SUBSTRING(@CNPJLoja,9,4) + '-'
	+ SUBSTRING(@CNPJLoja,13,2) as CNPJRepresentante,
	'0800 198915' as CanaisDeAtendimento,
	dataemi as DataEmissaoSeguro,
	
	ce_razao as Segurado,
	SUBSTRING(ce_cgc,1,3) + '.'
	+ SUBSTRING(ce_cgc,4,3) + '.'
	+ SUBSTRING(ce_cgc,7,3) + '-'
	+ SUBSTRING(ce_cgc,10,2) as CPF,
	ce_cgc as terste,
	ce_telefone as Telefone,
	ce_endereco as Endereco, 
	ce_bairro as Bairro, 
	ce_numero as Numero,
	ce_complemento as Complemento, 
	ce_municipio as Cidade, 
	ce_cep as CEP,
	ce_estado as UF,

	--CT_codigoEstipulanteGE as CodigoEstipulanteGE, 
	--CT_codigoInternoProdutoGE as CodigoInternoProdutoGE, 
	--CE_Razao as Razao, 
	--ce_inscricaoEstadual as InscricaoEstatual,

	--ce_email as Email,

	lo_municipio as EmissaoCidade,
	'dia ' + convert(varchar(10),day(dataemi)) + ' de ' + 
	convert(varchar(10),DATENAME(MONTH, dataemi)) + ' de ' + 
	convert(varchar(10),year(dataemi)) as EmissaoData

	from nfcapa, 
	fin_cliente,
	loja
	where numeroPed = @numeroPedido 
	and cliente = ce_codigoCliente
	and lojaorigem = lo_loja
	and tiponota = 'V'

end

/*
SP_FIN_GE_Monta_Campos_capa 3756
select * from nfcapa where numeroped = 3756
*/

GO
/****** Object:  StoredProcedure [dbo].[SP_FIN_GE_Monta_Campos_item]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
/*
SP_FIN_GE_Monta_Campos_item 3755
select * from nfitens where numeroped = 3749
SELECT * FROM temp_GE_Itens order by certificadoInicio
*/

CREATE PROCEDURE [dbo].[SP_FIN_GE_Monta_Campos_item]
	@numeroPedido numeric           
As 
Begin
	 
--	insert into #TemPGE 
	IF object_id('temp_GE_Itens') IS NOT NULL 
	BEGIN
		drop table temp_GE_Itens
	END

	SELECT 
	convert(varchar(20),
	REPLICATE ( '0' ,4 - LEN(controle.cts_codigoEstipulanteGE)) + RTRIM(controle.cts_codigoEstipulanteGE)
	+ '01' + REPLICATE ( '0' ,6 - LEN(loja.LO_Loja)) + RTRIM(loja.LO_Loja) + 
	REPLICATE ( '0' ,8 - LEN(itens.CertificadoInicio)) + RTRIM(itens.CertificadoInicio))
	as NumeroDoBilhete,

	itens.dataemi as DataEmissaoSeguro,
	itens.dataemi as DataCompraBem,
	itens.dataemi as PerildoGarantiaInicio,
	DATEADD (mm, cast(prod.pr_garantiaFabricante/30 as integer), itens.dataemi) as PerildoGarantiaFIM,
	DATEADD (dd, 01, DATEADD (mm, cast(prod.pr_garantiaFabricante/30 as integer), itens.dataemi)) as PerildoVigenciaInicio,
	DATEADD (mm, (itens.planoGarantia)-12, DATEADD (dd, 01, DATEADD (mm, cast(prod.pr_garantiaFabricante/30 as integer), itens.dataemi))) as PerildoVigenciaFIM,
	--DATEADD (mm, itens.planoGarantia, itens.dataemi) as PerildoVigenciaFIM,

	'R$ ' + convert(char(15),replace(convert(decimal(10,2), itens.VLUNIT),'.',',')) as limiteMaximoIndenizacao,

	prod.pr_descricao as ProdutoSegurado, 
	fornec.fo_nomeFantasia as Marca, 
	itens.referencia as Modelo, 
	'R$ ' + convert(char(15),replace(convert(decimal(10,2), itens.VLUNIT),'.',',')) as ValorProduto, 
	'R$ ' + convert(char(15),replace(convert(decimal(10,2), itens.ge_premioLiquido),'.',',')) as PremioLiquido, 
	'R$ ' + convert(char(15),replace(convert(decimal(10,2), itens.ge_iof),'.',',')) as IOF, 
	'R$ ' + convert(char(15),replace(convert(decimal(10,2), itens.ge_valorCustoSeguradora),'.',',')) as PremioTotal, 

	/*
	itens.qtdeGarantia as QTDEGarantia, 
	itens.item as Item, 
	itens.planoGarantia as PlanoGarantia, 
	itens.lojaOrigem as Loja, 
	cliente as Cliente,
	cast(prod.pr_garantiaFabricante/30 as integer) as GarantiaFabricante, 
	faixa.fpg_premioLiquido as CustoDaSegurandora,
		*/ 
	itens.certificadoInicio as CertificadoInicio, 
	itens.certificadoFim as CertificadoFim
	
	into temp_GE_Itens
	from nfitens itens, 
	produtoLoja prod, 
	FIN_faixapremioge faixa, 
	fornecedor fornec, 
	nfcapa capa,
	ControleSistema controle,
	loja loja

	where itens.numeroPed = @numeroPedido  
	and itens.garantiaEstendida = 'S' 
	and itens.certificadoInicio is not null 
	and prod.pr_referencia = itens.referencia  
	and itens.VLUNIT between faixa.fpg_faixainicial 
	and faixa.fpg_faixaFinal  
	and faixa.fpg_plano = itens.planoGarantia 
	and fornec.fo_codigoFornecedor = prod.pr_codigoFornecedor 
	and capa.numeroPed = @numeroPedido  
	and capa.nf = itens.nf
	and capa.garantiaEstendida = 'S' 
	and capa.tipoNota = 'V'
	and capa.LOJAORIGEM = loja.lo_loja
	order by CertificadoInicio

	--

	declare @contador int

	update temp_GE_Itens set CertificadoFim = '' where certificadoInicio = certificadofim
	set @contador = (select count(CertificadoInicio) from temp_GE_Itens where CertificadoFim not in (CertificadoInicio,''))

	while @contador <> 0
		Begin
		declare @certificadoInicio as char(12)
		declare @certificadoFIM as char(12)
		declare @NumeroCertificado as char(12)

		set @certificadoInicio = (select top 1 CertificadoInicio from temp_GE_Itens where CertificadoFim not in (CertificadoInicio,''))
		set @certificadoFIM = (select top 1 CertificadoFIM from temp_GE_Itens where CertificadoFim not in (CertificadoInicio,''))
		--set @NumeroCertificado = @certificadoInicio - @certificadoFIM

		while @certificadoInicio < @certificadoFIM
		Begin
			set @certificadoInicio = @certificadoInicio + 1
			insert into temp_GE_Itens select top 1 SUBSTRING(NumeroDoBilhete,1,12) + REPLICATE ( '0' ,8 - LEN(@CertificadoInicio)) + RTRIM(@CertificadoInicio), 
			DataEmissaoSeguro, DataCompraBem, PerildoGarantiaInicio, PerildoGarantiaFIM, 
			PerildoVigenciaInicio, PerildoVigenciaFIM, limiteMaximoIndenizacao, ProdutoSegurado, Marca, Modelo, 
			ValorProduto, PremioLiquido, IOF,PremioTotal, @CertificadoInicio, '' 
			from temp_GE_Itens where CertificadoFim not in (CertificadoInicio,'')
			print @certificadoInicio
		end 

		update temp_GE_Itens set CertificadoFim = '' where certificadoInicio <> @certificadoFIM and CertificadoFIM = @certificadoFIM
		set @contador = (select count(CertificadoInicio) from temp_GE_Itens where CertificadoFim not in (CertificadoInicio,''))
	end

--	insert into #tempGE select * from #tempGE where CertificadoInicio <> CertificadoFim

	--SELECT * FROM temp_GE_Itens order by certificadoInicio

end

/*

SP_FIN_GE_Monta_Campos_item 3755


exec SP_FIN_GE_Monta_Campos_item 3752
exec SP_FIN_GE_Monta_Campos_item2 3752

select * from nfitens where numeroped = 3749

*/

GO
/****** Object:  StoredProcedure [dbo].[SP_FIN_Grava_Cliente_Loja]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE                               PROCEDURE [dbo].[SP_FIN_Grava_Cliente_Loja]
			@NovoCodigo numeric,
            @Razao   varchar (40),
			@Cnpj varchar(20),
			@Pessoa char(1),
			@Situacao int,
			@PagamentoCarteira char(1),
			@InscricaoEstadual varchar(15),
			@Cep char(8),
			@Endereco varchar(40),
			@Numero varchar (10),
			@Municipio varchar(15),
			@CodigoMunicipio char(7),
			@Estado char(2),
			@Complemento varchar(15),
			@Bairro varchar(15),
			@Praca int,
			@Telefone varchar(15),
			@Celular varchar(15),
			@Fax varchar(15),
			@DataNascimento datetime,
			@RamoAtividade int,
			@Email varchar(50),
			@Segmento int,
			@EnderecoCobranca varchar(40),
			@NumeroCobranca varchar (5),
			@ComplementoCobranca varchar(15),
			@CepCobranca char(8),
			@BairroCobranca varchar(15),
			@MunicipioCobranca varchar(15),
			@EstadoCobranca char(2),
			@LimiteCredito float,
			@TipoCliente char(1),
			@ClienteFidelidade char(18),
			@Vendedor int,
			@Loja char(5)
As
	SET NOCOUNT ON 

Begin 
	                          
	     Insert Into fin_cliente(CE_CodigoCliente,CE_Razao, CE_CGC, CE_TipoPessoa,
			CE_Situacao,
			CE_PagamentoCarteira, CE_InscricaoEstadual, CE_Cep,
			CE_Endereco, CE_Numero, CE_Municipio, CE_CodigoMunicipio, 
			CE_Estado, CE_Complemento, CE_Bairro, CE_Praca, CE_Telefone, 
			CE_Celular, CE_Fax, CE_DataNasc, CE_RamoAtividade, CE_EMail, 
			CE_Segmento, CE_DataCadastro, CE_EnderecoCobranca, CE_NumeroCobranca, 
			CE_ComplCobranca, CE_CEPCobranca, CE_BairroCobranca,
			CE_MunicipioCobranca, CE_EstadoCobranca, CE_LimiteCredito, 
			CE_DataLimiteCredito,CE_TipoCliente,CE_ClienteFidelidade,
			CE_Vendedor, CE_Loja) 
		values (@NovoCodigo, @Razao, @Cnpj, @Pessoa, @Situacao,
			@PagamentoCarteira, @InscricaoEstadual, @Cep, @Endereco,
			@Numero, @Municipio, @CodigoMunicipio, @Estado, @Complemento, 
			@Bairro, @Praca, @Telefone, @Celular, @Fax, @DataNascimento,
			@RamoAtividade, @Email, @Segmento, convert(char(10),getdate(),102), 
			@EnderecoCobranca, @NumeroCobranca, @ComplementoCobranca, 
			@CepCobranca, @BairroCobranca, @MunicipioCobranca, 
			@EstadoCobranca, @LimiteCredito, convert(char(10),getdate(),102), @TipoCliente,
			@ClienteFidelidade, @Vendedor, @Loja)	            
	
	if @TipoCliente  = 'F'
	   begin
		Insert Into dmac..fin_cliente(CE_CodigoCliente,CE_Razao, CE_CGC, CE_TipoPessoa,
			CE_Situacao,
			CE_PagamentoCarteira, CE_InscricaoEstadual, CE_Cep,
			CE_Endereco, CE_Numero, CE_Municipio, CE_CodigoMunicipio, 
			CE_Estado, CE_Complemento, CE_Bairro, CE_Praca, CE_Telefone, 
	   		CE_Celular, CE_Fax, CE_DataNasc, CE_RamoAtividade, CE_EMail, 
			CE_Segmento, CE_DataCadastro, CE_EnderecoCobranca, CE_NumeroCobranca, 
			CE_ComplCobranca, CE_CEPCobranca, CE_BairroCobranca,
			CE_MunicipioCobranca, CE_EstadoCobranca, CE_LimiteCredito, 
			CE_DataLimiteCredito,CE_TipoCliente,CE_clienteFidelidade) 
		values (@NovoCodigo, @Razao, @Cnpj, @Pessoa, @Situacao,
			@PagamentoCarteira, @InscricaoEstadual, @Cep, @Endereco,
			@Numero, @Municipio, @CodigoMunicipio, @Estado, @Complemento, 
			@Bairro, @Praca, @Telefone, @Celular, @Fax, @DataNascimento,
			@RamoAtividade, @Email, @Segmento, convert(char(10),getdate(),102), 
			@EnderecoCobranca, @NumeroCobranca, @ComplementoCobranca, 
			@CepCobranca, @BairroCobranca, @MunicipioCobranca, 
			@EstadoCobranca, @LimiteCredito, convert(char(10),getdate(),102), @TipoCliente,
			@ClienteFidelidade)	            
	   end	
end

GO
/****** Object:  StoredProcedure [dbo].[SP_FIN_Ler_Cliente_Por_Parametro]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
create                PROCEDURE [dbo].[SP_FIN_Ler_Cliente_Por_Parametro]
	                @CNPJ   char (10)		

As 

Begin
    	 
	     	select * from fin_cliente where CE_CGC = @CNPJ

end

GO
/****** Object:  StoredProcedure [dbo].[SP_FIN_Ler_Clientes]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
create                PROCEDURE [dbo].[SP_FIN_Ler_Clientes]

As 

Begin
       	 
	     	select * from fin_cliente order by CE_CodigoCliente

end

GO
/****** Object:  StoredProcedure [dbo].[SP_FIN_Ler_Clientes_Por_Parametro_Cnpj]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO




CREATE               PROCEDURE [dbo].[SP_FIN_Ler_Clientes_Por_Parametro_Cnpj]
                   @Cnpj   varchar (20)

As 

Begin
    	 
	     	select * from fin_cliente where CE_CGC = @Cnpj-- and CE_CodigoCliente > 900000

end





GO
/****** Object:  StoredProcedure [dbo].[SP_FIN_Ler_Clientes_Por_Parametro_Codigo]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
create               PROCEDURE [dbo].[SP_FIN_Ler_Clientes_Por_Parametro_Codigo]
                   @CodigoCliente   int 

As 

Begin
     	 
	     	select * from fin_cliente where CE_CodigoCliente = @codigoCliente

end

GO
/****** Object:  StoredProcedure [dbo].[SP_FIN_Ler_Codigo_Municipio_Por_Parametro]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
create                 PROCEDURE [dbo].[SP_FIN_Ler_Codigo_Municipio_Por_Parametro]
                   @Municipio varchar(40)

As 

Begin
    	 
	     	select * from fin_municipio where Mun_Nome like @Municipio + '%'

end

GO
/****** Object:  StoredProcedure [dbo].[SP_FIN_Ler_Estado]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
create                 PROCEDURE [dbo].[SP_FIN_Ler_Estado]
			

As 

Begin
     

	select * from FIN_Estado order by UF_Estado

end


/*
 exec SP_FIN_Ler_Bancos
*/

GO
/****** Object:  StoredProcedure [dbo].[SP_FIN_Ler_Municipio]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE                 PROCEDURE [dbo].[SP_FIN_Ler_Municipio]
			@MunicipioCodigo char(7)
                  

As 

Begin
     

	select * from Fin_Municipio where Mun_Codigo like @MunicipioCodigo

end


/*
 SQL = "Select * from Municipio " _
            & "where Mun_Codigo Like  '" & rsClientePedido("CE_Mun_Codigo") & "'"
*/

GO
/****** Object:  StoredProcedure [dbo].[SP_FIN_Pesquisa_Cep]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
create                  PROCEDURE [dbo].[SP_FIN_Pesquisa_Cep]
			@Cep varchar(50)

As 

Begin

	select * from FIN_CEP where CEP=@Cep

end

GO
/****** Object:  StoredProcedure [dbo].[SP_FIN_Pesquisa_Cliente_Ficha_Financeira_Por_Codigo]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
create                  PROCEDURE [dbo].[SP_FIN_Pesquisa_Cliente_Ficha_Financeira_Por_Codigo]
			@Codigo varchar(10)

As 

Begin

		select * from clientefichafinanceira where codigo=@codigo

end

GO
/****** Object:  StoredProcedure [dbo].[SP_FIN_Pesquisa_Codigo]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
create                PROCEDURE [dbo].[SP_FIN_Pesquisa_Codigo]
                   @Codigo int
As 

Begin
	     	select CE_CodigoCliente from fin_cliente where CE_CodigoCliente=@Codigo
end

GO
/****** Object:  StoredProcedure [dbo].[SP_FIN_Pesquisa_Codigo_Cliente]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
create                 PROCEDURE [dbo].[SP_FIN_Pesquisa_Codigo_Cliente]
                   @Codigo int
As 

Begin
	     	select * from fin_cliente where CE_CodigoCliente=@Codigo
end

GO
/****** Object:  StoredProcedure [dbo].[SP_FIN_Pesquisa_Municipio]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
create                   PROCEDURE [dbo].[SP_FIN_Pesquisa_Municipio]

As 

Begin

	select * from Fin_Municipio

end

GO
/****** Object:  StoredProcedure [dbo].[SP_FIN_Pesquisa_Municipio_Por_Parametro]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO


CREATE                    PROCEDURE [dbo].[SP_FIN_Pesquisa_Municipio_Por_Parametro]
                    @Municipio char(60),
                    @UF CHAR(4)

As 

Begin

       select Mun_Codigo, Mun_UF from Fin_Municipio where Mun_Nome=@municipio and Mun_UF = @UF and Mun_Codigo > 0

end



GO
/****** Object:  StoredProcedure [dbo].[SP_FIN_Pesquisa_Ramo_Atividade_Por_Codigo]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
create                 PROCEDURE [dbo].[SP_FIN_Pesquisa_Ramo_Atividade_Por_Codigo]
			@CodigoCliente int

As 

Begin

	select CE_RamoAtividade, CE_Segmento from FIN_Cliente
					where CE_CodigoCliente=@CodigoCliente

end

GO
/****** Object:  StoredProcedure [dbo].[SP_FIN_Pesquisa_Ramo_Atividade_Por_Pessoa]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
create                 PROCEDURE [dbo].[SP_FIN_Pesquisa_Ramo_Atividade_Por_Pessoa]
			@TipoPessoa char(1)

As 

Begin

	select RMO_Codigo, RMO_DescricaoRamo from FIN_RamoAtividade 
					where RMO_Pessoa=@TipoPessoa order by RMO_DescricaoRamo

end

GO
/****** Object:  StoredProcedure [dbo].[SP_FIN_Pesquisa_Segmento_Por_Codigo_Cliente]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
create                 PROCEDURE [dbo].[SP_FIN_Pesquisa_Segmento_Por_Codigo_Cliente]
			@CodigoCliente int

As 

Begin

	select CE_RamoAtividade, CE_Segmento from FIN_Cliente 
				where CE_CodigoCliente =@CodigoCliente

end

GO
/****** Object:  StoredProcedure [dbo].[SP_FIN_Pesquisa_Segmento_Por_Ramo_Atividade]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
create                 PROCEDURE [dbo].[SP_FIN_Pesquisa_Segmento_Por_Ramo_Atividade]
			@RamoAtividade int

As 

Begin

	select * from FIN_Segmento where SEG_RamoAtividade=@RamoAtividade

end

GO
/****** Object:  StoredProcedure [dbo].[SP_FIN_Pesquisa_Ultimo_Numero_Cliente]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE                       PROCEDURE [dbo].[SP_FIN_Pesquisa_Ultimo_Numero_Cliente]

As 
	set nocount on
Begin
	declare @NovoCodigo int
	--update controlesistema set cts_sequenciaCliente = 981089
	select @NovoCodigo=((select max (CTS_SequenciaCliente) from controleSistema)+1)

	create table #Temp_Pesquisa_Ultimo_Numero_Cliente
			(Temp_NovoCodigo  int)

	insert into #Temp_Pesquisa_Ultimo_Numero_Cliente(Temp_NovoCodigo) values (@NovoCodigo)      	
	
	select (Temp_NovoCodigo)as UltNumCliente from #Temp_Pesquisa_Ultimo_Numero_Cliente
end

GO
/****** Object:  StoredProcedure [dbo].[SP_FIN_Situacao_Cliente]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
create              PROCEDURE [dbo].[SP_FIN_Situacao_Cliente]

As 

Begin

	select * from fin_SituacaoCliente order by SI_CodigoSituacao

end

GO
/****** Object:  StoredProcedure [dbo].[SP_GLB_Acerta_Cliente]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
--SELECT * FROM FIN_CLIENTE WHERE CE_CGC LIKE '%.%'
--SELECT * FROM FIN_Cliente WHERE CE_CodigoCliente = '925817'

--EXEC SP_GLB_Acerta_Cliente '925817'

CREATE Procedure [dbo].[SP_GLB_Acerta_Cliente]
	@CodigoCliente INT
As

Begin
	
	declare @CGC varchar(14)

	update fin_cliente set cE_cgc = substring(CE_CGC,2,14) 
	where substring(CE_CGC,1,1) = '0' and len(RTRIM(CE_CGC)) = 15
	and ce_codigoCliente = @CodigoCliente 

	select @CGC = CE_CGC from FIN_Cliente where 
	ce_codigoCliente = @CodigoCliente 

	exec @cgc = TiraLetras @cgc

	update FIN_Cliente set CE_CGC = @cgc where 
	ce_codigoCliente = @CodigoCliente

	-- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- 

	declare @IE varchar(20)

	select @IE = CE_InscricaoEstadual from FIN_Cliente where 
	ce_codigoCliente = @CodigoCliente 

	exec @IE = TiraLetras @IE

	update FIN_Cliente set CE_InscricaoEstadual = @IE where 
	ce_codigoCliente = @CodigoCliente

	-- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- 

	declare @telefone varchar(20)

	--update fin_cliente set CE_Telefone = replace(CE_Telefone,'-','') 
	--where CE_Telefone like '%-%'
	--and ce_codigoCliente = @CodigoCliente 

	select @telefone = CE_Telefone from FIN_Cliente where 
	ce_codigoCliente = @CodigoCliente

	exec @telefone = TiraLetras @telefone

	update FIN_Cliente set CE_Telefone = @telefone where 
	ce_codigoCliente = @CodigoCliente

	-- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- 

	select @telefone = ce_celular from FIN_Cliente where 
	ce_codigoCliente = @CodigoCliente

	exec @telefone = TiraLetras @telefone

	update FIN_Cliente set ce_celular = @telefone where 
	ce_codigoCliente = @CodigoCliente

	-- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- 

	select @telefone = CE_Fax from FIN_Cliente where 
	ce_codigoCliente = @CodigoCliente

	exec @telefone = TiraLetras @telefone

	update FIN_Cliente set CE_Fax = @telefone where 
	ce_codigoCliente = @CodigoCliente


	-- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- 

	DECLARE @cep as varchar(20)

	select @cep = CE_CEP from FIN_Cliente where 
	ce_codigoCliente = @CodigoCliente

	exec @cep = TiraLetras @cep

	update FIN_Cliente set CE_CEP = @cep where 
	ce_codigoCliente = @CodigoCliente


	-- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- 

	update FIN_Cliente set CE_CodigoMunicipio = Mun_Codigo from FIN_Municipio, FIN_Cliente 
	where CE_CodigoMunicipio is null and rtrim(CE_Municipio) = Mun_Nome
	and ce_codigoCliente = @CodigoCliente 

	update FIN_Cliente set CE_CodigoMunicipio = '3550308'  
	where CE_Municipio = 'SAO PAULO'
	and ce_codigoCliente = @CodigoCliente 

	-- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- 

end

--exec SP_CORRECAO_DIARIA

 
GO
/****** Object:  StoredProcedure [dbo].[SP_GLB_Importa_Cliente]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
/*

Atualiza Estoque atraves da Confirmação do ItensVenda

exec SP_GLB_Importa_Cliente 55555
*/

CREATE Procedure [dbo].[SP_GLB_Importa_Cliente]
	@codigoCliente  	int

	as

	Begin
	
	BEGIN TRANSACTION

	if @codigoCliente < 900000 
	begin
		delete fin_cliente where CE_CodigoCliente = @codigoCliente 
		insert into fin_cliente select top 1 * from [SVDMAC].[DMAC].[DBO].FIN_Cliente where CE_CodigoCliente = @codigoCliente
	end 

	COMMIT TRANSACTION

End

 
 
GO
/****** Object:  StoredProcedure [dbo].[SP_GLB_Valida_Cliente]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO



CREATE PROCEDURE [dbo].[SP_GLB_Valida_Cliente]
	@codigoCliente varchar(10)
as
BEGIN

	IF object_id('temp_Fin_Cliente_Erro') IS NOT NULL 
	BEGIN
		drop table temp_Fin_Cliente_Erro
	END

	CREATE TABLE [dbo].[temp_Fin_Cliente_Erro](
		[campoErrado] [varchar](20) NULL
	) ON [PRIMARY]
	
	declare @CE_CGC as varchar(15)
	declare @CE_InscricaoEstadual as varchar(18)
	declare @CE_Razao as varchar(60)
	declare @CE_Endereco as varchar(60)
	declare @CE_Bairro as varchar(40)
	declare @CE_Municipio as varchar(40)
	declare @CE_Estado as varchar(2)
	declare @CE_CEP as varchar(8)
	declare @CE_Telefone as varchar(9)
	declare @CE_Fax as varchar(9)
	declare @CE_EMail as varchar(60)
	declare @CE_Numero as varchar(10)
	declare @CE_Complemento as varchar(30)
	declare @CE_Celular as varchar(15)
	declare @CE_Mun_Codigo as varchar(7)
	declare @CE_CodigoMunicipio as varchar(7)
	declare @valida as char(1)

	select Top 1 
	@CE_CGC = CE_CGC,
	@CE_InscricaoEstadual = CE_InscricaoEstadual,
	@CE_Razao = CE_Razao,
	@CE_Endereco = CE_Endereco,
	@CE_Bairro = CE_Bairro,
	@CE_Municipio = CE_Municipio,
	@CE_Estado = CE_Estado,
	@CE_CEP = CE_CEP,
	@CE_Telefone = CE_Telefone,
	@CE_Fax = CE_Fax,
	@CE_EMail = CE_EMail,
	@CE_Numero = CE_Numero,
	@CE_Complemento = CE_Complemento,
	@CE_Celular = CE_Celular,
	@CE_Mun_Codigo = CE_Mun_Codigo,
	@CE_CodigoMunicipio = CE_CodigoMunicipio
	from fin_cliente where CE_CodigoCliente = @codigoCliente
	--sp_help fin_cliente
	--select * from fin_cliente where CE_CodigoCliente = '7'

	-- -- VALIDA CNPJ -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- 		  

	if LEN(@CE_CGC) > 11
	begin

		DECLARE @INDICE INT, @SOMA INT, @DIG1 INT, @DIG2 INT, @VAR1 INT, @VAR2 INT, @RESULTADO CHAR(1)

		SET @SOMA = 0
		SET @INDICE = 1
		SET @RESULTADO = 'N'
		SET @VAR1 = 5 

		WHILE (@INDICE <= 4)
		BEGIN
		  SET @Soma = @Soma + CONVERT(INT,SUBSTRING(@CE_CGC,@INDICE,1)) * @VAR1
		  SET @INDICE = @INDICE + 1 /* Navegando um-a-um até < = 4, as quatro primeira posições */
		  SET @VAR1 = @VAR1 - 1       /* subtraindo o algorítimo de 5 até 2 */
		END

		SET @VAR2 = 9
	    
		WHILE (@INDICE <= 12)
		BEGIN
		  SET @Soma = @Soma + CONVERT(INT,SUBSTRING(@CE_CGC,@INDICE,1)) * @VAR2
		  SET @INDICE = @INDICE + 1
		  SET @VAR2 = @VAR2 - 1            
		END

	   SET @DIG1 = (@soma % 11)

	   IF @DIG1 < 2
			SET @DIG1 = 0;
	   ELSE
			SET @DIG1 = 11 - (@soma % 11);

		SET @INDICE = 1
		SET @SOMA = 0
		SET @VAR1 = 6 
		SET @RESULTADO = 'N'

		WHILE (@INDICE <= 5)
		BEGIN
		  SET @Soma = @Soma + CONVERT(INT,SUBSTRING(@CE_CGC,@INDICE,1)) * @VAR1
		  SET @INDICE = @INDICE + 1 /* Navegando um-a-um até < = 5, as quatro primeira posições */
		  SET @VAR1 = @VAR1 - 1       /* subtraindo o algorítimo de 6 até 2 */
		END

	 
		SET @VAR2 = 9
		WHILE (@INDICE <= 13)
		BEGIN
		  SET @Soma = @Soma + CONVERT(INT,SUBSTRING(@CE_CGC,@INDICE,1)) * @VAR2
		  SET @INDICE = @INDICE + 1
		  SET @VAR2 = @VAR2 - 1            
		END

		   SET @DIG2 = (@soma % 11)
		   IF @DIG2 < 2
			   SET @DIG2 = 0;
		   ELSE /* SE O RESTO DA DIVISÃO NÃO FOR < 2*/
			   SET @DIG2 = 11 - (@soma % 11);

		--IF not (@DIG1 = SUBSTRING(@CE_CGC,LEN(@CE_CGC)-1,1)) AND (@DIG2 = SUBSTRING(@CE_CGC,LEN(@CE_CGC),1))
		  --insert into temp_Fin_Cliente_Erro values ('CGC')
		  
	end 			
		
	-- -- VALIDA CPF -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- 		  
		
	
	if LEN(@CE_CGC) <= 11
	begin
	
	DECLARE   @CE_CGC_TEMP VARCHAR(11),
			  @DIGITOS_IGUAIS CHAR(1)
	          
	  SET @RESULTADO = 'N'

	  /*
		  Verificando se os digitos são iguais
		  A Principio CPF com todos o números iguais são Inválidos
		  apesar de validar o Calculo do digito verificado
		  EX: O CPF 00000000000 é inválido, mas pelo calculo
		  Validaria
	  */

	  SET @CE_CGC_TEMP = SUBSTRING(@CE_CGC,1,1)

	  SET @INDICE = 1
	  SET @DIGITOS_IGUAIS = 'S'

	  WHILE (@INDICE <= 11)
	  BEGIN
		IF SUBSTRING(@CE_CGC,@INDICE,1) <> @CE_CGC_TEMP
		  SET @DIGITOS_IGUAIS = 'N'
		SET @INDICE = @INDICE + 1
	  END;

	  --Caso os digitos não sejão todos iguais Começo o calculo do digitos
	  IF @DIGITOS_IGUAIS = 'N' 
	  BEGIN
		--Cálculo do 1º dígito
		SET @SOMA = 0
		SET @INDICE = 1
		WHILE (@INDICE <= 9)
		BEGIN
		  SET @Soma = @Soma + CONVERT(INT,SUBSTRING(@CE_CGC,@INDICE,1)) * (11 - @INDICE);
		  SET @INDICE = @INDICE + 1
		END

		SET @DIG1 = 11 - (@SOMA % 11)

		IF @DIG1 > 9
		  SET @DIG1 = 0;

		-- Cálculo do 2º dígito }
		SET @SOMA = 0
		SET @INDICE = 1
		WHILE (@INDICE <= 10)
		BEGIN
		  SET @Soma = @Soma + CONVERT(INT,SUBSTRING(@CE_CGC,@INDICE,1)) * (12 - @INDICE);
		  SET @INDICE = @INDICE + 1
		END

		SET @DIG2 = 11 - (@SOMA % 11)

		IF @DIG2 > 9
		  SET @DIG2 = 0;

		-- Validando
		--IF (@DIG1 = SUBSTRING(@CE_CGC,LEN(@CE_CGC)-1,1)) AND (@DIG2 = SUBSTRING(@CE_CGC,LEN(@CE_CGC),1))
		  --insert into temp_Fin_Cliente_Erro values ('CGC')
	  END
	
	end	


	-- -- INSCRICAO ESTADUAL -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- 		  
	--exec SP_glb_Valida_Cliente 7
	
	IF @CE_InscricaoEstadual = 'ISENTO'
		select @valida = 1
	ELSE
		exec @valida = SP_GLB_VALIDA_IE @CE_Estado,@CE_InscricaoEstadual
	
	if @valida = 0
		insert into temp_Fin_Cliente_Erro values ('InscricaoEstadual')
		
	
	-- -- RAZAO -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- 		  
	--declare @CE_Razao as varchar(60)
	IF LEN(@CE_Endereco) < 5
	insert into temp_Fin_Cliente_Erro values ('Razao')
	
	-- -- ENDERECO -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- 		  
	--declare @CE_Endereco as varchar(60)
	
	IF LEN(@CE_Endereco) < 5
		insert into temp_Fin_Cliente_Erro values ('Endereco')
	
	-- -- BAIRRO -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- 		  
	--declare @CE_Bairro as varchar(40)
	
	IF LEN(@CE_Bairro) < 3
		insert into temp_Fin_Cliente_Erro values ('Bairro')
	
	-- -- MUNICIPIO -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- 		  
	--declare @CE_Municipio as varchar(40)
	IF LEN(@CE_Municipio) < 3
		insert into temp_Fin_Cliente_Erro values ('Municipio')
	
	-- -- ESTADO -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- 		  
	--declare @CE_Estado as varchar(2)
	
	IF LEN(@CE_Estado) <> 2
		insert into temp_Fin_Cliente_Erro values ('Estado')
	
	-- -- CEP -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- 		  
	--declare @CE_CEP as varchar(8)
	
	IF LEN(@CE_CEP) <> 8
		insert into temp_Fin_Cliente_Erro values ('CEP')
	
	-- -- TELEFONE, FAX, CELULAR -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- 		  
	--declare @CE_Telefone as varchar(9)
	
	IF LEN(rtrim(@CE_Telefone)) NOT BETWEEN 8 AND 12
		insert into temp_Fin_Cliente_Erro values ('Telefone')
	
	IF LEN(@CE_Fax)  NOT BETWEEN 8 AND 12
		--insert into temp_Fin_Cliente_Erro values ('FAX')
	
	IF LEN(@CE_Celular) NOT BETWEEN 8 AND 12
		insert into temp_Fin_Cliente_Erro values ('Celular')
	
	-- -- EMAIL -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- 		  
	--exec SP_glb_Valida_Cliente 7
	
	if len(@CE_EMail) > 0
		exec @valida = SP_GLB_Valida_Email @CE_EMail
	
	if (@valida) = 0
		--insert into temp_Fin_Cliente_Erro values ('Email')
				
	
	-- -- NUMERO -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- 		  
	--declare @CE_Numero as varchar(10)
	
	select @ce_numero = @ce_numero
	
	--PRINT 'numero'
	--PRINT @CE_Numero
	IF LEN(@CE_Numero) < 1 or LEN(@CE_Numero) > 6 or @CE_Numero < 1 or @CE_Numero is null
		insert into temp_Fin_Cliente_Erro values ('Numero')
	
	-- -- COMPLEMENTO -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- 		  
	--declare @CE_Complemento as varchar(30)
	--insert into temp_Fin_Cliente_Erro values ('Complemento')
	
	-- -- MUNICIPIO -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- 		  
	--declare @CE_Mun_Codigo as varchar(7)
	--declare @CE_CodigoMunicipio as varchar(7)
	IF LEN(@CE_Mun_Codigo) <> 7
		insert into temp_Fin_Cliente_Erro values ('CodigoMunicipio')
	
		
	-- -- VALIDA NUMERO -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- 		  		
	
	
	
  select * from temp_Fin_Cliente_Erro
  --drop table temp_Fin_Cliente_Erro      
	
END

/*

exec SP_glb_Valida_Cliente 55555

*/




  
GO
/****** Object:  StoredProcedure [dbo].[SP_GravaComplementoVenda]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER OFF
GO
/****** 

Grava ComplementoVenda 

*/

create         Procedure [dbo].[SP_GravaComplementoVenda]	
	@Pedido  		Char(10),
	@Codigo		        Integer,
	@Sequencia              Integer,
	@ValorCampo             varchar(8000)
As

Begin
        Declare @SQL                    varchar    (8000)
	
	Begin Transaction
/*
	  Delete ComplementoVenda
	    Where   COV_NumeroPedido = @Pedido and
		    COV_CodigoComplemento = @Codigo and
	            COV_SequenciaComplemento = @Sequencia

          Insert ComplementoVenda
                (COV_NumeroPedido, 
		 COV_CodigoComplemento,
	         COV_SequenciaComplemento,
		 COV_ValorComplemento)
          Values
                (@Pedido,
	         @Codigo,
		 @Sequencia,
                 @ValorCampo)
					  			
          if @codigo = 1 and @Sequencia=4 
              Delete ComplementoVenda
	      Where   COV_NumeroPedido = @Pedido and
		      COV_CodigoComplemento = 1 and
	              COV_SequenciaComplemento = 17
          
*/
	IF @Codigo = 1
	   Select @SQL = 'UPDATE NFCapa Set ' + LTrim(RTrim(@ValorCampo)) + ' Where NumeroPed = ' + @Pedido
	   Execute (@SQL)

	
	IF @Codigo = 2			
	   Select @SQL = 'UPDATE NFItens Set ' +LTRIM(RTRIM(@ValorCampo)) + ' Where NumeroPed = ' + @Pedido
	Execute (@SQL)	
	

	
	If @@ERROR = 0 
	  Begin 		
	  	Commit Transaction 
          end 	
	
	Else
		Rollback Transaction	
       End


/*

Exec SP_GravaComplementoVenda 1,1,1,'Valor'

*/

GO
/****** Object:  StoredProcedure [dbo].[SP_Inicializa_Inventario]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
/*
Procedure Totaliza Capa Nota Fiscal a partir do Numero do Pedido
*/

CREATE PROCEDURE [dbo].[SP_Inicializa_Inventario]
         @NroLocais int
as    
         DECLARE	@Contador	int,
                        @Cont           char(10),
                        @SQL            char(1000),
                        @Local          char(10)                    
BEGIN
 delete controleinv
 
 delete contagem
 
 delete dupla
 
 delete estoqueinv
 
 delete locais

 select @Contador = 1
 While @Contador <= @NroLocais
   Begin 
      select @Cont=@Contador
      select @Local = 'Local ' + @Cont 
      Select @SQL='Insert into locais(CL_CodigoLocal, CL_NomeLocal, CL_Situacao)  values(' +  LTrim(Rtrim(@Cont)) +
      ',' +''''+  LTrim(Rtrim(@Local))+''''+ ',' + '''A''' + ')'
      Exec (@SQL)
      set @Contador = @Contador + 1
    End

End

/*
  exec SP_Inicializa_Inventario 73
*/



GO
/****** Object:  StoredProcedure [dbo].[SP_LB]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO


CREATE  PROCEDURE [dbo].[SP_LB]

	    @Loja	Char(5),
	    @NumeroPedido	Char(7)

AS

	DECLARE		
	@SQL  char(4000)
 --   @NomeServidor   char(40),
 --   @Cliente        numeric
                        
BEGIN

       -- Select @NomeServidor = (Select LO_NomeServidor from Loja where lo_loja=ltrim(rtrim(@Loja)))

       Select @SQL ='Update nfcapa set liberaBloqueio = ' + '''' + 'S' + '''' + '
                     Where  numeroped = ' + '''' + @NumeroPedido + ''''  
	   Exec (@SQL)


	   Select @SQL ='if (select count(*) from nfcapa 
					 where liberaBloqueio = ''S''
                     and  numeroped = ''' + @NumeroPedido + ''') = 0
					 delete autorizacaodesconto 
					 where aut_numeroPedido = ''' + @NumeroPedido + ''' and aut_loja = ''' + ltrim(rtrim(@Loja)) + ''''

		print (@sql)
        Exec (@SQL)



END








GO
/****** Object:  StoredProcedure [dbo].[SP_LiberaMercadoriaLoja]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
/*
	Backup
*/

Create      Procedure [dbo].[SP_LiberaMercadoriaLoja]
	             @NF 	 numeric,
		     @Fornecedor numeric,
		     @Serie	 char(2),
		     @Loja	 char(5)

As


Declare		@SQL		   Char(500),
			@NomeServidor  char(50)
                
  select @NomeServidor = (Select LO_Nomeservidor from Loja where LO_Loja = rtrim(ltrim(@Loja)))
  
   select @SQL = 'update ' + LTrim(Rtrim(@NomeServidor)) + 'estoqueloja set EL_Estoque = CI_quantidade
                 from ItemNfcompra,' + LTrim(Rtrim(@NomeServidor)) + 'EstoqueLoja as E
		 where e.el_refeencia = ci_referencia and ci_notafiscal = ' + @NF + ' 
		 and ci_serie = ' + '''' + @Serie + '''' + ' 
                 and ci_loja = ' + '''' + Ltrim(RTrim(@Loja)) + '''' + ' and ci_fornecedor = ' + @fornecedor

  print @SQL

GO
/****** Object:  StoredProcedure [dbo].[SP_Replicar_Pedido_Venda]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
/*
 
*/ 
create  Procedure [dbo].[SP_Replicar_Pedido_Venda]
                  @NumeroPedido	varchar(10),
                  @NumeroPedidoNovo	varchar (10)
		 
As 

Begin
     Declare @SQL               char (500),
             @NomeTabela        char(20)

      Select @SQL = ' '
     Select @NomeTabela = ' '
     Select @NomeTabela = 'TempCapaPedido' + ltrim(rtrim(@NumeroPedido))

     Select @SQL = 'Select * Into ' +  @NomeTabela + ' from nfcapa Where NumeroPed = ' + @NumeroPedido
     Exec (@SQL)
     --print @sql
     Select @SQL=' '
     Select @SQL='Update ' + @NomeTabela + ' set TipoNota= ' + '''PD''' + ',Numeroped =' + @NumeroPedidoNovo 
     Exec (@SQL)
     Select @SQL='Insert into NFCapa Select * from  ' + @NomeTabela 
     Exec (@SQL)
     --  print @sql
     Select @SQL='Drop Table ' + @NomeTabela 
     Exec (@SQL)
    -- print @sql


     Select @SQL = ' '
     Select @NomeTabela = ' '
     Select @NomeTabela = 'TempItemPedido' + ltrim(rtrim(@NumeroPedido))
     Exec (@sql)
     Select @SQL = 'Select * Into ' +  ltrim(rtrim(@NomeTabela)) + ' from Nfitens Where NumeroPed = ' + @NumeroPedido
     Exec (@sql)
    -- print @sql
     Select @SQL='Update ' + @NomeTabela + ' set TipoNota= ' + '''PD''' + ',Numeroped =' + @NumeroPedidoNovo
     Exec (@SQL)
     Select @SQL='Insert into NFitens Select * from  ' + @NomeTabela 
     Exec (@SQL)
      -- print @sql
     Select @SQL='Drop Table ' + @NomeTabela 
     Exec (@SQL)
      --  print @sql
end         
  

/*
  Exec SP_Replicar_Pedido_Venda 293,325
 select * from nfcapa where numeroped=325
  select * from nfitens where numeroped=325
  delete nfcapa where numeroped=325
  delete nfitens where numeroped=325
 */

GO
/****** Object:  StoredProcedure [dbo].[SP_Totaliza_Capa_Nota_Fiscal]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
/*
Procedure Totaliza Capa Nota Fiscal
a partir do Numero do Pedido

*/
 
CREATE       PROCEDURE [dbo].[SP_Totaliza_Capa_Nota_Fiscal]
	@NumeroPedido		varchar(6)

AS  
	DECLARE		@ValorMercadoria	Float,
			@TotalICMS		Float,
			@TotalNF	        Float,
			@BaseICMS               Float,
                        @Desconto               Float,
                        @SubTotal               Float
                       
BEGIN
            Update NFITENS SET VALORICMS=(((vltotitem - desconto) * PR_ICMPDV) / 100)
                   From NFITENS,PRODUTO
                   Where PR_Referencia=Referencia and NUMEROPED=@NumeroPedido

            Select @ValorMercadoria = (Select sum(vltotitem) 
                                      From NFItens Where NUMEROPED = @NumeroPedido)

            Select @Desconto = (Select sum(desconto) 
                                      From NFItens Where NUMEROPED = @NumeroPedido)
                                      
            Select @TotalICMS = (Select sum(ValorICMS) 
                                      From NFItens Where NUMEROPED = @NumeroPedido)

            Select @TotalNF = (Select FRETECOBR 
                                      From NFCapa Where NUMEROPED = @NumeroPedido)

            Select @TotalNF = (@TotalNF + (Select sum(vltotitem - desconto)
                                     From NFITENS Where NUMEROPED = @NumeroPedido))


            Select @BaseICMS = (Select sum(BaseICMS) 
                                      From NFItens Where NUMEROPED = @NumeroPedido)

            Update NFCapa set VLRMERCADORIA=@ValorMercadoria,desconto=@Desconto,
                              TOTALNOTA=@TotalNF,BASEICMS=@BaseICMS,VLRICMS=@TotalICMS where NUMEROPED = @NumeroPedido  

End

GO
/****** Object:  StoredProcedure [dbo].[SP_Totaliza_Capa_Nota_Fiscal_Loja]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE         PROCEDURE [dbo].[SP_Totaliza_Capa_Nota_Fiscal_Loja]
	@NumeroPedido		varchar(6)

AS  
	DECLARE	@ValorMercadoria	Float,
			@TotalICMS			Float,
			@TotalNF	    	Float,
			@BaseICMS       	Float,
            @Desconto       	Float,
            @SubTotal       	Float
                       
BEGIN

			Update nfitens set VLTOTITEM = round((VLUNIT * QTDE),2) where numeroped = @NumeroPedido
			Update nfitens set VLUNIT2 =   round((VLUNIT * QTDE) - DESCONTO,2) where numeroped = @NumeroPedido
            Update NFITENS SET VALORICMS=(((vltotitem - desconto) * PR_ICMPDV) / 100),VLUNIT2 = (vltotitem - DESCONTO)
                   From NFITENS,PRODUTOLOJA
                   Where PR_Referencia=Referencia and NUMEROPED=@NumeroPedido and TIPONOTA <> 'T'

            Select @ValorMercadoria = round((Select sum(vltotitem) 
                                      From NFItens Where NUMEROPED = @NumeroPedido),2)

            Select @Desconto = round((Select sum(desconto) 
                                      From NFItens Where NUMEROPED = @NumeroPedido),2)
                                      
            Select @TotalICMS = round((Select sum(ValorICMS) 
                                      From NFItens Where NUMEROPED = @NumeroPedido),2)

            Select @TotalNF = round((Select FRETECOBR 
                                      From NFCapa Where NUMEROPED = @NumeroPedido),2)

            Select @TotalNF = round((@TotalNF + (Select sum(vltotitem - desconto)
                                     From NFITENS Where NUMEROPED = @NumeroPedido)),2)


            Select @BaseICMS = round((Select sum(BaseICMS) 
                                      From NFItens Where NUMEROPED = @NumeroPedido),2)

            Update NFCapa set VLRMERCADORIA=@ValorMercadoria,desconto=@Desconto,SUBTOTAL=@ValorMercadoria-@Desconto,
                              TOTALNOTA=@TotalNF,BASEICMS=@BaseICMS,VLRICMS=@TotalICMS where NUMEROPED = @NumeroPedido  and TIPONOTA <> 'T'

End

GO
/****** Object:  StoredProcedure [dbo].[SP_VDA_Atualiza_Promocao_Loja]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE  PROCEDURE [dbo].[SP_VDA_Atualiza_Promocao_Loja]


As 

Begin
        Declare @SQL                    char    (8000),
	        @Situacao               char    (03),
                @Conso                  char    (07),
                @Data                   char    (12),   
	    	@Loja_C                 char    (05),			
                @Mes                    char    (04),
                @Tipo                   char    (03),
                @NumeroPedido           int,
                @ContaReg               int,
                @NomeTabela             Varchar(50),
                @EstoqueMes             char(50),
                @ProdutoMes             char(50),
                @VendaMes               float,
                @vendaAnterior          float,
                @TicketMedioMes         float,
                @TicketMedioAnterior    float, 
                @TCursor_Loja           char(05),
                @TCursor_Ano            char(04),
                @TCursor_Mes            char(02),
                @TCursor_Visa            float,
                @TCursor_Mastercard      float,
                @TCursor_Amex            float,
                @TCursor_Dinners         float,
                @TCursor_OutroCartao     float, 
                @TCursor_Faturada        float,
                @TCursor_Financiada      float, 
                @TCursor_Dinheiro        float,
                @TCursor_TEF             float,
                @TCursor_Cheque          float,
                @TCursor_ChequePre       float,
                @TCursor_Avr              float 
   
  
   IF EXISTS (select * from sys.objects where name = 'TempPromocaoLoja')
      Drop Table TempPromocaoLoja

   
   Create	Table TempPromocaoLoja	(
		          Temp_Loja	     	    Char(5)	,
		          Temp_Referencia       Char(07), 
                  Temp_NroPromocao      numeric,
                  Temp_PrecoPromocao 	Float)
		        
		        
            Insert Into TempPromocaoLoja(Temp_Loja,Temp_Referencia,Temp_NroPromocao)
            select  PM_Controlelojas,pm_referencia,max(pm_numeropromocao) as Promocao from  promocaoaux 
                    group by PM_Controlelojas,pm_referencia order by pm_referencia

            Update TempPromocaoLoja set Temp_PrecoPromocao=PM_PromocaoAvista from TempPromocaoLoja,promocaoaux
                   where Temp_Loja=PM_ControleLojas and Temp_NroPromocao=PM_NumeroPromocao and 
                         Temp_Referencia=PM_Referencia
            
           
            Update produtoLoja set PR_Classe='P',PR_IndicePreco=3,PR_PrecoVenda1=Temp_PrecoPromocao
                   from produtoLoja as l,TempPromocaoLoja 
                   where l.PR_Referencia=Temp_Referencia 

end



/*
*/

GO
/****** Object:  StoredProcedure [dbo].[SP_VDA_Calcula_ICMS_Pobreza]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO


create PROCEDURE [dbo].[SP_VDA_Calcula_ICMS_Pobreza]
		@NF		 varchar(5)

AS
	


Begin

--select top 10 * from nfitens WHERE NF = 728 AND SERIE = 'NE'

	UPDATE nfitens
	SET
	aliqICMSFECP = UF_FECP,
	aliqICMSDest = UF_ICMSInterno , 
	aliqICMSInter = UF_ICMSInterEstadual , 
	ICMSInterpart = UF_Participacao, 
	valICMSRemet = (((((VLUNIT * QTDE) - I.DESCONTO) * UF_ICMSDifal)/100) * (100 - UF_Participacao))/100 , 
	valICMSDest = ((((((VLUNIT * QTDE) - I.DESCONTO) * UF_ICMSDifal)/100) * (UF_Participacao))/100)  , 
	valorICMSFECP = ((((VLUNIT * QTDE) - I.DESCONTO) * UF_FECP) / 100) 
	FROM nfitens AS I,nfcapa as c,fin_estado,PRODUTOloja, FIN_Cliente
	WHERE CE_Estado <> 'SP'
	and c.NF = @NF
	AND c.nf = i.nf 
	AND c.SERIE = i.SERIE 
	AND c.LOJAORIGEM = i.lojaorigem 
	AND PR_Referencia = i.Referencia 
	AND CE_Estado = UF_Estado
	and c.CLIENTE = CE_CodigoCliente
	and c.serie = 'NE'
	AND i.serie = 'NE'

	update nfcapa 
	set 
	valICMSRemet = (select sum(valICMSRemet) from NfItens as i where c.nf = i.nf aND c.Serie = i.Serie AND c.LojaOrigem = i.LojaOrigem ),
	valICMSDest = (select sum(valICMSDest) from NfItens as i where c.nf = i.nf aND c.Serie = i.Serie AND c.LojaOrigem = i.LojaOrigem ),
	valorICMSFECP = (select sum(valorICMSFECP) from NfItens as i where c.nf = i.nf aND c.Serie = i.Serie AND c.LojaOrigem = i.LojaOrigem ) 
	from nfcapa as c, FIN_Cliente
	where CE_Estado <> 'SP'
	and CE_CodigoCliente = c.cliente
	and c.NF = @NF
	and c.serie ='NE'
	and serie = 'NE'


	
END




GO
/****** Object:  StoredProcedure [dbo].[SP_VDA_Consulta_Itens_Venda]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
/*
Procedure que gera Titulos a Receber
a partir da tabela NFCapa

*/

CREATE                Procedure [dbo].[SP_VDA_Consulta_Itens_Venda] 


        		   
                            @Where                 char(2000),
			    @DataInicial           char(10),
			    @DataFinal 		   char(10)
As                          
			    
  Declare		    @SQL            	   Char(4000),
			    @Loja	           Char(5)
			    	  

	
          Begin
		select @Loja = (select CTS_Loja from ControleSistema)
		

	     Select @SQL = 'Select vi_quantidade, vi_referencia, vi_lojaOrigem,VI_NotaFiscal,VI_Serie,VC_Codigovendedor,VE_Nome,PR_Comprador,
	            PR_Descricao,PR_PrecoVenda1,PR_PrecoCusto1,CO_Nome,FO_NomeFantasia,ES_Estoque,ES_Venda
	            From [svcentralfer].[dmac].[dbo].CapaNFVenda,[svcentralfer].[dmac].[dbo].ItemNFVenda,[svcentralfer].[dmac].[dbo].Produto,[svcentralfer].[dmac].[dbo].Fornecedor,[svcentralfer].[dmac].[dbo].Vendedor,[svcentralfer].[dmac].[dbo].Comprador,[svcentralfer].[dmac].[dbo].Estoque
	            Where VC_dataemissao between ' + '''' + @DataInicial +''''+ ' and ' + '''' + @DataFinal + ''' and VC_TipoNota= ' + '''V''' + ' and VC_NotaFiscal=VI_NotaFiscal
                    and VC_Serie=VI_Serie and VC_Lojaorigem=VI_Lojaorigem and VC_CodigoVendedor = VE_CodigoVendedor and PR_CodigoFornecedor = FO_CodigoFornecedor 
                    and PR_Comprador = CO_CodigoComprador And VI_Referencia = PR_Referencia and VI_Referencia=ES_Referencia and ES_Loja = ' + '''' + @Loja + '''' + @Where  
		    
          
		 exec (@SQL)
		end
	
/*
exec SP_VDA_Consulta_Itens_Venda ' and fo_codigofornecedor = 001 and co_codigocomprador = 03 
and vc_codigovendedor = 2 and ve_codigovendedor = vc_codigovendedor','2012-10-14', '2012-10-15'


Select vi_lojaOrigem,VI_NotaFiscal,VI_Serie,VC_Codigovendedor,VE_Nome,PR_Comprador,
	PR_Descricao,PR_PrecoVenda1,PR_PrecoCusto1,CO_Nome,FO_RazaoSocial,ES_Estoque,ES_Venda
	From CapaNFVenda,ItemNFVenda,Produto,Fornecedor,Vendedor,Comprador,Estoque
	Where VC_dataemissao between '2012-10-14' and '2012-10-15' and VC_TipoNota= 'V' and VC_NotaFiscal=VI_NotaFiscal
        and VC_Serie=VI_Serie and VC_Lojaorigem=VI_Lojaorigem And VI_Referencia = PR_Referencia 
	and VI_Referencia=ES_Referencia  and fo_codigofornecedor = 001 and co_codigocomprador = 03 
	and vc_codigovendedor = 2 and ve_codigovendedor = vc_codigovendedor and ve_loja = vc_lojaorigem
	group by vi_lojaOrigem,VI_NotaFiscal,VI_Serie,VC_Codigovendedor,VE_Nome,PR_Comprador,
	PR_Descricao,PR_PrecoVenda1,PR_PrecoCusto1,CO_Nome,FO_RazaoSocial,ES_Estoque,ES_Venda

 exec SP_VDA_Consulta_Itens_Venda '' 
select * from capanfvenda where vc_dataemissao = '1999-01-02'
update capanfvenda set vc_serie = '00'
update itemnfvenda set vi_dataemissao = '2012/10/15' where vi_notafiscal = 61
select * from itemnfvenda where vi_referencia = '0010250'
select * from vendedor
select * from loja
select * from comprador
select * from fornecedor
select * from controlesistema
select pr_comprador, * from produto
insert into vendedor (Ve_Codigovendedor, Ve_Nome, Ve_RegistroFuncionario, Ve_Loja, Ve_Comissionado, Ve_situacao)
values ( 0, 'Jose', 443, 134, 'P', 'D')

update capanfvenda set Vc_notafiscal = 61


			    @DataInicial	   datetime,
                            @DataFinal	           datetime,
                            @Fornecedor            int,
                            @Comprador             int,
                            @Loja                  char(05),
                            @Vendedor              int,
                            @Processo


*/

GO
/****** Object:  StoredProcedure [dbo].[SP_VDA_Cria_Cupom]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO



CREATE PROCEDURE [dbo].[SP_VDA_Cria_Cupom]
	@pedido	 numeric

AS
	declare

		@espacoP as char(4),
		@CRT as char(2),

		--IDE
		@ide_NSU numeric,
		@ide_numerocaixa char(3),

		--DACFE
		@dacfe_impressora varchar(100),
		@dacfe_retornasp char(1),
		@dacfe_imprTef char(1),
		@dacfe_tipoImpr char(1),
		@dacfe_IMPRSOTEF char(3),

		--EMIT
		@emit_cnpj char(14),
		@emit_IE char(12),
		@emit_INDRATISSQN char(1),

		--DEST
		@dest_cgc char(14),

		--ICMSTOT
		@icmstot_VDESCSUBTOT numeric(10,2),
		@icmstot_VACRESSUBTOT numeric(10,2),
		@icmstot_VTOTTRIB numeric(10,2),

		--INFADIC
		@infadic_infclp varchar(5000),
		@infadic_infclpTEMP varchar(1000),

		--PAG
		@pag_TPAG char(2),
		@pag_VPAG numeric(10,2),
		@pag_CADMC char(3),

		--PROD
		@prod_cprod varchar(60),
		@prod_xprod varchar(120),
		@prod_ncm varchar(8),
		@prod_cfop varchar(4),
		@prod_ucom varchar(6),
		@prod_qcom float,
		@prod_vuncom numeric(10,2),
		@prod_indregra char(1),
		@prod_vDesc numeric(10,2),
		@proc_xcampodet varchar(20),
		@prod_XTEXTODET varchar(60),

		--ICMS
		@icms_orig char(1),
		@icms_cst char(2),
		@icms_picms numeric(10,2),
		@icms_csosn char(3),

		--PIS
		@pis_cst char(2),
		@pis_vbc numeric(10,2),
		@pis_ppis numeric(10,2),

		--COFIN
		@cofins_cst char(2),
		@cofins_vbc numeric(10,2),
		@cofins_pcofins numeric(10,2)



Begin


--insert into sat_capa (cp_nsu, cp_caixa, cp_impressora, cp_tipoImpr, cp_cnpjLoja, cp_inscricaoLoja, cp_cnpjCliente, cp_cpfCliente, cp_NomeCliente)
--	 select           @nsu,   @caixa, @impressora, 2, lo_cgc, LO_InscricaoEstadual,(case when len(CGCCLI ) = 14 then CGCCLI else '' end), (case when len(CGCCLI)=11 then CGCCLI else '' end), (case when len(NOMCLI) > 2 then NOMCLI else '' end) 
--	   from loja,nfcapa 
--	  where lo_loja= lojaorigem 

--insert into sat_itens (ci_nsu, ci_referencia, ci_decricao, ci_ncm, ci_cfop, ci_ucom, ci_qcom, ci_vuncom, ci_vdesconto, ci_origem, ci_cst, ci_pIcms, ci_ppis, ci_pcofins)
--	 select           @nsu, REFERENCIA, pr_Descricao, pr_ClasseFiscal, CFOP, pr_Unidade, QTDE, VLUNIT, DESCONTO, SUBSTRING(pr_cst,1,1), icms, VALORICMS,  PisCofins , 0
--	  from nfitens,produto
--	 where pr_referencia = referencia

	--(case when len(CGCCLI ) = 14 then CGCCLI else '' end), (case when len(CGCCLI)=11 then CGCCLI else '' end), (case when len(NOMCLI) > 2 then NOMCLI else '' end) 

	delete sat_nf  where snf_pedido = @pedido
	select @espacoP = '    '

			--@icmstot_VDESCSUBTOT float,
		--@icmstot_VACRESSUBTOT float,
		--@icmstot_VTOTTRIB float,

	select 
	@ide_NSU = CONCAT(right(year(dataemi),2),REPLICATE('0', 2 - LEN(month(dataemi))) + RTrim(month(dataemi)), NUMEROPED),
	@ide_numerocaixa = REPLICATE('0', 3 - LEN(nroCaixa)) + RTrim(nroCaixa), 
	@emit_cnpj = LO_CGC,  
	@emit_IE = LO_InscricaoEstadual,
	@dest_cgc = CPFNFP,
	@icmstot_VDESCSUBTOT = desconto,
	@icmstot_VACRESSUBTOT = 0,
	@icmstot_VTOTTRIB = ((TOTALNOTA * 26.25) / 100),
	@CRT = CTS_TipoEmpresa
	from nfcapa, loja, ControleSistema
	where 
	lo_loja= lojaorigem 
	and numeroped = @pedido

	-- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- 

	insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values('[IDE]','','',@pedido) 
	insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values(@espacoP + 'NSU','=',@ide_NSU,@pedido)		
	insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values(@espacoP + 'NUMEROCAIXA','=',@ide_numerocaixa,@pedido)		

	-- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- 

	SELECT 
		@dacfe_impressora = cs_dacfe,
		@dacfe_retornasp = '3',
		@dacfe_imprTef = 'IMPRTEF',
		@dacfe_tipoImpr = '7',
		@dacfe_IMPRSOTEF = 'não'
	FROM CONTROLESERIE where cs_nroCaixa = @ide_numerocaixa
	
	insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values('','','',@pedido) 
	insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values('[DACFE]','','',@pedido)	
	insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values(@espacoP + 'IMPRESSORA','=',@dacfe_impressora,@pedido)					
	insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values(@espacoP + 'IMPRESSORA','=',@dacfe_impressora,@pedido)					
	insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values(@espacoP + 'RETORNASP','=',@dacfe_retornasp,@pedido)					
	insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values(@espacoP + 'IMPRTEF','','',@pedido)					
	insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values(@espacoP + 'TIPOIMPR','=',@dacfe_tipoImpr,@pedido)					
	insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values(@espacoP + 'IMPRSOTEF','=',@dacfe_IMPRSOTEF,@pedido)					

	-- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- 

	--DESENV
	--select @emit_cnpj = '61099008000141'
	--select @emit_IE = '111111111111'
	select @emit_INDRATISSQN = 'N'

	insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values('','','',@pedido)							
	insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values('[EMIT]','','',@pedido)	
	insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values(@espacoP + 'CNPJ','=',@emit_cnpj,@pedido)					
	insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values(@espacoP + 'IE','=',@emit_IE,@pedido)		
	insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values(@espacoP + 'INDRATISSQN','=',@emit_INDRATISSQN,@pedido)		
				

	-- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- 

	if len(@dest_cgc) > 1
	begin
		insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values('','','',@pedido)							
		insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values('[DEST]','','',@pedido)							

		if len(@dest_cgc) >= 14 						
			insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values(@espacoP + 'CNPJ','=',@dest_cgc,@pedido)					

		if len(@dest_cgc) < 14
			insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values(@espacoP + 'CPF','=',@dest_cgc,@pedido)					
	end 

	-- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- 

	insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values('','','',@pedido) 
	insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values('[ICMSTOT]','','',@pedido) 
	insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values(@espacoP + 'VDESCSUBTOT','=',@icmstot_VDESCSUBTOT,@pedido)		
	insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values(@espacoP + 'VACRESSUBTOT','=',@icmstot_VACRESSUBTOT,@pedido)		
	insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values(@espacoP + 'VTOTTRIB','=',@icmstot_VTOTTRIB,@pedido)		

	-- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- 

	select @infadic_infclp = ''
	select @infadic_infclpTEMP = ''

	Create	Table #TEMPInfAdic	(  
	infadic_infclp varchar(5000))
	Insert Into #TEMPInfAdic(infadic_infclp)
	Select CNF_Carimbo from CarimboNotaFiscal where CNF_NumeroPed = @pedido
	Declare curInfAdic Insensitive Cursor For
	Select 	infadic_infclp from #TEMPInfAdic
	Open curInfAdic
	Fetch Next From curInfAdic Into @infadic_infclpTEMP                                  
	While @@Fetch_Status=0
	Begin

		--if (select count(CNF_Carimbo) from CarimboNotaFiscal where CNF_NumeroPed = @pedido) > 0
		select @infadic_infclp = @infadic_infclpTEMP + ' - '


	Fetch Next From curInfAdic Into 
	@infadic_infclpTEMP      

	End

	Close curInfAdic
	Deallocate curInfAdic
	
	insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values('','','',@pedido) 
	insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values('[INFADIC]','','',@pedido)	
	insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values(@espacoP + 'INFCPL','=',@infadic_infclp,@pedido)		


	-- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- 

	Create	Table #TEMPPag	(  
	pag_TPAG char(2),
	pag_VPAG numeric(10,2),
	pag_CADMC char(3))
	Insert Into #TEMPPag(pag_TPAG, pag_VPAG, pag_CADMC)
		Select MO_TipoPag, MC_Valor, MC_Agencia
		From MovimentoCaixa, Modalidade where MC_Pedido = @pedido AND 	MO_Grupo = MC_Grupo AND 	MO_Grupo < 11000
	Declare curPag Insensitive Cursor For
	Select 	pag_TPAG, pag_VPAG, pag_CADMC from #TEMPPag
	Open curPag
	Fetch Next From curPag Into @pag_TPAG, @pag_VPAG, @pag_CADMC                                  
	While @@Fetch_Status=0
	Begin

		insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values('','','',@pedido) 
		insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values('[PAG]','','',@pedido)	
		insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values(@espacoP + 'TPAG','=',@pag_TPAG,@pedido)	
		insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values(@espacoP + 'VPAG','=',@pag_VPAG,@pedido)	
		insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values(@espacoP + 'CADMC','=',@pag_CADMC,@pedido)		

	Fetch Next From curPag Into 
	@pag_TPAG, @pag_VPAG, @pag_CADMC      

	End

	Close curPag
	Deallocate curPag
	
	--select * from Modalidade

	--Tipo de TPAG
	--01 – Dinheiro;
	--02 – Cheque.
	--03 – Cartão de Crédito;
	--04 – Cartão de Débito;
	--05 – Crédito Loja;
	--10 – Vale Alimentação;
	--11 – Vale Refeição;
	--12 – Vale Presente;
	--13 – Vale Combustível;
	--99 – Outros.

	-- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- 

	Create	Table #TempItens	(  
	prod_cprod varchar(60),	prod_xprod varchar(120), prod_ncm varchar(8), prod_cfop varchar(4), prod_ucom varchar(6), prod_qcom float, 
	prod_vuncom numeric(10,2), prod_indregra char(1), icms_cst char(2), icms_orig CHAR(1), icms_picms float, 
	pis_vbc numeric(10,2), pis_cst char(2), pis_ppis numeric(10,2), 
	cofins_vbc numeric(10,2), cofins_cst char(2), cofins_pcofins numeric(10,2), prod_vDesc numeric(10,2), proc_xcampodet varchar(20), prod_XTEXTODET varchar(60))

	Insert Into #TempItens(prod_cprod, prod_xprod, prod_ncm, prod_cfop, prod_ucom, prod_qcom, prod_vuncom, prod_indregra, icms_cst, 
	icms_orig, icms_picms, 
	pis_vbc, pis_cst, pis_ppis, 
	cofins_vbc,cofins_cst,cofins_pcofins, prod_vDesc, proc_xcampodet, prod_XTEXTODET)

	Select pr_referencia, pr_descricao, pr_classeFiscal, cfop, 'PC', qtde, vlunit, 'A',  REPLICATE('0', 2 - LEN(CSTICMS)) + RTrim(CSTICMS),  '0',  ICMSAplicado, 
	VLUNIT2,'01', 1.65, 
	VLUNIT2,'01',7.60,
	DESCONTO, 'Cod. CEST',PR_CEST
	
	From produtoloja,nfitens where  pr_referencia = referencia and numeroped = @pedido
	Declare CurItens Insensitive Cursor For
	Select 	prod_cprod,prod_xprod,prod_ncm,
	prod_cfop,prod_ucom,prod_qcom,
	prod_vuncom,prod_indregra,icms_cst,
	icms_orig,icms_picms,pis_vbc, pis_cst, pis_ppis, 
	cofins_vbc,cofins_cst,cofins_pcofins, prod_vDesc, proc_xcampodet, prod_XTEXTODET from #TempItens
	Open CurItens
	Fetch Next From CurItens Into @prod_cprod, 	@prod_xprod, @prod_ncm, 
	@prod_cfop, @prod_ucom, @prod_qcom, 
	@prod_vuncom, @prod_indregra, @icms_cst, 
	@icms_orig, @icms_picms,	
	@pis_vbc, @pis_cst, @pis_ppis, 
	@cofins_vbc, @cofins_cst , @cofins_pcofins , @prod_vDesc, @proc_xcampodet, @prod_XTEXTODET                   
	While @@Fetch_Status=0
	Begin
	
		insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values('','','',@pedido) 
		insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values('[DET]','','',@pedido)	
		insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values('','','',@pedido)							
		insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values('[PROD]','','',@pedido)							
		insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values(@espacoP + 'CPROD','=',@prod_cprod,@pedido)					
		insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values(@espacoP + 'XPROD','=',@prod_xprod,@pedido)					
		insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values(@espacoP + 'NCM','=',@prod_ncm,@pedido)					
		insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values(@espacoP + 'CFOP','=',@prod_cfop,@pedido)					
		insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values(@espacoP + 'UCOM','=',@prod_ucom,@pedido)					
		insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values(@espacoP + 'QCOM','=',@prod_qcom,@pedido)					
		insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values(@espacoP + 'VUNCOM','=',@prod_vuncom,@pedido)					
		insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values(@espacoP + 'INDREGRA','=',@prod_indregra,@pedido)	
		insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values(@espacoP + 'VDESC','=',@prod_vDesc,@pedido)	
		
		
		if @prod_XTEXTODET <> ''
		begin
			insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values('','','',@pedido)							
			insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values('[OBSFISCODET]','','',@pedido)							
			insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values(@espacoP + 'XCAMPODET','=',@proc_xcampodet,@pedido)					
			insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values(@espacoP + 'XTEXTODET','=',@prod_XTEXTODET,@pedido)					
		end

		insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values('','','',@pedido)							
		insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values('[IMPOSTO]','','',@pedido)							

		IF @CRT <> 'SN'
		begin

			IF @icms_cst = '40' or @icms_cst = '41' or @icms_cst = '60'
			begin
				insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values('','','',@pedido)							
				insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values('[ICMS40]','','',@pedido)							
				insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values(@espacoP + 'ORIG','=',@icms_orig,@pedido)
				insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values(@espacoP + 'CST','=',@icms_cst,@pedido)
			end

			IF @icms_cst = '00' or @icms_cst = '20' or @icms_cst = '90'
			begin
				insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values('','','',@pedido)							
				insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values('[ICMS00]','','',@pedido)							
				insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values(@espacoP + 'ORIG','=',@icms_orig,@pedido)
				insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values(@espacoP + 'CST','=',@icms_cst,@pedido)
				insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values(@espacoP + 'PICMS','=',@icms_picms,@pedido)
			end

		end 
		ELSE
		BEGIN
			
			insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values('','','',@pedido)							
			insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values('[ICMSSN102]','','',@pedido)							
			insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values(@espacoP + 'ORIG','=',@icms_orig,@pedido)
			insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values(@espacoP + 'CSOSN','=',@icms_csosn,@pedido)

		end

		print 'PISALIQ'
		--select @pis_vbc = (@pis_vbc * @pis_ppis) / 100
		insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values('','','',@pedido)							
		insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values('[PISALIQ]','','',@pedido)							
		insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values(@espacoP + 'CST','=',@pis_cst,@pedido)
		insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values(@espacoP + 'VBC','=',@pis_vbc,@pedido)
		insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values(@espacoP + 'PPIS','=',@pis_ppis,@pedido)

		print 'COFINSALIQ'
		--select @cofins_vbc = (@cofins_vbc * @cofins_pcofins) / 100
		insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values('','','',@pedido)							
		insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values('[COFINSALIQ]','','',@pedido)							
		insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values(@espacoP + 'CST','=',@cofins_cst,@pedido)
		insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values(@espacoP + 'VBC','=',@cofins_vbc,@pedido)
		insert into SAT_NF(snf_Descricao,snf_Sinal,snf_Dados,snf_pedido) values(@espacoP + 'pcofins','=',@cofins_pcofins,@pedido)


	Fetch Next From CurItens Into 
	@prod_cprod, @prod_xprod, @prod_ncm, @prod_cfop, @prod_ucom, @prod_qcom, @prod_vuncom, @prod_indregra, 
	@icms_cst,@icms_orig,@icms_picms,@pis_vbc, @pis_cst, @pis_ppis, @cofins_vbc, @cofins_cst , @cofins_pcofins , @prod_vDesc , @proc_xcampodet, @prod_XTEXTODET

	End

	Close CurItens
	Deallocate CurItens

	-- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- 

	--select * from nfitens
	
END











GO
/****** Object:  StoredProcedure [dbo].[SP_VDA_Cria_NFe]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

CREATE PROCEDURE [dbo].[SP_VDA_Cria_NFe]

	@Loja		Char(5),
	@NF		    Numeric,
	@Serie		Char(2),
    @Carimbo    varchar(MAX)

AS

	DECLARE	@SQL        	char(4000),
			@CondPagto		Char(2),
			@CondPagtoNF	Char(2),
			@Parcelas       Char(2),
			@NroNF_NFe		Char(10),
			@Referencia		Char(7),
			@UFCliente		Char(2),
			@IDDEST			char(1),
			@finNFe			char(1),
            @CEPCliente     Char(8),
            @NomeServidor   char(40),
            @Cliente        char(7),
			@ClienteT       char(7),
			@IE				char(13),
			@Pessoa         char(1),
			@TipoEmissao    Char(1),
			@QtdeVolume     float,
			@TotalFrete     numeric,
			@PercFrete		float,
			@DiferencaFrete float,
			@Item			numeric,
			@tiponota		char(4),
			@Operacao		char(60),
			@cfop			numeric(18,0),
            @Hora           char(12),
            @Chave          char(8),
            @UFLoja         char(2),
			@EntradaSaida   char(1),
			@CRT			char(1)


                 
BEGIN

	exec sp_delete_nfe @loja, @nf, @Serie
	delete NFE_NFLojas 
	 where NFL_NroNFE = @nf
	
	Select @tiponota = (Select top 1 tiponota 
	                      from nfcapa 
	                     where LojaOrigem = @Loja 
	                       And NF = @NF 
	                       And Serie = @Serie)

	
	-- -- ACERTOS NFCAPA -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- 
	
	update NfItens set 
		   VALORICMS = round(((BASEICMS * ICMSAplicado) / 100),2) 
	 where nf = @nf 
	   and serie = @Serie
	   and @tiponota <> 'S'
	
	update NfCapa set 
		   vlrICMS = round((select SUM(VALORICMS) as total 
						      from NfItens 
						     where nf = @nf 
						       and serie = @Serie),2) 
	 where nf = @nf 
	   and serie = @Serie
	   and @tiponota <> 'S'       

	-- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- --


	if (SELECT top 1 cts_tipoempresa FROM ControleSistema) = 'SN'
		select @CRT = '1'
	
	if (SELECT top 1 cts_tipoempresa FROM ControleSistema) <> 'SN' 
		select @CRT = '3'

	print 'OK 1'
	Select @CondPagtoNF = (Select TOP 1 CondPag 
	                         from NFcapa 
	                        where LojaOrigem = @Loja 
	                          And NF = @NF 
	                          And Serie = @Serie)

	Select @Parcelas = (Select TOP 1 CP_parcelas 
	                      from CondicaoPagamento 
	                     Where CP_Codigo = @CondPagtoNF)
	                     
	SELECT @cfop = (Select TOP 1 CODOPER 
	                        from NFcapa 
	                       where LojaOrigem = @Loja 
	                         And NF = @NF 
	                         And Serie = @Serie)
	
	--Update ControleSup set CS_NumeroNFe = (CS_NumeroNFe + 1)
	
	Select @NroNF_NFe = @NF
	
	print 'OK 2'
	Select @UFCliente = (select ce_Estado 
	                       from NFCapa,FIN_cliente 
	                      where ce_codigocliente = cliente 
	                        and lojaorigem = @Loja 
	                        and NF = @Nf 
	                        and Serie = @serie)

	Select @Pessoa = (Select CE_TipoPessoa 
	                    from NFCapa,FIN_cliente 
	                   where ce_codigocliente = cliente 
	                     and lojaorigem = @Loja 
	                     and NF = @Nf 
	                     and Serie = @serie)
               

    Select @CEPCliente = (Select replicate('0',8 - len(CE_Cep)) + CE_Cep 
                            from NFCapa,FIN_cliente 
                           where ce_codigocliente = cliente 
                             and lojaorigem = @Loja 
                             and NF = @Nf 
                             and Serie = @serie)

	print 'OK 3'
    Select @QtdeVolume = (Select sum(qtde) 
                            from nfItens 
                           where LojaOrigem = @Loja 
                             And NF = @NF 
                             And Serie = @Serie)

	--select @EntradaSaida = (Select top 1 substring(codoper,1,1) from nfcapa where LojaOrigem = @Loja And NF = @NF And Serie = @Serie)
	
	print 'OK 3-1'	
	Select @Cliente = (Select top 1 cliente 
	                     from nfcapa 
	                    where LojaOrigem = @Loja 
	                      And NF = @NF 
	                      And Serie = @Serie 
	                      and tiponota <> 'T')
	
	print 'OK 3-2'
	Select @ClienteT = (Select top 1 lojat 
	                      from nfcapa 
	                     where LojaOrigem = @Loja 
	                       And NF = @NF 
	                       And Serie = @Serie 
	                       and TIPONOTA = 'T')
	
	print 'OK 3-3'	
    Select @TotalFrete = (Select fretecobr 
                            from NFCapa 
                           where lojaorigem = @Loja 
                             and NF = @Nf 
                             and Serie = @serie)
    
    print 'OK 3-4'    
    Select @PercFrete = (Select ((fretecobr * 100)/ vlrmercadoria) 
                           from NFCapa 
                          where lojaorigem = @Loja 
                            and NF = @Nf 
                            and Serie = @serie)
	print 'OK 3-5'	
	select @DiferencaFrete = (select ( @TotalFrete - (sum(((vltotitem - desconto) * @PercFrete) / 100))) 
	                            from NFitens
		                       where lojaorigem = @Loja 
		                         and NF = @Nf 
		                         and Serie = @serie)
	print 'OK 3-6'	                         
	Select @Item = (select top 1 Item 
	                  from nfitens 
	                 where lojaorigem = @Loja 
	                   and NF = @Nf 
	                   and Serie = @serie 
	                 order by Item)
    
    print 'OK 3-7'
    Select @UFLoja = (select distinct substring(convert(nvarchar(9),lo_codigoMunicipio),1,2)
                        from Loja,nfcapa 
                       where lojaorigem = @Loja 
                         and NF = @Nf 
                         and Serie = @serie 
                         and lojaorigem = lo_loja)
                         
    Select @Hora = CONVERT(varchar,GETDATE(),114)
    Select @Hora = replace(@Hora,':','')
    Select @Chave = substring(@hora,5,2) + substring(@hora,3,2) + substring(@hora,1,2) + substring(@hora,8,2)
      

-- SELECT Name + REPLICATE('*', 20 - LEN(Name)) FROM Employee
--	update nfcapa set fonecli = replace(fonecli,'-','') where LojaOrigem = @Loja And NF = @NF And Serie = @Serie
--      update nfcapa set fonecli = replace(fonecli,' ','') where LojaOrigem = @Loja And NF = @NF And Serie = @Serie
--	update nfcapa set fonecli = replace(fonecli,'.','') where LojaOrigem = @Loja And NF = @NF And Serie = @Serie
--	update nfcapa set fonecli = replace(fonecli,'(','') where LojaOrigem = @Loja And NF = @NF And Serie = @Serie
--	update nfcapa set fonecli = replace(fonecli,')','') where LojaOrigem = @Loja And NF = @NF And Serie = @Serie
--      update nfcapa set cepcli = ' ' where LojaOrigem = @Loja And NF = @NF And Serie = @Serie And len(cepcli)<7
	print 'OK 4'
	
	Update nfitens set 
	       CSTICMS = 60 
	  from nfitens, produtoloja 
	 where referencia = pr_referencia 
	   and pr_substituicaoTributaria = 'S' 
	   and LojaOrigem = @Loja 
	   AND Serie = @Serie 
	   AND NF = @NF 
	   and @tiponota <> 'S'
	
	print ('Update nfitens set CSTICMS = 60')

	print 'OK 5'
	
	Update nfitens set 
	       CSTICMS = 20 
	  from nfitens, produtoloja 
	 where referencia = pr_referencia 
	   and pr_substituicaoTributaria = 'N' 
	   and Pr_codigoreducaoicms > 0 
	   and LojaOrigem = @Loja 
	   AND Serie = @Serie 
	   AND NF = @NF
	   and @tiponota <> 'S'
	print ('Update nfitens set CSTICMS = 20')

	print 'OK 6'
	
	Update nfitens set 
	       CSTICMS = 00 
	  from nfitens, produtoloja 
	 where referencia = pr_referencia 
	   and pr_substituicaoTributaria = 'N' 
	   and Pr_codigoreducaoicms = 0 
	   and LojaOrigem = @Loja 
	   AND Serie = @Serie 
	   AND NF = @NF
	   and @tiponota <> 'S'

	   	Update nfitens set 
	       CSTICMS = 02
	  from nfitens
	 where LojaOrigem = @Loja 
	   AND Serie = @Serie 
	   AND NF = @NF
	   and @CRT = 1

	   	   	Update nfitens set 
	       CSTICMS = 02
	  from nfitens
	 where LojaOrigem = @Loja 
	   AND Serie = @Serie 
	   AND NF = @NF
	   and NUMEROPED = 1520
	   
	select @IDDEST = '1'

	if @Tiponota NOT IN ('E') 
		BEGIN

	IF @UFCliente = 'SP'
	   BEGIN
			IF @pessoa = 'F' or @pessoa = 'U' or @Pessoa = 'J' or @pessoa = 'O' and @Tiponota NOT IN ('S','E') 
				Begin
					
					Update nfitens set 
					       CFOP = 5102 
					  from nfitens, produtoloja 
					  where referencia = pr_referencia 
			           and pr_substituicaoTributaria = 'N' 
			           and LojaOrigem = @Loja 
			           AND Serie = @Serie 
			           AND NF = @NF
			           and @tiponota <> 'S'
			           print ('Update nfitens set CFOP = 5102')
			           
				end
			IF @pessoa = 'F' or @pessoa = 'U' or @pessoa = 'J' or @pessoa = 'O' and @Tiponota NOT IN ('S','E') 
				Begin
					
					Update nfitens set 
						   CFOP = 5405 
					  from nfitens, produtoloja 
					 where referencia = pr_referencia 
					   and pr_substituicaoTributaria = 'S' 
					   and LojaOrigem = @Loja 
					   AND Serie = @Serie 
					   AND NF = @NF
					   and @tiponota <> 'S'
					print ('Update nfitens set CFOP = 5405')
					
				end
		  --print @tiponot
		END

	IF @UFCliente <> 'SP'
		BEGIN
			set @IDDEST = '2'

			--comenta essa linha aqui FELIPE
			update nfitens 
			set ICMSAplicado = aliqICMSInter
			where LojaOrigem = @Loja 
			AND Serie = @Serie 
			AND NF = @NF
			and @tiponota = 'V'

		
			IF @pessoa = 'F' or @pessoa = 'U' or @Pessoa = 'J' or @pessoa = 'O' and @Tiponota NOT IN ('S','E') 
				Begin
					Update nfitens set 
					       CFOP = 6404 
					  from nfitens, produtoloja 
					 where referencia = pr_referencia 
					   and pr_substituicaoTributaria = 'S' 
					   and LojaOrigem = @Loja 
					   AND Serie = @Serie 
					   AND NF = @NF  
					   and @tiponota <> 'S'
					print ('Update nfitens set CFOP = 6404')
				end 
				
			IF @pessoa = 'F' or @pessoa = 'U' and @Tiponota NOT IN ('S','E') 
				Begin
					Update nfitens set 
					       CFOP = 6108 
					  from nfitens, produtoloja 
					 where referencia = pr_referencia 
					   and pr_substituicaoTributaria = 'N' 
					   and LojaOrigem = @Loja 
					   AND Serie = @Serie 
					   AND NF = @NF  
					   and @tiponota <> 'S'
					print ('Update nfitens set CFOP = 6108')
				end
				
			IF @Pessoa = 'J' or @pessoa = 'O' and @Tiponota NOT IN ('S','E') 
				Begin
					Update nfitens set 
					       CFOP = 6102 
					  from nfitens, produtoloja 
					 where referencia = pr_referencia 
					   and pr_substituicaoTributaria = 'N' 
					   and LojaOrigem = @Loja 
					   AND Serie = @Serie 
					   AND NF = @NF
					   and @tiponota <> 'S'
					print ('Update nfitens set CFOP = 6102')
				end
		END

	END

	IF rtrim(ltrim(@tiponota)) = 'T'
		Begin
			set @IDDEST = '1'
			Update nfitens set 
			       CFOP = 5409 
			  from nfitens, produtoloja 
			 where referencia = pr_referencia 
			   and pr_substituicaoTributaria = 'S' 
			   and LojaOrigem = @Loja 
			   AND Serie = @Serie 
			   AND NF = @NF
			print ('Update nfitens set CFOP = 5409 transferencia ST')

			Update nfitens set 
			       CFOP = 5152 
			  from nfitens, produtoloja 
			 where referencia = pr_referencia 
			   and pr_substituicaoTributaria = 'N' 
			   and LojaOrigem = @Loja 
			   AND Serie = @Serie 
			   AND NF = @NF
		end
	
			
		--update NFItens set 
		--       CFOP = (select codoper 
		--                 from NFCapa 
		--                where LojaOrigem = @Loja 
		--                  AND Serie = @Serie 
		--                  AND NF = @NF )	
		--  from NFItens 
		-- where LojaOrigem = @Loja 
		--   AND Serie = @Serie 
		--   AND NF = @NF
		
	Update nfcapa set codoper = (select top 1 CFOP from nfitens where LojaOrigem = @Loja AND Serie = @Serie AND NF = @NF) 
	from nfcapa where LojaOrigem = @Loja AND Serie = @Serie AND NF = @NF			
				
	print 'NF'
	print @CondPagtoNF
	
	If @CondPagtoNF = 1
	   Begin
		Select @CondPagto = 0
	   End
	   
	If @CondPagtoNF = 3
	   Begin
		Select @CondPagto = 2
	   End
	   
	If @CondPagtoNF between 4 and 199 
	   Begin
		Select @CondPagto = 1
	   End
	   
	If @CondPagtoNF = 2 or @CondPagtoNF >= 200 
           Begin		
                Select @CondPagto = 2
           End

	select @Operacao = (select top 1 cn_descricaooperacao 
	                      from codigooperacaonovo, NFCapa 
	                     where codoper = cn_codigooperacaonovo 
	                       and LojaOrigem = @Loja 
	                       AND Serie = @Serie 
	                       AND NF = @NF)
	
	if LTrim(Rtrim(@Operacao)) = ''	   
	Begin
		Select @Operacao = 'Venda.'
	End
	  
	/*
	FINNFE
	1 – NF-e normal
	2 – NF-e complementar
	3 – NF-e de ajuste
	4 – Devolução de mercadoria
	*/

	SET @finNFe = '1'
	if  @Tiponota <> 'E' 
	select @entradaSaida =  '1'

	if @cfop = '5202' or @cfop = '5411' or @cfop = '5553' or @cfop = '5909'  or @cfop = '6202' or @cfop = '6411' or @cfop = '6913' 
	begin
		select @entradaSaida = '1'
		select @finNFe = '4'
	end
	
	if @cfop = '1202' or @cfop = '1411' or @cfop = '2202' 
	begin
		select @entradaSaida = '0'
		select @finNFe = '4'
	end	
	
	if @cfop = '5918'  
	begin
		select @entradaSaida = '1'
		select @finNFe = '4'
	end	


	set @IE = (select top 1 ce_inscricaoEstadual 
	             from FIN_Cliente, NFCapa 
	            where cliente = CE_CodigoCliente 
	              and NF = @NF 
	              and serie = @Serie 
	              and LOJAORIGEM = @Loja)

	if @Pessoa = 'F' or @Pessoa = 'U' 
	begin
		set @pessoa = '9'
		set @IE = ''
	end 
	
	if @Pessoa = 'J' or @Pessoa = 'O' 
	begin
		set @pessoa = '1'	
		
		if @IE = 'ISENTO'
		begin
			set @pessoa = '9'	
			set @IE = ''	
		end 
		
	end 



	-- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- 
	
	Select @SQL = 'INSERT INTO NFe_ide (eLoja,eNF,eSerie,cUF,cNF,natOp,indPag,mod,serie,nNF,dEmi,dSaiEnt,hSaiEnt,
	tpNF,cMunFG,tpImp,tpEmis,cDV,tpAmb,finNFe,procEmi,verProc,dhCont,xJust,IDDEST,INDFINAL,INDPRES,refNFe) Select LojaOrigem AS eLoja,nf AS eNF,
	Serie as eSerie,'+''''+ LTrim(RTrim(@UFLoja))+'''' +' AS cUF,'+ LTrim(Rtrim(@NF)) +' As cNF,
	' + '''' + LTrim(RTrim(@Operacao)) + '''' + ' as natop,
	'+ @CondPagto +' As indPag,'+ '''55''' +' AS mod,'+'''1'''+' As serie,
	' + ''''+ LTrim(RTrim(@NroNF_NFe))+'''' +' AS nNF,dataemi As dEmi,DataEmi As dSaiEnt,
	Hora as hSaiEnt,' + '' + @entradaSaida + '' + ' As tpNF,LO_CodigoMunicipio As cMunFG,' + '''1''' + ' As tpImp,
	' + '''1''' + ' As tpEmis,'+ ''' ''' +' As cDV,' +'''2'''+ ' As tpAmb,' + '''' + @finNFe + '''' + ' As finNFe,
	' + '''3''' + ' As procEmi,'+ '''2.0.0''' +' As verProc,getdate() As dhCont,
	' + '''Erro no envio da Nota Fiscal Eletronica devido a problemas com Sefaz''' + ' As xJust, 
	''' + @IDDEST + ''' as IDDEST,''1'' as INDFINAL,''1'' as INDPRES, ChaveNFeDevolucao
	FROM NFCapa (NOLOCK), Loja (NOLOCK) 
	WHERE LojaOrigem = '+''''+ LTrim(Rtrim(@Loja)) +''''+' AND Serie = '+''''+ @Serie + '''' +
	' AND NF = '+ LTrim(Rtrim(@NF)) +' AND LojaOrigem = LO_Loja collate sql_latin1_general_cp1_ci_as'

	Print (@SQL)
	Exec (@SQL)
	
--select * from NFe_ide where eNF = '2049'


	-- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- 

	IF rtrim(ltrim(@tiponota)) = 'T'
		Select @SQL = 'INSERT INTO NFE_controle (eLoja,eNF,eSerie,danfe_IMPRESSORA,danfe_RETORNARESP,
		email_DESTINATARIO,email_ASSUNTO,email_MENSAGEM,email_EMAILEMITENTE,email_NOMEEMITENTE,email_ANEXOPDF,
		email_ANEXOXML,email_ANEXOPROTOCOLO,email_anexoadicional,email_COMPACTADO,email_RETORNARESP) 
		Select LojaOrigem AS eLoja,nf AS eNF,Serie as eSerie,CTS_DanfeImpressora AS danfe_IMPRESSORA,''3'' as danfe_RETORNARESP,
		'''' as email_DESTINATARIO,'''' as email_ASSUNTO,'''' AS email_MENSAGEM,
		''nfesaida@demeo.com.br'' email_EMAILEMITENTE,LO_NomeFantasia AS email_NOMEEMITENTE,''SIM'' as email_ANEXOPDF,
		''SIM'' as email_ANEXOXML,''SIM'' as email_ANEXOPROTOCOLO, ''NAO'' as email_anexoadicional,''NAO'' as email_COMPACTADO, ''1'' email_RETORNARESP
		FROM ControleSistema, NFCapa (NOLOCK), Loja (NOLOCK) 
		WHERE LojaOrigem = LO_loja and LojaOrigem = '+''''+ LTrim(Rtrim(@Loja)) +''''+' AND Serie = '+''''+ @Serie + '''' +
		' AND NF = '+ LTrim(Rtrim(@NF)) +' AND LojaOrigem = LO_Loja collate sql_latin1_general_cp1_ci_as'
	ELSE
		Select @SQL = 'INSERT INTO NFE_controle (eLoja,eNF,eSerie,danfe_IMPRESSORA,danfe_RETORNARESP,
		email_DESTINATARIO,email_ASSUNTO,email_MENSAGEM,email_EMAILEMITENTE,email_NOMEEMITENTE,email_ANEXOPDF,
		email_ANEXOXML,email_ANEXOPROTOCOLO,email_anexoadicional,email_COMPACTADO,email_RETORNARESP) 
		Select LojaOrigem AS eLoja,nf AS eNF,Serie as eSerie,CTS_DanfeImpressora AS danfe_IMPRESSORA,''3'' as danfe_RETORNARESP,
		ce_email as email_DESTINATARIO,''Nota Fiscal Eletrônica ' + LTrim(Rtrim(@NF)) + ' - '' + LO_NomeFantasia as email_ASSUNTO,''Olá '' + ltrim(rtrim(CE_Razao)) + '' 
		Você está recebendo uma cópia da DANFE e o arquivo XML'' AS email_MENSAGEM,
		''nfesaida@demeo.com.br'' email_EMAILEMITENTE,LO_NomeFantasia AS email_NOMEEMITENTE,''SIM'' as email_ANEXOPDF,
		''SIM'' as email_ANEXOXML,''SIM'' as email_ANEXOPROTOCOLO, ''NAO'' as email_anexoadicional,''NAO'' as email_COMPACTADO, ''1'' email_RETORNARESP
		FROM ControleSistema, NFCapa (NOLOCK), fin_Cliente, Loja (NOLOCK) 
		WHERE LojaOrigem = LO_loja and cliente = CE_CodigoCliente and LojaOrigem = '+''''+ LTrim(Rtrim(@Loja)) +''''+' AND Serie = '+''''+ @Serie + '''' +
		' AND NF = '+ LTrim(Rtrim(@NF)) +' AND LojaOrigem = LO_Loja collate sql_latin1_general_cp1_ci_as'

	Print (@SQL)
	Exec (@SQL)

	-- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- 
	/*
		1 – Simples Nacional (SN);
		2 – Simples Nacional – excesso de sublimite de receita bruta;
		3 – Regime Normal.
	*/






	Select @SQL = 'INSERT INTO NFe_emit(eLoja,eNF,eSerie,CNPJ,xNome,xFant,xLgr,nro,xCpl,xBairro,cMun,xMun,UF,
	CEP,cPais,xPais,fone,IE,IEST,IM,CNAE,CRT) SELECT LojaOrigem as eLoja,NF as eNF,Serie as eSerie,
	LO_CGC As CNPJ,LO_razao As xNome,LO_NomeFantasia As xFant,
	Lo_Endereco As xLgr,Lo_numero As nro,'''' As xCpl,LO_Bairro As xBairro,
	LO_CodigoMunicipio As cMun,LO_Municipio As xMun,LO_UF As UF,LO_CEP As CEP, 
	'+ '''1058''' +' As cPais, '+'''Brasil'''+' As xPais,LO_DDD + LO_Telefone As fone,
	LO_InscricaoEstadual As IE,'+''' '''+' As IEST,'+''' '''+' As IM,'+''' '''+' As CNAE, 
	'+''''+ @CRT +''''+' As CRT
	FROM Loja (NOLOCK), NFCapa (NOLOCK) WHERE LojaOrigem = LO_loja And 
	LojaOrigem = '+''''+ LTrim(Rtrim(@Loja)) +''''+' AND Serie = '+ '''' + @Serie + '''' +
	' AND NF = '+ LTrim(Rtrim(@NF))

	Print (@SQL)
	Exec (@SQL)

	IF rtrim(ltrim(@tiponota)) = 'T'
		Select @SQL = 'INSERT INTO NFe_dest (eLoja,eNF,eSerie,CNPJ,xNome,xLgr,nro,xCpl,xBairro,cMun,xMun,UF,CEP,cPais,
		xPais,fone,IE,ISUF,email,INDIEDEST) SELECT ' + '''' + LTrim(Rtrim(@Loja)) + '''' + ' as eLoja,' + '''' + LTrim(Rtrim(@NF)) + '''' + ' as eNF, ''NE'' as eSerie,
		(Case When len(lo_CGC) = 14 Then lo_cgc else substring(lo_cgc, 2, 14) end) as CNPJ,
		lo_razao As xNome, lo_endereco As xLgr, lo_numero As nro,'''' As xCpl,
		lo_bairro As xBairro, lo_codigomunicipio As cMun, lo_municipio As xMun, lo_uf As UF,
		lo_cep as CEP,
		''1058'' As cPais,'+'''Brasil'''+' AS xPais,lo_telefone As fone,
		lo_inscricaoEstadual as IE,
		'''' As ISUF,LO_emailoja as Email, ''' + '9' +  ''' as INDIEDEST
		FROM loja (nolock)
		WHERE lo_loja = '+''''+ @ClienteT +''''
	else
	--IF rtrim(ltrim(@tiponota)) = 'E'
	--	Select @SQL = 'INSERT INTO NFe_dest (eLoja,eNF,eSerie,CNPJ,xNome,xLgr,nro,xCpl,xBairro,cMun,xMun,UF,CEP,cPais,
	--	xPais,fone,IE,ISUF,email,INDIEDEST) SELECT ' + '''' + LTrim(Rtrim(@Loja)) + '''' + ' as eLoja,' + '''' + LTrim(Rtrim(@NF)) + '''' + ' as eNF, ''NE'' as eSerie,
	--	(Case When len(lo_CGC) = 14 Then lo_cgc else substring(lo_cgc, 2, 14) end) as CNPJ,
	--	lo_razao As xNome, lo_endereco As xLgr, lo_numero As nro,'''' As xCpl,
	--	lo_bairro As xBairro, lo_codigomunicipio As cMun, lo_municipio As xMun, lo_uf As UF,
	--	lo_cep as CEP,
	--	''1058'' As cPais,'+'''Brasil'''+' AS xPais,lo_telefone As fone,
	--	lo_inscricaoEstadual as IE,
	--	'''' As ISUF,LO_emailoja as Email, ''' + '9' +  ''' as INDIEDEST
	--	FROM loja (nolock)
	--	WHERE lo_loja = '+''''+ @Loja +''''
	--ELSE 
		Select @SQL = 'INSERT INTO NFe_dest (eLoja,eNF,eSerie,CNPJ,CPF,xNome,xLgr,nro,xCpl,xBairro,cMun,xMun,UF,CEP,cPais,
		xPais,fone,IE,ISUF,email, INDIEDEST)SELECT LojaOrigem as eLoja,NF as eNF,Serie as eSerie,
		(Case When len(CE_CGC) = 14 Then CE_CGC else '+''' '''+' end),
		(Case When len(CE_CGC) = 11 Then CE_CGC else '+''' '''+' end),
		CE_Razao As xNome,CE_Endereco As xLgr,CE_numero As nro,CE_Complemento As xCpl,
		CE_bairro As xBairro,CE_CodigoMunicipio As cMun,CE_Municipio As xMun,CE_Estado As UF,
		'+''''+ LTrim(Rtrim(@CEPCliente)) +''''+' as CEP,
		' + '''1058''' + ' As cPais,'+'''Brasil'''+' AS xPais,CE_telefone As fone,
		''' + @IE + ''' as IE,
		CE_InscricaoEstadualSuframa As ISUF,CE_email as Email, ''' + @pessoa +  ''' as INDIEDEST
		FROM NFCapa (NOLOCK),fin_Cliente (nolock)
		WHERE cliente = CE_CodigoCliente AND
		LojaOrigem = '+''''+ LTrim(Rtrim(@Loja)) +''''+
		' AND Serie = '+'''' + @Serie + ''''+' AND NF = '+ LTrim(Rtrim(@NF));		
	


	--Print @SQL-- select ce_cgc,* from fin_cliente where ce_codigocliente = 60046
	Print (@SQL)
	Exec (@SQL)

	-- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- 

	--select * from nfe_estrutura
	Select @SQL = 'INSERT INTO NFe_prod (eLoja,eNF,eSerie,H_nItem,I_cProd,I_cEAN,I_xProd,I_NCM,I_EXTIPI,I_CFOP,
	I_uCom,I_qCom,I_vUnCom,I_vProd,I_cEANTrib,I_uTrib,I_qTrib,I_vUnTrib,I_vFrete,I_vSeg,I_vDesc,I_vOutro,
	I_indTot,N_origICMS,N_CSTICMS,N_modBCICMS,N_vBCICMS,N_pRedBCICMS,N_pICMS,N_vICMS,N_modBCST,N_pMVAST,
	N_pRedBCST,N_vBCST,N_pICMSST,N_vICMSST,O_cIEnq,O_CNPJProd,O_cSelo,O_qSelo,O_cEnq,O_CSTIPI,
	O_vBCIPI,O_qUnid,O_vUnid,O_pIPI,O_vIPI,O_CSTIPINT,P_vBCII,P_vDespAdu,P_vII,P_vIOF,Q_CSTPIS,
	Q_vBCPIS,Q_pPIS,Q_qBCProdPIS,Q_vAliqProdPIS,Q_vPIS,R_vBCPISST,R_pPISST,R_qBCProdPISST,
	R_vAliqProdPISST,R_vPISST,S_CSTCOFINS,S_vBCCOFINS,S_pCOFINS,S_qBCProdCOFINS,S_vAliqProdCOFINS,
	S_vCOFINS,T_vBCCOFINSST,T_pCOFINSST,T_qBCProdCOFINSST,T_vAliqProdCOFINSST,T_vCOFINSST,
	U_vBCISSQN,U_vAliqISSQN,U_vISSQN,U_cMunFGISSQN,U_cListServ,U_cSitTrib,V_infAdProd,W_vFCPUFDest,
	W_pFCPUFDest, W_pICMSUFDest, W_pICMSInter, W_pICMSInterPart, W_vICMSUFRemet, W_vICMSUFDest, W_vBcUFDest, I_CEST, X_orig,X_CSOSN) 

	SELECT LojaOrigem as eLoja,NF as eNF,Serie as eSerie,ITEM As H_nItem,Referencia As I_cProd,
	'+''' '''+' As I_cEAN,PR_Descricao As I_xProd,PR_ClasseFiscal As I_NCM,'+''' '''+' As I_EXTIPI,
	CFOP As I_CFOP,PR_Unidade As I_uCom,QTDE As I_qCom,VLUnit As I_vUnCom,
	VLTotItem As I_vProd,'+''' '''+' As I_cEANTrib,PR_UNIDADE AS I_uTrib,QTDE aS I_qTrib,
	VLUnit as I_vUnTrib,

	(Case When Item = '+  LTrim(Rtrim(convert(char(20),@Item))) +' Then ((((vltotitem - desconto) * '+ LTrim(Rtrim(convert(char(20),@PercFrete))) +') / 100) + '+ LTrim(Rtrim(convert(char(20),@DiferencaFrete))) +') 
	else (((vltotitem - desconto) * '+ LTrim(Rtrim(convert(char(20),@PercFrete))) +') / 100) end),

	'+'''0'''+' as I_vSeg,desconto as I_vDesc,'+'''0'''+' as I_vOutro, 
	'+'''1'''+' I_indTot,'+ '''0''' +' as N_origICMS,CSTICMS as N_CSTICMS,'+ '''2''' +' as N_modBCICMS,
	baseicms as N_vBCICMS,PR_codigoReducaoICMS as N_pRedBCICMS,ICMSAplicado as N_pICMS,
	ValorICMS as N_vICMS,'+'''0'''+' as N_modBCST,'+'''0'''+' as N_pMVAST,'+'''0'''+' as N_pRedBCST,
	'+'''0'''+' as N_vBCST,'+'''0'''+' as N_pICMSST,'+'''0'''+' as N_vICMSST,
	'+''' '''+' as O_cIEnq,'+''' '''+' as O_CNPJProd,'+''' '''+' as O_cSelo,'+''' '''+' as O_qSelo,
	'+'''999'''+' as O_cEnq,'+'''50'''+' as O_CSTIPI, baseIPI as O_vBCIPI, qtde as O_qUnid,
	vlUnit as O_vUnid, aliqIPI as O_pIPI, vlIpi as O_vIPI,'+''' '''+' as O_CSTIPINT,
	'+'''0'''+' as P_vBCII,'+'''0'''+' as P_vDespAdu,'+'''0'''+' as P_vII,
	'+'''0'''+' as P_vIOF,'+'''01'''+' as Q_CSTPIS,

	(Case When Item = '+  LTrim(Rtrim(convert(char(20),@Item))) +' Then ( (vltotitem - desconto) + (((vltotitem - desconto) * 
	'+ LTrim(Rtrim(convert(char(20),@PercFrete))) +') / 100) + '+ LTrim(Rtrim(convert(char(20),@DiferencaFrete))) +') 
	else ((vltotitem - desconto) + (((vltotitem - desconto) * '+ LTrim(Rtrim(convert(char(20),@PercFrete))) +') / 100)) end), 

	'+'''1.65'''+' as Q_pPIS,'+'''0'''+' as Q_qBCProdPIS,'+'''0'''+' as Q_vAliqProdPIS,

	(Case When Item = '+  LTrim(Rtrim(convert(char(20),@Item))) +'
	Then ((((vltotitem - desconto) + (((vltotitem - desconto) * '+ LTrim(Rtrim(convert(char(20),@PercFrete))) +' ) /100) + 
	'+ LTrim(Rtrim(convert(char(20),@DiferencaFrete))) +' ) * 1.65)/100)
	else ((((vltotitem - desconto) + (((vltotitem - desconto) * '+ LTrim(Rtrim(convert(char(20),@PercFrete))) +' ) /100)) * 1.65)/100) end) as Q_vPIS,

	'+'''0'''+' as R_vBCPISST,'+'''0'''+' as R_pPISST,'+'''0'''+' as R_qBCProdPISST,
	'+'''0'''+' as R_vAliqProdPISST,'+'''0'''+' as R_vPISST,'+'''01'''+' as S_CSTCOFINS,

	(Case When Item = '+  LTrim(Rtrim(convert(char(20),@Item))) +' Then ( (vltotitem - desconto) + (((vltotitem - desconto) * 
	'+ LTrim(Rtrim(convert(char(20),@PercFrete))) +') / 100) + '+ LTrim(Rtrim(convert(char(20),@DiferencaFrete))) +') 
	else ((vltotitem - desconto) + (((vltotitem - desconto) * '+ LTrim(Rtrim(convert(char(20),@PercFrete))) +') / 100)) end),  

	'+'''7.60'''+' as S_pCOFINS,'+'''0'''+' as S_qBCProdCOFINS,'+'''0'''+' as S_vAliqProdCOFINS,

	(Case When Item = '+  LTrim(Rtrim(convert(char(20),@Item))) +'
	Then ((((vltotitem - desconto) + (((vltotitem - desconto) * '+ LTrim(Rtrim(convert(char(20),@PercFrete))) +' ) /100) + 
	'+ LTrim(Rtrim(convert(char(20),@DiferencaFrete))) +' ) * 7.60)/100)
	else ((((vltotitem - desconto) + (((vltotitem - desconto) * '+ LTrim(Rtrim(convert(char(20),@PercFrete))) +' ) /100)) * 7.60)/100) end),

	'+'''0'''+' as T_vBCCOFINSST,'+'''0'''+' as T_pCOFINSST,
	'+'''0'''+' as T_qBCProdCOFINSST,'+'''0'''+' as T_vAliqProdCOFINSST,
	'+'''0'''+' as T_vCOFINSST,'+'''0'''+' as U_vBCISSQN,'+'''0'''+' as U_vAliqISSQN,
	'+'''0'''+' as U_vISSQN,'+''' '''+' as U_cMunFGISSQN,'+''' '''+' as U_cListServ,
	'+''' '''+' as U_cSitTrib,'+''' '''+' as V_infAdProd,
	valorICMSFECP as W_vFCPUFDest, aliqICMSFECP as W_pFCPUFDest, aliqICMSDest as W_pICMSUFDest, 
	aliqICMSInter as W_pICMSInter, ICMSInterPart as W_pICMSInterPart, valICMSRemet as W_vICMSUFRemet,
	valICMSDest as W_vICMSUFDest,(case when valorICMSFECP = 0 then 0 else vlunit2 end) as W_vBcUFDest, pr_cest, '+'''0'''+','+'''102'''+'
	FROM produtoloja (NOLOCK), NFItens (NOLOCK) 
	WHERE PR_Referencia = Referencia AND LojaOrigem = '+''''+ LTrim(Rtrim(@Loja)) +''''+ 
	' AND Serie = '+'''' + @Serie + ''''+' AND NF = '+ LTrim(Rtrim(@NF)) +' Order by H_nItem'

	--select * from nfe_prod whe

	Print @SQL 
	Exec (@SQL)

	-- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- 

	Select @SQL = 'Insert into NFe_total (eLoja,eNF,eSerie,vBCICMS,vICMS,vBCST,vST,vProd,vFrete,vSeg,vDesc,vII,vIPI,
	vCOFINS,vOutro,vNF,vServ,vBCISSQ,vISS,vPIS,vCOFINSISSQ,vRetPIS,vRetCOFINS,vRetCSLL,vBCIRRF,
	vIRRF,vBCRetPrev,vRetPrev, vFCPUFDest, vICMSUFRemet, vICMSUFDest, vVICMSDESON)Select LojaOrigem as eLoja,NF as eNF,Serie as eSerie,

	(Case When baseicms is null Then 0 else baseicms end), 

	VlrICMS AS vICMS,0 as vBCST,0 as vST,
	vlrmercadoria as vProd,Fretecobr as vFrete,'+''' 0''' + ' as vSeg,Desconto as vDesc,
	'+ '''0''' +' as vII,totalipi as vIPI,(((Totalnota-totalipi) * 7.60)/100) as vCOFINS,0 as vOutro,
	TotalNota as vNF,'+ '''0''' +' as vServ,'+ '''0''' +' as vBCISSQ,'+ '''0''' +' as vISS,
	(((Totalnota - totalipi) * 1.65)/100) as vPIS,'+'''0'''+' as vCOFINSISSQ,'+ '''0''' +' as vRetPIS,
	'+ '''0''' +' as vRetCOFINS,'+ '''0''' +' as vRetCSLL,'+ '''0''' +' as vBCIRRF,
	'+ '''0''' +' as vIRRF,'+ '''0''' +' as vBCRetPrev,'+ '''0''' +' as vRetPrev, 
	valorICMSFECP as vFCPUFDest, valICMSRemet as vICMSUFRemet, valICMSDest as vICMSUFDest, 0
	from NFCapa(Nolock) 
	Where LojaOrigem = '+''''+ LTrim(Rtrim(@Loja)) +''''+' AND Serie = '+'''' + @Serie + ''''+
	' AND NF = '+ LTrim(Rtrim(@NF))

	Print @SQL -- baseicms as vBCICMS,
	Exec (@SQL)
	
	--select * from nfe_prod where enf = 3796
	--select * from NFItens where nf = 3796
	
	-- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- 

	Select @SQL = 'Insert into NFe_transp (eLoja,eNF,eSerie,modFrete,CNPJ,CPF,xNome,IE,xEnder,xMun,UF,vServ,vBCRet,pICMSRet,
	vICMSRet,CFOP,cMunFG,placa,UFveic,RNTC,qVol,esq,marca,nVol,pesoL,pesoB,nLacres)
	Select LojaOrigem as eLoja,NF as eNF,Serie as eSerie,TipoFrete as modFrete,'+''' '''+' As CNPJ,
	'+''' '''+' as CPF,'+''' '''+' as xNome,'+''' '''+' as IE,'+''' '''+' as xEnder,
	'+''' '''+' as xMun,'+''' '''+' as UF,'+ '''0'''+' as vServ,'+ '''0''' +' as vBCRet,
	'+ '''0''' +' as pICMSRet,'+ '''0''' +' as vICMSRet,'+''' '''+' as CFOP,'+''' '''+' as cMunFG,
	'+''' '''+' as placa,'+''' '''+' as UFveic,'+''' '''+' as RNTC,
	volume as qVol,'+'''VOLUME(S)'''+' as esq,'+''' '''+' as marca,
	'+ '''0''' +' as nVol,pesolq as pesoL,pesobr as pesoB,
	'+ '''0''' +' as nLacres FROM Loja(NOLOCK), NFCapa (NOLOCK)
	Where lojaOrigem = LO_loja And LojaOrigem = '+''''+ LTrim(Rtrim(@Loja)) +''''+' 
	AND Serie = '+'''' + @Serie + ''''+' AND NF = '+ LTrim(Rtrim(@NF))

	Print @SQL
	Exec (@SQL)

	-- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- 

	declare @descricao varchar(max)
	--declare @sequencia int
	--declare @sequenciaMaxima int

	SET @Carimbo = ''
	Declare Temp_Carimbo insensitive cursor for
			Select rtrim(LTRIM(CNF_Carimbo))
			  from CarimboNotaFiscal 
			 where CNF_Loja = @Loja 
			   and cnf_serie = @Serie 
			   and CNF_NF = @NF
			 order by CNF_TipoCarimbo desc, CNF_Sequencia 
	Open Temp_Carimbo
	Fetch Next From Temp_Carimbo Into @Descricao
	While @@Fetch_Status = 0  
		Begin

		set @Carimbo = @Carimbo + @Descricao + '  -  '
			Fetch Next From Temp_Carimbo Into @Descricao
		end
	close Temp_Carimbo
	Deallocate Temp_Carimbo

	--set @Carimbo = left(@Carimbo,len(@Carimbo)-2)

	Select @SQL = 'insert into NFe_infAdic (eLoja,eNF,eSerie,infAdFisco,infCpl,xCampoCont,
	xTextoCont,xCampoFisco,xTextoFisco,nProc,indProc) Select LojaOrigem as eLoja,
	NF as eNF,Serie as eSerie,'+''' '''+' as infAdFisco,''PEDIDO: '''+' + RTrim(LTrim(Convert(VarChar(10),numeroped)))+ 
	'+''', VENDEDOR: '''+' + RTrim(LTrim(Convert(VarChar(10),Vendedor)))+'+''', COND PAGTO: '''+' + 
	(Case When (RTrim(LTrim(cp_condicao))) is Null Then '+''' '''+' else cp_condicao end) + '+'''  -  ' + @Carimbo + '''' + ''+' as infcpl,
	'+'''E-MAIL'''+' as xCampoCont, Upper(LO_EmaiLoja) as xTextoCont,
	'+''' '''+' as xCampoFisco,'+''' '''+' as xTextoFisco,
	'+''' '''+' as nProc,'+''' '''+' as indProc from nfCapa(nolock),condicaopagamento(nolock),Loja(nolock)
	where cp_codigo = condpag and cp_id = 1 AND LojaOrigem = LO_Loja AND LojaOrigem = '+''''+ LTrim(Rtrim(@Loja)) +''''+' 
	AND Serie = '+'''' + @Serie + ''''+' AND NF = '+ LTrim(Rtrim(@NF))

	Print @SQL
	Exec (@SQL)


	Select @SQL = 'insert into NFE_DUP (eLoja,eNF,eSerie,nDup,dVend,vDup) 
	select dp_loja, dp_NotaFiscal, dp_serie,DP_SEQUENCIA, DP_DataVencimento, DP_ValorDuplicata  from Duplicata
	where dp_loja = '+''''+ LTrim(Rtrim(@Loja)) +''''+' 
	AND dp_serie = '+'''' + @Serie + ''''+' AND dp_NotaFiscal = '+ LTrim(Rtrim(@NF))

	Print @SQL
	Exec (@SQL)

	if (SELECT top 1 cts_tipoempresa FROM ControleSistema) = 'SN' 
	begin
		UPDATE nfe_prod set n_vICMS = 0
		UPDATE nfe_prod set n_vBCICMS = 0
		UPDATE nfe_prod set N_CSTICMS = 60
		UPDATE nfe_total set vICMS = 0
		UPDATE nfe_total set vBCICMS = 0
	end

	
	--select * from NFE_NFLojas where NFL_NroNFE = 11363 order by NFL_Sequencia
	--delete NFE_NFLojas where NFL_NroNFE = 11363 order by NFL_Sequencia

	/*
	DROP TABLE NFE_DUP
	select * from NFE_estrutura  where etr_rotulo = 'DUP'
	update NFE_estrutura set ETR_TABELA_DE = 'NFE_DUP' where etr_rotulo = 'DUP'
	update NFE_estrutura set ETR_TABELA_DE = '' where etr_rotulo = 'DUP' AND ETR_SEQUENCIA = 133

	*/
	
END





GO
/****** Object:  StoredProcedure [dbo].[SP_VDA_ImportaProduto_Solucao]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
/*

   Carrega Produto DMAC_Loja Para Solução

*/
Create     Procedure [dbo].[SP_VDA_ImportaProduto_Solucao]
        	                  
As

--NFCapa
	Declare 
                @SQL                    Char(4000),
                @Referencia             char(14),
                @Descricao              char(50),
                @PrecoVenda             float,
                @Unidade                char(02),
                @Qtde                   numeric,
                @ICMSPDV                float,
                @SitTributaria          char(02)
           
         

    
       Declare #TemP_Produto Insensitive Cursor For
 		Select  PR_Referencia,PR_Descricao,PR_PrecoVenda1,PR_Unidade,0,PR_ICMPDV,'00'
		From   Produto 
		
      
       Open  #TemP_Produto 

 	Fetch Next From  #TemP_Produto  into
              @Referencia,@Descricao,@PrecoVenda,@Unidade,@Qtde,@ICMSPDV,@SitTributaria

        While (@@Fetch_status = 0) 
           Begin
 
          -- EXEC DB_Loja..


 	Fetch Next From  #TemP_Produto  into
              @Referencia,@Descricao,@PrecoVenda,@Unidade,@Qtde,@ICMSPDV,@SitTributaria

        End


	If @@ERROR <> 0
	   Begin	
	   	Rollback Transaction		
	   	Return
	   End

	Close #TemP_Produto
	Deallocate #TemP_Produto
/*
exec SP_VDA_ImportaProduto_Solucao
*/

GO
/****** Object:  StoredProcedure [dbo].[SP_VDA_NFE_ChaveAcesso]    Script Date: 23/09/2016 10:42:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE                                      PROCEDURE [dbo].[SP_VDA_NFE_ChaveAcesso]

	@eLoja		Char(5),
	@eNF		Numeric,
	@eSerie		Char(2)

AS

	DECLARE		@SQL        	char(4000),
			@chaveAcesso    char(44),
                        @ChaveAcessoCompleta    char(44),
			@Pesos	        char(43),
                        @RestoDiv       int,
                        @Somatoria      int,
                        @DV		int,
                        @NumeroDiv	int,
                        @algarismo      int,
                        @cont           int,
                        @UF             char(2),
                        @AnoMes         char(4),
			@CNPJ           char(14),
			@modelo         char(2),
 			@Serie		char(3),
                        @Chave          char(8),
                        @NF             char(9),
			@TipoEmissao    char(1),
                        @Resto	        int

     
BEGIN
        

     Select @UF = (Select cUF from nfe_ide where eLoja = @eLoja And eNF = @eNF And eSerie = @eSerie)
     Select @AnoMes = (select substring(replace(CONVERT(varchar,GETDATE(),2),'.',''),1,4))
     Select @CNPJ = (Select CNPJ from nfe_emit where eLoja = @eLoja And eNF = @eNF And eSerie = @eSerie) 
     Select @modelo = (Select mod from nfe_ide where eLoja = @eLoja And eNF = @eNF And eSerie = @eSerie)    
     select @Serie = (Select replicate('0',3 - len(ltrim(rtrim(serie)))) + serie from nfe_ide where eLoja = @eLoja And 
                      eNF = @eNF And eSerie = @eSerie)    
     Select @NF = (Select replicate('0',9 - len(ltrim(rtrim(enf))))  + enf from nfe_ide where eLoja = @eLoja And 
                      eNF = @eNF And eSerie = @eSerie)
     Select @TipoEmissao = (Select tpEmis from nfe_ide where eLoja = @eLoja And eNF = @eNF And eSerie = @eSerie) 
     Select @Chave = (Select cNF from nfe_ide where eLoja = @eLoja And eNF = @eNF And eSerie = @eSerie)  
     Select @ChaveAcesso = @UF + @AnoMes + @CNPJ + @modelo + @Serie + @NF + @TipoEmissao + @Chave
     Select @Pesos = '4329876543298765432987654329876543298765432'
    
print @ChaveAcesso
-- 
--   35111060872124000350550010000000301 53131751 
--ok 35111060872124000350550010000003671 00000367



     Select @cont = 1
     Select @somatoria = 0

     while @cont < 44
       Begin
         select @resto = (convert(int,substring(@ChaveAcesso,@cont,1)) * convert(int,substring(@Pesos,@cont,1)))
         Select @Somatoria = @Somatoria + @Resto
       
        print convert(int,substring(@ChaveAcesso,@cont,1)) 
        print convert(int,substring(@Pesos,@cont,1))   
        print (convert(int,substring(@ChaveAcesso,@cont,1))*convert(int,substring(@Pesos,@cont,1)))
        print @Somatoria
        print ''


        select @Cont = @Cont + 1
       End

     Select @RestoDiv = @Somatoria % 11

     If @RestoDiv > 1
         Select @DV = 11 - @RestoDiv
     else
         Select @DV = 0
    

     Select @ChaveAcessoCompleta = rtrim(ltrim(@ChaveAcesso)) + convert(char(1),@DV)
  
print @ChaveAcessoCompleta
print @DV

     Select @SQL = 'update nfe_ide set cDV = ' + '''' + rtrim(ltrim(convert(char(2),@DV))) + '''' + 
                   ' ,ChaveAcesso = ' + '''' + @ChaveAcessoCompleta + '''' +
                   ' where eLoja = ' + '''' + rtrim(ltrim(@eLoja)) + '''' + 
                   ' And eNF = ' + '''' + rtrim(ltrim(convert(char(5),@eNF))) + '''' +
                   ' And eSerie = ' + '''' + @eSerie + ''''
    exec (@SQL)   
     
     Select @SQL = 'update nfcapa set ChaveNFe = ' + '''' + @ChaveAcessoCompleta + '''' + 
                   ' where LojaOrigem = ' + '''' + rtrim(ltrim(@eLoja)) + '''' + 
                   ' And NF = ' + '''' + rtrim(ltrim(convert(char(5),@eNF))) + '''' +
                   ' And Serie = ' + '''' + @eSerie + ''''

     exec (@SQL)

END
 

/*
exec SP_VDA_NFE_ChaveAcesso '134',7,'NE'
select '35',CUF,mod,serie,nnf,tpemis,cnf,cdv from nfe_ide where enf = 7
select cnpj from nfe_emit  where enf = 7


select * from nfe_ide where enf = 7
select * from nfe_emit  where enf = 7
select * from nfcapa where nf = 7


*/

GO
EXEC sys.sp_addextendedproperty @name=N'MS_Description', @value=N'0' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'GarantiaEstendida', @level2type=N'COLUMN',@level2name=N'ge_seqCancelamento'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_Description', @value=N'''1900-01-01 00:00:00.000''' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'GarantiaEstendida', @level2type=N'COLUMN',@level2name=N'ge_dataCancelamento'
GO
