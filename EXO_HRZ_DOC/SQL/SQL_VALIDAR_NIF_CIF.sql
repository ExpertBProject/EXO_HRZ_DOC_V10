CREATE FUNCTION [EXO_VALIDAR_NIF_CIF]
(
	@pLicTradNum NVARCHAR(32)
)
RETURNS INT
AS
BEGIN
	DECLARE @vLetraINI NVARCHAR(1)
	DECLARE @vLetraFIN NVARCHAR(1)
	DECLARE @vPais NVARCHAR(2)
	DECLARE @vLetra NVARCHAR(1)
	DECLARE @vNumero NVARCHAR(32)
	DECLARE @vAbcdario NVARCHAR(23)
	DECLARE @vDigitoControl NVARCHAR(1)
	DECLARE @vUnidades INT
	DECLARE @vLetras NVARCHAR(23)
	DECLARE @vI INT
	DECLARE @vJ INT
	DECLARE @vSumaPares INT
	DECLARE @vSumaImpares INT
	DECLARE @vSumaTotal INT
	DECLARE @vAuxNum1 NVARCHAR(7)
	DECLARE @vAuxNum2 INT
	DECLARE @vEsNumero INT
	DECLARE @vEsNumeroOTRO INT
	DECLARE @vEsNumeroINI INT
	DECLARE @vEsNumeroFIN INT
	DECLARE @vDato NVARCHAR(32)
	DECLARE @va INT
	DECLARE @vb INT
	DECLARE @vc INT
	DECLARE @vNIF_CIF INT
	DECLARE @pResultado INT

	select @vDato = ''
	select @vNumero = ''
	select @vLetra = ''

	IF LEN(LTRIM(RTRIM(@pLicTradNum))) > 2 BEGIN
		SELECT @vLetraINI = LEFT(SUBSTRING(LTRIM(RTRIM(@pLicTradNum)), 3, LEN(LTRIM(RTRIM(@pLicTradNum)))), 1)
		SELECT @vLetraFIN = RIGHT(SUBSTRING(LTRIM(RTRIM(@pLicTradNum)), 3, LEN(LTRIM(RTRIM(@pLicTradNum)))), 1)
		SELECT @vPais = LEFT(LTRIM(RTRIM(@pLicTradNum)), 2)
		
		IF isnumeric(@vLetraINI)=1 BEGIN
				SELECT @vEsNumeroINI=1
				SELECT @vEsNumeroFIN=0
		END
		ELSE BEGIN	
			SELECT @vEsNumeroINI=0		
			IF isnumeric(@vLetraFIN)=1  BEGIN
				SELECT @vEsNumeroFIN=1
			END 
			ELSE BEGIN
				SELECT @vEsNumeroFIN =0
			END
		END										
		IF @vEsNumeroINI = 0 BEGIN
			SELECT @vLetra = @vLetraINI
			SELECT @vNumero = SUBSTRING(LTRIM(RTRIM(@pLicTradNum)), 4, LEN(LTRIM(RTRIM(@pLicTradNum))) - 3)
			SELECT @vNumero = REPLACE(REPLACE(REPLACE(REPLACE(@vNumero, '.', ''), ',', ''), '-', ''), '+', '')
		END
		ELSE BEGIN			
			IF @vEsNumeroFIN = 0 BEGIN
				SELECT @vLetra = @vLetraFIN
				SELECT @vNumero = SUBSTRING(LTRIM(RTRIM(@pLicTradNum)), 3, LEN(LTRIM(RTRIM(@pLicTradNum))) - 3)
				SELECT @vNumero = REPLACE(REPLACE(REPLACE(REPLACE(@vNumero, '.', ''), ',', ''), '-', ''), '+', '')
			END
		END

		IF SUBSTRING(@vNumero, 1, LEN(@vNumero) -1) = '' BEGIN
			SELECT @vEsNumero=0
		END
		ELSE BEGIN
			IF isnumeric(@vnumero)=1 BEGIN
				SELECT @vEsNumero=1
			END
			ELSE BEGIN
				SELECT @vEsNumero=0
			END
		END
		
		IF @vLetraINI <> '' AND @vEsNumeroINI = 0 AND LEN(@vNumero) >= 8 AND LEN(@vNumero) <= 9 AND @vEsNumero = 1 BEGIN
			SELECT @vDigitoControl = RIGHT(@vNumero, 1)
			SELECT @vAuxNum1 = SUBSTRING(@vNumero, 1, 7)
			SELECT @vLetras = 'ABCDEFGHIJ'
			SELECT @vAbcdario = 'ABCDEFGHJPQRSUVNW'
						
			IF CHARINDEX ( @vLetra, @vAbcdario) > 0 BEGIN 
				SELECT @vSumaPares = 0
				SELECT @vSumaImpares = 0
				SELECT @vSumaTotal = 0
				SELECT @vI = 1

				WHILE @vI <= 7 BEGIN
					SELECT @vAuxNum2 = CAST(SUBSTRING(@vAuxNum1, @vI, 1) AS INT)
					IF @vI % 2 = 0 BEGIN
						SELECT @vSumaPares = @vSumaPares + @vAuxNum2
					END
					ELSE BEGIN
						SELECT @vAuxNum2 = @vAuxNum2 * 2										
						SELECT @vJ = 1

						WHILE @vJ <= LEN(CAST(@vAuxNum2 AS NVARCHAR(15))) BEGIN
							SELECT @vSumaImpares = @vSumaImpares + CAST(SUBSTRING(CAST(@vAuxNum2 AS NVARCHAR(15)), @vJ, 1) AS INT)
							SELECT @vJ = @vJ + 1
						END
					END									
					SELECT @vI = @vI + 1
				END							
				SELECT @vSumaTotal = @vSumaTotal + @vSumaPares + @vSumaImpares
				SELECT @vUnidades = CAST(SUBSTRING(CAST(@vSumaTotal AS NVARCHAR(15)), LEN(CAST(@vSumaTotal AS NVARCHAR(15))), 1) AS INT)

				IF @vUnidades <> 0 BEGIN
					SELECT @vUnidades = 10 - @vUnidades
				END

				IF @vLetra = 'A' OR @vLetra = 'B' OR @vLetra = 'E' OR @vLetra = 'H' BEGIN
					IF @vDigitoControl = CAST(@vUnidades AS NVARCHAR(1)) BEGIN
						SELECT @vDato = @vPais + @vLetra + @vNumero
					END
				END
				ELSE BEGIN
					IF @vLetra = 'K' OR @vLetra = 'P' OR @vLetra = 'Q' OR @vLetra = 'S' BEGIN
						IF @vDigitoControl = SUBSTRING(@vLetras, @vUnidades, 1) BEGIN
							SELECT @vDato = @vPais + @vLetra + @vNumero
						END
						ELSE BEGIN
							IF @vUnidades = 0 AND @vDigitoControl = 'J' BEGIN
								SELECT @vDato = @vPais + 'J' + @vNumero
							END
						END
					END
					ELSE BEGIN
						IF @vDigitoControl = CAST(@vUnidades AS NVARCHAR(1)) BEGIN
							SELECT @vDato = @vPais + @vLetra + @vNumero
						END
						ELSE BEGIN
							IF @vDigitoControl = SUBSTRING(@vLetras, @vUnidades, 1) BEGIN
								SELECT @vDato = @vPais +@vLetra + @vNumero
							END
							ELSE BEGIN
								IF @vUnidades = 0 AND @vDigitoControl = 'J' BEGIN
									SELECT @vDato = @vPais + 'J' + @vNumero
								END 
							END
						END
					END
				END					
			END
		END
		ELSE BEGIN
			IF @vNumero='' BEGIN
				SELECT @vEsNumeroOTRO=0
			END
			ELSE BEGIN
				IF isnumeric(@vEsNumero)=1 BEGIN
					SELECT @vEsNumeroOTRO=1
				END
				ELSE BEGIN
					SELECT @vEsNumeroOTRO=0
				END
			END
			IF @vLetraFIN <> '' AND @vEsNumeroFIN = 0 AND LEN(@vNumero) >= 7 AND LEN(@vNumero) <= 10 AND @vEsNumeroOTRO = 1 BEGIN
				SELECT @vLetras = 'TRWAGMYFPDXBNJZSQVHLCKE'
				SELECT @va = 0
				SELECT @vNIF_CIF = CAST(@vNumero AS INT)
				SELECT @vb = - 1

				WHILE (@vb <> 0) BEGIN
					SELECT @vb = FLOOR(@vNIF_CIF / 24)
					SELECT @vc = @vNIF_CIF - (24 * @vb)
					SELECT @va = @va + @vc
					SELECT @vNIF_CIF = @vb
				END

				SELECT @vb = FLOOR(@va / 23)
    			SELECT @vc = @va - (23 * @vb)			
				SELECT @vDato = @vPais +@vNumero + SUBSTRING(@vLetras, @vc + 1, 1)
			END
		END
	END
	
	IF @pLicTradNum <> '' AND @pLicTradNum = @vDato BEGIN
		SELECT @pResultado = 1
	END
	ELSE BEGIN
		SELECT @pResultado = 0
	END

	RETURN isnull(@pResultado, 0) 
END

