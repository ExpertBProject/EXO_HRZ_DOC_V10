CREATE FUNCTION EXO_VALIDAR_NIF_CIF(IN pLicTradNum NVARCHAR(32)) RETURNS pResultado INT
LANGUAGE SQLSCRIPT
AS
	-- Declare the return variable here
	vLetraINI NVARCHAR(1);
	vLetraFIN NVARCHAR(1);
	vPais NVARCHAR(2);
	vLetra NVARCHAR(1);
	vNumero NVARCHAR(32);
	vAbcdario NVARCHAR(23);
	vDigitoControl NVARCHAR(1);
	vUnidades INT;
	vLetras NVARCHAR(23);
	vI INT;
	vJ INT;
	vSumaPares INT;
	vSumaImpares INT;
	vSumaTotal INT;
	vAuxNum1 NVARCHAR(7);
	vAuxNum2 INT;
	vEsNumero INT;
	vEsNumeroOTRO INT;
	vEsNumeroINI INT;
	vEsNumeroFIN INT;
	vDato NVARCHAR(32);
	va INT;
	vb INT;
	vc INT;
	vNIF_CIF INT;
		
BEGIN
	vDato := '';

	vNumero := '';
	vLetra := '';
	
	-- Add the T-SQL statements to compute the return value here
	IF LENGTH(LTRIM(RTRIM(:pLicTradNum))) > 2 THEN
		-- Por lo menos la longitud del NIF/CIF tiene que ser de 3 caracteres porque 2 son obligatorios (el país)
		vLetraINI := LEFT(SUBSTRING(LTRIM(RTRIM(:pLicTradNum)), 3, LENGTH(LTRIM(RTRIM(:pLicTradNum)))), 1);
		vLetraFIN := RIGHT(SUBSTRING(LTRIM(RTRIM(:pLicTradNum)), 3, LENGTH(LTRIM(RTRIM(:pLicTradNum)))), 1);
		vPais := LEFT(LTRIM(RTRIM(:pLicTradNum)), 2);
		
		SELECT CASE WHEN :vLetraINI = '' THEN 0 ELSE LOCATE_REGEXPR(START '([[:digit:]]{1})' IN :vLetraINI GROUP 1) END INTO vEsNumeroINI FROM DUMMY;
		SELECT CASE WHEN :vLetraFIN = '' THEN 0 ELSE LOCATE_REGEXPR(START '([[:digit:]]{1})' IN :vLetraFIN GROUP 1) END INTO vEsNumeroFIN FROM DUMMY;		
						
		IF :vEsNumeroINI = 0 THEN
			-- Validar CIF
			vLetra := :vLetraINI;
			vNumero := SUBSTRING(LTRIM(RTRIM(:pLicTradNum)), 4, LENGTH(LTRIM(RTRIM(:pLicTradNum))) - 3);
			vNumero := REPLACE(REPLACE(REPLACE(REPLACE(:vNumero, '.', ''), ',', ''), '-', ''), '+', '');
		ELSE			
			IF :vEsNumeroFIN = 0 THEN
				-- Validar NIF
				vLetra := :vLetraFIN;
				vNumero := SUBSTRING(LTRIM(RTRIM(:pLicTradNum)), 3, LENGTH(LTRIM(RTRIM(:pLicTradNum))) - 3);
				vNumero := REPLACE(REPLACE(REPLACE(REPLACE(:vNumero, '.', ''), ',', ''), '-', ''), '+', '');	
			END IF;
		END IF;
		
		SELECT CASE WHEN SUBSTRING(:vNumero, 1, LENGTH(:vNumero) -1) = '' THEN 0 ELSE LOCATE_REGEXPR(START '([[:digit:]]{' || TO_VARCHAR(LENGTH(:vNumero) -1) || '})' IN SUBSTRING(:vNumero, 1, LENGTH(:vNumero) -1) GROUP 1) END INTO vEsNumero FROM DUMMY;
		
		IF :vLetraINI <> '' AND :vEsNumeroINI = 0 AND LENGTH(:vNumero) >= 8 AND LENGTH(:vNumero) <= 9 AND :vEsNumero = 1 THEN
			-- El formato del CIF es correcto
			vDigitoControl := RIGHT(:vNumero, 1);
			vAuxNum1 := SUBSTRING(:vNumero, 1, 7);
			vLetras := 'ABCDEFGHIJ';
			vAbcdario := 'ABCDEFGHJPQRSUVNW';
						
			IF LOCATE(:vAbcdario, :vLetra) > 0 THEN
				vSumaPares := 0;
				vSumaImpares := 0;
				vSumaTotal := 0;
				vI := 1;

				WHILE :vI <= 7 DO
					vAuxNum2 := CAST(SUBSTRING(:vAuxNum1, :vI, 1) AS INT);

					IF MOD(:vI, 2) = 0 THEN
						vSumaPares := :vSumaPares + :vAuxNum2;
					ELSE
						vAuxNum2 := :vAuxNum2 * 2;
											
						vJ := 1;

						WHILE :vJ <= LENGTH(CAST(:vAuxNum2 AS NVARCHAR(15))) DO
							vSumaImpares := :vSumaImpares + CAST(SUBSTRING(CAST(:vAuxNum2 AS NVARCHAR(15)), :vJ, 1) AS INT);

							vJ := :vJ + 1;
						END WHILE;
					END IF;
									
					vI := :vI + 1;
				END WHILE;
								
				vSumaTotal := :vSumaTotal + :vSumaPares + :vSumaImpares;
				vUnidades := CAST(SUBSTRING(CAST(:vSumaTotal AS NVARCHAR(15)), LENGTH(CAST(:vSumaTotal AS NVARCHAR(15))), 1) AS INT);

				IF :vUnidades <> 0 THEN
					vUnidades := 10 - :vUnidades;
				END IF;

				IF :vLetra = 'A' OR :vLetra = 'B' OR :vLetra = 'E' OR :vLetra = 'H' THEN
					IF :vDigitoControl = CAST(:vUnidades AS NVARCHAR(1)) THEN
						vDato := :vPais || :vLetra || :vNumero;
					END IF;
				ELSE
					IF :vLetra = 'K' OR :vLetra = 'P' OR :vLetra = 'Q' OR :vLetra = 'S' THEN
						IF :vDigitoControl = SUBSTRING(:vLetras, :vUnidades, 1) THEN
							vDato := :vPais || :vLetra || :vNumero;
						ELSE
							IF :vUnidades = 0 AND :vDigitoControl = 'J' THEN
								vDato = :vPais || 'J' || :vNumero;
							END IF;
						END IF;
					ELSE
						IF :vDigitoControl = CAST(:vUnidades AS NVARCHAR(1)) THEN
							vDato = :vPais || :vLetra || :vNumero;
						ELSE
							IF :vDigitoControl = SUBSTRING(:vLetras, :vUnidades, 1) THEN
								vDato = :vPais || :vLetra || :vNumero;
							ELSE
								IF :vUnidades = 0 AND :vDigitoControl = 'J' THEN
									vDato = :vPais || 'J' || :vNumero;
								END IF;
							END IF;
						END IF;
					END IF;
				END IF;					
				--SET vDato = @DigitoControl	+ ' ' + CAST(@Unidades AS NVARCHAR(1)) + ' ' + SUBSTRING(@Letras, @Unidades, 1)		
			END IF;
		ELSE
			SELECT CASE WHEN :vNumero = '' THEN 0 ELSE LOCATE_REGEXPR(START '([[:digit:]]{' || TO_VARCHAR(LENGTH(:vNumero)) || '})' IN :vNumero GROUP 1) END INTO vEsNumeroOTRO FROM DUMMY;
			
			IF :vLetraFIN <> '' AND :vEsNumeroFIN = 0 AND LENGTH(:vNumero) >= 7 AND LENGTH(:vNumero) <= 10 AND :vEsNumeroOTRO = 1 THEN
				-- El formato del NIF es correcto
				vLetras := 'TRWAGMYFPDXBNJZSQVHLCKE';

				va := 0;
				vNIF_CIF := CAST(:vNumero AS INT);
				vb := - 1;

				WHILE (:vb <> 0) DO
					vb := FLOOR(:vNIF_CIF / 24);
					vc := :vNIF_CIF - (24 * :vb);
					va := :va + :vc;
					vNIF_CIF := :vb;
				END WHILE;

				vb := FLOOR(:va / 23);
    			vc := :va - (23 * :vb);
				
				vDato := :vPais || :vNumero || SUBSTRING(:vLetras, :vc + 1, 1);
			END IF;
		END IF;
	END IF;
	
	IF :pLicTradNum <> '' AND :pLicTradNum = :vDato THEN
		pResultado := 1;
	ELSE
		pResultado := 0;
	END IF;
END;

