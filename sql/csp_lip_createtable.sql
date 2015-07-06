-- Written by: Fredrik Eriksson, Jonny Springare
-- Created: 2015-04-17

CREATE PROCEDURE [dbo].[csp_lip_createtable]
	@@tablename NVARCHAR(64)
	, @@localname_singular NVARCHAR(MAX)
	, @@localname_plural NVARCHAR(MAX)
	, @@errorMessage NVARCHAR(512) OUTPUT
	, @@idtable INT OUTPUT
	, @@iddescriptiveexpression INT OUTPUT
AS
BEGIN
	-- FLAG_EXTERNALACCESS --
	
	DECLARE	@return_value INT
	DECLARE	@idstringlocalname INT
	DECLARE	@idstring INT
	DECLARE	@transid UNIQUEIDENTIFIER
	DECLARE @sql NVARCHAR(300)
	DECLARE @currentPosition INT
	DECLARE @nextOccurance	 INT
	DECLARE @currentString NVARCHAR(256)
	DECLARE @currentLanguage NVARCHAR(8)
	DECLARE @currentLocalize NVARCHAR(256)
	DECLARE @isFirstLocalize BIT
	DECLARE @count INT
	
	SET @return_value =  NULL
	SET @idstringlocalname = NULL
	SET @idstring = NULL
	SET @@idtable = NULL
	SET @transid = NEWID()
	SET @@iddescriptiveexpression = NULL
	SET @sql = N''
	SET @isFirstLocalize = 1
	SET @@errorMessage = N''
	
	--Check if table already exists
	EXEC lsp_gettable @@name = @@tablename, @@count = @count OUTPUT
	
	IF  @count > 0 --Tablename already exists
	BEGIN
		SET @@idtable = -1
		SET @@iddescriptiveexpression = -1
		SET @@errorMessage = N'Table ''' + @@tablename + N''' already exists. Please verify that attributes for the table are correct.'
	END
	ELSE
	BEGIN
		EXEC @return_value = [dbo].[lsp_addtable]
			@@name = @@tablename
			, @@idtable = @@idtable OUTPUT
			, @@localname = @idstringlocalname OUTPUT
			, @@descriptive = @@iddescriptiveexpression OUTPUT
			, @@transactionid = @transid
			, @@user = 1
			
		-- Refresh ldc to make sure table is visible in LIME later on
		EXEC lsp_refreshldc

		--If return value is not 0, something went wrong and the table wasn't created
		IF @return_value <> 0
		BEGIN
			SET @@idtable = -1
			SET @@iddescriptiveexpression = -1
			SET @@errorMessage = N'Table ''' + @@tablename + N''' couldn''t be created'
		END
		ELSE
		BEGIN

			--Set localnames singular
			--Make sure @@localname_singular ends with ; in order to avoid infinite loop
			IF RIGHT(@@localname_singular, 1) <> N';'
			BEGIN
				SET @@localname_singular=@@localname_singular + N';'
			END
			
			SET @currentPosition = 0
			--Loop through localnames
			WHILE @currentPosition <= LEN(@@localname_singular) AND @return_value = 0
			BEGIN
				SET @nextOccurance = CHARINDEX(';', @@localname_singular, @currentPosition)
				IF @nextOccurance <> 0
				BEGIN
					SET @sql = N''
					SET @currentString = SUBSTRING(@@localname_singular, @currentPosition, @nextOccurance - @currentPosition)
					SET @currentLanguage=SUBSTRING(@currentString,0,CHARINDEX(':', @currentString))
					SET @currentLocalize=SUBSTRING(@currentString,CHARINDEX(':', @currentString)+1,LEN(@currentString)-CHARINDEX(':', @currentString))
					
					--Set local names for field
					SET @sql = N'UPDATE [string] 
					SET [' + @currentLanguage + N'] = ''' + @currentLocalize + N''''
					+ N' WHERE [idstring] = ' + CONVERT(NVARCHAR(12),@idstringlocalname)
					EXEC sp_executesql @sql
					
					SET @currentPosition = @nextOccurance+1
				END
			END
			--End localnames singular
			
			--Set localnames plural
			--Make sure @@localname_plural ends with ; in order to avoid infinite loop
			SET @currentPosition=0
			IF RIGHT(@@localname_plural, 1) <> N';'
			BEGIN
				SET @@localname_plural=@@localname_plural + N';'
			END
			
			SET @currentPosition = 0
			--Loop through localnames
			WHILE @currentPosition <= LEN(@@localname_plural) AND @return_value = 0
			BEGIN
				SET @nextOccurance = CHARINDEX(';', @@localname_plural, @currentPosition)
				IF @nextOccurance <> 0
				BEGIN
					SET @currentString = SUBSTRING(@@localname_plural, @currentPosition, @nextOccurance - @currentPosition)
					SET @currentLanguage=SUBSTRING(@currentString,0,CHARINDEX(':', @currentString))
					SET @currentLocalize=SUBSTRING(@currentString,CHARINDEX(':', @currentString)+1,LEN(@currentString)-CHARINDEX(':', @currentString))
					
					IF @isFirstLocalize = 1
					BEGIN
						EXEC @return_value = [dbo].[lsp_addstring]
							@@idcategory = 17
							, @@string = @currentLocalize
							, @@lang = @currentLanguage
							, @@idstring = @idstring OUTPUT
						SET @isFirstLocalize = 0
					END
					ELSE
					BEGIN
						EXEC @return_value = dbo.lsp_setstring
							@@idstring = @idstring
							, @@lang = @currentLanguage
							, @@string = @currentLocalize
					END
					
					SET @currentPosition = @nextOccurance+1
				END
			END

			EXEC @return_value = lsp_addattributedata
				@@owner	= N'table',
				@@idrecord = @@idtable,
				@@idrecord2 = NULL,
				@@name = N'localnameplural',
				@@value	=  @idstring
			--End localnames plural
			
			EXEC lsp_refreshldc
			
			--If return value is not 0, something went wrong while setting table attributes
			IF @return_value <> 0
			BEGIN
				SET @@errorMessage = N'Something went wrong while setting localnames for table ''' + @@tablename + N'''. Please check that table properties are correct.'
			END
		END
	END
END
