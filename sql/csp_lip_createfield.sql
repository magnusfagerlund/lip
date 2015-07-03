SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- Written by: Fredrik Eriksson
-- Created: 2015-04-16

CREATE PROCEDURE [dbo].[csp_lip_createfield]
	@@tablename NVARCHAR(64)
	, @@fieldname NVARCHAR(64)
	, @@localname NVARCHAR(2048)
	, @@separator NVARCHAR(2048) = N''
	, @@type NVARCHAR(64)
	, @@defaultvalue NVARCHAR(64) = N''
	, @@limedefaultvalue NVARCHAR(64) = N''
	, @@limereadonly INT = 0
	, @@invisible INT = 0
	, @@required INT = 0
	, @@width INT = NULL
	, @@height INT = NULL
	, @@length INT = NULL
	, @@fieldorder INT = 0 -- Default value 0 means it will be put last
	, @@isnullable INT = 0
	, @@errorMessage NVARCHAR(512) OUTPUT
	, @@idfield INT OUTPUT
AS
BEGIN

	-- FLAG_EXTERNALACCESS --

	DECLARE	@return_value INT
	DECLARE @idstringlocalname INT
	DECLARE @idcategory INT
	DECLARE @idstring INT
	DECLARE @idfieldtype INT
	DECLARE @count INT
	DECLARE @sql NVARCHAR(300)
	DECLARE @currentPosition INT
	DECLARE @nextOccurance	 INT
	DECLARE @currentString NVARCHAR(256)
	DECLARE @currentLanguage NVARCHAR(8)
	DECLARE @currentLocalize NVARCHAR(256)

	SET @return_value = NULL
	SET @@idfield = NULL
	SET @idstringlocalname = NULL
	SET @idcategory = NULL
	SET @idstring = NULL
	SET @@errorMessage = N''
	SET @sql = N''
	
	--Check if field already exists
	EXEC lsp_getfield @@table = @@tablename, @@name = @@fieldname, @@count = @count OUTPUT
	
	IF  @count > 0 --Fieldname already exists
	BEGIN
		SET @@errorMessage = N'Field ''' + @@fieldname + N''' already exists. Please verify that properties for the field are correct.'
	END
	ELSE --Field doesn't exist
	BEGIN
		-- Get field type
		SELECT @idfieldtype = idfieldtype
		FROM fieldtype
		WHERE name = @@type
			AND active = 1
			AND creatable = 1

		EXEC @return_value = [dbo].[lsp_addfield]
			@@table = @@tablename
			,@@name = @@fieldname
			,@@fieldtype = @idfieldtype
			,@@length = @@length
			,@@isnullable = @@isnullable
			,@@defaultvalue = @@defaultvalue OUTPUT
			,@@idfield = @@idfield OUTPUT
			,@@localname = @idstringlocalname OUTPUT
			,@@idcategory = @idcategory OUTPUT
			
		--If return value is not 0, something went wrong and the field wasn't created
		IF @return_value <> 0
		BEGIN
			SET @@errorMessage = N'Field ''' + @@fieldname + N''' couldn''t be created'
		END
		ELSE
		BEGIN
			SET @return_value = 0

			--Make sure @@localname ends with ; in order to avoid infinite loop
			IF RIGHT(@@localname, 1) <> N';'
			BEGIN
				SET @@localname=@@localname + N';'
			END
			
			SET @currentPosition = 0
			--Loop through localnames
			WHILE @currentPosition <= LEN(@@localname) AND @return_value = 0
			BEGIN
				SET @nextOccurance = CHARINDEX(';', @@localname, @currentPosition)
				IF @nextOccurance <> 0
				BEGIN
					SET @sql = N''
					SET @currentString = SUBSTRING(@@localname, @currentPosition, @nextOccurance - @currentPosition)
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
			
			--Set limereadonly attribute
			EXEC @return_value = [dbo].[lsp_setfieldattributevalue] @@idfield = @@idfield, @@name = N'limereadonly', @@valueint = @@limereadonly
			
			--Set Default value (interpreted by LIME)
			EXEC @return_value = [dbo].[lsp_setattributevalue] @@owner = N'field', @@idrecord = @@idfield, @@name = N'limedefaultvalue', @@value = @@limedefaultvalue	-- Default Value (interpreted by LIME Pro) 
			
			--Set invisible/visible
			EXEC @return_value = [dbo].[lsp_setattributevalue] @@owner = N'field', @@idrecord = @@idfield, @@name = N'invisible', @@valueint = @@invisible
			
			--Set required attribute
			EXEC @return_value = [dbo].[lsp_setfieldattributevalue] @@idfield = @@idfield, @@name = N'required', @@valueint = @@required
			
			--Set width
			IF @@width IS NOT NULL
			BEGIN
				EXEC @return_value = [dbo].[lsp_setfieldattributevalue] @@idfield = @@idfield, @@name = N'width', @@valueint = @@width
			END
			
			--Set height
			IF @@height IS NOT NULL
			BEGIN
				EXEC @return_value = [dbo].[lsp_setfieldattributevalue] @@idfield = @@idfield, @@name = N'height', @@valueint = @@height
			END
			
			--Set fieldorder, if not provided we use default value 0 which means it will be put last
			EXEC @return_value = [dbo].[lsp_setfieldorder] @@idfield = @@idfield, @@fieldorder = @@fieldorder
			
			--Create separator
			IF @@separator <> N''
			BEGIN
				SET @idstring = -1
				EXEC @return_value = [dbo].[lsp_setattributevalue] @@owner = N'field'
							, @@idrecord = @@idfield
							, @@name = 'separator'
							, @@value = 1
				EXEC @return_value = [dbo].[lsp_setattributevalue] @@owner = N'field'
											, @@idrecord = @@idfield
											, @@name = N'separatorlocalname'
											, @@value = @idstring output
											
				--Make sure @@localname ends with ; in order to avoid infinite loop
				IF RIGHT(@@separator, 1) <> N';'
				BEGIN
					SET @@separator=@@separator + N';'
				END
				
				SET @currentPosition = 0
				
				--Loop through localnames
				WHILE @currentPosition <= LEN(@@separator) AND @return_value = 0
				BEGIN
					SET @nextOccurance = CHARINDEX(';', @@separator, @currentPosition)
					IF @nextOccurance <> 0
					BEGIN
						SET @currentString = SUBSTRING(@@separator, @currentPosition, @nextOccurance - @currentPosition)
						SET @currentLanguage=SUBSTRING(@currentString,0,CHARINDEX(':', @currentString))
						SET @currentLocalize=SUBSTRING(@currentString,CHARINDEX(':', @currentString)+1,LEN(@currentString)-CHARINDEX(':', @currentString))
						EXEC @return_value = [dbo].[lsp_setattributevalue] @@owner = N'string'
										, @@idrecord = @idstring
										, @@name = @currentLanguage
										, @@value = @currentLocalize
						SET @currentPosition = @nextOccurance+1
					END
				END								
			END
			--End of creating separator
			
			-- Refresh ldc to make sure field is visible in LIME later on
			EXEC lsp_refreshldc
			
			--If return value is not 0, something went wrong while setting field attributes
			IF @return_value <> 0
			BEGIN
				SET @@errorMessage = N'Something went wrong while setting attributes for field ''' + @@fieldname + N'''. Please check that field properties are correct.'
			END
			
		END
	END	
END
