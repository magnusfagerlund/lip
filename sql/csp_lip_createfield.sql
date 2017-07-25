IF EXISTS (SELECT name FROM sysobjects WHERE name = 'csp_lip_createfield' AND UPPER(type) = 'P')
   DROP PROCEDURE [csp_lip_createfield]
GO
-- Written by: Fredrik Eriksson, Jonny Springare
-- Created: 2015-04-16
-- Last updated: 2017-03-23
CREATE PROCEDURE [dbo].[csp_lip_createfield]
	@@tablename NVARCHAR(64)
	, @@fieldname NVARCHAR(64)
	, @@fieldtype NVARCHAR(64)
	, @@defaultvalue NVARCHAR(125) = NULL --defaultvalue in lsp_addfield takes 128 chars, but only 125 can be set
	, @@length INT = NULL
	, @@isnullable INT = 0
	, @@errorMessage NVARCHAR(MAX) OUTPUT
	, @@warningMessage NVARCHAR(MAX) OUTPUT
	, @@idfield INT OUTPUT --idfield is set to -1 if field already exists
	, @@idcategory INT OUTPUT
	, @@idstringlocalname INT OUTPUT
AS
BEGIN
	-- FLAG_EXTERNALACCESS --

	DECLARE	@return_value INT
	DECLARE @idfieldtype INT
	DECLARE @count INT
	DECLARE @supportedFieldtypes NVARCHAR(MAX)
	DECLARE @linebreak NVARCHAR(2)
	DECLARE @existingFieldtype INT
	
	SET @return_value = NULL
	SET @@idfield = NULL
	SET @@idstringlocalname = NULL
	SET @@idcategory = NULL
	SET @@errorMessage = N''
	SET @@warningMessage = N''
	SET @linebreak = CHAR(13) + CHAR(10)
	SET @supportedFieldtypes = N'string;integer;decimal;time;link;yesno;set;option;formatedstring;color;relation;xml;file;sql;geography;html'
	--Not supported: user
	
	--Make sure @@length is set to NULL if fieldtype is not string
	IF @@fieldtype <> N'string' AND @@length IS NOT NULL
	BEGIN
		SET @@length = NULL
	END
	
	-- Get field type
	SELECT @idfieldtype = idfieldtype
	FROM fieldtype
	WHERE name = @@fieldtype
		AND active = 1
		AND creatable = 1
	
	--Check if field already exists
	EXEC lsp_getfield @@table = @@tablename, @@name = @@fieldname, @@fieldtype = @existingFieldtype OUTPUT, @@count = @count OUTPUT
	
	IF  @count > 0 --Fieldname already exists
	BEGIN
		SET @@idfield = -1
		IF @idfieldtype <> @existingFieldtype
		BEGIN
			SET @@warningMessage = @@warningMessage + N'Warning: field ''' + @@fieldname + N''' already exists and will not be re-created. Existing field''s type DID NOT match the fieldtype in the package.' + @linebreak
		END
		ELSE
		BEGIN
			SET @@warningMessage = @@warningMessage + N'Warning: field ''' + @@fieldname + N''' already exists and will not be re-created. Please verify that properties for the field are correct.' + @linebreak
		END
	END
	ELSE --Field doesn't exist
	BEGIN
		--Check if fieldtype exists
		IF (SELECT COUNT(*) FROM fieldtype WHERE name = @@fieldtype AND active = 1 AND creatable = 1) <> 1
		BEGIN
			SET @@errorMessage = @@errorMessage +  N'ERROR: ''' + @@fieldtype + N''' is not a valid fieldtype. Field ''' + @@fieldname + ''' couldn''t be created' + @linebreak
		END
		ELSE
		BEGIN
			--Check if fieldtype is implemented in LIP
			IF CHARINDEX(@@fieldtype, @supportedFieldtypes) = 0
			BEGIN
				SET @@errorMessage = @@errorMessage +  N'ERROR: fieldtype ''' + @@fieldtype + N''' is not implemented in LIP. Field ''' + @@fieldname + ''' couldn''t be created' + @linebreak
			END
			ELSE
			BEGIN				
				--Don't pass defaultvalue for option- and setfields
				IF @@fieldtype IN (N'option', N'set')
				BEGIN
					EXEC @return_value = [dbo].[lsp_addfield]
						@@table = @@tablename
						,@@name = @@fieldname
						,@@fieldtype = @idfieldtype
						,@@length = @@length
						,@@isnullable = @@isnullable
						,@@idfield = @@idfield OUTPUT
						,@@localname = @@idstringlocalname OUTPUT
						,@@idcategory = @@idcategory OUTPUT
				END
				ELSE
				BEGIN
					EXEC @return_value = [dbo].[lsp_addfield]
						@@table = @@tablename
						,@@name = @@fieldname
						,@@fieldtype = @idfieldtype
						,@@length = @@length
						,@@isnullable = @@isnullable
						,@@defaultvalue = @@defaultvalue OUTPUT
						,@@idfield = @@idfield OUTPUT
						,@@localname = @@idstringlocalname OUTPUT
						,@@idcategory = @@idcategory OUTPUT
				END
					
				--If return value is not 0, something went wrong and the field wasn't created
				IF @return_value <> 0
				BEGIN
					SET @@errorMessage = @@errorMessage + N'ERROR: field ''' + @@fieldname + N''' couldn''t be created' + @linebreak
				END
			END
		END
	END		
	SET @@idcategory = ISNULL(@@idcategory, 0)
END
