SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- Written by: Fredrik Eriksson
-- Created: 2015-04-16

CREATE PROCEDURE [dbo].[csp_lip_createfield]
	@@tablename NVARCHAR(64)
	, @@fieldname NVARCHAR(64)
	, @@localnameenus NVARCHAR(512)
	, @@localnamesv NVARCHAR(512) = @@localnameenus
	, @@localnameno NVARCHAR(512) = @@localnameenus
	, @@localnameda NVARCHAR(512) = @@localnameenus
	, @@localnamefi NVARCHAR(512) = @@localnameenus
	, @@type NVARCHAR(64)
	, @@defaultvalue NVARCHAR(64) = N''
	, @@limedefaultvalue NVARCHAR(64) = N''
	, @@limereadonly INT = 0
	, @@invisible INT = 0
	, @@required INT = 0
	, @@width INT = NULL
	, @@height INT = NULL
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

	SET @return_value = NULL
	SET @@idfield = NULL
	SET @idstringlocalname = NULL
	SET @idcategory = NULL
	SET @idstring = NULL
	SET @@errorMessage = N''
	
	--Check if field already exists
	EXEC lsp_getfield @@table = @@tablename, @@name = @@fieldname, @@count = @count OUTPUT
	
	IF  @count> 0 --Fieldname already exists
	BEGIN
		SET @@errorMessage = N'Field ''' + @@fieldname + N''' already exists. Verify that properties for the field are correct.'
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
			@@table = @@tablename,
			@@name = @@fieldname,
			@@fieldtype = @idfieldtype,
			@@defaultvalue = @@defaultvalue OUTPUT,
			@@idfield = @@idfield OUTPUT,
			@@localname = @idstringlocalname OUTPUT,
			@@idcategory = @idcategory OUTPUT
			
		--If return value is not 0, something went wrong and the field wasn't created
		IF @return_value <> 0
		BEGIN
			SET @@errorMessage = N'Field ''' + @@fieldname + N''' couldn''t be created'
		END
		ELSE
		BEGIN
			SET @return_value = 0
			UPDATE [string]
			SET sv = @@localnamesv
				, en_us = @@localnameenus
				, no = @@localnameno
				, da = @@localnameda
				, fi = @@localnamefi
			WHERE [idstring] = @idstringlocalname
			
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
