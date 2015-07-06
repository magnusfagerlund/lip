SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- Written by: Jonny Springare
-- Created: 2015-07-06

CREATE PROCEDURE [dbo].[csp_lip_settableattributes]
	@@tablename NVARCHAR(64)
	, @@idtable INT
	, @@iddescriptiveexpression INT
	, @@descriptive NVARCHAR(MAX) = N''
	, @@tableorder INT = 0
	, @@invisible INT = 0 --Default value 0 means not invisible
	, @@errorMessage NVARCHAR(512) OUTPUT
AS
BEGIN
	-- FLAG_EXTERNALACCESS --
	DECLARE	@return_value INT

	SET @return_value =  NULL
	SET @@errorMessage = N''
	
	--Set descriptive expression
	EXEC @return_value = [dbo].[lsp_setstring]
		@@idstring = @@iddescriptiveexpression
		, @@lang = N'ALL'
		, @@string = @@descriptive
	
	--Set table order
	EXEC @return_value = [dbo].[lsp_setattributevalue]
		@@owner = N'table'
		, @@idrecord = @@idtable
		, @@name = N'tableorder'
		, @@valueint = @@tableorder
		
	--Set invisible
	IF @@invisible = 1 OR @@invisible = 2
	BEGIN
		EXEC @return_value = [dbo].[lsp_setattributevalue]
			@@owner = N'table'
			, @@idrecord = @@idtable
			, @@name = N'invisible'
			, @@valueint = @@invisible
	END
	
	EXEC lsp_refreshldc
	
	--If return value is not 0, something went wrong while setting table attributes
	IF @return_value <> 0
	BEGIN
		SET @@errorMessage = N'Something went wrong while setting attributes for table ''' + @@tablename + N'''. Please check that table properties are correct.'
	END
END
