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
	, @@syscomment NVARCHAR(MAX) = NULL
	, @@label INT = NULL
	, @@actionpad NVARCHAR(MAX) = NULL
	, @@log BIT = NULL
	, @@errorMessage NVARCHAR(MAX) OUTPUT
	, @@warningMessage NVARCHAR(MAX) OUTPUT
AS
BEGIN
	-- FLAG_EXTERNALACCESS --
	DECLARE	@return_value INT
	DECLARE @linebreak NVARCHAR(2)
	
	SET @return_value =  NULL
	SET @@errorMessage = N''
	SET @@warningMessage = N''
	SET @linebreak = CHAR(13) + CHAR(10)
	
	--Set descriptive expression
	EXEC @return_value = [dbo].[lsp_setstring]
		@@idstring = @@iddescriptiveexpression
		, @@lang = N'ALL'
		, @@string = @@descriptive
		
	IF @return_value <> 0
	BEGIN
		SET @@warningMessage = @@warningMessage + N'Warning: couldn''t set descriptive expression for table ''' + @@tablename  + @linebreak
	END
	
	--Set table order
	EXEC @return_value = [dbo].[lsp_setattributevalue]
		@@owner = N'table'
		, @@idrecord = @@idtable
		, @@name = N'tableorder'
		, @@valueint = @@tableorder
		
	IF @return_value <> 0
	BEGIN
		SET @@warningMessage = @@warningMessage + N'Warning: couldn''t set table order for table ''' + @@tablename  + @linebreak
	END
		
	--Set invisible
	IF @@invisible = 1 OR @@invisible = 2
	BEGIN
		EXEC @return_value = [dbo].[lsp_setattributevalue]
			@@owner = N'table'
			, @@idrecord = @@idtable
			, @@name = N'invisible'
			, @@valueint = @@invisible
		IF @return_value <> 0
		BEGIN
			SET @@warningMessage = @@warningMessage + N'Warning: couldn''t set invisible attribute for table ''' + @@tablename  + @linebreak
		END
	END
	
	--Set comment
	IF @@syscomment IS NOT NULL
	BEGIN
		EXEC @return_value = [dbo].[lsp_setattributevalue]
			@@owner = N'table'
			, @@idrecord = @@idtable
			, @@name = N'syscomment'
			, @@value = @@syscomment
			
		IF @return_value <> 0
		BEGIN
			SET @@warningMessage = @@warningMessage + N'Warning: couldn''t set syscomment for table ''' + @@tablename  + @linebreak
		END
	END
	
	--Set label
	IF @@label IS NOT NULL
	BEGIN
		EXEC @return_value = [dbo].[lsp_setattributevalue]
			@@owner = N'table'
			, @@idrecord = @@idtable
			, @@name = N'label'
			, @@valueint = @@label
		IF @return_value <> 0
		BEGIN
			SET @@warningMessage = @@warningMessage + N'Warning: couldn''t set attribute ''label'' for table ''' + @@tablename  + @linebreak
		END
	END
	
	--Set actionpad
	IF @@actionpad IS NOT NULL
	BEGIN
		EXEC @return_value = [dbo].[lsp_setattributevalue]
			@@owner = N'table'
			, @@idrecord = @@idtable
			, @@name = N'actionpad'
			, @@value = @@actionpad
		IF @return_value <> 0
		BEGIN
			SET @@warningMessage = @@warningMessage + N'Warning: couldn''t set actionpad for table ''' + @@tablename  + @linebreak
		END
	END
	
	--Set logging option
	IF @@log IS NOT NULL
	BEGIN
		EXEC @return_value = [dbo].[lsp_setattributevalue]
			@@owner = N'table'
			, @@idrecord = @@idtable
			, @@name = N'log'
			, @@value = @@log
		IF @return_value <> 0
		BEGIN
			SET @@warningMessage = @@warningMessage + N'Warning: couldn''t set log attribute for table ''' + @@tablename  + @linebreak
		END
	END
	
	EXEC lsp_setdatabasetimestamp
	EXEC lsp_refreshldc
	
END
