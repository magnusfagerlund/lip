-- Written by: Fredrik Eriksson
-- Created: 2015-04-17

CREATE PROCEDURE [dbo].[csp_lip_createtable]
	@@tablename NVARCHAR(64)
	, @@localnamesingularenus NVARCHAR(512)
	, @@localnamesingularsv NVARCHAR(512) = @@localnamesingularenus
	, @@localnamesingularno NVARCHAR(512) = @@localnamesingularenus
	, @@localnamesingularda NVARCHAR(512) = @@localnamesingularenus
	, @@localnamesingularfi NVARCHAR(512) = @@localnamesingularenus
	, @@localnamepluralenus NVARCHAR(512)
	, @@localnamepluralsv NVARCHAR(512) = @@localnamepluralenus
	, @@localnamepluralno NVARCHAR(512) = @@localnamepluralenus
	, @@localnamepluralda NVARCHAR(512) = @@localnamepluralenus
	, @@localnamepluralfi NVARCHAR(512) = @@localnamepluralenus
	, @@idtable INT OUTPUT
AS
BEGIN

	-- FLAG_EXTERNALACCESS --
	
	DECLARE	@return_value INT
	DECLARE	@idstringlocalname INT
	DECLARE	@idstring INT
	DECLARE	@transid UNIQUEIDENTIFIER
	DECLARE	@descriptive INT
		
	SET @idstringlocalname = NULL
	SET @idstring = NULL
	SET @@idtable = NULL
	SET @transid = NEWID()
	SET @descriptive = NULL
	
	EXEC [dbo].[lsp_addtable]
		@@name = @@tablename
		, @@idtable = @@idtable OUTPUT
		, @@localname = @idstringlocalname OUTPUT
		, @@descriptive = @descriptive OUTPUT
		, @@transactionid = @transid
		, @@user = 1

	-- Set local name
	UPDATE [string]
	SET en_us = @@localnamesingularenus
		, sv = @@localnamesingularsv
		, [no] = @@localnamesingularno
		, da = @@localnamesingularda
		, fi = @@localnamesingularfi
	WHERE [idstring] = @idstringlocalname

	-- Set local name plural
	SET @idstring = NULL
	EXEC dbo.lsp_addstring
		@@idcategory = 17
		, @@string = @@localnamepluralsv
		, @@lang = 'sv'
		, @@idstring = @idstring OUTPUT

	EXEC dbo.lsp_setstring
		@@idstring = @idstring
		, @@lang = N'en_us'
		, @@string = @@localnamepluralenus
		
	EXEC dbo.lsp_setstring
		@@idstring = @idstring
		, @@lang = N'no'
		, @@string = @@localnamepluralno
		
	EXEC dbo.lsp_setstring
		@@idstring = @idstring
		, @@lang = N'da'
		, @@string = @@localnamepluralda
		
	EXEC dbo.lsp_setstring
		@@idstring = @idstring
		, @@lang = N'fi'
		, @@string = @@localnamepluralfi

	EXEC lsp_addattributedata
		@@owner	= N'table',
		@@idrecord = @@idtable,
		@@idrecord2 = NULL,
		@@name = N'localnameplural',
		@@value	=  @idstring
	
	-- Refresh ldc to make sure table is visible in LIME later on
	EXEC lsp_refreshldc
END
