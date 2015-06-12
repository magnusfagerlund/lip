SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- Written by: Fredrik Eriksson
-- Created: 2015-04-16

CREATE PROCEDURE [dbo].[csp_lip_createfield]
	@@tablename NVARCHAR(64)
	, @@fieldname NVARCHAR(64)
	, @@localnamesv NVARCHAR(512) = N''
	, @@localnameenus NVARCHAR(512) = N''
	, @@localnameno NVARCHAR(512) = N''
	, @@localnameda NVARCHAR(512) = N''
	, @@localnamefi NVARCHAR(512) = N''
	, @@type NVARCHAR(64) = N''
	, @@defaultvalue NVARCHAR(64) = N''
	, @@idfield INT OUTPUT
AS
BEGIN

	-- FLAG_EXTERNALACCESS --

	DECLARE	@return_value INT
	DECLARE @idstringlocalname INT
	DECLARE @idcategory INT
	DECLARE @idstring INT
	DECLARE @idfieldtype INT

	SET @return_value = NULL
	SET @@idfield = NULL
	SET @idstringlocalname = NULL
	SET @idcategory = NULL
	SET @idstring = NULL

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
		
	-- Check if all localnames has been set. Otherwise, replace with localname en_us
	SELECT @@localnamesv = CASE @@localnamesv WHEN '' THEN @@localnameenus ELSE @@localnamesv END
	SELECT @@localnameno = CASE @@localnameno WHEN '' THEN @@localnameenus ELSE @@localnameno END
	SELECT @@localnameda = CASE @@localnameda WHEN '' THEN @@localnameenus ELSE @@localnameda END
	SELECT @@localnamefi = CASE @@localnamefi WHEN '' THEN @@localnameenus ELSE @@localnamefi END
			
	UPDATE [string]
	SET sv = @@localnamesv
		, en_us = @@localnameenus
		, no = @@localnameno
		, da = @@localnameda
		, fi = @@localnamefi
	WHERE [idstring] = @idstringlocalname
	
	
	-- Refresh ldc to make sure field is visible in LIME later on
	EXEC lsp_refreshldc
	
END
