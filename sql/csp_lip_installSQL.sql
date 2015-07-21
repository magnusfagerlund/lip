SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- Written by: Jonny Springare
-- Created: 2015-07-09

CREATE PROCEDURE [dbo].[csp_lip_installSQL]
	@@sql NVARCHAR(MAX)
	, @@name NVARCHAR(128)
	, @@type NVARCHAR(128)
	, @@errormessage NVARCHAR(512) OUTPUT
AS
BEGIN
	-- FLAG_EXTERNALACCESS --
	SET @@errormessage = N''
	
	IF @@type = N'procedure'
	BEGIN
		--Check if stored procedure already exists and replace CREATE with ALTER then
		IF (SELECT COUNT(*) FROM dbo.[storedprocedureview] WHERE [name] = @@name) > 0
		BEGIN
			IF CHARINDEX(N'CREATE PROCEDURE', @@sql) > 0
			BEGIN
				SET @@sql = STUFF(@@sql, CHARINDEX(N'CREATE PROCEDURE', @@sql), LEN(N'CREATE PROCEDURE'), 'ALTER PROCEDURE')
			END
		END
	END
	ELSE IF @@type = N'function'
	BEGIN
		--Check if function already exists and replace CREATE with ALTER then
		IF (SELECT COUNT(*) FROM dbo.[functionview] WHERE [name] = @@name) > 0
		BEGIN
			IF CHARINDEX(N'CREATE FUNCTION', @@sql) > 0
			BEGIN
				SET @@sql = STUFF(@@sql, CHARINDEX(N'CREATE FUNCTION', @@sql), LEN(N'CREATE FUNCTION'), 'ALTER FUNCTION')
			END
		END
	END
	ELSE
	BEGIN
		SET @@errormessage= N'''' + @@type + N''' is not a valid SQL type for LIP.'
	END
	
	IF @@errormessage = N''
	BEGIN
		EXEC sp_executesql @@sql
		EXEC lsp_refreshldc
	END
END
