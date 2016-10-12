-- Written by: Jonny Springare
-- Created: 2016-10-12

CREATE PROCEDURE [dbo].[csp_lip_endInstallation]
AS
BEGIN
	-- FLAG_EXTERNALACCESS --
	
	IF ( OBJECT_ID('lsp_setdatabasetimestamp') > 0 )
	BEGIN
		EXEC lsp_setdatabasetimestamp
	END
	ELSE
	BEGIN
		EXEC lsp_refreshldc
	END
END
