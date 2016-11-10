-- Written by: Jonny Springare
-- Created: 2016-10-12
IF EXISTS (SELECT name FROM sysobjects WHERE name = 'csp_lip_endInstallation' AND UPPER(type) = 'P')
   DROP PROCEDURE [csp_lip_endInstallation]
GO
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
