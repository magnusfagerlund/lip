-- Written by: Jonny Springare
-- Created: 2015-07-09

CREATE PROCEDURE [dbo].[csp_lip_installSQL]
	@@sql NVARCHAR(MAX)
AS
BEGIN
	-- FLAG_EXTERNALACCESS --
	exec sp_executesql @@sql
END
