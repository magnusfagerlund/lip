SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- Written by: Jonny Springare
-- Created: 2015-07-09

CREATE PROCEDURE [dbo].[csp_lip_installSQL]
	@@sql NVARCHAR(MAX)
AS
BEGIN
	-- FLAG_EXTERNALACCESS --
	exec sp_executesql @@sql
END
