-- Written by: Jonny Springare
-- Created: 2015-12-18

CREATE PROCEDURE [dbo].[csp_lip_addRelations]
	@@table1 NVARCHAR(64)
	, @@field1 NVARCHAR(64) = NULL
	, @@table2 NVARCHAR(64)
	, @@field2 NVARCHAR(64) = NULL
	, @@errorMessage NVARCHAR(512) OUTPUT
AS
BEGIN

	-- FLAG_EXTERNALACCESS --
	DECLARE @idfield1 INT
	DECLARE @idtable1 INT
	DECLARE @fieldtype1 INT
	
	DECLARE @idfield2 INT
	DECLARE @idtable2 INT
	DECLARE @fieldtype2 INT
	
	DECLARE @fieldtypeRelation INT
	
	--Get id for fieldtype relation
	SELECT @fieldtypeRelation = idfieldtype
			FROM fieldtype
			WHERE name = N'relation'
				AND active = 1
				AND creatable = 1
	
	--Get id's
	EXEC lsp_getfield @@idfield=@idfield1 OUTPUT, @@name=@@field1, @@table=@@table1, @@fieldtype=@fieldtype1 OUTPUT
	EXEC lsp_gettable @@idtable=@idtable1 OUTPUT, @@name=@@table1
	
	EXEC lsp_getfield @@idfield=@idfield2 OUTPUT, @@name=@@field2, @@table=@@table2, @@fieldtype=@fieldtype2 OUTPUT
	EXEC lsp_gettable @@idtable=@idtable2 OUTPUT, @@name=@@table2
	
	--Check if fields exist
	IF @idfield1 IS NOT NULL
	BEGIN
		IF @idfield2 IS NOT NULL
		BEGIN
			--Check if the fields are relationfields
			IF @fieldtype1 = @fieldtypeRelation
			BEGIN
				IF @fieldtype2 = @fieldtypeRelation
					BEGIN
						--Check if the fields exist in table relationfield
						IF EXISTS (SELECT idrelationfield FROM relationfield WHERE idfield=@idfield1)
						BEGIN
							IF EXISTS (SELECT idrelationfield FROM relationfield WHERE idfield=@idfield1)
							BEGIN
								EXEC	lsp_addrelation
										@@idfield1 = @idfield1,
										@@idtable1 = @idtable1,
										@@idfield2 = @idfield2,
										@@idtable2 = @idtable2
							END
							ELSE
							BEGIN
								SET @@errorMessage = @@table2 + '.' + @@field2 + N' does not exist in table ''relationfield'', relation to ' + @@table1 + '.' + @@field1 + N' can''t be created.'
							END
						END
						ELSE
						BEGIN
							SET @@errorMessage = @@table1 + '.' + @@field1 + N' does not exist in table ''relationfield'', relation to ' + @@table2 + '.' + @@field2 + N' can''t be created.'
						END
					END
					ELSE
					BEGIN
						SET @@errorMessage = @@table2 + '.' + @@field2 + N' is not a relationfield/tab, relation to ' + @@table1 + '.' + @@field1 + N' can''t be created.'
						RETURN
					END
			END
			ELSE
			BEGIN
				SET @@errorMessage = @@table1 + '.' + @@field1 + N' is not a relationfield/tab, relation to ' + @@table2 + '.' + @@field2 + N' can''t be created.'
				RETURN
			END
		END
		ELSE
		BEGIN
			SET @@errorMessage = @@table2 + '.' + @@field2 + N' hasn''t been created, relation to ' + @@table1 + '.' + @@field1 + N' can''t be created.'
			RETURN
		END
	END
	ELSE
	BEGIN
		SET @@errorMessage = @@table1 + '.' + @@field1 + N' hasn''t been created, relation to ' + @@table2 + '.' + @@field2 + N' can''t be created.'
		RETURN
	END
	
END
