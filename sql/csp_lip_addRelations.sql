-- Written by: Jonny Springare
-- Created: 2015-12-18
IF EXISTS (SELECT name FROM sysobjects WHERE name = 'csp_lip_addRelations' AND UPPER(type) = 'P')
   DROP PROCEDURE [csp_lip_addRelations]
GO
CREATE PROCEDURE [dbo].[csp_lip_addRelations]
	@@table1 NVARCHAR(64)
	, @@field1 NVARCHAR(64) = NULL
	, @@table2 NVARCHAR(64)
	, @@field2 NVARCHAR(64) = NULL
	, @@createdfields NVARCHAR(MAX)
	, @@errorMessage NVARCHAR(MAX) OUTPUT
	, @@warningMessage NVARCHAR(MAX) OUTPUT
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
	
	DECLARE	@return_value INT
	DECLARE @linebreak NVARCHAR(2)
	SET @linebreak = CHAR(13) + CHAR(10)
	SET @@errorMessage = N''
	SET @@warningMessage = N''
	
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
			SET @@createdfields = N';' + @@createdfields
			--Check if we have created the fields during this installation
			IF CHARINDEX(N';' + CONVERT(nvarchar(max), @idfield1) + N';', @@createdfields) > 0 AND CHARINDEX(N';' + CONVERT(nvarchar(max), @idfield2) + N';', @@createdfields) > 0
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
									--Check if the fields are already in a relation
									IF (SELECT relatedidfield FROM relationfieldview WHERE idfield=@idfield1) IS NULL
									BEGIN
										IF (SELECT relatedidfield FROM relationfieldview WHERE idfield=@idfield2) IS NULL
										BEGIN
											EXEC @return_value = lsp_addrelation
													@@idfield1 = @idfield1,
													@@idtable1 = @idtable1,
													@@idfield2 = @idfield2,
													@@idtable2 = @idtable2
											IF @return_value <> 0
											BEGIN
												SET @@errorMessage = N'ERROR: couldn''t create relation between' + @@table2 + '.' + @@field2 + N' and ' + @@table1 + '.' + @@field1
											END
										END
										ELSE
										BEGIN
											SET @@errorMessage = N'ERROR: ' + @@table2 + '.' + @@field2 + N' is already in a relation, relation to ' + @@table1 + '.' + @@field1 + N' can''t be created.'
										END
									END
									ELSE
									BEGIN
										IF (SELECT relatedidfield FROM relationfieldview WHERE idfield=@idfield1) = @idfield2
										BEGIN
											SET @@warningMessage = N'Warning: Relation between ' + @@table1 + '.' + @@field1 + ' and ' + @@table2 + '.' + @@field2 + N' already exists.'
										END
										ELSE
										BEGIN
											SET @@errorMessage = N'ERROR: ' + @@table1 + '.' + @@field1 + N' is already in a relation, relation to ' + @@table2 + '.' + @@field2 + N' can''t be created.'
										END
									END
								END
								ELSE
								BEGIN
									SET @@errorMessage = N'ERROR: ' + @@table2 + '.' + @@field2 + N' does not exist in table ''relationfield'', relation to ' + @@table1 + '.' + @@field1 + N' can''t be created.'
								END
							END
							ELSE
							BEGIN
								SET @@errorMessage = N'ERROR: ' + @@table1 + '.' + @@field1 + N' does not exist in table ''relationfield'', relation to ' + @@table2 + '.' + @@field2 + N' can''t be created.'
							END
						END
						ELSE
						BEGIN
							SET @@errorMessage = N'ERROR: ' + @@table2 + '.' + @@field2 + N' is not a relationfield/tab, relation to ' + @@table1 + '.' + @@field1 + N' can''t be created.'
							RETURN
						END
				END
				ELSE
				BEGIN
					SET @@errorMessage = N'ERROR: ' + @@table1 + '.' + @@field1 + N' is not a relationfield/tab, relation to ' + @@table2 + '.' + @@field2 + N' can''t be created.'
					RETURN
				END
			END
			ELSE
			BEGIN
				SET @@errorMessage = N'ERROR: Cannot create relation between ' + @@table1 + '.' + @@field1 + N' and ' + @@table2 + '.' + @@field2 + N', since one of the fields already existed before installation.'
				RETURN
			END
		END
		ELSE
		BEGIN
			SET @@errorMessage = N'ERROR: ' + @@table2 + '.' + @@field2 + N' hasn''t been created during this installation, relation to ' + @@table1 + '.' + @@field1 + N' can''t be created.'
			RETURN
		END
	END
	ELSE
	BEGIN
		SET @@errorMessage = N'ERROR: ' + @@table1 + '.' + @@field1 + N' hasn''t been created during this installation, relation to ' + @@table2 + '.' + @@field2 + N' can''t be created.'
		RETURN
	END	
END
