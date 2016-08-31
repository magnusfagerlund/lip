-- Written by: Fredrik Eriksson, Jonny Springare
-- Created: 2015-04-16

CREATE PROCEDURE [dbo].[csp_lip_createfield]
	@@tablename NVARCHAR(64)
	, @@fieldname NVARCHAR(64)
	, @@localname NVARCHAR(MAX)
	, @@separator NVARCHAR(MAX) = N''
	, @@description NVARCHAR(MAX) = N'' --Tooltip
	, @@optionlist NVARCHAR(MAX) = N''
	, @@fieldtype NVARCHAR(64)
	, @@defaultvalue NVARCHAR(64) = NULL
	, @@limedefaultvalue NVARCHAR(64) = NULL
	, @@limereadonly INT = NULL
	, @@invisible INT = NULL
	, @@required INT = NULL
	, @@limerequiredforedit INT = NULL
	, @@width INT = NULL
	, @@height INT = NULL
	, @@length INT = NULL
	, @@newline INT = 2 -- Default value 2 means Fixed width
	, @@sql NVARCHAR(MAX) = NULL
	, @@onsqlupdate NVARCHAR(MAX) = NULL
	, @@onsqlinsert NVARCHAR(MAX) = NULL
	, @@formatsql BIT = NULL
	, @@comment NVARCHAR(MAX) = N''
	, @@syscomment NVARCHAR(MAX) = NULL
	, @@limevalidationrule NVARCHAR(MAX) = NULL
	, @@limevalidationtext NVARCHAR(MAX) = N''
	, @@fieldorder INT = 0 -- Default value 0 means it will be put last
	, @@isnullable INT = 0
	, @@type INT = 0
	, @@label INT = NULL
	, @@adlabel INT = NULL
	, @@relationtab BIT = 0
	, @@errorMessage NVARCHAR(MAX) OUTPUT
	, @@warningMessage NVARCHAR(MAX) OUTPUT
	, @@idfield INT OUTPUT --idfield is set to -1 if field already exists
AS
BEGIN
	-- FLAG_EXTERNALACCESS --

	DECLARE	@return_value INT
	DECLARE @idstringlocalname INT
	DECLARE @idcategory INT
	DECLARE @idstring INT
	DECLARE @idfieldtype INT
	DECLARE @count INT
	DECLARE @sql NVARCHAR(300)
	DECLARE @currentPosition INT
	DECLARE @nextOccurance	 INT
	DECLARE @currentString NVARCHAR(256)
	DECLARE @currentLanguage NVARCHAR(11)
	DECLARE @currentLocalize NVARCHAR(256)
	DECLARE @isFirstLocalize BIT
	DECLARE @currentOption NVARCHAR(MAX)
	DECLARE @nextOptionStarts INT
	DECLARE @nextOptionEnds INT
	DECLARE @currentPositionInOption INT
	DECLARE @supportedFieldtypes NVARCHAR(MAX)
	DECLARE @linebreak NVARCHAR(2)
	
	SET @return_value = NULL
	SET @@idfield = NULL
	SET @idstringlocalname = NULL
	SET @idcategory = NULL
	SET @idstring = NULL
	SET @@errorMessage = N''
	SET @@warningMessage = N''
	SET @linebreak = CHAR(13) + CHAR(10)
	SET @sql = N''
	SET @isFirstLocalize = 1
	SET @currentOption =N''
	SET @nextOptionStarts = 0
	SET @nextOptionEnds = 0
	SET @supportedFieldtypes = N'string;integer;decimal;time;link;yesno;set;option;formatedstring;color;relation;xml;file;sql;geography;html'
	--Not supported: user
	
	--Make sure @@length is set to NULL if fieldtype is not string
	IF @@fieldtype <> N'string' AND @@length IS NOT NULL
	BEGIN
		SET @@length = NULL
	END
	
	--Check if field already exists
	EXEC lsp_getfield @@table = @@tablename, @@name = @@fieldname, @@count = @count OUTPUT
	
	IF  @count > 0 --Fieldname already exists
	BEGIN
		SET @@idfield = -1
		SET @@warningMessage = @@warningMessage + N'Warning: field ''' + @@fieldname + N''' already exists and will not be re-created. Please verify that properties for the field are correct.' + @linebreak
	END
	ELSE --Field doesn't exist
	BEGIN
		--Check if fieldtype exists
		IF (SELECT COUNT(*) FROM fieldtype WHERE name = @@fieldtype AND active = 1 AND creatable = 1) <> 1
		BEGIN
			SET @@errorMessage = @@errorMessage +  N'ERROR: ''' + @@fieldtype + N''' is not a valid fieldtype. Field ''' + @@fieldname + ''' couldn''t be created' + @linebreak
		END
		ELSE
		BEGIN
			--Check if fieldtype is implemented in LIP
			IF CHARINDEX(@@fieldtype, @supportedFieldtypes) = 0
			BEGIN
				SET @@errorMessage = @@errorMessage +  N'ERROR: fieldtype ''' + @@fieldtype + N''' is not implemented in LIP. Field ''' + @@fieldname + ''' couldn''t be created' + @linebreak
			END
			ELSE
			BEGIN
				-- Get field type
				SELECT @idfieldtype = idfieldtype
				FROM fieldtype
				WHERE name = @@fieldtype
					AND active = 1
					AND creatable = 1

				EXEC @return_value = [dbo].[lsp_addfield]
					@@table = @@tablename
					,@@name = @@fieldname
					,@@fieldtype = @idfieldtype
					,@@length = @@length
					,@@isnullable = @@isnullable
					,@@defaultvalue = @@defaultvalue OUTPUT
					,@@idfield = @@idfield OUTPUT
					,@@localname = @idstringlocalname OUTPUT
					,@@idcategory = @idcategory OUTPUT
					

				
				-- Refresh ldc to make sure field is visible in LIME later on
				EXEC lsp_setdatabasetimestamp
				EXEC lsp_refreshldc
					
				--If return value is not 0, something went wrong and the field wasn't created
				IF @return_value <> 0
				BEGIN
					SET @@errorMessage = @@errorMessage + N'ERROR: field ''' + @@fieldname + N''' couldn''t be created' + @linebreak
				END
				ELSE
				BEGIN
					SET @return_value = 0
					
					--Set idcategory for textfield or decimal-field, since this isn't done by lsp_addfield (only setfields and optionfields)
					IF @@fieldtype = N'string' OR @@fieldtype = N'decimal'
					BEGIN
						EXEC @return_value =  lsp_setfieldattributevalue @@idfield = @@idfield, 
														 @@name = N'idcategory',
														 @@valueint = @idcategory OUTPUT
						IF @return_value <> 0
						BEGIN
							SET @@warningMessage = @@warningMessage + N'Warning: couldn''t set idcategory for field ''' + @@fieldname  + @linebreak
							SET @return_value = 0
						END
					END

					--Make sure @@localname ends with ; in order to avoid infinite loop
					IF RIGHT(@@localname, 1) <> N';'
					BEGIN
						SET @@localname=@@localname + N';'
					END
					
					SET @currentPosition = 0
					--Loop through localnames
					WHILE @currentPosition <= LEN(@@localname) AND @return_value = 0
					BEGIN
						SET @nextOccurance = CHARINDEX(';', @@localname, @currentPosition)
						IF @nextOccurance <> 0
						BEGIN
							SET @sql = N''
							SET @currentString = SUBSTRING(@@localname, @currentPosition, @nextOccurance - @currentPosition)
							SET @currentLanguage=SUBSTRING(@currentString,0,CHARINDEX(':', @currentString))
							SET @currentLocalize=SUBSTRING(@currentString,CHARINDEX(':', @currentString)+1,LEN(@currentString)-CHARINDEX(':', @currentString))
							
							--Set local names for field
							SET @sql = N'UPDATE [string] 
							SET [' + @currentLanguage + N'] = ''' + @currentLocalize + N''''
							+ N' WHERE [idstring] = ' + CONVERT(NVARCHAR(12),@idstringlocalname)
							EXEC sp_executesql @sql
							
							SET @currentPosition = @nextOccurance+1
						END
					END	
					
					--Set limereadonly attribute
					IF @@limereadonly IS NOT NULL
					BEGIN
						EXEC @return_value = [dbo].[lsp_setfieldattributevalue] @@idfield = @@idfield, @@name = N'limereadonly', @@valueint = @@limereadonly
						IF @return_value <> 0
						BEGIN
							SET @@warningMessage = @@warningMessage + N'Warning: couldn''t set attribute limereadonly for field ''' + @@fieldname + @linebreak
							SET @return_value = 0
						END
					END
					
					--Set Default value (interpreted by LIME)
					IF @@limedefaultvalue IS NOT NULL
					BEGIN
						EXEC @return_value = [dbo].[lsp_setattributevalue] @@owner = N'field', @@idrecord = @@idfield, @@name = N'limedefaultvalue', @@value = @@limedefaultvalue	-- Default Value (interpreted by LIME Pro) 
						IF @return_value <> 0
						BEGIN
							SET @@warningMessage = @@warningMessage + N'Warning: couldn''t set limedefaultvalue for field ''' + @@fieldname + @linebreak
							SET @return_value = 0
						END
					END
					
					--Set invisible/visible
					IF @@invisible IS NOT NULL
					BEGIN
						EXEC @return_value = [dbo].[lsp_setattributevalue] @@owner = N'field', @@idrecord = @@idfield, @@name = N'invisible', @@valueint = @@invisible
						IF @return_value <> 0
						BEGIN
							SET @@warningMessage = @@warningMessage + N'Warning: couldn''t set attribute invisible for field ''' + @@fieldname + @linebreak
							SET @return_value = 0
						END
					END
					
					--Set required attribute
					IF @@required IS NOT NULL
					BEGIN
						EXEC @return_value = [dbo].[lsp_setfieldattributevalue] @@idfield = @@idfield, @@name = N'required', @@valueint = @@required
						IF @return_value <> 0
						BEGIN
							SET @@warningMessage = @@warningMessage + N'Warning: couldn''t set attribute required for field ''' + @@fieldname + @linebreak
							SET @return_value = 0
						END
					END
					--Set attribute Required for editing in Lime
					IF @@limerequiredforedit IS NOT NULL
					BEGIN
						EXEC @return_value = [dbo].[lsp_setfieldattributevalue] @@idfield = @@idfield, @@name = N'limerequiredforedit', @@valueint = @@limerequiredforedit
						IF @return_value <> 0
						BEGIN
							SET @@warningMessage = @@warningMessage + N'Warning: couldn''t set attribute limerequiredforedit for field ''' + @@fieldname + @linebreak
							SET @return_value = 0
						END
					END
					
					--Set width
					IF @@width IS NOT NULL
					BEGIN
						EXEC @return_value = [dbo].[lsp_setfieldattributevalue] @@idfield = @@idfield, @@name = N'width', @@valueint = @@width
						IF @return_value <> 0
						BEGIN
							SET @@warningMessage = @@warningMessage + N'Warning: couldn''t set width for field ''' + @@fieldname + @linebreak
							SET @return_value = 0
						END
					END
					
					--Set height
					IF @@height IS NOT NULL
					BEGIN
						EXEC @return_value = [dbo].[lsp_setfieldattributevalue] @@idfield = @@idfield, @@name = N'height', @@valueint = @@height
						IF @return_value <> 0
						BEGIN
							SET @@warningMessage = @@warningMessage + N'Warning: couldn''t set height for field ''' + @@fieldname + @linebreak
							SET @return_value = 0
						END
					END
					
					--Set width properties
					EXEC @return_value = [dbo].[lsp_setfieldattributevalue] @@idfield = @@idfield, @@name = N'newline', @@valueint = @@newline
					IF @return_value <> 0
					BEGIN
						SET @@warningMessage = @@warningMessage + N'Warning: couldn''t set attribute newline for field ''' + @@fieldname + @linebreak
						SET @return_value = 0
					END
					
					--Set SQL Expression
					IF @@sql IS NOT NULL
					BEGIN
						EXEC @return_value = [dbo].[lsp_setfieldattributevalue] @@idfield = @@idfield, @@name = N'sql', @@value = @@sql
						IF @return_value <> 0
						BEGIN
							SET @@warningMessage = @@warningMessage + N'Warning: couldn''t set SQL-expression for field ''' + @@fieldname + @linebreak
							SET @return_value = 0
						END
					END
					
					--Set SQL for update
					IF @@onsqlupdate IS NOT NULL
					BEGIN
						EXEC @return_value = [dbo].[lsp_setfieldattributevalue] @@idfield = @@idfield, @@name = N'onsqlupdate', @@value = @@onsqlupdate
						IF @return_value <> 0
						BEGIN
							SET @@warningMessage = @@warningMessage + N'Warning: couldn''t set attribute ''SQL for update'' for field ''' + @@fieldname + @linebreak
							SET @return_value = 0
						END
					END
					
					--Set SQL for new
					IF @@onsqlinsert IS NOT NULL
					BEGIN
						EXEC @return_value = [dbo].[lsp_setfieldattributevalue] @@idfield = @@idfield, @@name = N'onsqlinsert', @@value = @@onsqlinsert
						IF @return_value <> 0
						BEGIN
							SET @@warningMessage = @@warningMessage + N'Warning: couldn''t set attribute ''SQL for new'' for field ''' + @@fieldname + @linebreak
							SET @return_value = 0
						END
					END
					
					--Set formatsql
					IF @@formatsql IS NOT NULL
					BEGIN
						EXEC @return_value = [dbo].[lsp_setfieldattributevalue] @@idfield = @@idfield, @@name = N'formatsql', @@valueint = @@formatsql
						IF @return_value <> 0
						BEGIN
							SET @@warningMessage = @@warningMessage + N'Warning: couldn''t set attribute ''formatsql'' for field ''' + @@fieldname + @linebreak
							SET @return_value = 0
						END
					END
					
					--Set syscomment (private comment)
					IF @@syscomment IS NOT NULL
					BEGIN
						EXEC @return_value = [dbo].[lsp_setfieldattributevalue] @@idfield = @@idfield, @@name = N'syscomment', @@value = @@syscomment
						IF @return_value <> 0
						BEGIN
							SET @@warningMessage = @@warningMessage + N'Warning: couldn''t set attribute ''syscomment'' for field ''' + @@fieldname + @linebreak
							SET @return_value = 0
						END
					END
					
					--Set limevalidationrule
					IF @@limevalidationrule IS NOT NULL
					BEGIN
						EXEC @return_value = [dbo].[lsp_setfieldattributevalue] @@idfield = @@idfield, @@name = N'limevalidationrule', @@value = @@limevalidationrule
						IF @return_value <> 0
						BEGIN
							SET @@warningMessage = @@warningMessage + N'Warning: couldn''t set attribute ''limevalidationrule'' for field ''' + @@fieldname + @linebreak
							SET @return_value = 0
						END
					END
					
					--Set label
					IF @@label IS NOT NULL
					BEGIN
						EXEC @return_value = [dbo].[lsp_setfieldattributevalue] @@idfield = @@idfield, @@name = N'label', @@valueint = @@label
						IF @return_value <> 0
						BEGIN
							SET @@warningMessage = @@warningMessage + N'Warning: couldn''t set attribute ''label'' for field ''' + @@fieldname + @linebreak
							SET @return_value = 0
						END
					END
					
					--Set AD-label
					IF @@adlabel IS NOT NULL
					BEGIN
						EXEC @return_value = [dbo].[lsp_setfieldattributevalue] @@idfield = @@idfield, @@name = N'adlabel', @@valueint = @@adlabel
						IF @return_value <> 0
						BEGIN
							SET @@warningMessage = @@warningMessage + N'Warning: couldn''t set attribute ''adlabel'' for field ''' + @@fieldname + @linebreak
							SET @return_value = 0
						END
					END
					
					--Set fieldorder, if not provided we use default value 0 which means it will be put last
					EXEC @return_value = [dbo].[lsp_setfieldorder] @@idfield = @@idfield, @@fieldorder = @@fieldorder
						IF @return_value <> 0
						BEGIN
							SET @@warningMessage = @@warningMessage + N'Warning: couldn''t set fieldorder for field ''' + @@fieldname + @linebreak
							SET @return_value = 0
						END
					
					--Set relation properties, if relationfield
					IF @@fieldtype = N'relation'
					BEGIN
						EXEC @return_value = lsp_setattributevalue @@owner = 'field', @@idrecord = @@idfield, @@name = 'relationmincount', @@value = 0
						IF @return_value <> 0
						BEGIN
							SET @@errorMessage = @@errorMessage +  N'ERROR: couldn''t set attribute ''relationmincount'' for field ''' + @@fieldname + @linebreak
							SET @return_value = 0
						END
						IF @@relationtab = 1
						BEGIN
							EXEC lsp_setattributevalue @@owner = 'field', @@idrecord = @@idfield, @@name = 'relationmaxcount', @@value = 1
							IF @return_value <> 0
							BEGIN
								SET @@errorMessage = @@errorMessage +  N'ERROR: couldn''t set attribute ''relationmaxcount'' for field ''' + @@fieldname + @linebreak
								SET @return_value = 0
							END
						END	
					END
					
					--Create separator
					SET @idstring = -1
					EXEC @return_value = [dbo].[lsp_setattributevalue] @@owner = N'field'
												, @@idrecord = @@idfield
												, @@name = N'separatorlocalname'
												, @@value = @idstring output							
					IF @return_value <> 0
					BEGIN
						SET @@errorMessage = @@errorMessage +  N'ERROR: couldn''t set attribute ''separatorlocalname'' for field ''' + @@fieldname + @linebreak
						SET @return_value = 0
					END
					
					IF @@separator <> N''
					BEGIN
						EXEC @return_value = [dbo].[lsp_setattributevalue] @@owner = N'field'
									, @@idrecord = @@idfield
									, @@name = 'separator'
									, @@value = 1
									
						IF @return_value <> 0
						BEGIN
							SET @@warningMessage = @@warningMessage + N'Warning: couldn''t set separator for field ''' + @@fieldname + @linebreak
							SET @return_value = 0
						END
						ELSE
						BEGIN	
							--Make sure @@localname ends with ; in order to avoid infinite loop
							IF RIGHT(@@separator, 1) <> N';'
							BEGIN
								SET @@separator=@@separator + N';'
							END
							
							SET @currentPosition = 0
							
							--Loop through localnames
							WHILE @currentPosition <= LEN(@@separator) AND @return_value = 0
							BEGIN
								SET @nextOccurance = CHARINDEX(';', @@separator, @currentPosition)
								IF @nextOccurance <> 0
								BEGIN
									SET @currentString = SUBSTRING(@@separator, @currentPosition, @nextOccurance - @currentPosition)
									SET @currentLanguage=SUBSTRING(@currentString,0,CHARINDEX(':', @currentString))
									SET @currentLocalize=SUBSTRING(@currentString,CHARINDEX(':', @currentString)+1,LEN(@currentString)-CHARINDEX(':', @currentString))
									EXEC @return_value = [dbo].[lsp_setattributevalue] @@owner = N'string'
													, @@idrecord = @idstring
													, @@name = @currentLanguage
													, @@value = @currentLocalize
									SET @currentPosition = @nextOccurance+1
								END
							END
							IF @return_value <> 0
							BEGIN
								SET @@warningMessage = @@warningMessage + N'Warning: couldn''t set separator localnames for field ''' + @@fieldname + @linebreak
								SET @return_value = 0								
							END
						END
					END
					--End of creating separator
					
					
					--Create limevalidationtext
					IF @@fieldtype <> N'file' AND @@fieldtype <> N'sql' AND @@fieldtype <> N'geography' AND @@fieldtype <> N'html'
					BEGIN
						SET @idstring = -1
						EXEC @return_value = [dbo].[lsp_setattributevalue] @@owner = N'field'
													, @@idrecord = @@idfield
													, @@name = N'limevalidationtext'
													, @@value = @idstring output
						IF @return_value <> 0
						BEGIN
							SET @@errorMessage = @@errorMessage + N'ERROR: couldn''t set ''limevalidationtext'' for field ''' + @@fieldname + @linebreak
							SET @return_value = 0
						END
						ELSE
						BEGIN
							IF @@limevalidationtext <> N''
							BEGIN	
								--Make sure @@limevalidationtext ends with ; in order to avoid infinite loop
								IF RIGHT(@@limevalidationtext, 1) <> N';'
								BEGIN
									SET @@limevalidationtext=@@limevalidationtext + N';'
								END
								
								SET @currentPosition = 0
								
								--Loop through localnames
								WHILE @currentPosition <= LEN(@@limevalidationtext) AND @return_value = 0
								BEGIN
									SET @nextOccurance = CHARINDEX(';', @@limevalidationtext, @currentPosition)
									IF @nextOccurance <> 0
									BEGIN
										SET @currentString = SUBSTRING(@@limevalidationtext, @currentPosition, @nextOccurance - @currentPosition)
										SET @currentLanguage=SUBSTRING(@currentString,0,CHARINDEX(':', @currentString))
										SET @currentLocalize=SUBSTRING(@currentString,CHARINDEX(':', @currentString)+1,LEN(@currentString)-CHARINDEX(':', @currentString))
										EXEC @return_value = [dbo].[lsp_setattributevalue] @@owner = N'string'
														, @@idrecord = @idstring
														, @@name = @currentLanguage
														, @@value = @currentLocalize
										SET @currentPosition = @nextOccurance+1
									END
								END	
								IF @return_value <> 0
								BEGIN
									SET @@warningMessage = @@warningMessage + N'Warning: couldn''t set limevalidationtext localnames for field ''' + @@fieldname + @linebreak
									SET @return_value = 0
								END							
							END
						END
					END
					--End of creating limevalidationtext
					
					--Create comment
					SET @idstring = -1
					EXEC @return_value = [dbo].[lsp_setattributevalue] @@owner = N'field'
												, @@idrecord = @@idfield
												, @@name = N'comment'
												, @@value = @idstring output
					IF @return_value <> 0
					BEGIN
						SET @@errorMessage = @@errorMessage + N'ERROR: couldn''t set attribute ''comment'' for field ''' + @@fieldname + @linebreak
						SET @return_value = 0
					END
					ELSE
					BEGIN
						IF @@comment <> N''
						BEGIN	
							--Make sure @@comment ends with ; in order to avoid infinite loop
							IF RIGHT(@@comment, 1) <> N';'
							BEGIN
								SET @@comment=@@comment + N';'
							END
							
							SET @currentPosition = 0
							
							--Loop through localnames
							WHILE @currentPosition <= LEN(@@comment) AND @return_value = 0
							BEGIN
								SET @nextOccurance = CHARINDEX(';', @@comment, @currentPosition)
								IF @nextOccurance <> 0
								BEGIN
									SET @currentString = SUBSTRING(@@comment, @currentPosition, @nextOccurance - @currentPosition)
									SET @currentLanguage=SUBSTRING(@currentString,0,CHARINDEX(':', @currentString))
									SET @currentLocalize=SUBSTRING(@currentString,CHARINDEX(':', @currentString)+1,LEN(@currentString)-CHARINDEX(':', @currentString))
									EXEC @return_value = [dbo].[lsp_setattributevalue] @@owner = N'string'
													, @@idrecord = @idstring
													, @@name = @currentLanguage
													, @@value = @currentLocalize
									SET @currentPosition = @nextOccurance+1
								END
							END
							IF @return_value <> 0
							BEGIN
								SET @@warningMessage = @@warningMessage + N'Warning: couldn''t set ''comment'' localnames for field ''' + @@fieldname + @linebreak
								SET @return_value = 0
							END						
						END
					END
					--End of creating comment
					
					
					--Create tooltip (description)
					
					--Check if a string for description/tooltip already exists for this field
					SET @idstring = (SELECT TOP 1 s.idstring 
									FROM string s 
										INNER JOIN attributedata a 
											ON s.idstring=a.value 
											AND a.name=N'description' 
											AND a.idrecord=@@idfield)
					
															
					IF @idstring IS NULL OR @idstring = -1
					BEGIN
						--Create a new description string if it doesn't exists. Make sure it is of "type" description, i.e. choose correct idcategory	
						DECLARE @idcategoryDescription INT
						SET @idcategoryDescription = (SELECT TOP 1 idcategory FROM category WHERE name = N'description')
						
						EXEC @return_value = [dbo].[lsp_addstring] @@idcategory = @idcategoryDescription, @@idstring = @idstring OUTPUT
						
						IF @return_value <> 0
						BEGIN
							SET @@errorMessage = @@errorMessage + N'ERROR: couldn''t set attribute ''description'' for field ''' + @@fieldname + @linebreak
						END	
						ELSE
						BEGIN
							EXEC @return_value = [dbo].[lsp_addattributedata] @@owner = N'field'
										, @@idrecord = @@idfield
										, @@name = 'description'
										, @@value = @idstring
							IF @return_value <> 0
							BEGIN
								SET @@errorMessage = @@errorMessage + N'ERROR: couldn''t set attribute ''description'' for field ''' + @@fieldname + @linebreak
							END	
						END
					END	
					IF @@description <> N'' AND @return_value = 0
					BEGIN					
						--Make sure @@description ends with ; in order to avoid infinite loop
						IF RIGHT(@@description, 1) <> N';'
						BEGIN
							SET @@description=@@description + N';'
						END
						
						SET @currentPosition = 0
						
						--Loop through localnames
						WHILE @currentPosition <= LEN(@@description) AND @return_value = 0
						BEGIN
							SET @nextOccurance = CHARINDEX(';', @@description, @currentPosition)
							IF @nextOccurance <> 0
							BEGIN
								SET @currentString = SUBSTRING(@@description, @currentPosition, @nextOccurance - @currentPosition)
								SET @currentLanguage=SUBSTRING(@currentString,0,CHARINDEX(':', @currentString))
								SET @currentLocalize=SUBSTRING(@currentString,CHARINDEX(':', @currentString)+1,LEN(@currentString)-CHARINDEX(':', @currentString))
								EXEC @return_value = [dbo].[lsp_setattributevalue] @@owner = N'string'
												, @@idrecord = @idstring
												, @@name = @currentLanguage
												, @@value = @currentLocalize
								SET @currentPosition = @nextOccurance+1
							END
						END		
						IF @return_value <> 0
						BEGIN
							SET @@warningMessage = @@warningMessage + N'Warning: couldn''t set ''description'' localnames for field ''' + @@fieldname + @linebreak
						END			
					END
					--End of creating tooltip
					
					--Create options
					IF @@optionlist <> N'' AND (@@fieldtype=N'option' OR @@fieldtype=N'set' OR @@fieldtype=N'string' OR @@fieldtype=N'decimal')
					BEGIN
						SET @idstring = -1


						--Make sure @@optionlist starts with [
						IF LEFT(@@optionlist, 1) <> N'['
						BEGIN
							SET @@optionlist= N'[' + @@optionlist
						END
						
						--Make sure @@optionlist ends with ] in order to avoid infinite loop
						IF RIGHT(@@optionlist, 1) <> N']'
						BEGIN
							SET @@optionlist=@@optionlist + N']'
						END
						
						SET @currentPosition = 0

						--Loop through options
						WHILE @currentPosition <= LEN(@@optionlist) AND @return_value = 0
						BEGIN
							SET @nextOptionStarts = CHARINDEX('[', @@optionlist, @currentPosition)
							
							IF @nextOptionStarts <> 0
							BEGIN
								SET @nextOptionEnds = CHARINDEX(']', @@optionlist, @nextOptionStarts)
								IF @nextOptionEnds <> 0
								BEGIN
									SET @currentOption = SUBSTRING(@@optionlist, @nextOptionStarts + 1, @nextOptionEnds - @nextOptionStarts - 1)
									
									--Make sure @@currentOption ends with ; in order to avoid infinite loop
									IF RIGHT(@currentOption, 1) <> N';'
									BEGIN
										SET @currentOption=@currentOption + N';'
									END
									
									SET @currentPositionInOption = 0
									SET @isFirstLocalize = 1
									SET @idstring = -1
									
									WHILE @currentPositionInOption <= LEN(@currentOption) AND @return_value = 0
										BEGIN
											SET @nextOccurance = CHARINDEX(';', @currentOption, @currentPositionInOption)
											IF @nextOccurance <> 0
											BEGIN
												SET @currentString = SUBSTRING(@currentOption, @currentPositionInOption, @nextOccurance - @currentPositionInOption)
												SET @currentLanguage=LOWER(SUBSTRING(@currentString,0,CHARINDEX(':', @currentString)))
												SET @currentLocalize=SUBSTRING(@currentString,CHARINDEX(':', @currentString)+1,LEN(@currentString)-CHARINDEX(':', @currentString))
												
												IF @isFirstLocalize = 1
												BEGIN
													IF @currentLanguage <> N'color' AND @currentLanguage <> N'default' AND @currentLanguage <> N'inactive'
													BEGIN
														EXEC @return_value = [dbo].[lsp_addstring]
															@@idcategory = @idcategory
															, @@string = @currentLocalize
															, @@lang = @currentLanguage
															, @@idstring = @idstring OUTPUT

														SET @isFirstLocalize = 0
													END
												END
												ELSE
												BEGIN
													IF @currentLanguage = N'color'
													BEGIN
														EXEC lsp_addattributedata
															@@owner	= N'string',
															@@idrecord = @idstring,
															@@idrecord2 = NULL,
															@@name = N'color',
															@@value	= @currentLocalize
													END
													ELSE
													BEGIN
														IF @currentLanguage = N'default'
														BEGIN
															IF LOWER(@currentLocalize) = N'true'
															BEGIN
																EXEC [dbo].[lsp_setfieldattributevalue] 
																	@@idfield = @@idfield
																	, @@name = N'defaultvalue'
																	, @@valueint = @idstring
															END
														END
														ELSE IF @currentLanguage = N'inactive'
														BEGIN
															EXEC lsp_addattributedata
																@@owner	= N'string',
																@@idrecord = @idstring,
																@@idrecord2 = NULL,
																@@name = N'inactive',
																@@value	= @currentLocalize
														END
														ELSE
														BEGIN
															EXEC @return_value = [dbo].[lsp_setstring]
																	@@idstring = @idstring
																	, @@lang = @currentLanguage
																	, @@string = @currentLocalize
														END
													END
												END
														
												SET @currentPositionInOption = @nextOccurance+1
											END
										END		
								END	
								SET @currentPosition = @nextOptionEnds+1
							END
						END	
						IF @return_value <> 0
						BEGIN
							SET @@warningMessage = @@warningMessage + N'Warning: couldn''t create options for field ''' + @@fieldname + @linebreak
						END								
					END
					
					IF @@fieldtype IN (N'time', N'option', N'decimal', N'document',N'integer')
					BEGIN
						--Set type for timefield or optionlist
						EXEC @return_value = [dbo].[lsp_setfieldattributevalue] @@idfield = @@idfield, @@name = N'type', @@valueint = @@type
						IF @return_value <> 0
						BEGIN
							SET @@errorMessage = @@errorMessage + N'ERROR: couldn''t set attribute ''type'' for field ''' + @@fieldname + @linebreak
							SET @return_value = 0
						END	
					END
					
					EXEC lsp_setdatabasetimestamp
					EXEC lsp_refreshldc
				END	
			END
		END
	END		
END
