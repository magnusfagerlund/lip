GO
/****** Object:  StoredProcedure [dbo].[csp_lip_createfield]    Script Date: 09/04/2015 09:28:45 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- Written by: Fredrik Eriksson, Jonny Springare
-- Created: 2015-04-16

CREATE PROCEDURE [dbo].[csp_lip_createfield]
	@@tablename NVARCHAR(64)
	, @@fieldname NVARCHAR(64)
	, @@localname NVARCHAR(MAX)
	, @@separator NVARCHAR(MAX) = N''
	, @@tooltip NVARCHAR(MAX) = N''
	, @@optionlist NVARCHAR(MAX) = N''
	, @@fieldtype NVARCHAR(64)
	, @@defaultvalue NVARCHAR(64) = NULL
	, @@limedefaultvalue NVARCHAR(64) = NULL
	, @@limereadonly INT = 0
	, @@invisible INT = 0
	, @@required INT = NULL
	, @@limerequiredforedit INT = 0
	, @@width INT = NULL
	, @@height INT = NULL
	, @@length INT = NULL
	, @@newline INT = 2 -- Default value 2 means Fixed width
	, @@sql NVARCHAR(MAX) = N''
	, @@onsqlupdate NVARCHAR(MAX) = N''
	, @@onsqlinsert NVARCHAR(MAX) = N''
	, @@fieldorder INT = 0 -- Default value 0 means it will be put last
	, @@isnullable INT = 0
	, @@type INT = 0
	, @@relationtab BIT = 0
	, @@errorMessage NVARCHAR(512) OUTPUT
	, @@idfield INT OUTPUT
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
	DECLARE @currentLanguage NVARCHAR(8)
	DECLARE @currentLocalize NVARCHAR(256)
	DECLARE @isFirstLocalize BIT
	DECLARE @currentOption NVARCHAR(MAX)
	DECLARE @nextOptionStarts INT
	DECLARE @nextOptionEnds INT
	DECLARE @currentPositionInOption INT
	DECLARE @supportedFieldtypes NVARCHAR(MAX)

	SET @return_value = NULL
	SET @@idfield = NULL
	SET @idstringlocalname = NULL
	SET @idcategory = NULL
	SET @idstring = NULL
	SET @@errorMessage = N''
	SET @sql = N''
	SET @isFirstLocalize = 1
	SET @currentOption =N''
	SET @nextOptionStarts = 0
	SET @nextOptionEnds = 0
	SET @supportedFieldtypes = N'string;integer;decimal;time;link;yesno;set;option;formatedstring;color;relation'
	--Not supported: geography;html;xml;file;user;sql
	
	--Make sure @@length is set to NULL if fieldtype is not string
	IF @@fieldtype <> N'string' AND @@length IS NOT NULL
	BEGIN
		SET @@length = NULL
	END
	
	--Check if field already exists
	EXEC lsp_getfield @@table = @@tablename, @@name = @@fieldname, @@count = @count OUTPUT
	
	IF  @count > 0 --Fieldname already exists
	BEGIN
		SET @@errorMessage = N'Field ''' + @@fieldname + N''' already exists and will not be re-created. Please verify that properties for the field are correct.'
	END
	ELSE --Field doesn't exist
	BEGIN
		--Check if fieldtype exists
		IF (SELECT COUNT(*) FROM fieldtype WHERE name = @@fieldtype AND active = 1 AND creatable = 1) <> 1
		BEGIN
			SET @@errorMessage = N'''' + @@fieldtype + N''' is not a valid fieldtype. Field ''' + @@fieldname + ''' couldn''t be created'
		END
		ELSE
		BEGIN
			--Check if fieldtype is implemented in LIP
			IF CHARINDEX(@@fieldtype, @supportedFieldtypes) = 0
			BEGIN
				SET @@errorMessage = N'Fieldtype ''' + @@fieldtype + N''' is not implemented in LIP. Field ''' + @@fieldname + ''' couldn''t be created'
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
				EXEC lsp_refreshldc
					
				--If return value is not 0, something went wrong and the field wasn't created
				IF @return_value <> 0
				BEGIN
					SET @@errorMessage = N'Field ''' + @@fieldname + N''' couldn''t be created'
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
					EXEC @return_value = [dbo].[lsp_setfieldattributevalue] @@idfield = @@idfield, @@name = N'limereadonly', @@valueint = @@limereadonly
					
					--Set Default value (interpreted by LIME)
					IF @@limedefaultvalue IS NOT NULL
					BEGIN
						EXEC @return_value = [dbo].[lsp_setattributevalue] @@owner = N'field', @@idrecord = @@idfield, @@name = N'limedefaultvalue', @@value = @@limedefaultvalue	-- Default Value (interpreted by LIME Pro) 
					END
					
					--Set invisible/visible
					EXEC @return_value = [dbo].[lsp_setattributevalue] @@owner = N'field', @@idrecord = @@idfield, @@name = N'invisible', @@valueint = @@invisible
					
					--Set required attribute
					IF @@required IS NOT NULL
					BEGIN
						EXEC @return_value = [dbo].[lsp_setfieldattributevalue] @@idfield = @@idfield, @@name = N'required', @@valueint = @@required
					END
					--Set attribute Required for editing in Lime
					EXEC @return_value = [dbo].[lsp_setfieldattributevalue] @@idfield = @@idfield, @@name = N'limerequiredforedit', @@valueint = @@limerequiredforedit
					
					--Set width
					IF @@width IS NOT NULL
					BEGIN
						EXEC @return_value = [dbo].[lsp_setfieldattributevalue] @@idfield = @@idfield, @@name = N'width', @@valueint = @@width
					END
					
					--Set height
					IF @@height IS NOT NULL
					BEGIN
						EXEC @return_value = [dbo].[lsp_setfieldattributevalue] @@idfield = @@idfield, @@name = N'height', @@valueint = @@height
					END
					
					--Set width properties
					EXEC @return_value = [dbo].[lsp_setfieldattributevalue] @@idfield = @@idfield, @@name = N'newline', @@valueint = @@newline
					
					--Set SQL Expression
					EXEC @return_value = [dbo].[lsp_setfieldattributevalue] @@idfield = @@idfield, @@name = N'sql', @@value = @@sql
					
					--Set SQL for update
					EXEC @return_value = [dbo].[lsp_setfieldattributevalue] @@idfield = @@idfield, @@name = N'onsqlupdate', @@value = @@onsqlupdate
					
					--Set SQL for new
					EXEC @return_value = [dbo].[lsp_setfieldattributevalue] @@idfield = @@idfield, @@name = N'onsqlinsert', @@value = @@onsqlinsert
					
					--Set fieldorder, if not provided we use default value 0 which means it will be put last
					EXEC @return_value = [dbo].[lsp_setfieldorder] @@idfield = @@idfield, @@fieldorder = @@fieldorder
					
					--Set relation properties, if relationfield
					IF @@fieldtype = N'relation'
					BEGIN
						EXEC lsp_setattributevalue @@owner = 'field', @@idrecord = @@idfield, @@name = 'relationmincount', @@value = 0
						IF @@relationtab = 1
						BEGIN
							EXEC lsp_setattributevalue @@owner = 'field', @@idrecord = @@idfield, @@name = 'relationmaxcount', @@value = 1
						END	
					END
					
					--Create separator
					IF @@separator <> N''
					BEGIN
						SET @idstring = -1
						EXEC @return_value = [dbo].[lsp_setattributevalue] @@owner = N'field'
									, @@idrecord = @@idfield
									, @@name = 'separator'
									, @@value = 1
						EXEC @return_value = [dbo].[lsp_setattributevalue] @@owner = N'field'
													, @@idrecord = @@idfield
													, @@name = N'separatorlocalname'
													, @@value = @idstring output
													
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
					END
					--End of creating separator
					
					--Create tooltip
					IF @@tooltip <> N''
					BEGIN
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
							
							EXEC @return_value = [dbo].[lsp_addattributedata] @@owner = N'field'
										, @@idrecord = @@idfield
										, @@name = 'description'
										, @@value = @idstring
						END	
													
						--Make sure @@localname ends with ; in order to avoid infinite loop
						IF RIGHT(@@tooltip, 1) <> N';'
						BEGIN
							SET @@tooltip=@@tooltip + N';'
						END
						
						SET @currentPosition = 0
						
						--Loop through localnames
						WHILE @currentPosition <= LEN(@@tooltip) AND @return_value = 0
						BEGIN
							SET @nextOccurance = CHARINDEX(';', @@tooltip, @currentPosition)
							IF @nextOccurance <> 0
							BEGIN
								SET @currentString = SUBSTRING(@@tooltip, @currentPosition, @nextOccurance - @currentPosition)
								SET @currentLanguage=SUBSTRING(@currentString,0,CHARINDEX(':', @currentString))
								SET @currentLocalize=SUBSTRING(@currentString,CHARINDEX(':', @currentString)+1,LEN(@currentString)-CHARINDEX(':', @currentString))
								EXEC @return_value = [dbo].[lsp_setattributevalue] @@owner = N'string'
												, @@idrecord = @idstring
												, @@name = @currentLanguage
												, @@value = @currentLocalize
								SET @currentPosition = @nextOccurance+1
							END
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
													IF @currentLanguage <> N'color' AND @currentLanguage <> N'default'
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
					END
					
					IF @@fieldtype = N'time' OR @@fieldtype = N'option'
					BEGIN
						--Set type for timefield or optionlist
						EXEC [dbo].[lsp_setfieldattributevalue] @@idfield = @@idfield, @@name = N'type', @@valueint = @@type
					END
					
					EXEC lsp_refreshldc
					
					--If return value is not 0, something went wrong while setting field attributes
					IF @return_value <> 0
					BEGIN
						SET @@errorMessage = N'Something went wrong while setting attributes for field ''' + @@fieldname + N'''. Please check that field properties are correct.'
					END
				END	
			END
		END
	END	
END
