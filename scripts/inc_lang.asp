<%
	'Core language file. If other language was choosen this file will be loaded first to ensure that all language variables is specified
	
	'browse.asp
	Dim langOldBrowser  : langOldBrowser = "Your browser doesn't support this method. Copy the database path manually"
	Dim langBrowse  : langBrowse = "Browse for database"
	Dim langOnlyMDB  : langOnlyMDB = "Only files with .MDB extention will be shown here"
	Dim langDriveSelector  : langDriveSelector = "Drive selector"
	Dim langCannotAccessFolder  : langCannotAccessFolder = "Cannot access the folder. You probably have no rights to view it. Error: "
	Dim langCurrentPath  : langCurrentPath = "Current path:"
	Dim langRemovable  : langRemovable = "Removable"
	Dim langFixed  : langFixed = "Fixed"
	Dim langNetwork  : langNetwork = "Network"
	Dim langUnknown  : langUnknown = "Unknown"

	'common
	Dim langPrev  : langPrev = "Prev"
	Dim langNext  : langNext = "Next"
	Dim langSortAscending  : langSortAscending = "Sort Ascending"
	Dim langSortDescending  : langSortDescending = "Sort Descending"
	Dim langSubmit  : langSubmit = "Submit"
	Dim langReset  : langReset = "Restore"
	Dim langNo  : langNo = "No"
	Dim langYes  : langYes = "Yes"
	Dim langCopyToClipboard  : langCopyToClipboard = "Copy to clipboard"
	Dim langCancel  : langCancel = "Cancel"
	Dim langTablesList  : langTablesList = "Tables"
	Dim langBack  : langBack = "Back"		'new in version 2.1
	Dim langTop : langTop = "Top"	'new in version 2.1
	Dim langShow : langShow = "Show"	'new in version 2.2
	Dim langHide : langHide = "Hide"	'new in version 2.2
	Dim langClose : langClose = "Close"
	
	'data.asp
	Dim langDataForTable  : langDataForTable = "&nbsp;:&nbsp;Data"
	Dim langAddRecord  : langAddRecord = "Add new record"
	Dim langRefreshTable  : langRefreshTable = "Refresh table"
	Dim langXMLExport  : langXMLExport = "XML Export"
	Dim langXMLExportAlt  : langXMLExportAlt = "Export table content as XML file"
	Dim langExcelExport  : langExcelExport = "Excel Export"
	Dim langExcelExportAlt  : langExcelExportAlt = "Export as delimited text file"
	Dim langNoPrimaryKey  : langNoPrimaryKey = "You cannot add/update data in this table. Please set the primary key first"
	Dim langEditRecord  : langEditRecord = "Edit the record"
	Dim langDeleteRecord  : langDeleteRecord = "Delete the record"
	Dim langBinaryData  : langBinaryData = "Binary data"
	Dim langSureToDeleteRecord  : langSureToDeleteRecord = "Are you sure you want to delete record with Primary Key(s)"
	Dim langPageSize  : langPageSize = "Page size:"
	Dim langFilter  : langFilter = "Filter:"
	Dim langCaptionData : langCaptionData = "Table Data"

	'database.asp
	Dim langDatabaseSelection  : langDatabaseSelection = "Database Select"
	Dim langDatabaseSelectionAlt  : langDatabaseSelectionAlt = "Please select a database to work"
	Dim langEnterPath  : langEnterPath = "To start working, select a database using Browse button or from the list below. You can also create a new blank database by typing a name of non-existent file and selecting ""Create blank"" checkbox"
	Dim langCurrentDatabase  : langCurrentDatabase = "Database:"
	Dim langCreateNew  : langCreateNew = "Create blank database"
	Dim langCreateNewAlt  : langCreateNewAlt = "(check if you want to create a blank database with path specified)"
	Dim langDatabaseCompacted  : langDatabaseCompacted = "Database compacted and repaired successfully"
	Dim langBackupCreated  : langBackupCreated = "Backup file was created successfully"
	Dim langBackupRestored  : langBackupRestored = "Database was restored from backup copy"
	Dim langDatabaseOptions  : langDatabaseOptions = "Database options"
	Dim langAffectCurrent  : langAffectCurrent = "These actions will affect the current database"
	Dim langCompactRepair  : langCompactRepair = "Compact and Repair database"
	Dim langCompactRepairAlt  : langCompactRepairAlt = "Compacts and repairs (if needed) an MS Access database."
	Dim langConvert2000  : langConvert2000 = "Convert to Access 2000"
	Dim langConvert2000Alt  : langConvert2000Alt = "Converts Access 97 database to Access 2000, then compacts and repairs it."
	Dim langMakeBackup  : langMakeBackup = "Make backup copy of database"
	Dim langMakeBackupAlt  : langMakeBackupAlt = "Creates a backup copy of your database with same file name and extention '.bak'. Any previous backups will be overwritten"
	Dim langRestoreBackup  : langRestoreBackup = "Restore from backup"
	Dim langRestoreBackupAlt  : langRestoreBackupAlt = "If backup file was created before, performing this action will restore your database from backup copy. The database will be overwritten!"
	Dim langDatabasePassword  : langDatabasePassword = "Database password:"
	Dim langDatabasePath  : langDatabasePath = "Path:"
	Dim langNewDatabasePassword  : langNewDatabasePassword = "Set database password"
	Dim langNewDatabasePasswordAlt  : langNewDatabasePasswordAlt = "Set a new database password or leave blank to remove the current one"
	Dim langNewPasswordSet  : langNewPasswordSet = "Database password has been changed successfully"
	Dim langRemoveDBPathAlt  : langRemoveDBPathAlt = "Remove database's path from list"
	Dim langActions  : langActions = "Actions"
	Dim langProperties  : langProperties = "Properties"
	Dim langFileSize  : langFileSize = "File size:"
	Dim langSizeAfterCompact  : langSizeAfterCompact = "Size after compact:"
	Dim langNewPassword  : langNewPassword = "New password"
	Dim langRetypeNewPassword  : langRetypeNewPassword = "Re-type password"
	Dim langChangePassword  : langChangePassword = "Change password"
	Dim langPasswordsMismatch  : langPasswordsMismatch = "Passwords mismatch"
	Dim langLocaleIdentifier  : langLocaleIdentifier = "Locale Identifier:"	'new in version 2.1
	Dim langDatabaseType  : langDatabaseType = "Database type:"			'new in version 2.1
	Dim langOpenDatabase : langOpenDatabase = "Open database"	'new in version 2.1
	Dim langSelectExistingDatabase : langSelectExistingDatabase = "Select existing database"	'new in version 2.1
	Dim langChangeLocaleID : langChangeLocaleID = "Change Locale"
	Dim langChangeLocaleIDAlt : langChangeLocaleIDAlt = "Allows changing a language of the database"
	Dim langNewLocaleID : langNewLocaleID = "New locale (language):"
	Dim langBrowseButton : langBrowseButton = "Browse"
	Dim langCaptionDatabase : langCaptionDatabase = "Databases"
	
	'default.asp
	Dim langWelcome  : langWelcome = "Welcome"
	Dim langWelcomeHeader  : langWelcomeHeader = "Welcome to Stp Database Administrator"
	Dim langWelcomeNote  : langWelcomeNote = "Stp Database Administrator allows you to manage your MSAccess databases through the Web, using only your browser, from any place, at any time."
	Dim langWelcomeNote2  : langWelcomeNote2 = "To start working with your databases, type your administrator password in the text box below and click Enter."
	Dim langEnterPassword  : langEnterPassword = "Password:"
	Dim langVersion  : langVersion = "Version"
	Dim langSubmitBug  : langSubmitBug = "Submit a bug"
	Dim langCaptionHome : langCaptionHome = "Home"
	Dim langLogOff : langLogOff = "Log Off"	'new in version 2.2
	
	'export_db.asp
	Dim langDatabaseExport  : langDatabaseExport = "Database Export"
	Dim langDatabaseExportAlt  : langDatabaseExportAlt = "Generate SQL script for selected tables, views and/or stored procedures"
	Dim langDatabaseExportNote  : langDatabaseExportNote = "Use Shift and Ctrl to select multiple tables, views and/or stored procedures"
	Dim langOptions  : langOptions = "Options"
	Dim langIncludeRelations  : langIncludeRelations = "Include relations"
	Dim langGenerateSQLScript  : langGenerateSQLScript = "Generate SQL script"
	Dim langSQLScriptNote  : langSQLScriptNote = "Generated SQL script for selected objects in database"
	Dim langProcedures  : langProcedures = "Procedures"
	Dim langViews  : langViews = "Views"
	
	'export_csv.asp
	Dim langPleaseDefineExp  : langPleaseDefineExp = "Please define the column and rows delimiters, or use default and click Export"
	Dim langTab  : langTab = "Tab"
	Dim langSpace  : langSpace = "Space"
	Dim langOther  : langOther = "Other"
	Dim langColumnDelimiter  : langColumnDelimiter = "Column Delimiter"
	Dim langTextQualifier  : langTextQualifier = "Text qualifier"
	Dim langNoFieldNames  : langNoFieldNames = "No field names"
	Dim langCaptionExportCSV : langCaptionExportCSV = "Excel Export"
	
	'ftquery.asp
	Dim langFreeTypeQuery  : langFreeTypeQuery = "Free-Type Query : Script"
	Dim langFreeTypeQueryAlt  : langFreeTypeQueryAlt = "Free-Type Query allows you to make your own SQL statement and get results from it (if there are any results returned)"
	Dim langTotalRecords  : langTotalRecords = "Total records returned:"
	Dim langRunIt  : langRunIt = "Execute"
	Dim langTypeSQL  : langTypeSQL = "Type your SQL statement in a box below. You can run several queries at once, separating each of them with semicolon (;)"	'updated in version 1.8
	Dim langUseTransaction  : langUseTransaction = "Run in one transaction"		'new in version 2.0
	Dim langIgnoreErrors  : langIgnoreErrors = "Ignore errors"				'new in version 2.0
	Dim langFTQResults  : langFTQResults = "Free-Type Query : Results"
	Dim langRecordsAffected  : langRecordsAffected = "Records affected:"
	Dim langCaptionFreeTypeQuery : langCaptionFreeTypeQuery = "Free-Type Query"
	
	'import_db.asp, new in version 2.1
	Dim langImportDatabase  : langImportDatabase = "Import from database"
	Dim langImportDatabaseAlt  : langImportDatabaseAlt = "Allows you to import tables, views and stored prodecures from another Access database"
	Dim langImportDatabaseWelcome  : langImportDatabaseWelcome = "Welcome to database import Wizard! The Wizard will allow you to import tables, views and stored procedures from another Access database. To begin please specify the external database, where you want to import from"
	Dim langImportDatabaseNote  : langImportDatabaseNote = "Choose external database"
	Dim langSelectExternalTables  : langSelectExternalTables = "Now you have an option to select external tables, views and/or stored procedures you want to import into your database. If you want to import data as well, make sure to not import many tables with a lot of data at once."
	Dim langImport  : langImport = "Import"
	Dim langImportSuccess  : langImportSuccess = "All selected tables, views and stored procedures were imported successfully"
	Dim langIncludeData  : langIncludeData = "Including data"
	Dim langPathToExternalDatabase : langPathToExternalDatabase = "Path to external database"
	Dim langExternalDBPassword : langExternalDBPassword = "External database password"
	Dim langCaptionImportDB : langCaptionImportDB = "Database Import"
	
	'linked.asp, new in version 2.3
	Dim langCaptionLinkedTable : langCaptionLinkedTable = "Link table Wizard"
	Dim langLinkedDatabaseSelect : langLinkedDatabaseSelect = "Please specify the database from which you want to link a table"
	Dim langTableToLink : langTableToLink = "Table to link"
	Dim langAliasName : langAliasName = "Link as (leave blank for the same name)"
	
	'lookup.asp, new in version 2.3
	Dim langLookup : langLookup = "Lookup"
	Dim langLookupAlt : langLookupAlt = "Select a value to insert"
	Dim langNoLookupValues : langNoLookupValues = "No related values available"
	
	'qlist.asp
	Dim langEnterQParams  : langEnterQParams = "Please enter the parameters of procedure divided by commas. Remember to enclose text parameters in single quotation marks."
	Dim langStoredProceduresList  : langStoredProceduresList = "Stored Procedures"
	Dim langSPName  : langSPName = "Name"
	Dim langSPCode  : langSPCode = "Code"
	Dim langSPActions  : langSPActions = "Actions"
	Dim langCreateProcedure  : langCreateProcedure = "Create a new procedure"
	Dim langUpdateProcedure  : langUpdateProcedure = "Update procedure"
	Dim langCreateProcedureNote  : langCreateProcedureNote = "Note, if you won't add any parameter in your SQL statement, then a new View will be created instead"
	Dim langProcedureName  : langProcedureName = "Procedure name:"
	Dim langSQLStatement  : langSQLStatement = "SQL Statement"
	Dim langParams1stWay  : langParams1stWay = "Parameters can be defined in 2 ways. First way, by adding PARAMETERS clause in your SQL statement with all parameters and thier types listed. For example:"
	Dim langParams2ndWay  : langParams2ndWay = "The second way, when parameters are determined on-the-fly. If you will add a parameter that is not recognized as a column name or SQL reserved word, it will be threated as parameter."
	Dim langDeleteProcedurePrompt  : langDeleteProcedurePrompt = "Are you sure you want to delete a stored procedure"
	Dim langSPExecute  : langSPExecute = "Execute Stored Procedure"
	Dim langSPEdit  : langSPEdit = "Edit Procedure"
	Dim langSPDelete  : langSPDelete = "Delete Stored Procedure"
	
	'recedit.asp
	Dim langRecord  : langRecord = "record"
	Dim langAutoNumberNote  : langAutoNumberNote = "Note that AutoNumber fields you cannot edit as they are updated automatically. Also columns of type Binary won't be shown here"
	Dim langRecEditNote  : langRecEditNote = "You can cycle through records using Update+Next/Update+Prev and Add+Next buttons. To add/update the record and return to table, click Add/Update button"
	Dim langUpdate  : langUpdate = "Update"
	Dim langAdd  : langAdd = "Add"
	Dim langRecordUpdated  : langRecordUpdated = "Record added/updated successfully"
	Dim langFirst  : langFirst = "First"
	Dim langLast  : langLast = "Last"
	Dim langCaptionRecEdit : langCaptionRecEdit = "Record Edit"
	Dim langEdit : langEdit = "Edit"
	
	'relations.asp
	Dim langRelationsNote  : langRelationsNote = "Each of relationships described also in more readable form."
	Dim langPrimaryIndex  : langPrimaryIndex = "Primary Index"
	Dim langPrimaryTable  : langPrimaryTable = "Primary Table"
	Dim langPrimaryColumn  : langPrimaryColumn = "Primary Column"
	Dim langForeignIndex  : langForeignIndex = "Foreign Index"
	Dim langForeignTable  : langForeignTable = "Foreign Table"
	Dim langForeignColumn  : langForeignColumn = "Foreign Column"
	Dim langDeleteRelationship  : langDeleteRelationship = "Delete this relationship"
	Dim langIfFieldChanged  : langIfFieldChanged = "If field <b><i>$PK_COLUMN_NAME</i></b> has been changed in <b><i>$PK_TABLE_NAME</i></b>, then field <b><i>$FK_COLUMN_NAME</i></b> in <b><i>$FK_TABLE_NAME</i></b> will be "
	Dim langIfFieldDeleted  : langIfFieldDeleted = "If record with <b><i>$PK_COLUMN_NAME</i></b> has been deleted from <b><i>$PK_TABLE_NAME</i></b>, then all records with same <b><i>$FK_COLUMN_NAME</i></b> in <b><i>$FK_TABLE_NAME</i></b> "
	Dim langChangedAlso  : langChangedAlso = "changed also."
	Dim langSetToNull  : langSetToNull = "set to null."
	Dim langSetToDefault  : langSetToDefault = "set to default value."
	Dim langWillBeDeleted  : langWillBeDeleted = "will be deleted."
	Dim langCreateRelationship  : langCreateRelationship = "Create new relationship"
	Dim langForeignIndexName  : langForeignIndexName = "Foreign index name"
	Dim langOnUpdate  : langOnUpdate = "On update:&nbsp;"
	Dim langNoAction  : langNoAction = "No Action"
	Dim langOnDelete  : langOnDelete = "On delete:&nbsp;"
	Dim langDelete  : langDelete = "Delete"
	Dim langRelations  : langRelations = "Relationships"
	
	'settings.asp
	Dim langSettings  : langSettings = "Settings"
	Dim langSettingsNotAvailable  : langSettingsNotAvailable = "Sorry, the Settings page is not availeble since you haven't specified an XML file path. If you wish to do it now, please open <font color=""green"">scripts/inc_config.asp</font> file in any text editor, such as Notepad and change the value of <font color=""green"">DBA_cfgProfilePath</font>"
	Dim langSessionVariables  : langSessionVariables = "Session Variables"
	Dim langUsername  : langUsername = "Username"
	Dim langUserPassword  : langUserPassword = "User's Password"
	Dim langDBPath  : langDBPath = "Database path"
	Dim langDBPassword  : langDBPassword = "Database password"
	Dim langSaveDBPaths  : langSaveDBPaths = "Save database paths?"
	Dim langOtherSettings  : langOtherSettings = "Other settings"
	Dim langSaveSuccess  : langSaveSuccess = "Settings were saved successfully"
	Dim langRecordsPerPage  : langRecordsPerPage = "Records per page"	'new in version 2.1
	Dim langLanguage : langLanguage = "Language"	'new in version 2.1
	Dim langSessionTimeout : langSessionTimeout = "Session timeout (min.)"		'new in version 2.2
	Dim langMaxTimeout : langMaxTimeout = "Max timeout (24 hrs.)"	'new in version 2.2
	Dim langDefault : langDefault = "Default"	'new in version 2.2
	Dim langShowSysTables : langShowSysTables = "Show system tables" 'new in version 2.2
	
	'structure.asp
	Dim langTableIndexes  : langTableIndexes = "&nbsp;:&nbsp;Indexes"
	Dim langIndexName  : langIndexName = "Index Name"
	Dim langColumn  : langColumn = "Column"
	Dim langUnique  : langUnique = "Unique"
	Dim langIndexedUnique  : langIndexedUnique = "Yes, unique index"
	Dim langIndxedDuplicates  : langIndxedDuplicates = "Yes, allow duplicates"
	Dim langPrimaryColumnAlt  : langPrimaryColumnAlt = "Primary Key column"
	Dim langUniqueIndexAlt  : langUniqueIndexAlt = "The index is unique"
	Dim langDeleteIndexAlt  : langDeleteIndexAlt = "Delete the index"
	Dim langCreateIndex  : langCreateIndex = "Create index"
	Dim langTableStructure  : langTableStructure = "&nbsp:&nbsp;Structure"
	Dim langOrdinal  : langOrdinal = "Ordinal"
	Dim langName  : langName = "Name"
	Dim langDataType  : langDataType = "Data type"
	Dim langNullable  : langNullable = "Nullable"
	Dim langMaxLength  : langMaxLength = "Max. length"
	Dim langDefaultValue2  : langDefaultValue2 = "Default Value"
	Dim langDescription  : langDescription = "Description"
	Dim langEditField  : langEditField = "Edit column"
	Dim langRemovePK  : langRemovePK = "Remove Primary Key"
	Dim langSetAsPK  : langSetAsPK = "Set as Primary Key"
	Dim langDeleteField  : langDeleteField = "Delete column"
	Dim langAddNewColumn  : langAddNewColumn = "Add new column"
	Dim langCreateTableQuery  : langCreateTableQuery = "&nbsp;:&nbsp;SQL"
	Dim langCreateTableQueryAlt  : langCreateTableQueryAlt = "CREATE TABLE SQL script that can be used to re-created this table"
	Dim langAreYouSureToDelete  : langAreYouSureToDelete = "Are you sure you want to delete column $name and all indexes to it?"
	Dim langAllowZeroLength  : langAllowZeroLength = "Allow zero-length"
	Dim langUnicodeCompress  : langUnicodeCompress = "Compress Unicode"
	Dim langIndexed  : langIndexed = "Indexed"
	Dim langCaptionTableStructure : langCaptionTableStructure = "Table Structure"
	
	'tablelist.asp
	Dim langTableName  : langTableName = "Table name"
	Dim langCreated  : langCreated = "Created"
	Dim langModified  : langModified = "Modified"
	Dim langViewTableStructureAlt  : langViewTableStructureAlt = "View table's structure"
	Dim langViewTableDataAlt  : langViewTableDataAlt = "View table's data"
	Dim langDeleteTableAlt  : langDeleteTableAlt = "Delete the table"
	Dim langSureToDeleteTable  : langSureToDeleteTable = "Are you sure you want to delete table $table_name?"
	Dim langCreateNewTable  : langCreateNewTable = "Create new table"
	Dim langNewTableName  : langNewTableName = "New table name:&nbsp;"
	Dim langTableNavigateAlt  : langTableNavigateAlt = "Cycle through records in table"
	Dim langCaptionTablesList : langCaptionTablesList = "Tables List"
	Dim langRenameTableAlt : langRenameTableAlt = "Rename table"
	Dim langAddLinkedTable : langAddLinkedTable = "Add linked table"
	
	'vlist.asp
	Dim langRunViewAlt  : langRunViewAlt = "Run view"
	Dim langDeleteViewAlt  : langDeleteViewAlt = "Delete view"
	Dim langCreateNewView  : langCreateNewView = "Create a new view"
	Dim langViewName  : langViewName = "View name:"
	Dim langSureToDeleteView  : langSureToDeleteView = "Are you sure you want to delete view $name?"
	Dim langUpdateView  : langUpdateView = "Update view"
	Dim langEditView  : langEditView = "Edit View"
	Dim langCaptionViews : langCaptionViews = "Views"
%>