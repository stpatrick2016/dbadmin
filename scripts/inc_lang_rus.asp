<%
	'browse.asp
	Const langOldBrowser = "��� ������� �� ������������ ������ �����. ��� ����� ����������� ���� � ������."
	Const langBrowse = "����� ���� ������"
	Const langOnlyMDB = "������ ����� � ����������� MDB ����� ���������� �����"
	Const langDriveSelector = "����� �����"
	Const langCannotAccessFolder = "��� ������� � �����. ������: "
	Const langCurrentPath = "������� ����"
	Const langRemovable = "������� ����"
	Const langFixed = "������� ����"
	Const langNetwork = "����"
	Const langUnknown = "�����������"
	
	'data.asp
	Const langDataForTable = "������ ��� �������:"
	Const langAddRecord = "�������� ����� ������"
	Const langRefreshTable = "�������� ������"
	Const langXMLExport = "������� � XML"
	Const langXMLExportAlt = "�������������� ������ � XML �������"
	Const langExcelExport = "������� ��� Excel"
	Const langExcelExportAlt = "������������� ������ ��� Excel"
	Const langNoPrimaryKey = "�� �� ������ �������� ������ � ���� �������. �������� ���� �� �����������"
	Const langEditRecord = "������������� ������"
	Const langDeleteRecord = "������� ������"
	Const langBinaryData = "�������� ����"
	Const langSureToDeleteRecord = "�� ������� ��� ������ ������� ������ � ������ (�������)"
	Const langPageSize = "������ ��������"

	'database.asp
	Const langCouldnotSaveConfig = "�� ������� ��������� ���� ������������:"
	Const langDatabaseNotExists = "��������� ���� ������ �� ����������"
	Const langDatabaseSelection = "����� ���� ������"
	Const langEnterPath = "������� ���� � ���� ������ ��� ������� �� ������ ����� ��� ������"
	Const langCurrentDatabase = "������� ���� ������:"
	Const langOtherDatabase = "������"
	Const langCreateNew = "������� �����"
	Const langCreateNewAlt = "(�������� ���� �� ������ ������� ����� ���� ������ � ��������� �����)"
	Const langRemovePath = "������� �� ������"
	Const langDatabaseCompacted = "���� ������ ���������� � ����� ������"
	Const langBackupCreated = "������� ���� ������ ������� ������"
	Const langBackupRestored = "���� ������ ������������ � �������� �����"
	Const langDatabaseOptions = "����� ���� ������"
	Const langAffectCurrent = "��� ����� �������� � ������� ����� ������"
	Const langCompactRepair = "����� � ��������� ���� ������"
	Const langCompactRepairAlt = "������� � ���������� (���� ���� �������������) ���� ������ Access 2000. ���� �������� ��� �������� ��� ����� ������ Access 97 , �� ��� ����� �������������� � Access 2000."
	Const langCompactRepair97 = "����� � ��������� ���� ������ Access 97"
	Const langCompactRepair97Alt = "������� � ���������� (���� ���� �������������) ���� ������ Access 97. ����������� ��� �������� ���� �� �� ������ �������������� ���� ���� ������ � Access 2000"
	Const langMakeBackup = "Make backup copy of database"
	Const langMakeBackupAlt = "������� ����� ���� ������ � ��� �� ������ � ����������� "".bak"". ����� ���������� ����� ����� ������������"
	Const langRestoreBackup = "������������ �� �����"
	Const langRestoreBackupAlt = "���� ����� ���� ������ ���� ������� �����, ��� �������� ����������� (�������) ������� ���� ������ �� �����. ������� ���� ������ ����� <font color=""red"" style=""color:red"">������������</font>!"
	Const langDatabasePassword = "������ ���� ������"

	'common
	Const langDatabaseAdministration = "����������������� ���� ������"
	Const langPrev = "����."
	Const langNext = "����."
	Const langSortAscending = "����������� �� ������������"
	Const langSortDescending = "����������� �� ���������"
	Const langSubmit = "�����������"
	
	'export_csv.asp
	Const langPleaseDefineExp = "����������, ���������� ����������� ��� ����������� ���������� � ������� �������"
	Const langTab = "Tab"
	Const langSpace = "������"
	Const langOther = "������"
	Const langColumnDelimiter = "����������� �����"
	Const langTextQualifier = "������������ ������"
	Const langNoFieldNames = "��� �������� �����"
	
	'ftquery.asp
	Const langFreeTypeQuery = "SQL ������"
	Const langFreeTypeQueryAlt = "SQL ������ ��������� ��������� ����� SQL ������ � ����������� �������� (���� ������� �������)"
	Const langTotalRecords = "����� ������� ����������:"
	Const langTypeSQL = "�������� SQL ������. �� ������ ������ ��������� ��������, �������� �� ������ � ������� (;)"	'updated in version 1.8
	Const langRunIt = "���������"
	Const langUseTransaction = "��������� � ����� ����������"		'new in version 1.8
	Const langIgnoreErrors = "������������ ������"					'new in version 1.8
	
	'inc_nav.asp
	Const langDatabase = "���� ������"
	Const langTablesList = "�������"
	Const langProcedures = "���������"
	Const langViews = "����"
	Const langRelations = "���������"
	Const langVisitStpWorks = "�������� ���� StPWorks!"
	Const langStPWorks = "StP Works"
	Const langVisitDBAdmin = "�������� ��������������"
	Const langCheckUpdate = "��������� ����� ������"
	
	'main.asp
	Const langWelcomeNote = "����� ���������� � StP Database Administrator - ��������� ����������������� ��� ������ ����� ��������.<br> ���������� ������� ������ ��������������<br> ���� ��� ������ ��� ����� ��������� ���������, �� �������� ������� ������."
	Const langEnterPassword = "������� ������:"
	Const langPasswordsMismatch = "������ �� ���������"
	Const langLoggedIn = "�� ������� ����� � �������<BR>���� �� ������ �������� ������, �������� ����� ������ ����"
	Const langNewPassword = "����� ������"
	Const langRetypeNewPassword = "����������� ����� ������"
	Const langChangePassword = "�������� ������"
	
	'qlist.asp
	Const langEnterQParams = "���������� ������� ��������� ��������� �������� �� ��������. �� �������� ��������� ������� ������ ��������� ����������."
	Const langStoredProceduresList = "������ ��������"
	Const langSPName = "��������"
	Const langSPCode = "������"
	Const langSPActions = "��������"
	Const langCreateProcedure = "������� ���������"
	Const langCreateProcedureNote = "���� �� ������ ��������� �� ����� �������, �� ����� ��� ����� ������ ������ ���������"
	Const langProcedureName = "�������� ���������:"
	Const langSQLStatement = "SQL ������"
	Const langParams1stWay = "��������� ����� ���� ���������� ����� ���������. ������ ������ - �������� ������ PARAMETERS � SQL ���� �� ����� ����������� � �� ������. ��������:"
	Const langParams2ndWay = "������ ������ - ����� ��������� ������������ �� ����� ����������. ���� �� ������� �����-���� ����� ������� �� ����� ��������� ��� �������, ���� ��� ����������������� �����, �� �� ����� ��������� ������� ����������."
	Const langDeleteProcedurePrompt = "�� ������� ��� ������ ������� ���������"
	
	'recedit.asp
	Const langRecord = "������"
	Const langAutoNumberNote = "���� AutoNumber �� ����� ���� �������� ��� ��� ��� ����������� �������������. ����� ���� ��������� ���� �� ��������������"
	Const langRecEditNote = "�� ������ ������������� � ���������/���������� ������ �� ����������� � �������."
	Const langUpdate = "��������"
	Const langAdd = "��������"
	Const langReset = "������������"
	Const langCancel = "������"
	Const langRecordUpdated = "������ ��������/��������� �������"
	Const langFirst = "������"
	Const langLast = "���������"
	
	'relations.asp
	Const langRelationsNote = "��������� ����� ��������� ������� � ����� ����������� �����.<br>������ ��������� ������� ����� ���� �����-���� �������� ����� ����������."
	Const langPrimaryIndex = "�������� ������"
	Const langPrimaryTable = "�������� �������"
	Const langPrimaryColumn = "�������� ����"
	Const langForeignIndex = "�������������� ������"
	Const langForeignTable = "�������������� �������"
	Const langForeignColumn = "�������������� ����"
	Const langDeleteRelationship = "�������"
	Const langIfFieldChanged = "���� ���� <b><i>$PK_COLUMN_NAME</i></b> �������� � <b><i>$PK_TABLE_NAME</i></b>, ����� ���� <b><i>$FK_COLUMN_NAME</i></b> � <b><i>$FK_TABLE_NAME</i></b> ����� "
	Const langIfFieldDeleted = "���� ������ � <b><i>$PK_COLUMN_NAME</i></b> ����� ������� �� <b><i>$PK_TABLE_NAME</i></b>, ����� ��� ������ � ����� ��������� � <b><i>$FK_COLUMN_NAME</i></b> � <b><i>$FK_TABLE_NAME</i></b> "
	Const langChangedAlso = "����� ��������."
	Const langSetToNull = "�������� �� null."
	Const langSetToDefault = "����������� � ����������� ��������."
	Const langWillBeDeleted = "�������."
	Const langCreateRelationship = "������� ���������"
	Const langForeignIndexName = "�������� �������"
	Const langOnUpdate = "��� ���������:&nbsp;"
	Const langNoAction = "��� ���������"
	Const langSetToNull2 = "���������� � Null"
	Const langOnDelete = "��� ��������:&nbsp;"
	Const langDelete = "�������"
	
	'structure.asp
	Const langTableIndexes = "������� � �������"
	Const langShowHideIndexes = "��������/�������� ������"
	Const langIndexName = "�������� �������"
	Const langColumn = "����"
	Const langUnique = "����������"
	Const langAction = "��������"
	Const langPrimaryColumnAlt = "�������� ����"
	Const langUniqueIndexAlt = "������ ��������"
	Const langDeleteIndexAlt = "������� ������"
	Const langCreateIndex = "������� ������"
	Const langTableStructure = "��������� �������:&nbsp;"
	Const langOrdinal = "���������� �����"
	Const langName = "��������"
	Const langDataType = "��� ������"
	Const langNullable = "�������"
	Const langMaxLength = "����. �����"
	Const langDefaultValue2 = "����������� ��������"
	Const langDescription = "��������"
	Const langEditField = "������������� ����"
	Const langRemovePK = "������� ����"
	Const langSetAsPK = "���������� ��� ��������"
	Const langDeleteField = "������� ����"
	Const langAddNewColumn = "�������� ����"
	Const langCreateTableQuery = "������ �������� ��������"
	Const langCreateTableQueryAlt = "����������� ������� �������� �� ��������<BR>������� ������� ������ ���� ������� ��������"
	Const langAreYouSureToDelete = "�� ������� ��� ������ ������� ���� $name � ��� ��� �������?"
	
	'tablelist.asp
	Const langTableName = "�������� �������"
	Const langViewTableStructure = "���������"
	Const langViewTableStructureAlt = "�����������/�������� ��������� �������"
	Const langViewTableData = "������"
	Const langViewTableDataAlt = "�����������/�������� ������ � �������"
	Const langDeleteTable = "�������"
	Const langDeleteTableAlt = "������� �������"
	Const langSureToDeleteTable = "�� ������� ��� ������ ������� ������� $table_name?"
	Const langCreateNewTable = "������� �������"
	Const langNewTableName = "�������� �������:&nbsp;"
	Const langTableNavigate = "������"
	Const langTableNavigateAlt = "������� �������� ������� � ������� �/��� ���������"
	
	'vlist.asp
	Const langViewsList = "������ �����"
	Const langCode = "������"
	Const langRunViewAlt = "���������"
	Const langDeleteViewAlt = "������� ���"
	Const langCreateNewView = "������� �����"
	Const langViewName = "��������:"
	Const langSureToDeleteView = "�� ������� ��� ������ ������� ��� $name?"
%>