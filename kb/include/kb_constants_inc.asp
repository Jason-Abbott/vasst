<%
'  Application Settings --------------------------------------------------

' user experience
Const g_sEMAIL_FROM = "info@sundancemediagroup.com"
Const g_sEMAIL_SUBJECT = "Thank you for registering"
Const g_sORG_NAME = "VASST"
Const g_sAPP_NAME = "Knowledge Base"
Const g_VALIDATION_CODE_LENGTH = 16
Const g_MIN_PASSWORD_LENGTH = 6		' minimum password length
Const g_PURGE_TEMP_AFTER = 1		' hours after which temp folders may be deleted
Const g_MAX_FILE_KB = 150			' maximum size of uploaded file in kilobytes
Const g_MAX_PICTURE_KB = 30			' maximum size of user picture
Const g_bFREE_PLUGINS_ONLY = false	' allow use of free plugins only?
Const g_sREPLACE_SPACE_WITH = "_"	' replace spaces in uploaded file names with this character

' files and directories
Const g_sDB_LOCATION = "data/kb.mdb"
Const g_sDB_LOG_LOCATION = "data/kb-log.mdb"
Const g_sFILES_DIR = "files"		' disable http read access or obfuscate name
Const g_sDOWNLOAD_DIR = "download"
Const g_sDELETE_DIR = "deleted"		' under files dir
Const g_sDENY_DIR = "denied"		' under files dir
Const g_sDEFAULT_PAGE = "kb_files.asp"

' object properties
Const g_sDEFAULT_HEADER = "<embed width='494' height='102' src='/images/sundancemenu.swf'><noembed>You'll need FLASH to view this menu</noembed>"
Const g_sSQL_DATE_DELIMIT = "#"		' Access uses #, MSSSQL uses '
Const g_sDB_CONNECT = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="
Const g_sFILE_SYSTEM_OBJECT = "Scripting.FileSystemObject"
Const g_sEMAIL_OBJECT =  "ASPMAIL.ASPMailCtrl.1" '"CDONTS.NewMail"
Const g_sEMAIL_SERVER = "mail.sisna.com" '"mail.webott.com"
Const g_sUSER_COOKIE = "userid"
Const g_sTIME_COOKIE = "offset"
Const g_sSITE_COOKIE = "site"

' Error Messages ---------------------------------------------------------

Const g_sMSG_FILE_EXISTS = "A file with that name already exists"
Const g_sMSG_NO_UPLOAD_DIR = "The upload directory does not exist"
Const g_sMSG_NO_FILE_OR_PATH = "Please specify a file to upload"
Const g_sMSG_AFTER_UPLOAD = "Thank you for your contribution.<br>You will be automatically notified when it is available for download."
Const g_sMSG_FILE_TOO_BIG = "The selected file exceeds the maximum size of "
Const g_sMSG_NEED_VALIDATION = "In a few moments you should receive an e-mail containing a validation code.<br>Enter that code here to finalize your registration."
Const g_sMSG_REGISTER_THANKS = "Thank you for signing up"
Const g_sMSG_CODE_MISMATCH = "The code you entered does not match the one we sent"
Const g_sMSG_FILE_EDIT = "Thank you.  Your changes have been saved."
Const g_sMSG_AFTER_VOTING = "Thank you.  Your vote has been recorded."
Const g_sMSG_HTML_LIMIT = "Only basic HTML is allowed: <b>bold</b>, <i>italic</i>, <u>underline</u> and <a href=""javascript:alert('that tickles');"">links</a>"
Const g_sMSG_ADMIN_ADD_USER = "The user has been added"
Const g_sMSG_MULTI_SELECT = "hold Ctrl while clicking to select multiple"
Const g_sMSG_MEDIA_HINT = "If several files used, combine into single .zip file"

' Session Variables ------------------------------------------------------

Const g_sCACHE_HEADER = "kbHeader"
Const g_sCACHE = "kbFiles"
Const g_sSESSION = "user"
Const g_USER_ID = 0
Const g_USER_TYPE = 1
Const g_USER_STATUS = 2
Const g_USER_IP = 3
Const g_USER_NAME = 4
Const g_USER_ITEM_SORT = 5
Const g_USER_ITEMS_PER_PAGE = 6
Const g_USER_TIME_SHIFT = 7
Const g_USER_FILE_FORMAT = 8
Const g_USER_MSG = 9
Const g_USER_SITE = 10

' Database ID values -----------------------------------------------------

Const g_USER_ADMIN = 4
Const g_USER_VERIFIED = 3
Const g_USER_UNVERIFIED = 2

Const g_STATUS_PENDING = 1
Const g_STATUS_APPROVED = 2
Const g_STATUS_REJECTED = 3
Const g_STATUS_DISABLED = 4

Const g_FORMAT_NTSC = 1
Const g_FORMAT_PAL = 2
Const g_FORMAT_NA = 3

Const g_SORT_NAME_ASC = 1
Const g_SORT_NAME_DESC = 2
Const g_SORT_DATE_ASC = 3
Const g_SORT_DATE_DESC = 4
Const g_SORT_OWNER_ASC = 5
Const g_SORT_OWNER_DESC = 6

Const g_ITEM_FILE = 1
Const g_ITEM_TUTORIAL = 2
Const g_ITEM_FORUM = 3
Const g_ITEM_PLUGIN = 4

Const g_SOFTWARE_VEGAS = 1

' logged activities
Const g_ACT_APPROVE_UPLOAD = 1
Const g_ACT_FILE_UPLOAD = 2
Const g_ACT_LOGIN = 3
Const g_ACT_FILE_DOWNLOAD = 4
Const g_ACT_DELETE_FILE = 5
Const g_ACT_REGISTER = 6
Const g_ACT_VALIDATE_REGISTRATION = 7
Const g_ACT_CREATE_CONTEST = 8
Const g_ACT_VOTE = 9
Const g_ACT_RANK_FILE = 10
Const g_ACT_RANK_TUTORIAL = 11
Const g_ACT_DISABLE_USER = 12
Const g_ACT_FAILED_LOGIN = 13
Const g_ACT_EMAILED_PASSWORD = 14
Const g_ACT_BAD_VALIDATION = 15
Const g_ACT_REGISTER_DUPLICATE_EMAIL = 16
Const g_ACT_EDIT_FILE_ENTRY = 17
Const g_ACT_COMPACT_DATABASE = 18
Const g_ACT_BACKUP_DATABASE = 19
Const g_ACT_UNAUTHORIZED_ACCESS = 20
Const g_ACT_COOKIE_LOGIN = 21
Const g_ACT_DENY_UPLOAD = 22
Const g_ACT_TOO_LARGE_FILE_UPLOAD = 23
Const g_ACT_SAVE_CONTEST = 24
Const g_ACT_CREATE_USER = 25
Const g_ACT_EMAILED_CODE = 26

' Colors -----------------------------------------------------------------

Const g_sCOLOR_EDGE = "#6699FF"		' 102,153,255
Const g_sCOLOR_TITLE = "#feb800"

' Array Indices ----------------------------------------------------------

Const g_PLUGIN_ID = 0
Const g_PLUGIN_NAME = 1
Const g_PLUGIN_PACKAGE = 2
Const g_PLUGIN_URL = 3
Const g_PLUGIN_ITEM_ID = 4

Const g_CAT_ID = 0
Const g_CAT_NAME = 1
Const g_CAT_ITEM_ID = 2

Const g_FILTER_PAGE = 0
Const g_FILTER_SORT = 1
Const g_FILTER_CATEGORY = 2
Const g_FILTER_SOFTWARE = 3
Const g_FILTER_AUTHOR = 4

' ADO Constants ----------------------------------------------------------
	
' cursors
Const adOpenForwardOnly = 0
Const adOpenKeyset = 1
Const adOpenDynamic = 2
Const adOpenStatic = 3

' cursor location
Const adUseServer = 2
Const adUseClient = 3

' locks
Const adLockReadOnly = 1
Const adLockPessimistic = 2
Const adLockOptimistic = 3
Const adLockBatchOptimistic = 4

' commands
Const adCmdUnknown = &H0008
Const adCmdText = &H0001
Const adCmdTable = &H0002
Const adCmdStoredProc = &H0004
Const adCmdFile = &H0100
Const adCmdTableDirect = &H0200
Const adExecuteNoRecords = &H00000080

' filters
Const adFilterNone = 0

' formats
Const adClipString = 2

' data types
Const adEmpty = 0
Const adTinyInt = 16
Const adSmallInt = 2
Const adInteger = 3
Const adBigInt = 20
Const adSingle = 4
Const adDouble = 5
Const adCurrency = 6
Const adDecimal = 14
Const adNumeric = 131
Const adBoolean = 11
Const adVariant = 12
Const adGUID = 72
Const adDate = 7
Const adDBDate = 133
Const adDBTime = 134
Const adDBTimeStamp = 135
Const adBSTR = 8
Const adChar = 129
Const adVarChar = 200
Const adLongVarChar = 201
Const adWChar = 130
Const adVarWChar = 202
Const adLongVarWChar = 203
Const adBinary = 128
Const adVarBinary = 204
Const adLongVarBinary = 205

' CDO Constants ----------------------------------------------------------

' format (.BodyFormat)
Const CdoBodyFormatHTML = 0
Const CdoBodyFormatText = 1

' priority (.Importance)
Const CdoLow = 0
Const CdoNormal = 1
Const CdoHigh = 2
%>