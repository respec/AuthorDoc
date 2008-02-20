Option Strict Off
Option Explicit On

Imports MapWinUtility

Module modGlobal
    'Copyright 2000-2008 by AQUA TERRA Consultants

    Public pAppName As String = "AuthorDoc"
    Public Const pSourceExtension As String = ".txt"
	
	'Variables mostly for conversion
    Public pBaseName As String
    Public pProjectFileName As String 'file containing list of source files
    Public pCurrentFilename As String 'current file in frmMain, txtMain
	
	'Global Const RTF_START = "{\rtf1\ansi\deff0{\fonttbl{\f0\fswiss MS Sans Serif;}}\pard\plain\fs17 "
	'Global Const RTF_BOLD = "\plain\fs17\b "
	'Global Const RTF_ITALIC = "\plain\fs17\i "
	'Global Const RTF_UNDERLINE = "\plain\fs17\ul "
	'Global Const RTF_BOLD = "\b "
	'Global Const RTF_ITALIC = "\i "
	'Global Const RTF_UNDERLINE = "\ul "
	'
	'Global Const RTF_BOLD_END = "\b0 "
	'Global Const RTF_ITALIC_END = "\i0 "
	'Global Const RTF_UNDERLINE_END = "\ul0 "
	
	'Global Const RTF_PLAIN = "\plain\fs17 "
	'Global Const RTF_PARAGRAPH = "\par "
	'Global Const RTF_END = "}"
	
    'Labels for popup context menu, set in here or in frmMain Form_Load
    Public pCaptureNew As String = "Capture New Image"
    Public pCaptureReplace As String = "Capture Replacement Image"
    Public pBrowseImage As String
    Public pViewImage As String
    Public pSelectLink As String
    Public pDeleteTag As String
	
    Public pKeywords As Collection
    Public pFileKeywords As Collection
    Public pFindTimeout As Single
End Module