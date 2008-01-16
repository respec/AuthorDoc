Option Strict Off
Option Explicit On

Imports atcUtility

Friend Class frmConvert
	Inherits System.Windows.Forms.Form
	'Copyright 2000 by AQUA TERRA Consultants
	Private TargetFormat As Short
	Private Const SectionConvert As String = "Convert Window"
	
	Private Sub cmdConvert_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdConvert.Click
		Dim Index As Short = cmdConvert.GetIndex(eventSender)
		Dim RememberProjectFileName As String
		Dim RememberBaseName As String
		Dim PreviewProjectFile As Short
		
		Dim contents As Boolean
		Dim list As Boolean
		Dim timestamps As Boolean
		Dim UpNext As Boolean
		Dim id As Boolean
		Dim makeProject As Boolean
		
		If ContentsCheck.CheckState = 1 Then contents = True Else contents = False
		If TimestampCheck.CheckState = 1 Then timestamps = True Else timestamps = False
		If UpNextCheck.CheckState = 1 Then UpNext = True Else UpNext = False
		If chkID.CheckState = 1 Then id = True Else id = False
		If ProjectCheck.CheckState = 1 Then makeProject = True Else makeProject = False
		
		SaveSetting(My.Application.Info.Title, SectionConvert, "Contents", CStr(contents))
		SaveSetting(My.Application.Info.Title, SectionConvert, "Timestamps", CStr(timestamps))
		SaveSetting(My.Application.Info.Title, SectionConvert, "UpNext", CStr(UpNext))
		SaveSetting(My.Application.Info.Title, SectionConvert, "ID", CStr(id))
		SaveSetting(My.Application.Info.Title, SectionConvert, "Project", CStr(makeProject))
		SaveSetting(My.Application.Info.Title, SectionConvert, "TargetFormat", CStr(TargetFormat))
		Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
		
		If Index = 1 Then 'Preview
			PreviewProjectFile = FreeFile
			RememberBaseName = BaseName
			RememberProjectFileName = ProjectFileName
            ProjectFileName = IO.Path.GetDirectoryName(CurrentFilename) & "\PreviewProject.txt"
			BaseName = FilenameOnly(ProjectFileName)
			FileOpen(PreviewProjectFile, ProjectFileName, OpenMode.Output)
			PrintLine(PreviewProjectFile, FilenameOnly(CurrentFilename))
			FileClose(PreviewProjectFile)
		End If
		Convert(TargetFormat, contents, timestamps, UpNext, id, makeProject)
		If Index = 1 Then 'Preview
			Kill(ProjectFileName)
			BaseName = RememberBaseName
			ProjectFileName = RememberProjectFileName
		End If
		Beep()
		Me.Close()
	End Sub
	
	Private Sub frmConvert_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Dim setting As Object
		SetUnInitialized()
		Text1.Text = ""
		'UPGRADE_WARNING: Couldn't resolve default property of object setting. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		setting = GetSetting(My.Application.Info.Title, SectionConvert, "TargetFormat", CStr(modConvert.outputType.tPRINT))
		If IsNumeric(setting) Then
			optTargetFormat(setting).Checked = True
			'UPGRADE_WARNING: Couldn't resolve default property of object setting. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			optTargetFormat_CheckedChanged(optTargetFormat.Item(CInt(setting)), New System.EventArgs())
		End If
		Select Case GetSetting(My.Application.Info.Title, SectionConvert, "Contents", CStr(0))
			Case CStr(True) : ContentsCheck.CheckState = System.Windows.Forms.CheckState.Checked
			Case CStr(False) : ContentsCheck.CheckState = System.Windows.Forms.CheckState.Unchecked
		End Select
		
		Select Case GetSetting(My.Application.Info.Title, SectionConvert, "Timestamps", CStr(0))
			Case CStr(True) : TimestampCheck.CheckState = System.Windows.Forms.CheckState.Checked
			Case CStr(False) : TimestampCheck.CheckState = System.Windows.Forms.CheckState.Unchecked
		End Select
		
		Select Case GetSetting(My.Application.Info.Title, SectionConvert, "UpNext", CStr(0))
			Case CStr(True) : UpNextCheck.CheckState = System.Windows.Forms.CheckState.Checked
			Case CStr(False) : UpNextCheck.CheckState = System.Windows.Forms.CheckState.Unchecked
		End Select
		
		Select Case GetSetting(My.Application.Info.Title, SectionConvert, "ID", CStr(0))
			Case CStr(True) : chkID.CheckState = System.Windows.Forms.CheckState.Checked
			Case CStr(False) : chkID.CheckState = System.Windows.Forms.CheckState.Unchecked
		End Select
		
		Select Case GetSetting(My.Application.Info.Title, SectionConvert, "Project", CStr(0))
			Case CStr(True) : ProjectCheck.CheckState = System.Windows.Forms.CheckState.Checked
			Case CStr(False) : ProjectCheck.CheckState = System.Windows.Forms.CheckState.Unchecked
		End Select
		
	End Sub
	
	'UPGRADE_WARNING: Event optTargetFormat.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub optTargetFormat_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optTargetFormat.CheckedChanged
		If eventSender.Checked Then
			Dim Index As Short = optTargetFormat.GetIndex(eventSender)
			TargetFormat = Index
			ContentsCheck.CheckState = System.Windows.Forms.CheckState.Checked
			ContentsCheck.Enabled = True
			Select Case Index
				Case modConvert.outputType.tASCII
					UpNextCheck.CheckState = System.Windows.Forms.CheckState.Unchecked
					UpNextCheck.Enabled = False
					TimestampCheck.CheckState = System.Windows.Forms.CheckState.Unchecked
					TimestampCheck.Enabled = False
					ProjectCheck.Enabled = True
					ProjectCheck.CheckState = System.Windows.Forms.CheckState.Checked
					chkID.CheckState = System.Windows.Forms.CheckState.Unchecked
					chkID.Enabled = False
					
				Case modConvert.outputType.tPRINT
					UpNextCheck.CheckState = System.Windows.Forms.CheckState.Unchecked
					UpNextCheck.Enabled = False
					
					TimestampCheck.Enabled = True
					
					ProjectCheck.CheckState = System.Windows.Forms.CheckState.Unchecked
					ProjectCheck.Enabled = False
					
					chkID.CheckState = System.Windows.Forms.CheckState.Unchecked
					chkID.Enabled = False
				Case modConvert.outputType.tHELP
					UpNextCheck.Enabled = True
					UpNextCheck.CheckState = System.Windows.Forms.CheckState.Checked
					
					TimestampCheck.CheckState = System.Windows.Forms.CheckState.Unchecked
					TimestampCheck.Enabled = False
					
					ProjectCheck.Enabled = True
					ProjectCheck.CheckState = System.Windows.Forms.CheckState.Unchecked
					
					chkID.Enabled = True
					chkID.CheckState = System.Windows.Forms.CheckState.Unchecked
				Case modConvert.outputType.tHTMLHELP
					UpNextCheck.Enabled = True
					UpNextCheck.CheckState = System.Windows.Forms.CheckState.Unchecked
					
					TimestampCheck.CheckState = System.Windows.Forms.CheckState.Unchecked
					TimestampCheck.Enabled = False
					
					ProjectCheck.Enabled = True
					ProjectCheck.CheckState = System.Windows.Forms.CheckState.Checked
					
					chkID.Enabled = True
					chkID.CheckState = System.Windows.Forms.CheckState.Unchecked
				Case modConvert.outputType.tHTML
					UpNextCheck.Enabled = True
					UpNextCheck.CheckState = System.Windows.Forms.CheckState.Checked
					
					TimestampCheck.CheckState = System.Windows.Forms.CheckState.Unchecked
					TimestampCheck.Enabled = False
					
					ProjectCheck.CheckState = System.Windows.Forms.CheckState.Unchecked
					ProjectCheck.Enabled = False
					
					chkID.CheckState = System.Windows.Forms.CheckState.Unchecked
					chkID.Enabled = False
			End Select
		End If
	End Sub
End Class