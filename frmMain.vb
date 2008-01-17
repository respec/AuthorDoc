Option Strict Off
Option Explicit On

Imports VB = Microsoft.VisualBasic
Imports MapWinUtility
Imports atcUtility

Friend Class frmMain
	Inherits System.Windows.Forms.Form
    'Copyright 2000-2008 by AQUA TERRA Consultants

    Dim mnuRecent As New ArrayList
    Dim path As String
    Dim CurrentFileContents As String 'What was last saved or retrieved from pCurrentFilename
    Dim MaxUndo As Integer = 10
    Dim Undos(MaxUndo) As String
    Dim UndoCursor(MaxUndo) As Integer
	Dim UndoPos As Integer
	Dim UndosAvail As Integer
    Dim Undoing As Boolean
	Dim Changed As Boolean 'True if txtMain.Text has been edited
	Dim ProjectChanged As Boolean
	Dim ViewFormatting As Boolean
	Dim FormatWhileTyping As Boolean
	Dim txtMainButton As Integer
	Dim AbortAction As Boolean
	
	Dim tagName As String
	Dim openTagPos, closeTagPos As Integer 'current tag being edited
	Dim NodeLinking As Integer 'Index in tree of file containing link being edited
	
	Private SashDragging As Boolean
	Private Const SectionMainWin As String = "Main Window"
	Private Const SectionRecentFiles As String = "Recent Files"
	Private Const MaxRecentFiles As Short = 6
	
	Private Sub cmdFind_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles cmdFind.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		If KeyAscii >= 32 And KeyAscii < 127 Then
			txtFind.Focus()
			txtFind.Text = Chr(KeyAscii)
			txtFind.SelectionStart = 1
		End If
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub cmdFind_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles cmdFind.MouseUp
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		Static Finding As Boolean
		Dim searchThrough, SearchFor As String
		Dim searchPos, selStart, startNodeIndex As Integer
		If Button = VB6.MouseButtonConstants.RightButton Then
			fraFind.Visible = False
			frmMain_Resize(Me, New System.EventArgs())
		ElseIf cmdFind.Text = "Stop" Then 
			Finding = False
		Else
			'Dim StartTime As Single
			Finding = True
			cmdFind.Text = "Stop"
			'StartTime = Timer
			searchThrough = txtMain.Text
			If txtFind.Text = "" And txtMain.SelectionLength > 0 Then txtFind.Text = txtMain.SelectedText
			If txtFind.Text <> "" Then
				SearchFor = UnEscape(txtFind.Text)
				selStart = txtMain.SelectionStart
				searchPos = txtMain.SelectionStart + txtMain.SelectionLength
                searchPos = txtMain.Find(SearchFor, searchPos, RichTextBoxFinds.None)
				startNodeIndex = tree1.SelectedItem.Index
				If searchPos < 0 And Finding Then
					If QuerySave <> MsgBoxResult.Cancel Then
NextNode: 
						If tree1.SelectedItem Is Nothing Then
							tree1_NodeClick(tree1, New AxComctlLib.ITreeViewEvents_NodeClickEvent(tree1.Nodes(1)))
						ElseIf tree1.SelectedItem.Index < tree1.Nodes.Count Then 
							tree1_NodeClick(tree1, New AxComctlLib.ITreeViewEvents_NodeClickEvent(tree1.Nodes(tree1.SelectedItem.Index + 1)))
						Else
							tree1_NodeClick(tree1, New AxComctlLib.ITreeViewEvents_NodeClickEvent(tree1.Nodes(1)))
						End If
						searchPos = txtMain.Find(SearchFor, 0)
						If searchPos < 0 And tree1.SelectedItem.Index <> startNodeIndex Then
							'If Timer - StartTime < FindTimeout Then
							System.Windows.Forms.Application.DoEvents()
							If Finding Then GoTo NextNode
						End If
					End If
				End If
			End If
		End If
		cmdFind.Text = "Find"
	End Sub
	
	Private Function UnEscape(ByVal Source As String) As String
        Dim retval As String = ""
		Dim ch As String
		Dim chpos, lastchpos As Integer
		chpos = 1
		lastchpos = Len(Source)
		While chpos <= lastchpos
			ch = Mid(Source, chpos, 1)
			If ch = "\" Then
				chpos = chpos + 1
				If chpos > lastchpos Then
                    retval &= ch
                Else
                    ch = Mid(Source, chpos, 1)
                    Select Case LCase(ch)
                        Case "c" : retval &= vbCrLf
                        Case "n" : retval &= vbLf
                        Case "r" : retval &= vbCr
                        Case "t" : retval &= vbTab
                        Case "\" : retval &= ch
                        Case Else : retval &= "^" & ch
                    End Select
                End If
            Else
                retval &= ch
			End If
			chpos = chpos + 1
		End While
		UnEscape = retval
	End Function
	
	Private Sub cmdReplace_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles cmdReplace.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		If KeyAscii >= 32 And KeyAscii < 127 Then
			txtReplace.Focus()
			txtReplace.Text = Chr(KeyAscii)
			txtReplace.SelectionStart = 1
		End If
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub cmdReplace_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles cmdReplace.MouseUp
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		Dim startNodeIndex As Integer
		Dim searchedBeyondStart As Boolean
		Dim FindText, ReplaceText As String
		If Button = VB6.MouseButtonConstants.RightButton Then
			fraFind.Visible = False
			frmMain_Resize(Me, New System.EventArgs())
		Else
			FindText = LCase(UnEscape(txtFind.Text))
			ReplaceText = UnEscape(txtReplace.Text)
			startNodeIndex = tree1.SelectedItem.Index
			searchedBeyondStart = False
			If LCase(txtMain.SelectedText) = FindText Then
NextReplace: 
				txtMain.SelectedText = ReplaceText
			End If
			cmdFind_MouseUp(cmdFind, New System.Windows.Forms.MouseEventArgs(Button * &H100000, 0, VB6.TwipsToPixelsX(x), VB6.TwipsToPixelsY(y), 0))
			If startNodeIndex <> tree1.SelectedItem.Index Then searchedBeyondStart = True
			If Shift > 0 Then
				If Not searchedBeyondStart Or startNodeIndex <> tree1.SelectedItem.Index Then
					If LCase(txtMain.SelectedText) = FindText Then GoTo NextReplace
				End If
			End If
		End If
    End Sub

    Private Sub frmMain_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        Dim setting As Object
        Dim rf As Integer
        BrowseImage = "Use Other Image (File)"
        ViewImage = "View image"
        SelectLink = "Link to Page (select)"
        DeleteTag = "Delete"
        mnuContext(0).Text = DeleteTag
        txtMain.Text = ""

        'UPGRADE_ISSUE: App property App.HelpFile was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
        'App.HelpFile = GetSetting(pAppName, "Files", "Help", My.Application.Info.DirectoryPath & "\AuthorDoc.chm")
        pBaseName = GetSetting(pAppName, "Defaults", "BaseName", "")
        path = GetSetting(pAppName, "Defaults", "Path", CurDir())
        ViewFormatting = CBool(GetSetting(pAppName, "Defaults", "ViewFormatting", CStr(True)))
        FormatWhileTyping = CBool(GetSetting(pAppName, "Defaults", "FormatWhileTyping", CStr(False)))
        mnuAutoParagraph.Checked = CBool(GetSetting(pAppName, "Defaults", "AutoParagraph", CStr(False)))
        setting = GetSetting(pAppName, "Defaults", "FindTimeout", CStr(2))
        If IsNumeric(setting) Then FindTimeout = setting
        setting = GetSetting(pAppName, SectionMainWin, "Width")
        If IsNumeric(setting) Then Width = VB6.TwipsToPixelsX(setting)
        setting = GetSetting(pAppName, SectionMainWin, "Height")
        If IsNumeric(setting) Then Height = VB6.TwipsToPixelsY(setting)
        setting = GetSetting(pAppName, SectionMainWin, "Left")
        If IsNumeric(setting) Then Left = VB6.TwipsToPixelsX(setting)
        setting = GetSetting(pAppName, SectionMainWin, "Top")
        If IsNumeric(setting) Then Top = VB6.TwipsToPixelsY(setting)
        setting = GetSetting(pAppName, SectionMainWin, "TreeWidth")
        If IsNumeric(setting) Then
            sash.Left = VB6.TwipsToPixelsX(setting)
            SashDragging = True
            sash_MouseMove(sash, New System.Windows.Forms.MouseEventArgs(1 * &H100000, 0, 0, 0, 0))
            SashDragging = False
        End If
        For rf = MaxRecentFiles To 1 Step -1
            setting = GetSetting(pAppName, SectionRecentFiles, CStr(rf), "")
            If IO.File.Exists(setting) Then AddRecentFile(CStr(setting))
        Next rf

        mnuFormatting.Checked = ViewFormatting
        mnuFormatWhileTyping.Checked = FormatWhileTyping
        cdlgOpen.FileName = path & "\" & pBaseName & pSourceExtension
        cdlgSave.FileName = path & "\" & pBaseName & pSourceExtension
        cdlgImageOpen.FileName = path
        If IO.Directory.Exists(path) Then ChDir(path)
        If IO.File.Exists(cdlgOpen.FileName) Then
            Me.Show()
            Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
            OpenProject((cdlgOpen.FileName), tree1)
            If tree1.Nodes.Count > 0 Then tree1_NodeClick(tree1, New AxComctlLib.ITreeViewEvents_NodeClickEvent(tree1.Nodes(1)))
            Me.Cursor = System.Windows.Forms.Cursors.Default
        End If
    End Sub

    Private Sub frmMain_Resize(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Resize
        Dim newWidth As Integer
        If VB6.PixelsToTwipsY(Height) > 800 Then sash.Height = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Height) - 753) 'menu height
        tree1.Height = sash.Height
        'If fraFind.Visible Then
        '	txtMain.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(fraFind.Top) + VB6.PixelsToTwipsY(fraFind.Height))
        'Else
        '	txtMain.Top = fraFind.Top
        'End If
        'If VB6.PixelsToTwipsY(sash.Height) > VB6.PixelsToTwipsY(txtMain.Top) Then txtMain.Height = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(sash.Height) - VB6.PixelsToTwipsY(txtMain.Top))

        txtMain.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(sash.Left) + VB6.PixelsToTwipsX(sash.Width))
        'fraFind.Left = txtMain.Left
        newWidth = VB6.PixelsToTwipsX(Width) - VB6.PixelsToTwipsX(txtMain.Left) - 100
        If newWidth > 0 Then
            txtMain.Width = VB6.TwipsToPixelsX(newWidth)
            'If fraFind.Visible Then
            '	fraFind.Width = VB6.TwipsToPixelsX(newWidth)
            '	If (newWidth - 324 - VB6.PixelsToTwipsX(cmdFind.Width) - VB6.PixelsToTwipsX(cmdReplace.Width)) > 100 Then
            '		txtFind.Width = VB6.TwipsToPixelsX((newWidth - VB6.PixelsToTwipsX(cmdFind.Width) - VB6.PixelsToTwipsX(cmdReplace.Width) - 324) / 2)
            '		cmdReplace.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(txtFind.Left) + VB6.PixelsToTwipsX(txtFind.Width) + 108)
            '		txtReplace.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(cmdReplace.Left) + VB6.PixelsToTwipsX(cmdReplace.Width) + 108)
            '		txtReplace.Width = txtFind.Width
            '	End If
            'End If
        End If
    End Sub

    Private Sub frmMain_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If QuerySave() = MsgBoxResult.Cancel Then
            e.Cancel = True
        ElseIf QuerySaveProject() = MsgBoxResult.Cancel Then
            e.Cancel = True
        Else
            'UPGRADE_ISSUE: App property App.HelpFile was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
            'SaveSetting(pAppName, "Files", "Help", App.HelpFile)

            SaveSetting(pAppName, "Defaults", "BaseName", pBaseName)
            SaveSetting(pAppName, "Defaults", "Path", path)
            SaveSetting(pAppName, "Defaults", "FindTimeout", CStr(FindTimeout))
            SaveSetting(pAppName, "Defaults", "ViewFormatting", CStr(ViewFormatting))
            SaveSetting(pAppName, "Defaults", "FormatWhileTyping", CStr(FormatWhileTyping))
            SaveSetting(pAppName, "Defaults", "AutoParagraph", CStr(mnuAutoParagraph.Checked))

            SaveSetting(pAppName, SectionMainWin, "Width", CStr(VB6.PixelsToTwipsX(Width)))
            SaveSetting(pAppName, SectionMainWin, "Height", CStr(VB6.PixelsToTwipsY(Height)))
            SaveSetting(pAppName, SectionMainWin, "Left", CStr(VB6.PixelsToTwipsX(Left)))
            SaveSetting(pAppName, SectionMainWin, "Top", CStr(VB6.PixelsToTwipsY(Top)))
            SaveSetting(pAppName, SectionMainWin, "TreeWidth", CStr(VB6.PixelsToTwipsX(sash.Left)))
            Dim rf As Integer
            For rf = mnuRecent.Count - 1 To 1 Step -1
                SaveSetting(pAppName, SectionRecentFiles, CStr(rf), mnuRecent(rf).Tag)
            Next rf
            While GetSetting(pAppName, SectionRecentFiles, CStr(rf)) <> ""
                SaveSetting(pAppName, SectionRecentFiles, CStr(rf), "")
                rf += 1
            End While

            For Each lOpenForm As Form In My.Application.OpenForms
                If Not lOpenForm Is Me Then lOpenForm.Close()
            Next
        End If
    End Sub

    Private Sub fraFind_MouseDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles fraFind.MouseDown
        Dim Button As Short = eventArgs.Button \ &H100000
        Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
        Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
        Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
        If Button = VB6.MouseButtonConstants.RightButton Or Shift = System.Windows.Forms.Keys.ShiftKey Then
            fraFind.Visible = False
            frmMain_Resize(Me, New System.EventArgs())
        End If
    End Sub

    Public Sub mnuAutoSave_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuAutoSave.Click
        mnuAutoSave.Checked = Not mnuAutoSave.Checked
    End Sub

    Public Sub mnuContext_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuContext.Click
        Dim Index As Short = mnuContext.GetIndex(eventSender)
        ContextAction(mnuContext(Index).Text)
    End Sub

    Public Sub ContextAction(ByRef cmd As String)
        Dim filename, PathName As String
        Select Case cmd
            Case pCaptureReplace
                filename = ReplaceString(SubTagValue("src"), "/", "\")
                filename = IO.Path.GetDirectoryName(path & "\" & NodeFile()) & "\" & filename
                frmCapture.Filename = filename
                frmCapture.Show()
            Case pCaptureNew, BrowseImage
                cdlgOpen.ShowDialog()
                cdlgSave.FileName = cdlgOpen.FileName
                filename = cdlgOpen.FileName
                If Len(filename) > 0 Then
                    PathName = IO.Path.GetDirectoryName(path & "\" & NodeFile())
                    filename = HTMLRelativeFilename(filename, PathName)
                End If
                If closeTagPos > openTagPos + 4 Then
                    EditSubTag("src", filename)
                Else
                    txtMain.Text = VB.Left(txtMain.Text, txtMain.SelectionStart) & "<img src=""" & filename & """>" & Mid(txtMain.Text, txtMain.SelectionStart + 1)
                End If
                If cmd = pCaptureNew Then
                    frmCapture.Filename = filename
                    frmCapture.Show()
                End If
            Case ViewImage
                filename = ReplaceString(SubTagValue("src"), "/", "\")
                filename = IO.Path.GetDirectoryName(path & "\" & NodeFile()) & "\" & filename
                If IO.File.Exists(filename) Then OpenFile(filename)
            Case DeleteTag
                If closeTagPos > openTagPos + 4 Then txtMain.Text = VB.Left(txtMain.Text, openTagPos - 1) & Mid(txtMain.Text, closeTagPos + 1)
            Case SelectLink
                NodeLinking = tree1.SelectedItem.Index
                Me.Cursor = System.Windows.Forms.Cursors.UpArrow
            Case Else : MsgBox("Unrecognized menu item: " & cmd, MsgBoxStyle.OkOnly, "AuthorDoc")
        End Select
    End Sub

    Public Sub mnuConvert_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuConvert.Click
        If QuerySave() <> MsgBoxResult.Cancel Then
            If QuerySaveProject() <> MsgBoxResult.Cancel Then frmConvert.Show()
        End If
    End Sub

    Public Sub mnuCopy_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuCopy.Click
        My.Computer.Clipboard.SetText(txtMain.SelectedText)
    End Sub

    Public Sub mnuCut_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuCut.Click
        My.Computer.Clipboard.SetText(txtMain.SelectedText)
        txtMain.SelectedText = ""
    End Sub

    Private Sub mnuEditProject_Click()
        If tree1.Visible Then
            If QuerySaveProject() <> MsgBoxResult.Cancel Then
                LoadTextboxFromFile(IO.Path.GetDirectoryName(pProjectFileName), FilenameOnly(pProjectFileName), "." & FileExt(pProjectFileName), txtMain)
                tree1.Visible = False
            End If
        Else
            If QuerySave() <> MsgBoxResult.Cancel Then
                tree1.Visible = True
                mnuRecent_Click(mnuRecent.Item(1), New System.EventArgs())
            End If
        End If
    End Sub

    Public Sub mnuExit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuExit.Click
        Close()
    End Sub

    Public Sub mnuFind_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuFind.Click
        If fraFind.Visible Then
            cmdFind_MouseUp(cmdFind, New System.Windows.Forms.MouseEventArgs(VB6.MouseButtonConstants.LeftButton * &H100000, 0, 0, 0, 0))
        Else
            fraFind.Visible = True
            frmMain_Resize(Me, New System.EventArgs())
        End If
    End Sub

    Public Sub mnuFindSelection_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuFindSelection.Click
        '    Case 6 'Control-F = find
        If Not fraFind.Visible Then
            fraFind.Visible = True
            frmMain_Resize(Me, New System.EventArgs())
        End If
        Dim SelEnd, selStart, txtLen As Integer
        If Len(txtMain.SelectedText) < 1 Then
            txtLen = Len(txtMain.Text)
            SelEnd = txtMain.SelectionStart
            selStart = txtMain.SelectionStart
            Do While selStart > 0
                If IsAlphaNumeric(Mid(txtMain.Text, selStart, 1)) Then
                    selStart = selStart - 1
                Else
                    Exit Do
                End If
            Loop
            Do While SelEnd <= txtLen
                If IsAlphaNumeric(Mid(txtMain.Text, SelEnd + 1, 1)) Then
                    SelEnd = SelEnd + 1
                Else
                    Exit Do
                End If
            Loop
            txtMain.SelectionStart = selStart
            txtMain.SelectionLength = SelEnd - selStart
        End If
        txtFind.Text = txtMain.SelectedText
        cmdFind.Focus()
    End Sub

    Public Sub mnuFormatting_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuFormatting.Click
        mnuFormatting.Checked = Not mnuFormatting.Checked
        ViewFormatting = mnuFormatting.Checked
        If ViewFormatting Then
            FormatText(txtMain)
        Else
            txtMain.Text = txtMain.Text
            txtMain.Refresh()
        End If
    End Sub

    Public Sub mnuFormatWhileTyping_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuFormatWhileTyping.Click
        mnuFormatWhileTyping.Checked = Not mnuFormatWhileTyping.Checked
        FormatWhileTyping = mnuFormatWhileTyping.Checked
        If FormatWhileTyping Then
            If Not ViewFormatting Then
                mnuFormatting_Click(mnuFormatting, New System.EventArgs())
            Else
                FormatText(txtMain)
            End If
        End If
    End Sub

    Public Sub mnuHelpAbout_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuHelpAbout.Click
        logger.msg("AuthorDoc" & vbCr & "Version " & My.Application.Info.Version.Major & "." & My.Application.Info.Version.Minor & vbCr & "Aqua Terra Consultants", MsgBoxStyle.OkOnly, "About AuthorDoc")
    End Sub

    Public Sub mnuHelpContents_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuHelpContents.Click
        'Dim newHelpfile As String
        ''UPGRADE_ISSUE: MSComDlg.CommonDialog control cdlg was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E047632-2D91-44D6-B2A3-0801707AF686"'
        ''UPGRADE_ISSUE: App property App.HelpFile was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
        'newHelpfile = OpenFile(App.HelpFile, cdlg)
        ''UPGRADE_ISSUE: App property App.HelpFile was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
        'If newHelpfile <> App.HelpFile Then
        '    If IO.File.Exists(newHelpfile) Then
        '        'UPGRADE_ISSUE: App property App.HelpFile was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
        '        App.HelpFile = newHelpfile
        '        'UPGRADE_ISSUE: App property App.HelpFile was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
        '        SaveSetting(pAppName, "Files", "Help", App.HelpFile)
        '    End If
        'End If
    End Sub

    Public Sub mnuHelpWebsite_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuHelpWebsite.Click
        ShellExecute(Me.Handle.ToInt32, "Open", "http://hspf.com/pub/authordoc", CStr(0), CStr(0), 0)
    End Sub

    Public Sub mnuLinkSection_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuLinkSection.Click
        GetCurrentTag()
        If tagName = "a" Then
            txtMain.SelectionStart = openTagPos + 9
        Else
            mnuLink_Click(mnuLink, New System.EventArgs())
            GetCurrentTag()
        End If
        ContextAction(SelectLink)
    End Sub

    Public Sub mnuNewProject_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuNewProject.Click
        Dim f As Short

        If QuerySave = MsgBoxResult.Cancel Then Exit Sub
        On Error GoTo ErrNew
        cdlgSave.ShowDialog()
        cdlgOpen.FileName = cdlgSave.FileName
        If Len(cdlgOpen.FileName) > 0 Then
            path = IO.Path.GetDirectoryName((cdlgOpen.FileName))
            ChDir(path)
            If Not IO.Directory.Exists(path) Then MkDir(path)
            f = FreeFile()
            FileOpen(f, cdlgOpen.FileName, OpenMode.Output)
            FileClose(f)
            OpenProject((cdlgOpen.FileName), tree1)
            mnuNewSection.Enabled = True
            ProjectChanged = False
            If tree1.Nodes.Count > 0 Then tree1_NodeClick(tree1, New AxComctlLib.ITreeViewEvents_NodeClickEvent(tree1.Nodes(1)))
        End If
        Exit Sub
ErrNew:
        MsgBox("Error creating new project:" & vbCr & Err.Description)
    End Sub

    Public Sub mnuNewSection_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuNewSection.Click
        Dim found As Boolean
        Dim nodNum As Integer
        Dim key, ThisName, keypath As String
        Dim filename As String
        Dim f As Short

        cdlgOpen.ShowDialog()
        cdlgSave.FileName = cdlgOpen.FileName
        filename = cdlgOpen.FileName
        If Len(filename) > Len(path) Then
            If UCase(VB.Left(filename, Len(path))) <> UCase(path) Then
                MsgBox("Files must be in the same directory as or a subdirectory of the project file's directory.", MsgBoxStyle.OKOnly)
            Else
                If UCase(VB.Right(filename, Len(pSourceExtension))) <> UCase(pSourceExtension) Then
                    filename = filename & pSourceExtension
                End If

                If Not IO.File.Exists(filename) Then
                    keypath = IO.Path.GetDirectoryName(filename)
                    If Not IO.Directory.Exists(keypath) Then IO.Directory.CreateDirectory(keypath)
                    f = FreeFile()
                    FileOpen(f, filename, OpenMode.Output)
                    FileClose(f)
                End If

                ThisName = FilenameOnly(filename)
                key = VB.Left(filename, Len(filename) - Len(pSourceExtension)) 'trim extension .txt
                key = "N" & Mid(key, Len(path) + 2)
                keypath = IO.Path.GetDirectoryName(Mid(key, 2))
                If tree1.Nodes.Count = 0 Then 'This is the first node
                    tree1.Nodes.Add(, , key, ThisName)
                ElseIf keypath = IO.Path.GetDirectoryName(NodeFile) Then  'place after selected sibling
                    tree1.Nodes.Add(tree1.SelectedItem, ComctlLib.TreeRelationshipConstants.tvwNext, key, ThisName)
                Else
                    nodNum = tree1.Nodes.Count
                    found = False
                    While nodNum >= 1 And Not found 'Look for last sibling
                        If IO.Path.GetDirectoryName(NodeFile(nodNum)) = keypath Then
                            tree1.Nodes.Add(tree1.Nodes(nodNum).key, ComctlLib.TreeRelationshipConstants.tvwNext, key, ThisName)
                            found = True
                        End If
                        nodNum = nodNum - 1
                    End While
                    If Not found Then tree1.Nodes.Add(tree1.SelectedItem, ComctlLib.TreeRelationshipConstants.tvwChild, key, ThisName)
                End If
                pCurrentFilename = cdlgOpen.FileName
                ProjectChanged = True
                tree1.Nodes(key).EnsureVisible()
            End If
        End If
    End Sub

    Private Function NodeFile(Optional ByRef nodNum As Integer = 0) As String
        If IsNothing(nodNum) OrElse nodNum = 0 Then
            nodNum = tree1.SelectedItem.Index
        End If
        If nodNum < 1 Then nodNum = 1
        NodeFile = Mid(tree1.Nodes(nodNum).Key, 2)
    End Function

    Public Sub mnuOpenProject_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuOpenProject.Click
        If IO.Directory.Exists(path) Then
            ChDriveDir(path)
        End If

        cdlgOpen.FileName = ""
        cdlgSave.FileName = "" 'BaseName
        cdlgOpen.ShowDialog()
        cdlgSave.FileName = cdlgOpen.FileName
        If Len(cdlgOpen.FileName) > 0 Then
            AddRecentFile((cdlgOpen.FileName))
            mnuRecent_Click(mnuRecent.Item(0), New System.EventArgs())
        End If
    End Sub

    Public Sub mnuOptions_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuOptions.Click
        frmOptions.Show()
    End Sub

    Public Sub mnuPaste_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuPaste.Click
        txtMain.SelectedText = My.Computer.Clipboard.GetText
    End Sub

    Public Sub mnuRecent_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        Dim newFilePath As String = eventSender.Tag
        If newFilePath.Length > 0 Then
            If QuerySaveProject() <> MsgBoxResult.Cancel Then
                If QuerySave() <> MsgBoxResult.Cancel Then
                    path = IO.Path.GetDirectoryName(newFilePath)
                    Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
                    OpenProject(newFilePath, tree1)
                    mnuNewSection.Enabled = True
                    ProjectChanged = False
                    If tree1.Nodes.Count > 0 Then tree1_NodeClick(tree1, New AxComctlLib.ITreeViewEvents_NodeClickEvent(tree1.Nodes(1)))
                    Me.Cursor = System.Windows.Forms.Cursors.Default
                End If
            End If
        End If
    End Sub

    Public Sub mnuRevert_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuRevert.Click
        txtMain.Text = CurrentFileContents
    End Sub

    Public Sub mnuSaveFile_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuSaveFile.Click
        Dim f As Short 'file handle

        f = FreeFile()
        'Kill pCurrentFilename
        FileOpen(f, pCurrentFilename, OpenMode.Output)
        PrintLine(f, txtMain.Text)
        FileClose(f)
        SetFileChanged(False)
        If pCurrentFilename = pProjectFileName Then
            Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
            OpenProject(pProjectFileName, tree1)
            Me.Cursor = System.Windows.Forms.Cursors.Default
        End If
    End Sub

    Public Sub mnuSaveProject_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuSaveProject.Click
        cdlgOpen.FileName = pProjectFileName
        cdlgSave.FileName = pProjectFileName
        cdlgSave.ShowDialog()
        cdlgOpen.FileName = cdlgSave.FileName
        If Len(cdlgOpen.FileName) > 0 Then
            SaveProject((cdlgOpen.FileName), tree1)
        End If
    End Sub

    Public Sub mnuImage_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuImage.Click
        Dim startPos As Integer
        Dim filename, PathName As String
        startPos = txtMain.SelectionStart
        txtMain.SelectionLength = 0
        txtMain.SelectedText = "<img src="""">"
        txtMain.SelectionStart = startPos + 10
        cdlgImageOpen.ShowDialog()
        filename = cdlgImageOpen.FileName
        If Len(filename) > 0 Then
            PathName = IO.Path.GetDirectoryName(path & "\" & NodeFile)
            filename = HTMLRelativeFilename(filename, PathName)
        End If
        txtMain.SelectedText = filename
    End Sub

    Public Sub mnuBold_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuBold.Click
        InsertTag("b")
    End Sub

    Public Sub mnuFigure_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuFigure.Click
        InsertTag("figure")
    End Sub

    Public Sub mnuItalic_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuItalic.Click
        InsertTag("i")
    End Sub

    Public Sub mnuPRE_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuPRE.Click
        InsertTag("pre")
    End Sub

    Public Sub mnuOL_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuOL.Click
        ListTag("ol")
    End Sub

    Private Sub testTextImage()
        '    Dim formatTxt As String
        '    Dim FormatStart, FormatEnd As Integer
        '    FormatStart = InStr(txtMain.Text, Asterisks80)
        '    FormatEnd = InStrRev(txtMain.Text, Asterisks80)
        '    If FormatEnd > FormatStart Then CardImage(Mid(txtMain.Text, FormatStart, FormatEnd - FormatStart))
    End Sub

    Public Sub mnuTextImage_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuTextImage.Click
        mnuTextImage.Checked = Not mnuTextImage.Checked
        If mnuTextImage.Checked Then
            testTextImage()
        Else
            frmSample.Visible = False
        End If
    End Sub

    Public Sub mnuUL_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuUL.Click
        ListTag("ul")
    End Sub

    Public Sub mnuUnderline_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuUnderline.Click
        InsertTag("u")
    End Sub

    Public Sub mnuLink_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuLink.Click
        Dim startPos As Integer
        startPos = txtMain.SelectionStart
        InsertTag("a")
        txtMain.SelectionStart = startPos + 2
        txtMain.SelectedText = " href=""#"""
        txtMain.SelectionStart = startPos + 9
    End Sub

    Private Sub ListTag(ByRef aTag As String)
        Dim startPos, endPos As Integer
        With txtMain
            startPos = .SelectionStart
            endPos = startPos + .SelectionLength
            InsertTag(aTag)
            .SelectionStart = startPos + 4
            .SelectedText = vbCrLf & "<li>"
            If endPos = startPos Then
                .SelectionStart = startPos + 10
                .SelectedText = vbCrLf
                .SelectionStart = startPos + 10
            Else
                startPos = .SelectionStart
                endPos = endPos + 9
                While startPos < endPos
                    startPos = InStr(startPos + 1, .Text, vbCrLf)
                    If startPos = 0 Or startPos >= endPos Then
                        startPos = endPos
                    Else
                        .SelectionStart = startPos + 1
                        .SelectedText = "<li>"
                        endPos = endPos + 4
                    End If
                End While
            End If
        End With
    End Sub

    Private Sub InsertTag(ByRef aTag As String)
        Dim startTag, endtag As String
        Dim startPos, endPos As Integer
        With txtMain
            startPos = .SelectionStart
            endPos = startPos + .SelectionLength

            Select Case LCase(aTag)
                Case "keyword", "indexword"
                    startTag = "<" & aTag & "="""
                    endtag = """>"
                Case Else
                    startTag = "<" & aTag & ">"
                    endtag = "</" & aTag & ">"
            End Select

            If .SelectionLength = 0 Then
                .SelectedText = startTag & endtag
                .SelectionStart = startPos + Len(startTag)
            Else
                .SelectionStart = endPos
                .SelectedText = endtag
                .SelectionStart = startPos
                .SelectedText = startTag
                .SelectionStart = endPos + Len(startTag & endtag)
            End If
        End With
    End Sub

    Private Sub sash_MouseDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles sash.MouseDown
        Dim Button As Short = eventArgs.Button \ &H100000
        Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
        Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
        Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
        SashDragging = True
    End Sub

    Private Sub sash_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles sash.MouseMove
        Dim Button As Short = eventArgs.Button \ &H100000
        Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
        Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
        Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
        Dim newLeftWidth As Integer
        If SashDragging And (VB6.PixelsToTwipsX(sash.Left) + x) > 100 And (VB6.PixelsToTwipsX(sash.Left) + x < VB6.PixelsToTwipsX(Width) - 100) Then
            sash.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(sash.Left) + x)
            newLeftWidth = VB6.PixelsToTwipsX(sash.Left)
            If newLeftWidth > 1000 Then tree1.Width = VB6.TwipsToPixelsX(newLeftWidth)
            frmMain_Resize(Me, New System.EventArgs())
        End If
    End Sub

    Private Sub sash_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles sash.MouseUp
        Dim Button As Short = eventArgs.Button \ &H100000
        Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
        Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
        Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
        SashDragging = False
    End Sub

    Private Sub Timer1_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Timer1.Tick
        'UPGRADE_ISSUE: Timer property Timer1.tag was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        If IsNumeric(Timer1.Tag) Then
            'UPGRADE_ISSUE: Timer property Timer1.tag was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
            tree1.SelectedItem = tree1.Nodes(CShort(Timer1.Tag))
            If txtMain.Text <> WholeFileString(pCurrentFilename) Then SetFileChanged(True)
        End If
        Timer1.Enabled = False
    End Sub

    Private Sub TimerSlowAction_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles TimerSlowAction.Tick
        TimerSlowAction.Enabled = False
        AbortAction = True
    End Sub

    Private Sub tree1_AfterLabelEdit(ByVal eventSender As System.Object, ByVal eventArgs As AxComctlLib.ITreeViewEvents_AfterLabelEditEvent) Handles tree1.AfterLabelEdit
        Dim OldFilePath As String
        With tree1.SelectedItem
            OldFilePath = path & "\" & Mid(.Key, 2) & pSourceExtension
            If IO.File.Exists(OldFilePath) Then
                Select Case MsgBox("Rename file '" & OldFilePath & "' to '" & eventArgs.newString & "?", MsgBoxStyle.YesNoCancel)
                    Case MsgBoxResult.No : .Text = eventArgs.newString : .Key = "N" & .FullPath
                    Case MsgBoxResult.Yes : .Text = eventArgs.newString : .Key = "N" & .FullPath
                        'Name OldFilePath As path & "\" & .fullpath & pSourceExtension
                        Rename(OldFilePath, IO.Path.GetDirectoryName(OldFilePath) & "\" & .Text & pSourceExtension)
                    Case MsgBoxResult.Cancel
                        eventArgs.cancel = True
                End Select
            End If
        End With
    End Sub

    Private Sub tree1_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxComctlLib.ITreeViewEvents_KeyDownEvent) Handles tree1.KeyDownEvent
        Select Case eventArgs.keyCode
            Case System.Windows.Forms.Keys.Delete : tree1.Nodes.Remove(tree1.SelectedItem.Index)
                'Case vbKeyInsert:
                ' Dim nod As ComctlLib.Node = tree1.Nodes.add(tree1.SelectedItem, tvwPrevious, "NewFile", "NewFile")
        End Select
    End Sub

    'A horrible hack to get around the tree control's penchant for changing
    'the selected node after we have lost control
    Private Sub DelaySetNode(ByRef nodeNum As Integer)
        'UPGRADE_ISSUE: Timer property Timer1.tag was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        Timer1.Tag = nodeNum
        Timer1.Enabled = True
    End Sub

    Private Sub tree1_NodeClick(ByVal eventSender As System.Object, ByVal eventArgs As AxComctlLib.ITreeViewEvents_NodeClickEvent) Handles tree1.NodeClick
        Dim filename, fullpath As String
        Dim inClick As Boolean
        If Not inClick And Not Timer1.Enabled Then
            inClick = True
            If NodeLinking > 0 Then
                fullpath = "c:\" & IO.Path.GetDirectoryName(NodeFile(NodeLinking))
                filename = HTMLRelativeFilename("c:\" & Mid(eventArgs.node.Key, 2), fullpath)
                EditSubTag("href", filename)
                DelaySetNode(NodeLinking)
                NodeLinking = 0
                Me.Cursor = System.Windows.Forms.Cursors.Default
            Else
                filename = Mid(eventArgs.node.Key, 2)
                If QuerySave() = MsgBoxResult.Cancel Then 'Should move focus back to old node here
                    DelaySetNode(1)
                Else
                    LoadTextboxFromFile(path, filename, pSourceExtension, txtMain)
                    If tree1.SelectedItem Is Nothing Then
                        tree1.SelectedItem = eventArgs.node
                    ElseIf tree1.SelectedItem.Index <> eventArgs.node.Index Then
                        tree1.SelectedItem = eventArgs.node
                    End If
                End If
            End If
        End If
        inClick = False
    End Sub

    Public Sub LoadTextboxFromFile(ByRef fullpath As String, ByRef filename As String, ByRef ext As String, ByRef txtBox As System.Windows.Forms.RichTextBox)
        Static LastAnswer As MsgBoxResult
        Dim altExt, altpath As String
        Dim thisAnswer As MsgBoxResult
        If Not IO.File.Exists(fullpath & "\" & filename & ext) Then 'Check for files named .html or pSourceExtension
            If LCase(ext) = LCase(pSourceExtension) Then altExt = ".html" Else altExt = pSourceExtension
            altpath = fullpath & "\" & filename & altExt
            If IO.File.Exists(altpath) Then
                If altExt = pSourceExtension Then
                    ext = pSourceExtension
                Else
                    If LastAnswer = 0 Then
                        thisAnswer = MsgBox("File " & filename & pSourceExtension & " was not found, use " & filename & ".html instead?", MsgBoxStyle.YesNoCancel, path)
                        If thisAnswer = MsgBoxResult.Cancel Then Exit Sub
                        LastAnswer = MsgBox("Treat other missing files the same way?", MsgBoxStyle.YesNo)
                        If LastAnswer = MsgBoxResult.Yes Then LastAnswer = thisAnswer Else LastAnswer = 0
                    Else
                        thisAnswer = LastAnswer
                    End If
                    If thisAnswer = MsgBoxResult.Yes Then FileCopy(altpath, fullpath & "\" & filename & ext)
                End If
            End If
        End If
        ReadFile(fullpath & "\" & filename & ext, txtBox)
    End Sub

    Private Sub ReadFile(ByRef filename As String, ByRef txtBox As System.Windows.Forms.RichTextBox)
        Dim f As Short 'file handle
        Dim FileLength As Integer
        f = FreeFile()
        On Error GoTo nofile
OpenFile:
        FileOpen(f, filename, OpenMode.Input)
        On Error GoTo 0
        FileLength = LOF(f)
        If txtBox.Name = "txtMain" Then
            pCurrentFilename = filename
            Text = pCurrentFilename
            CurrentFileContents = InputString(f, FileLength)
            txtBox.Text = CurrentFileContents
            If ViewFormatting Then FormatText(txtBox)
        Else
            txtBox.Text = InputString(f, FileLength)
        End If
        FileClose(f)
        SetFileChanged(False)
        Exit Sub
nofile:
        txtBox.Text = "(no file)"
        If MsgBox("File '" & filename & "' does not exist. Create it?", MsgBoxStyle.YesNo, "Missing file") = MsgBoxResult.Yes Then
            Err.Clear()
            On Error Resume Next
            If Not IO.Directory.Exists(IO.Path.GetDirectoryName(filename)) Then
                IO.Directory.CreateDirectory(IO.Path.GetDirectoryName(filename))
            End If
            On Error GoTo errCreate
            FileOpen(f, filename, OpenMode.Output)
            PrintLine(f, "")
            FileClose(f)
            GoTo OpenFile
        Else
            If txtBox.Name = "txtMain" Then
                pCurrentFilename = filename
                Text = pCurrentFilename
            End If
            GoTo endsub
        End If
errCreate:
        MsgBox("Could not create file '" & filename & "'" & vbCr & Err.Description)
endsub:
        SetFileChanged(False)
    End Sub

    Private Sub SetFileChanged(ByRef newValue As Boolean)
        If Changed <> newValue Then
            Changed = newValue
            mnuSaveFile.Enabled = Changed
            If Changed Then
                Text = pCurrentFilename & " (edited)"
            Else
                Text = pCurrentFilename
            End If
        End If
    End Sub

    Private Function RTF_START(ByRef txtBox As System.Windows.Forms.RichTextBox) As Object
        'RTF_START = "{\rtf1\ansi\deff0{\fonttbl{\f0\fswiss MS Sans Serif;}}\pard\plain\fs17 "
        'UPGRADE_WARNING: Couldn't resolve default property of object RTF_START. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        RTF_START = "{\rtf1\ansi\deff0{\fonttbl{\f0\fswiss MS Sans Serif;}{\f1\froman\fcharset2 Symbol;}{\f2\f" & txtBox.Font.Name & ";}{\f3\fmodern Courier New;}}"
    End Function

    Private Sub FormatTextSelection(ByRef txtBox As System.Windows.Forms.RichTextBox, ByRef startPos As Integer, ByRef endPos As Integer, ByRef aCommand As String)
        txtBox.SelectionStart = startPos
        txtBox.SelectionLength = endPos - startPos
        Select Case aCommand
            Case "bold" : txtBox.Font = VB6.FontChangeBold(txtBox.SelectionFont, True)
            Case "italic" : txtBox.SelectionFont = VB6.FontChangeItalic(txtBox.SelectionFont, True)
            Case "underline" : txtBox.SelectionFont = VB6.FontChangeUnderline(txtBox.SelectionFont, True)
            Case "bullet" : txtBox.SelectionBullet = True
            Case Else : txtBox.SelectionFont = VB6.FontChangeName(txtBox.SelectionFont, aCommand)
        End Select
    End Sub

    Private Function FormatText(ByRef txtBox As System.Windows.Forms.RichTextBox) As Object
        Dim txt, lTag As String
        Dim nextch, maxch As Integer
        Dim closeTag As Integer
        Dim selStart, SelLength As Integer
        txtBox.Visible = False
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
        selStart = txtBox.SelectionStart
        SelLength = txtBox.SelectionLength
        AbortAction = False
        TimerSlowAction.Enabled = True
        txt = txtBox.Text

        'clear formatting
        txtBox.SelectionStart = 0
        txtBox.SelectionLength = Len(txt)
        txtBox.Font = VB6.FontChangeBold(txtBox.SelectionFont, False)
        txtBox.SelectionFont = VB6.FontChangeItalic(txtBox.SelectionFont, False)
        txtBox.SelectionFont = VB6.FontChangeUnderline(txtBox.SelectionFont, False)
        'txtBox.SelBullet = False

        maxch = Len(txt)

        While nextch < maxch And Not AbortAction
            nextch = InStr(nextch + 1, txt, "<")
            If nextch = 0 Then
                nextch = maxch
            Else
                lTag = Mid(txt, nextch + 1, 2)
                closeTag = InStr(nextch + 1, txt, "</" & lTag)
                If closeTag > 0 Then
                    Select Case LCase(lTag)
                        Case "h>", "b>" : FormatTextSelection(txtBox, nextch + 2, closeTag - 1, "bold")
                        Case "i>" : FormatTextSelection(txtBox, nextch + 2, closeTag - 1, "italic")
                        Case "u>" : FormatTextSelection(txtBox, nextch + 2, closeTag - 1, "underline")
                        Case "pr" : FormatTextSelection(txtBox, nextch + 4, closeTag - 1, "Courier New")
                    End Select
                    'Else
                    '  If LCase(tag) = "li" Then FormatTextSelection txtBox, nextch + 2, nextch + 3, "bullet"
                End If
            End If
        End While
        If AbortAction Then
            If InStr(Text, "(formatting aborted)") = 0 Then Text = Text & " (formatting aborted)"
        End If
        txtBox.SelectionStart = selStart
        txtBox.SelectionLength = SelLength
        TimerSlowAction.Enabled = False
        txtBox.Visible = True
        txtBox.Focus()
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Return Nothing 'TODO: should this be a sub? all calls do not use return argument!
    End Function

    'Private Function FormatTextOld(txtBox As RichTextBox)
    '  Dim rtf$
    '  Dim nextch&, maxch&
    '  Dim openTag&, closeTag&, parenlevel&, spacepos&
    '  AbortAction = False
    '  TimerSlowAction.Enabled = True
    '  rtf = ReplaceString(txtBox.Text, "\", "\\")
    '  rtf = ReplaceString(txtBox.Text, "{", "\{")
    '  rtf = ReplaceString(txtBox.Text, "}", "\}")
    '
    '  If Not AbortAction Then rtf = ReplaceString(rtf, "<h", RTF_BOLD & "<h")
    '  If Not AbortAction Then rtf = ReplaceString(rtf, "</h", RTF_BOLD_END & "</h")
    '
    '  If Not AbortAction Then rtf = ReplaceString(rtf, "<u>", "<u>" & RTF_UNDERLINE)
    '  If Not AbortAction Then rtf = ReplaceString(rtf, "<b>", "<b>" & RTF_BOLD)
    '  If Not AbortAction Then rtf = ReplaceString(rtf, "<i>", "<i>" & RTF_ITALIC)
    '
    '  If Not AbortAction Then rtf = ReplaceString(rtf, "</u>", RTF_UNDERLINE_END & "</u>")
    '  If Not AbortAction Then rtf = ReplaceString(rtf, "</b>", RTF_BOLD_END & "</b>")
    '  If Not AbortAction Then rtf = ReplaceString(rtf, "</i>", RTF_ITALIC_END & "</i>")
    '
    '  If Not AbortAction Then rtf = ReplaceString(rtf, vbCrLf, RTF_PARAGRAPH)
    '
    '  'make sure text ends with a newline
    '  If Right(rtf, 2 * Len(RTF_PARAGRAPH)) <> RTF_PARAGRAPH & RTF_PARAGRAPH Then
    '    rtf = rtf & RTF_PARAGRAPH & RTF_PARAGRAPH
    '  End If
    '  If AbortAction And InStr(caption, "(formatting aborted)") = 0 Then
    '    caption = caption & " (formatting aborted)"
    '  End If
    '  rtf = RTF_START(txtBox) & rtf & RTF_END
    '
    '  If rtf <> txtBox.TextRTF Then
    '    Dim selStart&, SelLength&
    '    selStart = txtBox.selStart
    '    SelLength = txtBox.SelLength
    '    txtBox.TextRTF = rtf
    '    txtBox.selStart = selStart
    '    txtBox.SelLength = SelLength
    '  End If
    '  TimerSlowAction.Enabled = False
    'End Function

    Private Function QuerySaveProject() As MsgBoxResult
        Dim retval As MsgBoxResult
        retval = MsgBoxResult.Yes
        If ProjectChanged Then
            retval = MsgBox("Save changes to " & pProjectFileName & "?", MsgBoxStyle.YesNoCancel)
            If retval = MsgBoxResult.Yes Then SaveProject(pProjectFileName, tree1)
            ProjectChanged = False
        End If
        QuerySaveProject = retval
    End Function

    Private Function QuerySave() As MsgBoxResult
        Dim retval As MsgBoxResult
        retval = MsgBoxResult.Yes
        If Changed Then
            If Not mnuAutoSave.Checked Then
                retval = MsgBox("Save changes to " & pCurrentFilename & "?", MsgBoxStyle.YesNoCancel)
            End If
            If retval = MsgBoxResult.Yes Then mnuSaveFile_Click(mnuSaveFile, New System.EventArgs())
        End If
        QuerySave = retval
    End Function

    Private Sub txtFind_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtFind.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.Return Then cmdFind_MouseUp(cmdFind, New System.Windows.Forms.MouseEventArgs(VB6.MouseButtonConstants.LeftButton * &H100000, 0, 0, 0, 0))
    End Sub

    Private Sub txtMain_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtMain.KeyUp
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        If FormatWhileTyping Then FormatText(txtMain)
    End Sub

    Private Sub txtReplace_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtReplace.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.Return Then cmdReplace_MouseUp(cmdReplace, New System.Windows.Forms.MouseEventArgs(VB6.MouseButtonConstants.LeftButton * &H100000, 0, 0, 0, 0))
    End Sub

    Private Sub txtMain_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtMain.TextChanged
        Static InChange As Boolean
        If Not InChange And Not Undoing Then
            InChange = True

            Undos(UndoPos) = txtMain.Text
            UndoCursor(UndoPos) = txtMain.SelectionStart
            UndoPos = UndoPos + 1
            If UndoPos > MaxUndo Then UndoPos = 0
            If UndosAvail < MaxUndo Then UndosAvail = UndosAvail + 1

            If CurrentFileContents <> txtMain.Text Then
                If Not Changed Then SetFileChanged(True)
            Else
                If Changed Then SetFileChanged(False)
            End If
            mnuSaveFile.Enabled = Changed
            If mnuTextImage.Checked Then testTextImage()
            InChange = False
        End If
    End Sub

    Private Sub txtMain_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtMain.Click
        Dim mnuItem As Integer
        Dim filename, PathName As String
        Dim txt As String
        txt = txtMain.Text
        System.Windows.Forms.Application.DoEvents()
        For mnuItem = mnuContext.Count - 1 To 1 Step -1
            mnuContext.Unload(mnuItem)
        Next mnuItem
        GetCurrentTag() 'txt, txtMain.SelStart, tagName, openTagPos, closeTagPos
        Dim hashPos As Integer
        If openTagPos < closeTagPos Then
            Select Case tagName
                Case "img"
                    AddContextMenuItem(pCaptureReplace)
                    AddContextMenuItem(pCaptureNew)
                    AddContextMenuItem(BrowseImage)
                    AddContextMenuItem(ViewImage)
                    'filename = SubTagValue("src")
                    'filename = ReplaceString(filename, "/", "\")
                    'pathname = AbsolutePath(filename, IO.Path.GetDirectoryName(path & "\" & NodeFile))
                    'If Len(Dir(pathname)) > 0 Then frmSample.SetImage pathname
                Case "a"
                    AddContextMenuItem(SelectLink)
                    filename = SubTagValue("href")
                    hashPos = InStr(filename, "#")
                    If hashPos > 0 Then filename = VB.Left(filename, hashPos - 1)
                    If Len(filename) > 0 Then
                        filename = ReplaceString(filename, "/", "\")
                        If VB.Left(filename, 1) = "\" Then
                            PathName = path & filename
                        Else
                            PathName = IO.Path.GetDirectoryName(path & "\" & NodeFile()) & "\" & filename
                        End If
                        If IO.File.Exists(PathName) Then
                            frmSample.SetText(PathName)
                        ElseIf IO.File.Exists(PathName & pSourceExtension) Then
                            frmSample.SetText(PathName & pSourceExtension)
                        End If
                    End If
            End Select
            'If txtMainButton = vbRightButton Then PopupMenu mnuContextTop
        End If
    End Sub

    Private Sub txtMain_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtMain.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        Dim oldStart As Integer
        Select Case KeyAscii
            Case 26 'Control-Z = undo
                If UndosAvail > 0 Then
                    Undoing = True
                    UndoPos = UndoPos - 1
                    If UndoPos < 0 Then UndoPos = MaxUndo
                    txtMain.Text = Undos(UndoPos)
                    txtMain.SelectionStart = UndoCursor(UndoPos)
                    UndosAvail = UndosAvail - 1
                    Undoing = False
                End If
            Case 13
                If mnuAutoParagraph.Checked Then
                    oldStart = txtMain.SelectionStart
                    txtMain.Text = VB.Left(txtMain.Text, oldStart) & "<p>" & Mid(txtMain.Text, oldStart + 1)
                    txtMain.SelectionStart = oldStart + 3
                End If
        End Select
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtMain_MouseDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles txtMain.MouseDown
        Dim Button As Short = eventArgs.Button \ &H100000
        Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
        Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
        Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
        'txtMainButton = Button
    End Sub

    'Search in string txt for a tag that encloses start character position
    'Sets tagName to lowercase of first word in tag
    'Sets openTagPos, closeTagPos to string index of < and > of tag in txt
    Private Sub GetCurrentTag() 'txt$, start&, tagName$, openTagPos&, closeTagPos&)
        Dim txt As String
        Dim start As Integer
        txt = txtMain.Text
        start = txtMain.SelectionStart
        If start < 1 Then Exit Sub
        openTagPos = InStrRev(txt, "<", start)
        If openTagPos > 0 Then
            closeTagPos = InStrRev(txt, ">", start)
            If closeTagPos < openTagPos Then 'we are in a tag
                closeTagPos = InStr(start, txt, ">")
            End If
        End If
        Dim endNamePos As Integer
        If openTagPos > 0 And openTagPos <= start And closeTagPos >= start Then
            endNamePos = InStr(openTagPos, txt, " ")
            If endNamePos = 0 Or endNamePos > closeTagPos Then endNamePos = closeTagPos
            tagName = LCase(Mid(txt, openTagPos + 1, endNamePos - openTagPos - 1))
        Else
            openTagPos = 0
            closeTagPos = 0
            tagName = ""
        End If
    End Sub

    'Uses current tag delimited by openTagPos and closeTagPos
    'Sets subtagName$, value$
    'QuotedStringFromTag( st, v, 1 ) when the current tag is <img src="foo.png">
    'will result in st="src", v="foo.png"
    'Private Sub QuotedStringFromTag(subtagName$, value$, Optional stringNum& = 1)
    '  Dim valueStart&, valueEnd&, subtagStart&, num&
    '  Dim txt$
    '  txt = txtMain.Text
    '
    '  valueStart = InStr(openTagPos, txt, """") + 1
    '  num = 1
    '  While num < stringNum And valueStart > 0
    '    valueStart = InStr(valueStart, txt, """") + 1 'find close quote
    '    If valueStart > 0 Then
    '      valueStart = InStr(valueStart, txt, """") + 1 'find next open quote
    '    End If
    '    num = num + 1
    '  Wend
    '  valueEnd = InStr(valueStart, txt, """")
    '  If valueStart > 0 And valueEnd > valueStart And valueEnd < closeTagPos Then
    '    value = Mid(txt, valueStart, valueEnd - valueStart)
    '    subtagStart = InStrRev(txt, " ", valueStart)
    '    If subtagStart < openTagPos Then subtagStart = openTagPos
    '    subtagName = Mid(txt, subtagStart + 1, valueStart - subtagStart - 3)
    '  Else
    '    subtagName = ""
    '    value = ""
    '  End If
    '
    'End Sub

    'Uses current tag delimited by openTagPos and closeTagPos
    'If subtagName does not exist in the current tag, "" is returned.
    'SubTagValue( "src" ) when the current tag is <img src="foo.png">
    'will return foo.png
    Private Function SubTagValue(ByRef subtagName As String) As String
        Dim subtagStart, valueStart, valueEnd, selStart As Integer
        Dim retval As String = ""
        Dim txt, lTag As String
        txt = txtMain.Text
        selStart = txtMain.SelectionStart
        lTag = LCase(Mid(txt, openTagPos, closeTagPos - openTagPos + 1))
        subtagStart = InStr(1, lTag, LCase(subtagName))
        If subtagStart = 0 Then
            retval = ""
        Else
            valueStart = subtagStart + Len(subtagName) + 1
            If Mid(lTag, valueStart, 1) = """" Then
                valueStart = valueStart + 1
                valueEnd = InStr(valueStart, lTag, """")
            Else
                valueEnd = InStr(valueStart + 1, lTag, " ")
                If valueEnd = 0 Then valueEnd = Len(lTag)
            End If
            If valueEnd > valueStart Then retval = Mid(lTag, valueStart, valueEnd - valueStart)
        End If
        SubTagValue = retval
    End Function

    'Uses current tag delimited by openTagPos and closeTagPos
    'Modifies txtMain.Text, replacing current value of subtagName with NewValue
    'If subtagName does not exist in the current tag, it is added at the end
    'EditSubTag( "src", "bar.gif" ) when the current tag is <img src="foo.png">
    'will result in <img src="bar.gif">
    Private Sub EditSubTag(ByRef subtagName As String, ByRef newValue As String)
        Dim valueEnd, valueStart, subtagStart As Integer
        Dim txt, lTag As String
        txt = txtMain.Text
        lTag = LCase(Mid(txt, openTagPos, closeTagPos - openTagPos + 1))
        subtagStart = InStr(1, lTag, LCase(subtagName))
        If subtagStart = 0 Then
            txtMain.Text = VB.Left(txt, closeTagPos - 1) & " " & LCase(subtagName) & "=" & newValue & Mid(txt, closeTagPos)
        Else
            'subtagStart = subtagStart + openTagPos
            valueStart = subtagStart + Len(subtagName) + 1
            If Mid(lTag, valueStart, 1) = """" Then
                valueEnd = InStr(valueStart + 1, lTag, """")
            Else
                valueEnd = InStr(valueStart + 1, lTag, " ")
                If valueEnd = 0 Then valueEnd = Len(lTag)
            End If
            txtMain.Text = VB.Left(txt, openTagPos + valueStart - 1) & newValue & Mid(txt, openTagPos + valueEnd - 1)
            closeTagPos = InStr(openTagPos + 1, txtMain.Text, ">")
        End If
        txtMain.SelectionStart = openTagPos + 1
        txtMain_Click(txtMain, New System.EventArgs())
        txtMain.SelectionStart = closeTagPos + 1
    End Sub

    Private Sub AddContextMenuItem(ByRef newItem As String)
        Dim mnuItem As Integer
        mnuItem = mnuContext.Count
        mnuContext.Load(mnuItem)
        mnuContext(mnuItem).Text = newItem
    End Sub

    Private Sub txtMain_SelectionChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtMain.SelectionChanged
        Dim lastSelStart As Integer
        If txtMain.SelectionStart <> lastSelStart Then

            lastSelStart = txtMain.SelectionStart
        End If
    End Sub

    Friend Sub AddRecentFile(ByRef FilePath As String)
        Dim rf, rfMove As Integer
        Dim match As Boolean = False
        rf = 0
        For Each lRecent As ToolStripMenuItem In mnuRecent 'as While Not match And rf <= mnuRecent.Count - 1
            If lRecent.Tag.ToString.ToUpper = FilePath.ToUpper Then
                match = True
                Exit For
            End If
            rf += 1
        Next ' End While
        If match Then 'move file to top of list
            For rfMove = rf To 1 Step -1
                mnuRecent(rfMove).Tag = mnuRecent(rfMove - 1).Tag
                mnuRecent(rfMove).Text = "&" & rfMove + 1 & " " & FilenameOnly(mnuRecent(rfMove).Tag)
            Next rfMove
            mnuRecent(0).tag = FilePath
            mnuRecent(0).text = "&1 " & FilenameOnly(FilePath)
        Else 'Add file to list
            mnuRecentSeparator.Visible = True
            Dim lToolStripMenuItem As New ToolStripMenuItem
            With lToolStripMenuItem
                .Tag = FilePath
                .Visible = True
            End With
            mnuRecent.Insert(0, lToolStripMenuItem)
            mnuFile.DropDownItems.Insert(mnuFile.DropDownItems.IndexOf(mnuRecentSeparator) + 1, lToolStripMenuItem)
            AddHandler lToolStripMenuItem.Click, AddressOf mnuRecent_Click

            Dim lRecentIndex As Integer = 1
            For lRecentIndex = mnuRecent.Count - 1 To 0 Step -1
                If lRecentIndex >= MaxRecentFiles Then
                    mnuFile.DropDownItems.Remove(mnuRecent.Item(lRecentIndex))
                    mnuRecent.RemoveAt(lRecentIndex)
                Else
                    mnuRecent(lRecentIndex).Text = "&" & lRecentIndex + 1 & " " & FilenameOnly(mnuRecent(lRecentIndex).Tag)
                End If
            Next
        End If
    End Sub
End Class