Option Strict Off
Option Explicit On

Imports VB = Microsoft.VisualBasic
Imports MapWinUtility
Imports atcUtility

Friend Class frmMain
	Inherits System.Windows.Forms.Form
    'Copyright 2000-2008 by AQUA TERRA Consultants

    Private mMnuRecent As New ArrayList
    Private mPath As String
    Private mCurrentFileContents As String 'What was last saved or retrieved from pCurrentFilename
    Private mMaxUndo As Integer = 10
    Private mUndos(mMaxUndo) As String
    Private mUndoCursor(mMaxUndo) As Integer
    Private mUndoPos As Integer
    Private mUndosAvail As Integer
    Private mUndoing As Boolean
    Private mChanged As Boolean 'True if txtMain.Text has been edited
    Private mProjectChanged As Boolean
    Private mViewFormatting As Boolean
    Private mFormatWhileTyping As Boolean
    Private mAbortAction As Boolean

    Private mTagName As String
    Private mOpenTagPos, mCloseTagPos As Integer 'current tag being edited
    Private mNodeLinking As Integer 'Index in tree of file containing link being edited

    Private mSashDragging As Boolean
    Private Const cSectionMainWin As String = "Main Window"
    Private Const cSectionRecentFiles As String = "Recent Files"
    Private Const cMaxRecentFiles As Integer = 6

    Private Sub cmdFind_KeyPress(ByVal aEventSender As System.Object, _
                                 ByVal aEventArgs As System.Windows.Forms.KeyPressEventArgs) Handles cmdFind.KeyPress
        Dim lKeyAscii As Integer = Asc(aEventArgs.KeyChar)
        If lKeyAscii >= 32 And lKeyAscii < 127 Then
            txtFind.Focus()
            txtFind.Text = Chr(lKeyAscii)
            txtFind.SelectionStart = 1
        End If
        aEventArgs.KeyChar = Chr(lKeyAscii)
        If lKeyAscii = 0 Then
            aEventArgs.Handled = True
        End If
    End Sub

    Private Sub cmdFind_MouseUp(ByVal aEventSender As System.Object, _
                                ByVal aEventArgs As System.Windows.Forms.MouseEventArgs) Handles cmdFind.MouseUp
        Dim lButton As Integer = aEventArgs.Button \ &H100000
        Dim lShift As Integer = System.Windows.Forms.Control.ModifierKeys \ &H10000
        Dim lX As Single = VB6.PixelsToTwipsX(aEventArgs.X)
        Dim lY As Single = VB6.PixelsToTwipsY(aEventArgs.Y)
        Static lFinding As Boolean
        Dim lSearchThrough, lSearchFor As String
        Dim lSearchPos, lSelStart, lStartNodeIndex As Integer
        If lButton = VB6.MouseButtonConstants.RightButton Then
            fraFind.Visible = False
            frmMain_Resize(Me, New System.EventArgs())
        ElseIf cmdFind.Text = "Stop" Then
            lFinding = False
        Else
            lFinding = True
            cmdFind.Text = "Stop"
            lSearchThrough = txtMain.Text
            If txtFind.Text = "" And txtMain.SelectionLength > 0 Then txtFind.Text = txtMain.SelectedText
            If txtFind.Text <> "" Then
                lSearchFor = UnEscape(txtFind.Text)
                lSelStart = txtMain.SelectionStart
                lSearchPos = txtMain.SelectionStart + txtMain.SelectionLength
                lSearchPos = txtMain.Find(lSearchFor, lSearchPos, RichTextBoxFinds.None)
                lStartNodeIndex = tree1.SelectedNode.Index
                If lSearchPos < 0 And lFinding Then
                    If QuerySave() <> MsgBoxResult.Cancel Then
NextNode:
                        If tree1.SelectedNode Is Nothing Then
                            tree1.SelectedNode = tree1.Nodes(0)
                        ElseIf tree1.SelectedNode.Index < tree1.Nodes.Count Then
                            tree1.SelectedNode = tree1.SelectedNode.NextVisibleNode
                        Else
                            tree1.SelectedNode = tree1.Nodes(0)
                        End If
                        lSearchPos = txtMain.Find(lSearchFor, 0)
                        If lSearchPos < 0 And tree1.SelectedNode.Index <> lStartNodeIndex Then
                            System.Windows.Forms.Application.DoEvents()
                            If lFinding Then GoTo NextNode
                        End If
                    End If
                End If
            End If
        End If
        cmdFind.Text = "Find"
    End Sub

    Private Function UnEscape(ByVal aSource As String) As String
        Dim lRetVal As String = ""
        Dim lCharPos As Integer = 1
        Dim lLastCharPos As Integer = aSource.Length
        While lCharPos <= lLastCharPos
            Dim lChar As String = Mid(aSource, lCharPos, 1)
            If lChar = "\" Then
                lCharPos += 1
                If lCharPos > lLastCharPos Then
                    lRetVal &= lChar
                Else
                    lChar = Mid(aSource, lCharPos, 1)
                    Select Case LCase(lChar)
                        Case "c" : lRetVal &= vbCrLf
                        Case "n" : lRetVal &= vbLf
                        Case "r" : lRetVal &= vbCr
                        Case "t" : lRetVal &= vbTab
                        Case "\" : lRetVal &= lChar
                        Case Else : lRetVal &= "^" & lChar
                    End Select
                End If
            Else
                lRetVal &= lChar
            End If
            lCharPos += 1
        End While
        Return lRetVal
    End Function

    Private Sub cmdReplace_KeyPress(ByVal aEventSender As System.Object, _
                                    ByVal aEventArgs As System.Windows.Forms.KeyPressEventArgs) Handles cmdReplace.KeyPress
        Dim lKeyAscii As Integer = Asc(aEventArgs.KeyChar)
        If lKeyAscii >= 32 And lKeyAscii < 127 Then
            txtReplace.Focus()
            txtReplace.Text = Chr(lKeyAscii)
            txtReplace.SelectionStart = 1
        End If
        aEventArgs.KeyChar = Chr(lKeyAscii)
        If lKeyAscii = 0 Then
            aEventArgs.Handled = True
        End If
    End Sub

    Private Sub cmdReplace_MouseUp(ByVal aEventSender As System.Object, _
                                   ByVal aEventArgs As System.Windows.Forms.MouseEventArgs) Handles cmdReplace.MouseUp
        Dim lButton As Integer = aEventArgs.Button \ &H100000
        Dim lShift As Integer = System.Windows.Forms.Control.ModifierKeys \ &H10000
        Dim lX As Single = VB6.PixelsToTwipsX(aEventArgs.X)
        Dim lY As Single = VB6.PixelsToTwipsY(aEventArgs.Y)

        If lButton = VB6.MouseButtonConstants.RightButton Then
            fraFind.Visible = False
            frmMain_Resize(Me, New System.EventArgs())
        Else
            Dim lFindText As String = UnEscape(txtFind.Text).ToLower
            Dim lReplaceText As String = UnEscape(txtReplace.Text)
            Dim lStartNodeIndex As Integer = tree1.SelectedNode.Index
            Dim lSearchedBeyondStart As Boolean = False
            If txtMain.SelectedText.ToLower = lFindText Then
NextReplace:
                txtMain.SelectedText = lReplaceText
            End If
            cmdFind_MouseUp(cmdFind, New System.Windows.Forms.MouseEventArgs(lButton * &H100000, 0, VB6.TwipsToPixelsX(lx), VB6.TwipsToPixelsY(lY), 0))
            If lStartNodeIndex <> tree1.SelectedNode.Index Then lSearchedBeyondStart = True
            If lShift > 0 Then
                If Not lSearchedBeyondStart Or lStartNodeIndex <> tree1.SelectedNode.Index Then
                    If LCase(txtMain.SelectedText) = lFindText Then GoTo NextReplace
                End If
            End If
        End If
    End Sub

    Private Sub frmMain_Load(ByVal aEventSender As System.Object, _
                             ByVal aEventArgs As System.EventArgs) Handles MyBase.Load
        pBrowseImage = "Use Other Image (File)"
        pViewImage = "View image"
        pSelectLink = "Link to Page (select)"
        pDeleteTag = "Delete"
        mnuContext(0).Text = pDeleteTag
        txtMain.Text = ""

        'UPGRADE_ISSUE: App property App.HelpFile was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
        'App.HelpFile = GetSetting(pAppName, "Files", "Help", My.Application.Info.DirectoryPath & "\AuthorDoc.chm")
        pBaseName = GetSetting(pAppName, "Defaults", "BaseName", "")
        mPath = GetSetting(pAppName, "Defaults", "Path", CurDir())
        mViewFormatting = CBool(GetSetting(pAppName, "Defaults", "ViewFormatting", CStr(True)))
        mFormatWhileTyping = CBool(GetSetting(pAppName, "Defaults", "FormatWhileTyping", CStr(False)))
        mnuAutoParagraph.Checked = CBool(GetSetting(pAppName, "Defaults", "AutoParagraph", CStr(False)))

        Dim lSetting As Object = GetSetting(pAppName, "Defaults", "FindTimeout", CStr(2))
        If IsNumeric(lSetting) Then pFindTimeout = lSetting
        lSetting = GetSetting(pAppName, cSectionMainWin, "Width")
        If IsNumeric(lSetting) Then Width = VB6.TwipsToPixelsX(lSetting)
        lSetting = GetSetting(pAppName, cSectionMainWin, "Height")
        If IsNumeric(lSetting) Then Height = VB6.TwipsToPixelsY(lSetting)
        lSetting = GetSetting(pAppName, cSectionMainWin, "Left")
        If IsNumeric(lSetting) Then Left = VB6.TwipsToPixelsX(lSetting)
        lSetting = GetSetting(pAppName, cSectionMainWin, "Top")
        If IsNumeric(lSetting) Then Top = VB6.TwipsToPixelsY(lSetting)
        lSetting = GetSetting(pAppName, cSectionMainWin, "TreeWidth")
        If IsNumeric(lSetting) Then
            sash.Left = VB6.TwipsToPixelsX(lSetting)
            mSashDragging = True
            sash_MouseMove(sash, New System.Windows.Forms.MouseEventArgs(1 * &H100000, 0, 0, 0, 0))
            mSashDragging = False
        End If
        For lRecentFileIndex As Integer = cMaxRecentFiles To 1 Step -1
            lSetting = GetSetting(pAppName, cSectionRecentFiles, CStr(lRecentFileIndex), "")
            If IO.File.Exists(lSetting) Then AddRecentFile(CStr(lSetting))
        Next lRecentFileIndex

        mnuFormatting.Checked = mViewFormatting
        mnuFormatWhileTyping.Checked = mFormatWhileTyping
        cdlgOpen.FileName = mPath & "\" & pBaseName & pSourceExtension
        cdlgSave.FileName = mPath & "\" & pBaseName & pSourceExtension
        cdlgImageOpen.FileName = mPath
        If IO.Directory.Exists(mPath) Then ChDir(mPath)
        If IO.File.Exists(cdlgOpen.FileName) Then
            Me.Show()
            Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
            OpenProject((cdlgOpen.FileName), tree1)
            If tree1.Nodes.Count > 0 Then tree1.SelectedNode = tree1.Nodes(0)
            Me.Cursor = System.Windows.Forms.Cursors.Default
        End If
    End Sub

    Private Sub frmMain_Resize(ByVal aEventSender As System.Object, _
                               ByVal aEventArgs As System.EventArgs) Handles MyBase.Resize
        If VB6.PixelsToTwipsY(Height) > 800 Then
            sash.Height = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Height) - 753) 'menu height
        End If
        tree1.Height = sash.Height

        txtMain.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(sash.Left) + VB6.PixelsToTwipsX(sash.Width))
        Dim lNewWidth As Integer = VB6.PixelsToTwipsX(Width) - VB6.PixelsToTwipsX(txtMain.Left) - 100
        If lNewWidth > 0 Then
            txtMain.Width = VB6.TwipsToPixelsX(lNewWidth)
        End If
    End Sub

    Private Sub frmMain_FormClosing(ByVal aSender As Object, _
                                    ByVal aE As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If QuerySave() = MsgBoxResult.Cancel Then
            aE.Cancel = True
        ElseIf QuerySaveProject() = MsgBoxResult.Cancel Then
            aE.Cancel = True
        Else
            'UPGRADE_ISSUE: App property App.HelpFile was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
            'SaveSetting(pAppName, "Files", "Help", App.HelpFile)
            SaveSetting(pAppName, "Defaults", "BaseName", pBaseName)
            SaveSetting(pAppName, "Defaults", "Path", mPath)
            SaveSetting(pAppName, "Defaults", "FindTimeout", CStr(pFindTimeout))
            SaveSetting(pAppName, "Defaults", "ViewFormatting", CStr(mViewFormatting))
            SaveSetting(pAppName, "Defaults", "FormatWhileTyping", CStr(mFormatWhileTyping))
            SaveSetting(pAppName, "Defaults", "AutoParagraph", CStr(mnuAutoParagraph.Checked))

            SaveSetting(pAppName, cSectionMainWin, "Width", CStr(VB6.PixelsToTwipsX(Width)))
            SaveSetting(pAppName, cSectionMainWin, "Height", CStr(VB6.PixelsToTwipsY(Height)))
            SaveSetting(pAppName, cSectionMainWin, "Left", CStr(VB6.PixelsToTwipsX(Left)))
            SaveSetting(pAppName, cSectionMainWin, "Top", CStr(VB6.PixelsToTwipsY(Top)))
            SaveSetting(pAppName, cSectionMainWin, "TreeWidth", CStr(VB6.PixelsToTwipsX(sash.Left)))
            Dim lRecentFileIndex As Integer
            For lRecentFileIndex = mMnuRecent.Count - 1 To 1 Step -1
                SaveSetting(pAppName, cSectionRecentFiles, CStr(lRecentFileIndex), mMnuRecent(lRecentFileIndex).Tag)
            Next lRecentFileIndex
            While GetSetting(pAppName, cSectionRecentFiles, CStr(lRecentFileIndex)) <> ""
                SaveSetting(pAppName, cSectionRecentFiles, CStr(lRecentFileIndex), "")
                lRecentFileIndex += 1
            End While

            For Each lOpenForm As Form In My.Application.OpenForms
                If Not lOpenForm Is Me Then lOpenForm.Close()
            Next
        End If
    End Sub

    Private Sub fraFind_MouseDown(ByVal aEventSender As System.Object, _
                                  ByVal aEventArgs As System.Windows.Forms.MouseEventArgs) Handles fraFind.MouseDown
        Dim lButton As Integer = aEventArgs.Button \ &H100000
        Dim lShift As Integer = System.Windows.Forms.Control.ModifierKeys \ &H10000
        If lButton = VB6.MouseButtonConstants.RightButton Or lShift = System.Windows.Forms.Keys.ShiftKey Then
            fraFind.Visible = False
            frmMain_Resize(Me, New System.EventArgs())
        End If
    End Sub

    Public Sub mnuAutoSave_Click(ByVal aEventSender As System.Object, _
                                 ByVal aEventArgs As System.EventArgs) Handles mnuAutoSave.Click
        mnuAutoSave.Checked = Not mnuAutoSave.Checked
    End Sub

    Public Sub mnuContext_Click(ByVal aEventSender As System.Object, _
                                ByVal aEventArgs As System.EventArgs) Handles mnuContext.Click
        Dim Index As Integer = mnuContext.GetIndex(aEventSender)
        ContextAction(mnuContext(Index).Text)
    End Sub

    Public Sub ContextAction(ByRef aCommand As String)
        Dim lFilename, lPathName As String
        Select Case aCommand
            Case pCaptureReplace
                lFilename = ReplaceString(SubTagValue("src"), "/", "\")
                lFilename = IO.Path.GetDirectoryName(mPath & "\" & NodeFile()) & "\" & lFilename
                frmCapture.Filename = lFilename
                frmCapture.Show()
            Case pCaptureNew, pBrowseImage
                cdlgOpen.ShowDialog()
                cdlgSave.FileName = cdlgOpen.FileName
                lFilename = cdlgOpen.FileName
                If Len(lFilename) > 0 Then
                    lPathName = IO.Path.GetDirectoryName(mPath & "\" & NodeFile())
                    lFilename = HTMLRelativeFilename(lFilename, lPathName)
                End If
                If mCloseTagPos > mOpenTagPos + 4 Then
                    EditSubTag("src", lFilename)
                Else
                    txtMain.Text = VB.Left(txtMain.Text, txtMain.SelectionStart) & "<img src=""" & lFilename & """>" & Mid(txtMain.Text, txtMain.SelectionStart + 1)
                End If
                If aCommand = pCaptureNew Then
                    frmCapture.Filename = lFilename
                    frmCapture.Show()
                End If
            Case pViewImage
                lFilename = ReplaceString(SubTagValue("src"), "/", "\")
                lFilename = IO.Path.GetDirectoryName(mPath & "\" & NodeFile()) & "\" & lFilename
                If IO.File.Exists(lFilename) Then OpenFile(lFilename)
            Case pDeleteTag
                If mCloseTagPos > mOpenTagPos + 4 Then txtMain.Text = VB.Left(txtMain.Text, mOpenTagPos - 1) & Mid(txtMain.Text, mCloseTagPos + 1)
            Case pSelectLink
                mNodeLinking = tree1.SelectedNode.Index
                Me.Cursor = System.Windows.Forms.Cursors.UpArrow
            Case Else
                Logger.Msg("Unrecognized menu item: " & aCommand, MsgBoxStyle.OkOnly, "AuthorDoc")
        End Select
    End Sub

    Public Sub mnuConvert_Click(ByVal aEventSender As System.Object, _
                                ByVal aEventArgs As System.EventArgs) Handles mnuConvert.Click
        If QuerySave() <> MsgBoxResult.Cancel Then
            If QuerySaveProject() <> MsgBoxResult.Cancel Then frmConvert.Show()
        End If
    End Sub

    Public Sub mnuCopy_Click(ByVal aEventSender As System.Object, _
                             ByVal aEventArgs As System.EventArgs) Handles mnuCopy.Click
        My.Computer.Clipboard.SetText(txtMain.SelectedText)
    End Sub

    Public Sub mnuCut_Click(ByVal aEventSender As System.Object, _
                            ByVal aEventArgs As System.EventArgs) Handles mnuCut.Click
        My.Computer.Clipboard.SetText(txtMain.SelectedText)
        txtMain.SelectedText = ""
    End Sub

    Private Sub mnuEditProject_Click()
        If tree1.Visible Then
            If QuerySaveProject() <> MsgBoxResult.Cancel Then
                LoadTextboxFromFile(IO.Path.GetDirectoryName(pProjectFileName), IO.Path.GetFileNameWithoutExtension(pProjectFileName), "." & FileExt(pProjectFileName), txtMain)
                tree1.Visible = False
            End If
        Else
            If QuerySave() <> MsgBoxResult.Cancel Then
                tree1.Visible = True
                mnuRecent_Click(mMnuRecent.Item(1), New System.EventArgs())
            End If
        End If
    End Sub

    Public Sub mnuExit_Click(ByVal aEventSender As System.Object, _
                             ByVal aEventArgs As System.EventArgs) Handles mnuExit.Click
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
        mViewFormatting = mnuFormatting.Checked
        If mViewFormatting Then
            FormatText(txtMain)
        Else
            txtMain.Text = txtMain.Text
            txtMain.Refresh()
        End If
    End Sub

    Public Sub mnuFormatWhileTyping_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuFormatWhileTyping.Click
        mnuFormatWhileTyping.Checked = Not mnuFormatWhileTyping.Checked
        mFormatWhileTyping = mnuFormatWhileTyping.Checked
        If mFormatWhileTyping Then
            If Not mViewFormatting Then
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
        If mTagName = "a" Then
            txtMain.SelectionStart = mOpenTagPos + 9
        Else
            mnuLink_Click(mnuLink, New System.EventArgs())
            GetCurrentTag()
        End If
        ContextAction(pSelectLink)
    End Sub

    Public Sub mnuNewProject_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuNewProject.Click
        If QuerySave() = MsgBoxResult.Cancel Then Exit Sub
        Try
            cdlgSave.ShowDialog()
            cdlgOpen.FileName = cdlgSave.FileName
            If cdlgOpen.FileName.Length > 0 Then
                mPath = IO.Path.GetDirectoryName((cdlgOpen.FileName))
                ChDir(mPath)
                If Not IO.Directory.Exists(mPath) Then MkDir(mPath)
                Dim f As Integer = FreeFile()
                FileOpen(f, cdlgOpen.FileName, OpenMode.Output)
                FileClose(f)
                OpenProject((cdlgOpen.FileName), tree1)
                mnuNewSection.Enabled = True
                mProjectChanged = False
                If tree1.Nodes.Count > 0 Then tree1.SelectedNode = tree1.Nodes(0)
            End If
        Catch ex As Exception
            Logger.Msg("Error creating new project:" & vbCr & ex.Message)
        End Try
    End Sub

    Public Sub mnuNewSection_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuNewSection.Click
        Dim found As Boolean
        Dim nodNum As Integer
        Dim key, ThisName, keypath As String
        Dim filename As String
        Dim f As Integer

        cdlgOpen.ShowDialog()
        cdlgSave.FileName = cdlgOpen.FileName
        filename = cdlgOpen.FileName
        If Len(filename) > Len(mPath) Then
            If UCase(VB.Left(filename, Len(mPath))) <> UCase(mPath) Then
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

                ThisName = IO.Path.GetFileNameWithoutExtension(filename)
                key = VB.Left(filename, Len(filename) - Len(pSourceExtension)) 'trim extension .txt
                key = "N" & Mid(key, Len(mPath) + 2)
                keypath = IO.Path.GetDirectoryName(Mid(key, 2))
                If tree1.Nodes.Count = 0 Then 'This is the first node
                    tree1.Nodes.Add(key, ThisName)
                ElseIf keypath = IO.Path.GetDirectoryName(NodeFile) Then  'place after selected sibling
                    tree1.Nodes.Insert(tree1.SelectedNode.Index + 1, key, ThisName)
                Else
                    nodNum = tree1.Nodes.Count
                    found = False
                    While nodNum >= 1 And Not found 'Look for last sibling
                        If IO.Path.GetDirectoryName(NodeFile(nodNum)) = keypath Then
                            tree1.Nodes.Insert(tree1.Nodes(nodNum).Index + 1, key, ThisName)
                            found = True
                        End If
                        nodNum = nodNum - 1
                    End While
                    If Not found Then tree1.SelectedNode.Nodes.Add(key, ThisName)
                End If
                pCurrentFilename = cdlgOpen.FileName
                mProjectChanged = True
                tree1.Nodes(key).EnsureVisible()
            End If
        End If
    End Sub

    Private Function NodeFile(Optional ByRef nodNum As Integer = 0) As String
        If IsNothing(nodNum) OrElse nodNum = 0 Then
            nodNum = tree1.SelectedNode.Index
        End If
        If nodNum < 1 Then nodNum = 1
        NodeFile = Mid(tree1.Nodes(nodNum).FullPath, 2)
    End Function

    Public Sub mnuOpenProject_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuOpenProject.Click
        If IO.Directory.Exists(mPath) Then
            ChDriveDir(mPath)
        End If

        cdlgOpen.FileName = ""
        cdlgSave.FileName = "" 'BaseName
        cdlgOpen.ShowDialog()
        cdlgSave.FileName = cdlgOpen.FileName
        If Len(cdlgOpen.FileName) > 0 Then
            AddRecentFile((cdlgOpen.FileName))
            mnuRecent_Click(mMnuRecent.Item(0), New System.EventArgs())
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
                    mPath = IO.Path.GetDirectoryName(newFilePath)
                    Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
                    OpenProject(newFilePath, tree1)
                    mnuNewSection.Enabled = True
                    mProjectChanged = False
                    If tree1.Nodes.Count > 0 Then tree1.SelectedNode = tree1.Nodes(0)
                    Me.Cursor = System.Windows.Forms.Cursors.Default
                End If
            End If
        End If
    End Sub

    Public Sub mnuRevert_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuRevert.Click
        txtMain.Text = mCurrentFileContents
    End Sub

    Public Sub mnuSaveFile_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuSaveFile.Click
        Dim f As Integer 'file handle

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
            PathName = IO.Path.GetDirectoryName(mPath & "\" & NodeFile)
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
        Dim Button As Integer = eventArgs.Button \ &H100000
        Dim Shift As Integer = System.Windows.Forms.Control.ModifierKeys \ &H10000
        Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
        Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
        mSashDragging = True
    End Sub

    Private Sub sash_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles sash.MouseMove
        Dim Button As Integer = eventArgs.Button \ &H100000
        Dim Shift As Integer = System.Windows.Forms.Control.ModifierKeys \ &H10000
        Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
        Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
        Dim newLeftWidth As Integer
        If mSashDragging And (VB6.PixelsToTwipsX(sash.Left) + x) > 100 And (VB6.PixelsToTwipsX(sash.Left) + x < VB6.PixelsToTwipsX(Width) - 100) Then
            sash.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(sash.Left) + x)
            newLeftWidth = VB6.PixelsToTwipsX(sash.Left)
            If newLeftWidth > 1000 Then tree1.Width = VB6.TwipsToPixelsX(newLeftWidth)
            frmMain_Resize(Me, New System.EventArgs())
        End If
    End Sub

    Private Sub sash_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles sash.MouseUp
        Dim Button As Integer = eventArgs.Button \ &H100000
        Dim Shift As Integer = System.Windows.Forms.Control.ModifierKeys \ &H10000
        Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
        Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
        mSashDragging = False
    End Sub

    Private Sub TimerSlowAction_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles TimerSlowAction.Tick
        TimerSlowAction.Enabled = False
        mAbortAction = True
    End Sub

    Private Sub tree1_AfterLabelEdit(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.NodeLabelEditEventArgs) Handles tree1.AfterLabelEdit
        Dim OldFilePath As String
        With tree1.SelectedNode
            OldFilePath = mPath & "\" & Mid(.FullPath, 2) & pSourceExtension
            If IO.File.Exists(OldFilePath) Then
                Select Case MsgBox("Rename file '" & OldFilePath & "' to '" & eventArgs.Label & "?", MsgBoxStyle.YesNoCancel)
                    Case MsgBoxResult.No : .Text = eventArgs.Label ': .Key = "N" & .FullPath
                    Case MsgBoxResult.Yes : .Text = eventArgs.Label ': .Key = "N" & .FullPath
                        'Name OldFilePath As path & "\" & .fullpath & pSourceExtension
                        Rename(OldFilePath, IO.Path.GetDirectoryName(OldFilePath) & "\" & .Text & pSourceExtension)
                    Case MsgBoxResult.Cancel
                        eventArgs.CancelEdit = True
                End Select
            End If
        End With
    End Sub

    Private Sub tree1_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles tree1.KeyDown
        Select Case eventArgs.keyCode
            Case System.Windows.Forms.Keys.Delete : tree1.Nodes.Remove(tree1.SelectedNode)
        End Select
    End Sub

    Private Sub tree1_AfterSelect(ByVal sender As System.Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles tree1.AfterSelect
        Dim filename, fullpath As String
        Dim inClick As Boolean
        If Not inClick Then 'And Not Timer1.Enabled Then
            inClick = True
            If mNodeLinking > 0 Then
                fullpath = "c:\" & IO.Path.GetDirectoryName(NodeFile(mNodeLinking))
                filename = HTMLRelativeFilename(e.Node.FullPath, fullpath)
                EditSubTag("href", filename)

                '                DelaySetNode(mNodeLinking)
                tree1.SelectedNode = tree1.Nodes(mNodeLinking)
                If txtMain.Text <> WholeFileString(pCurrentFilename) Then SetFileChanged(True)

                mNodeLinking = 0
                Me.Cursor = System.Windows.Forms.Cursors.Default
            Else
                If e.Node Is tree1.Nodes(0) Then
                    filename = e.Node.FullPath
                Else
                    filename = e.Node.FullPath.Substring(tree1.Nodes(0).FullPath.Length)
                End If
                If QuerySave() = MsgBoxResult.Cancel Then 'Should move focus back to old node here
                    'DelaySetNode(1)
                    tree1.SelectedNode = tree1.Nodes(0)
                    If txtMain.Text <> WholeFileString(pCurrentFilename) Then SetFileChanged(True)
                Else
                    LoadTextboxFromFile(mPath, filename, pSourceExtension, txtMain)
                    If tree1.SelectedNode Is Nothing Then
                        tree1.SelectedNode = e.Node
                    ElseIf tree1.SelectedNode.Index <> e.Node.Index Then
                        tree1.SelectedNode = e.Node
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
        If Not IO.File.Exists(IO.Path.Combine(fullpath, filename & ext)) Then 'Check for files named .html or pSourceExtension
            If LCase(ext) = LCase(pSourceExtension) Then altExt = ".html" Else altExt = pSourceExtension
            altpath = fullpath & "\" & filename & altExt
            If IO.File.Exists(altpath) Then
                If altExt = pSourceExtension Then
                    ext = pSourceExtension
                Else
                    If LastAnswer = 0 Then
                        thisAnswer = MsgBox("File " & filename & pSourceExtension & " was not found, use " & filename & ".html instead?", MsgBoxStyle.YesNoCancel, mPath)
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
        Dim f As Integer 'file handle
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
            mCurrentFileContents = InputString(f, FileLength)
            txtBox.Text = mCurrentFileContents
            If mViewFormatting Then FormatText(txtBox)
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
        If mChanged <> newValue Then
            mChanged = newValue
            mnuSaveFile.Enabled = mChanged
            If mChanged Then
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
        mAbortAction = False
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

        While nextch < maxch And Not mAbortAction
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
        If mAbortAction Then
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
        If mProjectChanged Then
            retval = MsgBox("Save changes to " & pProjectFileName & "?", MsgBoxStyle.YesNoCancel)
            If retval = MsgBoxResult.Yes Then SaveProject(pProjectFileName, tree1)
            mProjectChanged = False
        End If
        QuerySaveProject = retval
    End Function

    Private Function QuerySave() As MsgBoxResult
        Dim retval As MsgBoxResult
        retval = MsgBoxResult.Yes
        If mChanged Then
            If Not mnuAutoSave.Checked Then
                retval = MsgBox("Save changes to " & pCurrentFilename & "?", MsgBoxStyle.YesNoCancel)
            End If
            If retval = MsgBoxResult.Yes Then mnuSaveFile_Click(mnuSaveFile, New System.EventArgs())
        End If
        QuerySave = retval
    End Function

    Private Sub txtFind_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtFind.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.Return Then cmdFind_MouseUp(cmdFind, New System.Windows.Forms.MouseEventArgs(VB6.MouseButtonConstants.LeftButton * &H100000, 0, 0, 0, 0))
    End Sub

    Private Sub txtMain_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtMain.KeyUp
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        If mFormatWhileTyping Then FormatText(txtMain)
    End Sub

    Private Sub txtReplace_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtReplace.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.Return Then cmdReplace_MouseUp(cmdReplace, New System.Windows.Forms.MouseEventArgs(VB6.MouseButtonConstants.LeftButton * &H100000, 0, 0, 0, 0))
    End Sub

    Private Sub txtMain_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtMain.TextChanged
        Static InChange As Boolean
        If Not InChange And Not mUndoing Then
            InChange = True

            mUndos(mUndoPos) = txtMain.Text
            mUndoCursor(mUndoPos) = txtMain.SelectionStart
            mUndoPos = mUndoPos + 1
            If mUndoPos > mMaxUndo Then mUndoPos = 0
            If mUndosAvail < mMaxUndo Then mUndosAvail = mUndosAvail + 1

            If mCurrentFileContents <> txtMain.Text Then
                If Not mChanged Then SetFileChanged(True)
            Else
                If mChanged Then SetFileChanged(False)
            End If
            mnuSaveFile.Enabled = mChanged
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
        If mOpenTagPos < mCloseTagPos Then
            Select Case mTagName
                Case "img"
                    AddContextMenuItem(pCaptureReplace)
                    AddContextMenuItem(pCaptureNew)
                    AddContextMenuItem(pBrowseImage)
                    AddContextMenuItem(pViewImage)
                    'filename = SubTagValue("src")
                    'filename = ReplaceString(filename, "/", "\")
                    'pathname = AbsolutePath(filename, IO.Path.GetDirectoryName(path & "\" & NodeFile))
                    'If Len(Dir(pathname)) > 0 Then frmSample.SetImage pathname
                Case "a"
                    AddContextMenuItem(pSelectLink)
                    filename = SubTagValue("href")
                    hashPos = InStr(filename, "#")
                    If hashPos > 0 Then filename = VB.Left(filename, hashPos - 1)
                    If Len(filename) > 0 Then
                        filename = ReplaceString(filename, "/", "\")
                        If VB.Left(filename, 1) = "\" Then
                            PathName = mPath & filename
                        Else
                            PathName = IO.Path.GetDirectoryName(mPath & "\" & NodeFile()) & "\" & filename
                        End If
                        If IO.File.Exists(PathName) Then
                            frmSample.SetText(PathName)
                        ElseIf IO.File.Exists(PathName & pSourceExtension) Then
                            frmSample.SetText(PathName & pSourceExtension)
                        End If
                    End If
            End Select
        End If
    End Sub

    Private Sub txtMain_KeyPress(ByVal aEventSender As System.Object, ByVal aEventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtMain.KeyPress
        Dim lKeyAscii As Integer = Asc(aEventArgs.KeyChar)
        Select Case lKeyAscii
            Case 26 'Control-Z = undo
                If mUndosAvail > 0 Then
                    mUndoing = True
                    mUndoPos = mUndoPos - 1
                    If mUndoPos < 0 Then mUndoPos = mMaxUndo
                    txtMain.Text = mUndos(mUndoPos)
                    txtMain.SelectionStart = mUndoCursor(mUndoPos)
                    mUndosAvail = mUndosAvail - 1
                    mUndoing = False
                End If
            Case 13
                If mnuAutoParagraph.Checked Then
                    Dim lSelectionStartOriginal As Integer = txtMain.SelectionStart
                    txtMain.Text = VB.Left(txtMain.Text, lSelectionStartOriginal) & "<p>" & Mid(txtMain.Text, lSelectionStartOriginal + 1)
                    txtMain.SelectionStart = lSelectionStartOriginal + 3
                End If
        End Select
        aEventArgs.KeyChar = Chr(lKeyAscii)
        If lKeyAscii = 0 Then
            aEventArgs.Handled = True
        End If
    End Sub

    Private Sub txtMain_MouseDown(ByVal aEventSender As System.Object, ByVal aEventArgs As System.Windows.Forms.MouseEventArgs) Handles txtMain.MouseDown
        Dim lButton As Integer = aEventArgs.Button \ &H100000
        Dim lShift As Integer = System.Windows.Forms.Control.ModifierKeys \ &H10000
        Dim x As Single = VB6.PixelsToTwipsX(aEventArgs.X)
        Dim y As Single = VB6.PixelsToTwipsY(aEventArgs.Y)
    End Sub

    'Search in string txt for a tag that encloses start character position
    'Sets tagName to lowercase of first word in tag
    'Sets openTagPos, closeTagPos to string index of < and > of tag in txt
    Private Sub GetCurrentTag()
        Dim lText As String = txtMain.Text
        Dim lSelectionStart As Integer = txtMain.SelectionStart
        If lSelectionStart < 1 Then Exit Sub
        mOpenTagPos = InStrRev(lText, "<", lSelectionStart)
        If mOpenTagPos > 0 Then
            mCloseTagPos = InStrRev(lText, ">", lSelectionStart)
            If mCloseTagPos < mOpenTagPos Then 'we are in a tag
                mCloseTagPos = InStr(lSelectionStart, lText, ">")
            End If
        End If
        Dim lEndNamePos As Integer
        If mOpenTagPos > 0 And mOpenTagPos <= lSelectionStart And mCloseTagPos >= lSelectionStart Then
            lEndNamePos = InStr(mOpenTagPos, lText, " ")
            If lEndNamePos = 0 Or lEndNamePos > mCloseTagPos Then lEndNamePos = mCloseTagPos
            mTagName = LCase(Mid(lText, mOpenTagPos + 1, lEndNamePos - mOpenTagPos - 1))
        Else
            mOpenTagPos = 0
            mCloseTagPos = 0
            mTagName = ""
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
    Private Function SubTagValue(ByRef aSubtagName As String) As String
        Dim lSubTagValue As String = ""
        Dim lText As String = txtMain.Text
        'Dim lSelectionStart As Integer = txtMain.SelectionStart
        Dim lTag As String = Mid(lText, mOpenTagPos, mCloseTagPos - mOpenTagPos + 1).ToLower
        Dim aSubtagStart As Integer = InStr(1, lTag, aSubtagName.ToLower)
        If aSubtagStart = 0 Then
            lSubTagValue = ""
        Else
            Dim lValueStart As Integer = aSubtagStart + aSubtagName.Length + 1
            Dim lValueEnd As Integer
            If Mid(lTag, lValueStart, 1) = """" Then
                lValueStart = lValueStart + 1
                lValueEnd = InStr(lValueStart, lTag, """")
            Else
                lValueEnd = InStr(lValueStart + 1, lTag, " ")
                If lValueEnd = 0 Then lValueEnd = lTag.Length
            End If
            If lValueEnd > lValueStart Then
                lSubTagValue = Mid(lTag, lValueStart, lValueEnd - lValueStart)
            End If
        End If
        Return lSubTagValue
    End Function

    'Uses current tag delimited by openTagPos and closeTagPos
    'Modifies txtMain.Text, replacing current value of subtagName with NewValue
    'If subtagName does not exist in the current tag, it is added at the end
    'EditSubTag( "src", "bar.gif" ) when the current tag is <img src="foo.png">
    'will result in <img src="bar.gif">
    Private Sub EditSubTag(ByRef aSubtagName As String, ByRef aNewValue As String)
        Dim lTxt As String = txtMain.Text
        Dim lTag As String = Mid(lTxt, mOpenTagPos, mCloseTagPos - mOpenTagPos + 1).ToLower
        Dim lSubtagStart As Integer = InStr(1, lTag, aSubtagName.ToLower)
        If lSubtagStart = 0 Then
            txtMain.Text = VB.Left(lTxt, mCloseTagPos - 1) & " " & aSubtagName.ToLower & "=" & aNewValue & Mid(lTxt, mCloseTagPos)
        Else
            'lsubtagStart += openTagPos
            Dim lValueStart As Integer = lSubtagStart + aSubtagName.Length + 1
            Dim lValueEnd As Integer
            If Mid(lTag, lValueStart, 1) = """" Then
                lValueEnd = InStr(lValueStart + 1, lTag, """")
            Else
                lValueEnd = InStr(lValueStart + 1, lTag, " ")
                If lValueEnd = 0 Then lValueEnd = Len(lTag)
            End If
            txtMain.Text = VB.Left(lTxt, mOpenTagPos + lValueStart - 1) & aNewValue & Mid(lTxt, mOpenTagPos + lValueEnd - 1)
            mCloseTagPos = InStr(mOpenTagPos + 1, txtMain.Text, ">")
        End If
        txtMain.SelectionStart = mOpenTagPos + 1
        txtMain_Click(txtMain, New System.EventArgs())
        txtMain.SelectionStart = mCloseTagPos + 1
    End Sub

    Private Sub AddContextMenuItem(ByRef aNewItem As String)
        Dim lMnuItem As Integer = mnuContext.Count
        mnuContext.Load(lMnuItem)
        mnuContext(lMnuItem).Text = aNewItem
    End Sub

    Private Sub txtMain_SelectionChanged(ByVal aEventSender As System.Object, ByVal aEventArgs As System.EventArgs) Handles txtMain.SelectionChanged
        Dim lLastSelStart As Integer
        If txtMain.SelectionStart <> lLastSelStart Then
            lLastSelStart = txtMain.SelectionStart
        End If
    End Sub

    Friend Sub AddRecentFile(ByRef aFilePath As String)
        Dim lMatch As Boolean = False
        Dim lRecentFileIndex As Integer = 0
        For Each lRecent As ToolStripMenuItem In mMnuRecent 'as While Not match And rf <= mnuRecent.Count - 1
            If lRecent.Tag.ToString.ToUpper = aFilePath.ToUpper Then
                lMatch = True
                Exit For
            End If
            lRecentFileIndex += 1
        Next ' End While
        If lMatch Then 'move file to top of list
            For lRecentFileMove As Integer = lRecentFileIndex To 1 Step -1
                mMnuRecent(lRecentFileMove).Tag = mMnuRecent(lRecentFileMove - 1).Tag
                mMnuRecent(lRecentFileMove).Text = "&" & lRecentFileMove + 1 & " " & IO.Path.GetFileNameWithoutExtension(mMnuRecent(lRecentFileMove).Tag)
            Next lRecentFileMove
            mMnuRecent(0).tag = aFilePath
            mMnuRecent(0).text = "&1 " & IO.Path.GetFileNameWithoutExtension(aFilePath)
        Else 'Add file to list
            mnuRecentSeparator.Visible = True
            Dim lToolStripMenuItem As New ToolStripMenuItem
            With lToolStripMenuItem
                .Tag = aFilePath
                .Visible = True
            End With
            mMnuRecent.Insert(0, lToolStripMenuItem)
            mnuFile.DropDownItems.Insert(mnuFile.DropDownItems.IndexOf(mnuRecentSeparator) + 1, lToolStripMenuItem)
            AddHandler lToolStripMenuItem.Click, AddressOf mnuRecent_Click

            For lRecentIndex As Integer = mMnuRecent.Count - 1 To 0 Step -1
                If lRecentIndex >= cMaxRecentFiles Then
                    mnuFile.DropDownItems.Remove(mMnuRecent.Item(lRecentIndex))
                    mMnuRecent.RemoveAt(lRecentIndex)
                Else
                    mMnuRecent(lRecentIndex).Text = "&" & lRecentIndex + 1 & " " & IO.Path.GetFileNameWithoutExtension(mMnuRecent(lRecentIndex).Tag)
                End If
            Next
        End If
    End Sub
End Class