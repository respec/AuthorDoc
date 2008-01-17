Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Imports atcUtility

Module modFileIO
    'Copyright 2000-2008 by AQUA TERRA Consultants
	
	Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Integer, ByVal dwMilliseconds As Integer) As Integer
	
	Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Integer) As Integer
	
	Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Integer, ByVal bInheritHandle As Integer, ByVal dwProcessId As Integer) As Integer
	
	
	Private Const INFINITE As Short = -1
	Private Const SYNCHRONIZE As Integer = &H100000
	Private NconvertPath As String
	
	Sub RunNconvert(ByRef cmdline As String)
		Dim ret, iTask, pHandle As Integer
		
		If NconvertPath = "" Then FindNconvert()
		
		iTask = Shell(NconvertPath & " " & cmdline, AppWinStyle.Hide)
		pHandle = OpenProcess(SYNCHRONIZE, False, iTask)
		ret = WaitForSingleObject(pHandle, INFINITE)
		ret = CloseHandle(pHandle)
		
	End Sub
	
	Sub FindNconvert()
		NconvertPath = GetSetting("Nconvert", "Paths", "ExePath", "")
        If Not IO.File.Exists(NconvertPath) Then
            NconvertPath = FindFile("Find Nconvert.exe to perform conversion", "Nconvert.exe")
            If IO.File.Exists(NconvertPath) Then
                SaveSetting("Nconvert", "Paths", "ExePath", NconvertPath)
            End If
        End If
    End Sub
	
	Sub OpenProject(ByRef filename As String, ByRef t As AxComctlLib.AxTreeView)
		Dim f As Short 'file handle
		Dim buf As String 'input buffer, contains current line
		Dim ThisName As String 'file name of current source file, minus extension
		Dim key As String 'unique ID for tree control
		Dim SectionName(50) As String 'Array of current section names for each level
		Dim SectionLevel As Integer 'Level of current source file, according to indentation
		Dim lvl As Integer 'Level in loop that constructs keys
		Dim nod As ComctlLib.Node 'Node inserted into tree control
		Dim dotpos As Integer 'position of . in filename
		Dim StartTime As Integer
		StartTime = VB.Timer()
		
		On Error GoTo OpenError
		
		f = FreeFile()
		frmMain.Cursor = System.Windows.Forms.Cursors.WaitCursor
		t.Visible = False
        If Not IO.File.Exists(filename) Then
            If MsgBox("File not found. Create new project file '" & filename & "'?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                'UPGRADE_ISSUE: GoSub statement is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C5A1A479-AB8B-4D40-AAF4-DB19A2E5E77F"'

                pProjectFileName = filename
                pBaseName = FilenameOnly(filename)
                t.Nodes.Clear()
                t.Nodes.Add(, , "N" & pBaseName, pBaseName)
                t.Nodes(1).Expanded = True

                FileOpen(f, filename, OpenMode.Output)
                FileClose(f)
            End If
        Else
            'UPGRADE_ISSUE: GoSub statement is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C5A1A479-AB8B-4D40-AAF4-DB19A2E5E77F"'

            pProjectFileName = filename
            pBaseName = FilenameOnly(filename)
            t.Nodes.Clear()
            t.Nodes.Add(, , "N" & pBaseName, pBaseName)
            t.Nodes(1).Expanded = True

            FileOpen(f, filename, OpenMode.Input)
            While Not EOF(f) ' Loop until end of file.
                buf = LineInput(f)
                ThisName = LTrim(buf)
                If ThisName <> "" Then
                    SectionLevel = (Len(buf) - Len(ThisName)) / 2 + 1 '2 spaces indentation per level
                    ThisName = RTrim(ThisName)
                    key = ThisName
                    SectionName(SectionLevel) = ThisName
                    If SectionLevel = 1 Then
                        nod = t.Nodes.Add("N" & pBaseName, ComctlLib.TreeRelationshipConstants.tvwChild, "N" & key, ThisName)
                    Else
                        For lvl = SectionLevel - 1 To 1 Step -1
                            key = SectionName(lvl) & "\" & key
                        Next lvl
                        On Error GoTo skip
                        nod = t.Nodes.Add("N" & Left(key, Len(key) - Len(ThisName) - 1), ComctlLib.TreeRelationshipConstants.tvwChild, "N" & key, ThisName)
                        If Not nod.Parent.Expanded Then nod.Parent.Expanded = True
                    End If
                End If
            End While
            FileClose(f)
        End If
        frmMain.AddRecentFile(pProjectFileName)
		t.Visible = True
		frmMain.Cursor = System.Windows.Forms.Cursors.Default
		If t.Nodes.Count > 0 Then t.Nodes(1).EnsureVisible()
		Exit Sub
		
OpenError: 
		MsgBox("Error reading project file '" & filename & "'" & vbCr & Err.Description)
		On Error Resume Next
		FileClose(f)
skip: 
		Debug.Print("Duplicate key in tree: " & key)
		Resume Next
    End Sub
	
	Public Sub SaveProject(ByRef filename As String, ByRef t As AxComctlLib.AxTreeView)
		Dim outfile As Short 'file handle
		Dim nodNum As Integer 'Node number (we go sequentially through nodes)
		Dim nod As ComctlLib.Node 'Node of the tree being written
		
		'Mark all as need to be saved
		For nodNum = 1 To t.Nodes.Count
			t.Nodes.Item(nodNum).tag = True
		Next 
		t.Nodes.Item(1).tag = False
		
		outfile = FreeFile()
		FileOpen(outfile, filename, OpenMode.Output)
		nod = t.Nodes.Item(1).Child
		While Not nod Is Nothing
			WriteProjectSection(nod, outfile)
			nod = nod.Next
		End While
		FileClose(outfile)
	End Sub
	
	Private Sub WriteProjectSection(ByRef nod As ComctlLib.Node, ByRef outfile As Short)
		Dim ThisName As String 'file name of current source file, minus extension
		Dim pos As Integer 'position of directory delimiter '\' in node key for counting levels
		Dim kid As ComctlLib.Node 'nod's child
		
		If nod.tag Then
			If Not nod.Parent Is Nothing Then
				If nod.Parent.tag Then
					WriteProjectSection(nod.Parent, outfile) 'Write parent first
					Exit Sub 'Writing parent will lead to doing this node, so we are done
				End If
			End If
			nod.tag = False
			ThisName = ""
			pos = InStr(nod.key, "\")
			While pos > 0 And pos < Len(nod.key)
				ThisName = ThisName & "  "
				pos = InStr(pos + 1, nod.key, "\")
			End While
			
			ThisName = ThisName & nod.Text
			PrintLine(outfile, ThisName)
			kid = nod.Child
			For pos = 1 To nod.Children
				WriteProjectSection(kid, outfile)
				kid = kid.Next
			Next 
            kid = Nothing
		End If
	End Sub
End Module