Option Strict Off
Option Explicit On
Imports atcUtility
Imports MapWinUtility

Module modFileIO
    'Copyright 2000-2008 by AQUA TERRA Consultants
	
	Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Integer, ByVal dwMilliseconds As Integer) As Integer
	
	Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Integer) As Integer
	
	Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Integer, ByVal bInheritHandle As Integer, ByVal dwProcessId As Integer) As Integer
	
	
    Private Const INFINITE As Integer = -1
	Private Const SYNCHRONIZE As Integer = &H100000
    Private mNconvertPath As String

    Sub RunNconvert(ByRef aCommandLine As String)
        If mNconvertPath = "" Then FindNconvert()

        Dim iTask As Integer = Shell(mNconvertPath & " " & aCommandLine, AppWinStyle.Hide)
        Dim lHandle As Integer = OpenProcess(SYNCHRONIZE, False, iTask)
        Dim lResult As Integer = WaitForSingleObject(lHandle, INFINITE)
        lResult = CloseHandle(lHandle)
    End Sub

    Sub FindNconvert()
        mNconvertPath = GetSetting("Nconvert", "Paths", "ExePath", "")
        If Not IO.File.Exists(mNconvertPath) Then
            mNconvertPath = FindFile("Find Nconvert.exe to perform conversion", "Nconvert.exe")
            If IO.File.Exists(mNconvertPath) Then
                SaveSetting("Nconvert", "Paths", "ExePath", mNconvertPath)
            End If
        End If
    End Sub
	
    Sub OpenProject(ByRef aFileName As String, ByRef aTreeView As AxComctlLib.AxTreeView)
        frmMain.Cursor = System.Windows.Forms.Cursors.WaitCursor
        aTreeView.Visible = False
        If Not IO.File.Exists(aFileName) Then
            If Logger.Msg("File not found. Create new project file '" & aFileName & "'?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                IO.File.Create(aFileName)
            End If
        End If
        If IO.File.Exists(aFileName) Then
            pProjectFileName = aFileName
            pBaseName = FilenameOnly(aFileName)
            aTreeView.Nodes.Clear()
            aTreeView.Nodes.Add(, , "N" & pBaseName, pBaseName)
            aTreeView.Nodes(1).Expanded = True

            Dim lSectionName(50) As String 'Array of current section names for each level
            Dim lSectionLevel As Integer 'Level of current source file, according to indentation
            Dim lLevel As Integer 'Level in loop that constructs keys

            For Each lLine As String In LinesInFile(aFileName)
                Dim lThisName As String = lLine.TrimStart  'file name of current source file, minus extension
                If lThisName.Length > 0 Then
                    lSectionLevel = (lLine.Length - lThisName.Length) / 2 + 1 '2 spaces indentation per level
                    lThisName = lThisName.TrimEnd
                    Dim lKey As String = lThisName 'unique ID for tree control
                    lSectionName(lSectionLevel) = lThisName
                    Dim lNode As ComctlLib.Node 'Node inserted into tree control
                    If lSectionLevel = 1 Then
                        lNode = aTreeView.Nodes.Add("N" & pBaseName, ComctlLib.TreeRelationshipConstants.tvwChild, "N" & lKey, lThisName)
                    Else
                        For lLevel = lSectionLevel - 1 To 1 Step -1
                            lKey = lSectionName(lLevel) & "\" & lKey
                        Next lLevel
                        Try
                            lNode = aTreeView.Nodes.Add("N" & Left(lKey, Len(lKey) - Len(lThisName) - 1), ComctlLib.TreeRelationshipConstants.tvwChild, "N" & lKey, lThisName)
                            If Not lNode.Parent.Expanded Then lNode.Parent.Expanded = True
                        Catch
                            Debug.Print("Duplicate key in tree: " & lKey)
                        End Try
                    End If
                End If
            Next
            frmMain.AddRecentFile(pProjectFileName)
        End If
        aTreeView.Visible = True
        frmMain.Cursor = System.Windows.Forms.Cursors.Default
        If aTreeView.Nodes.Count > 0 Then aTreeView.Nodes(1).EnsureVisible()
        Exit Sub
    End Sub
	
    Public Sub SaveProject(ByRef aFileName As String, ByRef aTreeView As AxComctlLib.AxTreeView)
        Dim lNode As ComctlLib.Node 'Node of the tree being written

        'Mark all except first as need to be saved
        For Each lNode In aTreeView.Nodes
            If Not lNode Is Nothing Then lNode.Tag = True
        Next
        aTreeView.Nodes.Item(1).Tag = False

        Dim lOutWriter As IO.StreamWriter = New IO.StreamWriter(aFileName)
        lNode = aTreeView.Nodes.Item(1).Child
        While Not lNode Is Nothing
            WriteProjectSection(lNode, lOutWriter)
            lNode = lNode.Next
        End While
        lOutWriter.Close()
    End Sub
	
    Private Sub WriteProjectSection(ByRef aNode As ComctlLib.Node, ByRef aOutWriter As IO.StreamWriter)
        If aNode.Tag Then
            If Not aNode.Parent Is Nothing Then
                If aNode.Parent.Tag Then
                    WriteProjectSection(aNode.Parent, aOutWriter) 'Write parent first
                    Exit Sub 'Writing parent will lead to doing this node, so we are done
                End If
            End If
            aNode.Tag = False
            Dim lThisName As String = "" 'file name of current source file, minus extension
            Dim lPosition As Integer 'position of directory delimiter '\' in node key for counting levels
            lPosition = InStr(aNode.Key, "\")
            While lPosition > 0 And lPosition < aNode.Key.Length
                lThisName &= "  "
                lPosition = InStr(lPosition + 1, aNode.Key, "\")
            End While

            lThisName = lThisName & aNode.Text
            aOutWriter.WriteLine(lThisName)
            Dim lChild As ComctlLib.Node = aNode.Child
            For lPosition = 1 To aNode.Children
                WriteProjectSection(lChild, aOutWriter)
                lChild = lChild.Next
            Next
            lChild = Nothing
        End If
    End Sub
End Module