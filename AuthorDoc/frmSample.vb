Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Imports atcUtility

Friend Class frmSample
	Inherits System.Windows.Forms.Form
    'Copyright 2000-2008 by AQUA TERRA Consultants
	
	Private Declare Function GetModuleUsage Lib "Kernel" (ByVal hModule As Short) As Short
	Private Const SW_HIDE As Short = 0 'Hidden Window
	Private Const SW_NORMAL As Short = 1 'Normal Window
	Private Const SW_MAXIMIZE As Short = 3 'Maximized Window
	Private Const SW_MINIMIZE As Short = 6 'Minimized Window
	
    Public Sub SetImage(ByRef aFilename As String)
        Me.Text = aFilename
        Select Case IO.Path.GetExtension(aFilename).ToUpper
            Case ".BMP", ".GIF", ".PNG", ".JPG"
                Try
                    img.Image = System.Drawing.Image.FromFile(aFilename)
                    img.Visible = True
                    txt.Visible = False
                    frmMain.Visible = True
                    Me.Show()
                    Me.Left = frmMain.Left + frmMain.Width
                    Me.BringToFront()
                    'Me.Width = 
                    frmMain.Activate()
                Catch
                    Me.Visible = False
                End Try
                'Case Else
                '    lFilename = IO.Path.Combine(IO.Path.GetTempPath, IO.Path.GetFileNameWithoutExtension(aFilename) & ".bmp")
                '    ' -D = delete original, -quiet = no output, -o = output filename
                '    Dim lCmdline As String = "-o """ & lFilename & """ -out bmp """ & aFilename & """"
                '    RunNconvert(lCmdline)
                '    Try
                '        img.Image = System.Drawing.Image.FromFile(lFilename)
                '        Kill(lFilename)
                '    Catch
                '        Debug.Print(Err.Description)
                '    End Try
        End Select
    End Sub

    Public Sub SetText(ByRef aFullPath As String)
        Dim lFilename As String = IO.Path.GetFileNameWithoutExtension(aFullPath)
        Me.Text = lFilename
        txt.Visible = True
        img.Visible = False
        Dim lDotPos As Integer = InStrRev(aFullPath, ".")
        Dim lExt As String
        If lDotPos > 0 Then lExt = Mid(aFullPath, lDotPos) Else lExt = ""
        frmMain.LoadTextboxFromFile(IO.Path.GetDirectoryName(aFullPath), lFilename, lExt, txt)
        StickToMainForm()
        Me.Show()
        frmMain.Activate()
    End Sub

    Public Sub StickToMainForm()
        Me.Location = New System.Drawing.Point(frmMain.Left + frmMain.Width, frmMain.Top)
    End Sub

End Class