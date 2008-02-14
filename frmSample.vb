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
        Dim lFilename As String

        Me.Text = aFilename
        Select Case UCase(VB.Right(aFilename, 3))
            Case "BMP", "GIF"
                img.Image = System.Drawing.Image.FromFile(aFilename)
            Case Else
                lFilename = IO.Path.Combine(IO.Path.GetTempPath, FilenameOnly(aFilename) & ".bmp")
                ' -D = delete original, -quiet = no output, -o = output filename
                Dim lCmdline As String = "-o """ & lFilename & """ -out bmp """ & aFilename & """"
                RunNconvert(lCmdline)
                On Error GoTo ErrLoad
                img.Image = System.Drawing.Image.FromFile(lFilename)
                Kill(lFilename)
        End Select
        img.Visible = True
        txt.Visible = False
        frmMain.Visible = True
        Me.Show()
        Me.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(frmMain.Left) + VB6.PixelsToTwipsX(frmMain.Width))
        frmMain.Activate()
        Exit Sub

ErrLoad:
        Debug.Print(Err.Description)
        Resume Next
    End Sub
	
	Public Sub SetText(ByRef fullpath As String)
		Dim dotpos As Integer
		Dim Filename, ext As String
		Filename = FilenameOnly(fullpath)
		Me.Text = Filename
		txt.Visible = True
		img.Visible = False
		dotpos = InStrRev(fullpath, ".")
		If dotpos > 0 Then ext = Mid(fullpath, dotpos) Else ext = ""
        frmMain.LoadTextboxFromFile(IO.Path.GetDirectoryName(fullpath), Filename, ext, txt)
		Me.Show()
		frmMain.Activate()
	End Sub
	
	'UPGRADE_WARNING: Event frmSample.Resize may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub frmSample_Resize(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Resize
		If ClientRectangle.Width > 112 And VB6.PixelsToTwipsY(Height) > 375 Then
			img.Width = ClientRectangle.Width
			img.Height = ClientRectangle.Height
			txt.Width = VB6.PixelsToTwipsX(Width) - 108
			txt.Height = VB6.PixelsToTwipsY(Height) - 372
		End If
	End Sub
End Class