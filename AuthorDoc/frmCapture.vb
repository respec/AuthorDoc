Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Imports atcUtility

Friend Class frmCapture
	Inherits System.Windows.Forms.Form
    'Copyright 2000-2008 by AQUA TERRA Consultants
	
	Public Filename As String
	
	Private Sub cmdCapture_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCapture.Click
        Dim lDelaySeconds As Double = 0.1
		If IsNumeric(txtDelay.Text) Then
            lDelaySeconds = Math.Max(CSng(txtDelay.Text), 0.0001) 'be sure > 0
        End If
        TimerDelay.Interval = 1000 * lDelaySeconds
		TimerDelay.Enabled = True
		Me.Hide()
		frmMain.Hide()
		frmSample.Hide()
	End Sub
	
	Private Sub TimerDelay_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles TimerDelay.Tick
        Dim tempFilename As String
		Dim cmdline As String

		TimerDelay.Enabled = False
		If optWindow.Checked Then
            'pictCapture.Image = CaptureActiveWindow()
		Else
            'pictCapture.Image = CaptureScreen()
		End If
		If Len(Filename) < 1 Then
            Using lCdlg As New SaveFileDialog
                lCdlg.Title = "Save As..."
                lCdlg.ShowDialog()
                Filename = lCdlg.FileName
            End Using
        End If
		If Len(Filename) > 0 Then
			If LCase(VB.Right(Filename, 4)) = ".bmp" Then
				'UPGRADE_WARNING: SavePicture was upgraded to System.Drawing.Image.Save and has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
				pictCapture.Image.Save(Filename)
				frmSample.SetImage(Filename)
			Else
                tempFilename = IO.Path.Combine(IO.Path.GetTempPath, IO.Path.GetFileNameWithoutExtension(Filename) & ".bmp")
				'UPGRADE_WARNING: SavePicture was upgraded to System.Drawing.Image.Save and has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
				pictCapture.Image.Save(tempFilename)
				frmSample.SetImage(tempFilename)
                If IO.File.Exists(Filename) Then Kill(Filename)
				' -D = delete original, -quiet = no output, -o = output filename
				cmdline = "-D -o """ & Filename & """ -out " & VB.Right(Filename, 3) & " """ & tempFilename & """"
				RunNconvert(cmdline)
				'Kill tempFilename
			End If
		End If
		Beep()
		frmMain.Show()
	End Sub
End Class