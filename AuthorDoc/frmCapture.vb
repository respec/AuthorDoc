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
        If MySampleForm IsNot Nothing Then MySampleForm.Hide()
	End Sub
	
	Private Sub TimerDelay_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles TimerDelay.Tick
        'Dim tempFilename As String
        'Dim cmdline As String

		TimerDelay.Enabled = False
		If optWindow.Checked Then
            SendKeys.Send("%{PRTSC}")
        Else
            SendKeys.Send("{PRTSC}")
        End If

        System.Windows.Forms.Application.DoEvents()

        If Clipboard.ContainsImage() Then
            If String.IsNullOrEmpty(Filename) Then
                Using lCdlg As New SaveFileDialog
                    lCdlg.Title = "Save As..."
                    lCdlg.ShowDialog()
                    Filename = lCdlg.FileName
                End Using
            End If
            If Not String.IsNullOrEmpty(Filename) Then

                Dim Screenshot As Image = Clipboard.GetImage()
                Screenshot.Save(Filename, System.Drawing.Imaging.ImageFormat.Png)

                If MySampleForm Is Nothing Then MySampleForm = New frmSample
                MySampleForm.SetImage(Filename)
                'Else
                '    tempFilename = IO.Path.Combine(IO.Path.GetTempPath, IO.Path.ChangeExtension(Filename, ".bmp"))
                '    pictCapture.Image.Save(tempFilename)
                '    frmSample.SetImage(tempFilename)
                '    If IO.File.Exists(Filename) Then Kill(Filename)
                '    ' -D = delete original, -quiet = no output, -o = output filename
                '    cmdline = "-D -o """ & Filename & """ -out " & VB.Right(Filename, 3) & " """ & tempFilename & """"
                '    RunNconvert(cmdline)
                '    'Kill tempFilename
                'End If
            End If
            Beep()
        End If
        frmMain.Show()
    End Sub
End Class