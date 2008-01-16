<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmSample
#Region "Windows Form Designer generated code "
	<System.Diagnostics.DebuggerNonUserCode()> Public Sub New()
		MyBase.New()
		'This call is required by the Windows Form Designer.
		InitializeComponent()
	End Sub
	'Form overrides dispose to clean up the component list.
	<System.Diagnostics.DebuggerNonUserCode()> Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
		If Disposing Then
			If Not components Is Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(Disposing)
	End Sub
	'Required by the Windows Form Designer
	Private components As System.ComponentModel.IContainer
	Public ToolTip1 As System.Windows.Forms.ToolTip
	Public cdlgOpen As System.Windows.Forms.OpenFileDialog
	Public cdlgSave As System.Windows.Forms.SaveFileDialog
	Public cdlgFont As System.Windows.Forms.FontDialog
	Public cdlgColor As System.Windows.Forms.ColorDialog
	Public cdlgPrint As System.Windows.Forms.PrintDialog
	Public WithEvents img As System.Windows.Forms.PictureBox
	Public WithEvents txt As System.Windows.Forms.RichTextBox
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmSample))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.cdlgOpen = New System.Windows.Forms.OpenFileDialog
		Me.cdlgSave = New System.Windows.Forms.SaveFileDialog
		Me.cdlgFont = New System.Windows.Forms.FontDialog
		Me.cdlgColor = New System.Windows.Forms.ColorDialog
		Me.cdlgPrint = New System.Windows.Forms.PrintDialog
		Me.img = New System.Windows.Forms.PictureBox
		Me.txt = New System.Windows.Forms.RichTextBox
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		Me.Text = "Sample"
		Me.ClientSize = New System.Drawing.Size(292, 225)
		Me.Location = New System.Drawing.Point(4, 27)
		Me.Icon = CType(resources.GetObject("frmSample.Icon"), System.Drawing.Icon)
		Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultLocation
		Me.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.BackColor = System.Drawing.SystemColors.Control
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable
		Me.ControlBox = True
		Me.Enabled = True
		Me.KeyPreview = False
		Me.MaximizeBox = True
		Me.MinimizeBox = True
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.ShowInTaskbar = True
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "frmSample"
		Me.img.BackColor = System.Drawing.Color.White
		Me.img.Font = New System.Drawing.Font("Courier New", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.img.Size = New System.Drawing.Size(261, 211)
		Me.img.Location = New System.Drawing.Point(0, 0)
		Me.img.TabIndex = 0
		Me.img.Dock = System.Windows.Forms.DockStyle.None
		Me.img.CausesValidation = True
		Me.img.Enabled = True
		Me.img.ForeColor = System.Drawing.SystemColors.ControlText
		Me.img.Cursor = System.Windows.Forms.Cursors.Default
		Me.img.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.img.TabStop = True
		Me.img.Visible = True
		Me.img.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
		Me.img.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.img.Name = "img"
		Me.txt.Size = New System.Drawing.Size(281, 191)
		Me.txt.Location = New System.Drawing.Point(0, 0)
		Me.txt.TabIndex = 1
		Me.txt.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.Vertical
		Me.txt.RTF = resources.GetString("txt.TextRTF")
		Me.txt.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txt.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txt.Name = "txt"
		Me.Controls.Add(img)
		Me.Controls.Add(txt)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class