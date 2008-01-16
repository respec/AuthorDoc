<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmOptions
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
	Public WithEvents txtFont As System.Windows.Forms.TextBox
	Public WithEvents txtFindTimeout As System.Windows.Forms.TextBox
	Public WithEvents Command1 As System.Windows.Forms.Button
	Public WithEvents txtTreeIndent As System.Windows.Forms.TextBox
	Public WithEvents Label3 As System.Windows.Forms.Label
	Public WithEvents Label2 As System.Windows.Forms.Label
	Public WithEvents Label1 As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmOptions))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.txtFont = New System.Windows.Forms.TextBox
		Me.txtFindTimeout = New System.Windows.Forms.TextBox
		Me.Command1 = New System.Windows.Forms.Button
		Me.txtTreeIndent = New System.Windows.Forms.TextBox
		Me.Label3 = New System.Windows.Forms.Label
		Me.Label2 = New System.Windows.Forms.Label
		Me.Label1 = New System.Windows.Forms.Label
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		Me.Text = "Options"
		Me.ClientSize = New System.Drawing.Size(313, 188)
		Me.Location = New System.Drawing.Point(4, 27)
		Me.Icon = CType(resources.GetObject("frmOptions.Icon"), System.Drawing.Icon)
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
		Me.Name = "frmOptions"
		Me.txtFont.AutoSize = False
		Me.txtFont.Size = New System.Drawing.Size(201, 24)
		Me.txtFont.Location = New System.Drawing.Point(80, 90)
		Me.txtFont.TabIndex = 6
		Me.txtFont.Text = "Text1"
		Me.txtFont.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtFont.AcceptsReturn = True
		Me.txtFont.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtFont.BackColor = System.Drawing.SystemColors.Window
		Me.txtFont.CausesValidation = True
		Me.txtFont.Enabled = True
		Me.txtFont.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtFont.HideSelection = True
		Me.txtFont.ReadOnly = False
		Me.txtFont.Maxlength = 0
		Me.txtFont.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtFont.MultiLine = False
		Me.txtFont.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtFont.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtFont.TabStop = True
		Me.txtFont.Visible = True
		Me.txtFont.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtFont.Name = "txtFont"
		Me.txtFindTimeout.AutoSize = False
		Me.txtFindTimeout.Size = New System.Drawing.Size(111, 24)
		Me.txtFindTimeout.Location = New System.Drawing.Point(170, 60)
		Me.txtFindTimeout.TabIndex = 3
		Me.txtFindTimeout.Text = "Text1"
		Me.txtFindTimeout.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtFindTimeout.AcceptsReturn = True
		Me.txtFindTimeout.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtFindTimeout.BackColor = System.Drawing.SystemColors.Window
		Me.txtFindTimeout.CausesValidation = True
		Me.txtFindTimeout.Enabled = True
		Me.txtFindTimeout.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtFindTimeout.HideSelection = True
		Me.txtFindTimeout.ReadOnly = False
		Me.txtFindTimeout.Maxlength = 0
		Me.txtFindTimeout.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtFindTimeout.MultiLine = False
		Me.txtFindTimeout.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtFindTimeout.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtFindTimeout.TabStop = True
		Me.txtFindTimeout.Visible = True
		Me.txtFindTimeout.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtFindTimeout.Name = "txtFindTimeout"
		Me.Command1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.Command1.Text = "Ok"
		Me.Command1.Size = New System.Drawing.Size(111, 31)
		Me.Command1.Location = New System.Drawing.Point(100, 140)
		Me.Command1.TabIndex = 2
		Me.Command1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Command1.BackColor = System.Drawing.SystemColors.Control
		Me.Command1.CausesValidation = True
		Me.Command1.Enabled = True
		Me.Command1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Command1.Cursor = System.Windows.Forms.Cursors.Default
		Me.Command1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Command1.TabStop = True
		Me.Command1.Name = "Command1"
		Me.txtTreeIndent.AutoSize = False
		Me.txtTreeIndent.Size = New System.Drawing.Size(111, 24)
		Me.txtTreeIndent.Location = New System.Drawing.Point(170, 27)
		Me.txtTreeIndent.TabIndex = 0
		Me.txtTreeIndent.Text = "Text1"
		Me.txtTreeIndent.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtTreeIndent.AcceptsReturn = True
		Me.txtTreeIndent.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtTreeIndent.BackColor = System.Drawing.SystemColors.Window
		Me.txtTreeIndent.CausesValidation = True
		Me.txtTreeIndent.Enabled = True
		Me.txtTreeIndent.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtTreeIndent.HideSelection = True
		Me.txtTreeIndent.ReadOnly = False
		Me.txtTreeIndent.Maxlength = 0
		Me.txtTreeIndent.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtTreeIndent.MultiLine = False
		Me.txtTreeIndent.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtTreeIndent.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtTreeIndent.TabStop = True
		Me.txtTreeIndent.Visible = True
		Me.txtTreeIndent.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtTreeIndent.Name = "txtTreeIndent"
		Me.Label3.Text = "Font"
		Me.Label3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label3.Size = New System.Drawing.Size(41, 21)
		Me.Label3.Location = New System.Drawing.Point(30, 90)
		Me.Label3.TabIndex = 5
		Me.Label3.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label3.BackColor = System.Drawing.SystemColors.Control
		Me.Label3.Enabled = True
		Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label3.UseMnemonic = True
		Me.Label3.Visible = True
		Me.Label3.AutoSize = False
		Me.Label3.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label3.Name = "Label3"
		Me.Label2.Text = "Find Timeout"
		Me.Label2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label2.Size = New System.Drawing.Size(111, 21)
		Me.Label2.Location = New System.Drawing.Point(30, 60)
		Me.Label2.TabIndex = 4
		Me.Label2.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label2.BackColor = System.Drawing.SystemColors.Control
		Me.Label2.Enabled = True
		Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label2.UseMnemonic = True
		Me.Label2.Visible = True
		Me.Label2.AutoSize = False
		Me.Label2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label2.Name = "Label2"
		Me.Label1.Text = "Tree Indent"
		Me.Label1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label1.Size = New System.Drawing.Size(111, 21)
		Me.Label1.Location = New System.Drawing.Point(30, 30)
		Me.Label1.TabIndex = 1
		Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label1.BackColor = System.Drawing.SystemColors.Control
		Me.Label1.Enabled = True
		Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label1.UseMnemonic = True
		Me.Label1.Visible = True
		Me.Label1.AutoSize = False
		Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label1.Name = "Label1"
		Me.Controls.Add(txtFont)
		Me.Controls.Add(txtFindTimeout)
		Me.Controls.Add(Command1)
		Me.Controls.Add(txtTreeIndent)
		Me.Controls.Add(Label3)
		Me.Controls.Add(Label2)
		Me.Controls.Add(Label1)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class