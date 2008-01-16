<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmCapture
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
	Public WithEvents pictCapture As System.Windows.Forms.PictureBox
	Public WithEvents TimerDelay As System.Windows.Forms.Timer
	Public WithEvents cmdCapture As System.Windows.Forms.Button
	Public WithEvents txtDelay As System.Windows.Forms.TextBox
	Public WithEvents optScreen As System.Windows.Forms.RadioButton
	Public WithEvents optWindow As System.Windows.Forms.RadioButton
	Public WithEvents lblDelay As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmCapture))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.pictCapture = New System.Windows.Forms.PictureBox
		Me.TimerDelay = New System.Windows.Forms.Timer(components)
		Me.cmdCapture = New System.Windows.Forms.Button
		Me.txtDelay = New System.Windows.Forms.TextBox
		Me.optScreen = New System.Windows.Forms.RadioButton
		Me.optWindow = New System.Windows.Forms.RadioButton
		Me.lblDelay = New System.Windows.Forms.Label
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		Me.Text = "Capture"
		Me.ClientSize = New System.Drawing.Size(144, 154)
		Me.Location = New System.Drawing.Point(4, 27)
		Me.Icon = CType(resources.GetObject("frmCapture.Icon"), System.Drawing.Icon)
		Me.MaximizeBox = False
		Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultLocation
		Me.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.BackColor = System.Drawing.SystemColors.Control
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable
		Me.ControlBox = True
		Me.Enabled = True
		Me.KeyPreview = False
		Me.MinimizeBox = True
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.ShowInTaskbar = True
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "frmCapture"
		Me.pictCapture.Size = New System.Drawing.Size(31, 21)
		Me.pictCapture.Location = New System.Drawing.Point(110, 10)
		Me.pictCapture.TabIndex = 5
		Me.pictCapture.Visible = False
		Me.pictCapture.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.pictCapture.Dock = System.Windows.Forms.DockStyle.None
		Me.pictCapture.BackColor = System.Drawing.SystemColors.Control
		Me.pictCapture.CausesValidation = True
		Me.pictCapture.Enabled = True
		Me.pictCapture.ForeColor = System.Drawing.SystemColors.ControlText
		Me.pictCapture.Cursor = System.Windows.Forms.Cursors.Default
		Me.pictCapture.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.pictCapture.TabStop = True
		Me.pictCapture.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
		Me.pictCapture.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.pictCapture.Name = "pictCapture"
		Me.TimerDelay.Enabled = False
		Me.TimerDelay.Interval = 1
		Me.cmdCapture.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdCapture.Text = "Capture"
		Me.AcceptButton = Me.cmdCapture
		Me.cmdCapture.Size = New System.Drawing.Size(81, 31)
		Me.cmdCapture.Location = New System.Drawing.Point(20, 110)
		Me.cmdCapture.TabIndex = 4
		Me.cmdCapture.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdCapture.BackColor = System.Drawing.SystemColors.Control
		Me.cmdCapture.CausesValidation = True
		Me.cmdCapture.Enabled = True
		Me.cmdCapture.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdCapture.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdCapture.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdCapture.TabStop = True
		Me.cmdCapture.Name = "cmdCapture"
		Me.txtDelay.AutoSize = False
		Me.txtDelay.Size = New System.Drawing.Size(21, 24)
		Me.txtDelay.Location = New System.Drawing.Point(20, 70)
		Me.txtDelay.TabIndex = 2
		Me.txtDelay.Text = "5"
		Me.txtDelay.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtDelay.AcceptsReturn = True
		Me.txtDelay.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtDelay.BackColor = System.Drawing.SystemColors.Window
		Me.txtDelay.CausesValidation = True
		Me.txtDelay.Enabled = True
		Me.txtDelay.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtDelay.HideSelection = True
		Me.txtDelay.ReadOnly = False
		Me.txtDelay.Maxlength = 0
		Me.txtDelay.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtDelay.MultiLine = False
		Me.txtDelay.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtDelay.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtDelay.TabStop = True
		Me.txtDelay.Visible = True
		Me.txtDelay.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtDelay.Name = "txtDelay"
		Me.optScreen.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optScreen.Text = "Screen"
		Me.optScreen.Size = New System.Drawing.Size(91, 16)
		Me.optScreen.Location = New System.Drawing.Point(20, 40)
		Me.optScreen.TabIndex = 1
		Me.optScreen.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.optScreen.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optScreen.BackColor = System.Drawing.SystemColors.Control
		Me.optScreen.CausesValidation = True
		Me.optScreen.Enabled = True
		Me.optScreen.ForeColor = System.Drawing.SystemColors.ControlText
		Me.optScreen.Cursor = System.Windows.Forms.Cursors.Default
		Me.optScreen.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.optScreen.Appearance = System.Windows.Forms.Appearance.Normal
		Me.optScreen.TabStop = True
		Me.optScreen.Checked = False
		Me.optScreen.Visible = True
		Me.optScreen.Name = "optScreen"
		Me.optWindow.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optWindow.Text = "Window"
		Me.optWindow.Size = New System.Drawing.Size(91, 16)
		Me.optWindow.Location = New System.Drawing.Point(20, 20)
		Me.optWindow.TabIndex = 0
		Me.optWindow.Checked = True
		Me.optWindow.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.optWindow.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optWindow.BackColor = System.Drawing.SystemColors.Control
		Me.optWindow.CausesValidation = True
		Me.optWindow.Enabled = True
		Me.optWindow.ForeColor = System.Drawing.SystemColors.ControlText
		Me.optWindow.Cursor = System.Windows.Forms.Cursors.Default
		Me.optWindow.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.optWindow.Appearance = System.Windows.Forms.Appearance.Normal
		Me.optWindow.TabStop = True
		Me.optWindow.Visible = True
		Me.optWindow.Name = "optWindow"
		Me.lblDelay.Text = "sec. delay"
		Me.lblDelay.Size = New System.Drawing.Size(71, 21)
		Me.lblDelay.Location = New System.Drawing.Point(50, 75)
		Me.lblDelay.TabIndex = 3
		Me.lblDelay.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblDelay.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblDelay.BackColor = System.Drawing.SystemColors.Control
		Me.lblDelay.Enabled = True
		Me.lblDelay.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblDelay.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblDelay.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblDelay.UseMnemonic = True
		Me.lblDelay.Visible = True
		Me.lblDelay.AutoSize = False
		Me.lblDelay.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblDelay.Name = "lblDelay"
		Me.Controls.Add(pictCapture)
		Me.Controls.Add(cmdCapture)
		Me.Controls.Add(txtDelay)
		Me.Controls.Add(optScreen)
		Me.Controls.Add(optWindow)
		Me.Controls.Add(lblDelay)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class