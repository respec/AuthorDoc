<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmFind
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
	Public WithEvents cmdClose As System.Windows.Forms.Button
	Public WithEvents chkCase As System.Windows.Forms.CheckBox
	Public WithEvents chkScope As System.Windows.Forms.CheckBox
	Public WithEvents cmdReplaceAll As System.Windows.Forms.Button
	Public WithEvents cmdReplace As System.Windows.Forms.Button
	Public WithEvents cmdFind As System.Windows.Forms.Button
	Public WithEvents txtReplace As System.Windows.Forms.TextBox
	Public WithEvents txtFind As System.Windows.Forms.TextBox
	Public WithEvents Label1 As System.Windows.Forms.Label
	Public WithEvents lblFind As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmFind))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.cmdClose = New System.Windows.Forms.Button
		Me.chkCase = New System.Windows.Forms.CheckBox
		Me.chkScope = New System.Windows.Forms.CheckBox
		Me.cmdReplaceAll = New System.Windows.Forms.Button
		Me.cmdReplace = New System.Windows.Forms.Button
		Me.cmdFind = New System.Windows.Forms.Button
		Me.txtReplace = New System.Windows.Forms.TextBox
		Me.txtFind = New System.Windows.Forms.TextBox
		Me.Label1 = New System.Windows.Forms.Label
		Me.lblFind = New System.Windows.Forms.Label
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		Me.Text = "Find"
		Me.ClientSize = New System.Drawing.Size(409, 157)
		Me.Location = New System.Drawing.Point(4, 23)
		Me.Icon = CType(resources.GetObject("frmFind.Icon"), System.Drawing.Icon)
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
		Me.Name = "frmFind"
		Me.cmdClose.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.CancelButton = Me.cmdClose
		Me.cmdClose.Text = "Done"
		Me.cmdClose.Size = New System.Drawing.Size(71, 21)
		Me.cmdClose.Location = New System.Drawing.Point(320, 120)
		Me.cmdClose.TabIndex = 9
		Me.cmdClose.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdClose.BackColor = System.Drawing.SystemColors.Control
		Me.cmdClose.CausesValidation = True
		Me.cmdClose.Enabled = True
		Me.cmdClose.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdClose.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdClose.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdClose.TabStop = True
		Me.cmdClose.Name = "cmdClose"
		Me.chkCase.Text = "Match Case"
		Me.chkCase.Size = New System.Drawing.Size(171, 21)
		Me.chkCase.Location = New System.Drawing.Point(110, 110)
		Me.chkCase.TabIndex = 8
		Me.chkCase.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkCase.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkCase.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkCase.BackColor = System.Drawing.SystemColors.Control
		Me.chkCase.CausesValidation = True
		Me.chkCase.Enabled = True
		Me.chkCase.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkCase.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkCase.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkCase.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkCase.TabStop = True
		Me.chkCase.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkCase.Visible = True
		Me.chkCase.Name = "chkCase"
		Me.chkScope.Text = "Search Whole Project"
		Me.chkScope.Size = New System.Drawing.Size(161, 21)
		Me.chkScope.Location = New System.Drawing.Point(110, 80)
		Me.chkScope.TabIndex = 7
		Me.chkScope.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkScope.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkScope.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkScope.BackColor = System.Drawing.SystemColors.Control
		Me.chkScope.CausesValidation = True
		Me.chkScope.Enabled = True
		Me.chkScope.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkScope.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkScope.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkScope.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkScope.TabStop = True
		Me.chkScope.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkScope.Visible = True
		Me.chkScope.Name = "chkScope"
		Me.cmdReplaceAll.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdReplaceAll.Text = "Replace All"
		Me.cmdReplaceAll.Size = New System.Drawing.Size(71, 41)
		Me.cmdReplaceAll.Location = New System.Drawing.Point(320, 70)
		Me.cmdReplaceAll.TabIndex = 6
		Me.cmdReplaceAll.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdReplaceAll.BackColor = System.Drawing.SystemColors.Control
		Me.cmdReplaceAll.CausesValidation = True
		Me.cmdReplaceAll.Enabled = True
		Me.cmdReplaceAll.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdReplaceAll.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdReplaceAll.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdReplaceAll.TabStop = True
		Me.cmdReplaceAll.Name = "cmdReplaceAll"
		Me.cmdReplace.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdReplace.Text = "Replace"
		Me.cmdReplace.Size = New System.Drawing.Size(71, 21)
		Me.cmdReplace.Location = New System.Drawing.Point(320, 40)
		Me.cmdReplace.TabIndex = 5
		Me.cmdReplace.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdReplace.BackColor = System.Drawing.SystemColors.Control
		Me.cmdReplace.CausesValidation = True
		Me.cmdReplace.Enabled = True
		Me.cmdReplace.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdReplace.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdReplace.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdReplace.TabStop = True
		Me.cmdReplace.Name = "cmdReplace"
		Me.cmdFind.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdFind.Text = "Find"
		Me.AcceptButton = Me.cmdFind
		Me.cmdFind.Size = New System.Drawing.Size(71, 21)
		Me.cmdFind.Location = New System.Drawing.Point(320, 10)
		Me.cmdFind.TabIndex = 4
		Me.cmdFind.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdFind.BackColor = System.Drawing.SystemColors.Control
		Me.cmdFind.CausesValidation = True
		Me.cmdFind.Enabled = True
		Me.cmdFind.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdFind.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdFind.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdFind.TabStop = True
		Me.cmdFind.Name = "cmdFind"
		Me.txtReplace.AutoSize = False
		Me.txtReplace.Size = New System.Drawing.Size(201, 24)
		Me.txtReplace.Location = New System.Drawing.Point(110, 40)
		Me.txtReplace.TabIndex = 3
		Me.txtReplace.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtReplace.AcceptsReturn = True
		Me.txtReplace.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtReplace.BackColor = System.Drawing.SystemColors.Window
		Me.txtReplace.CausesValidation = True
		Me.txtReplace.Enabled = True
		Me.txtReplace.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtReplace.HideSelection = True
		Me.txtReplace.ReadOnly = False
		Me.txtReplace.Maxlength = 0
		Me.txtReplace.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtReplace.MultiLine = False
		Me.txtReplace.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtReplace.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtReplace.TabStop = True
		Me.txtReplace.Visible = True
		Me.txtReplace.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtReplace.Name = "txtReplace"
		Me.txtFind.AutoSize = False
		Me.txtFind.Size = New System.Drawing.Size(201, 24)
		Me.txtFind.Location = New System.Drawing.Point(110, 10)
		Me.txtFind.TabIndex = 2
		Me.txtFind.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtFind.AcceptsReturn = True
		Me.txtFind.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtFind.BackColor = System.Drawing.SystemColors.Window
		Me.txtFind.CausesValidation = True
		Me.txtFind.Enabled = True
		Me.txtFind.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtFind.HideSelection = True
		Me.txtFind.ReadOnly = False
		Me.txtFind.Maxlength = 0
		Me.txtFind.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtFind.MultiLine = False
		Me.txtFind.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtFind.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtFind.TabStop = True
		Me.txtFind.Visible = True
		Me.txtFind.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtFind.Name = "txtFind"
		Me.Label1.Text = "Replace With:"
		Me.Label1.Size = New System.Drawing.Size(101, 21)
		Me.Label1.Location = New System.Drawing.Point(10, 40)
		Me.Label1.TabIndex = 1
		Me.Label1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
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
		Me.lblFind.Text = "Find Text:"
		Me.lblFind.Size = New System.Drawing.Size(101, 21)
		Me.lblFind.Location = New System.Drawing.Point(10, 10)
		Me.lblFind.TabIndex = 0
		Me.lblFind.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblFind.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblFind.BackColor = System.Drawing.SystemColors.Control
		Me.lblFind.Enabled = True
		Me.lblFind.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblFind.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblFind.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblFind.UseMnemonic = True
		Me.lblFind.Visible = True
		Me.lblFind.AutoSize = False
		Me.lblFind.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblFind.Name = "lblFind"
		Me.Controls.Add(cmdClose)
		Me.Controls.Add(chkCase)
		Me.Controls.Add(chkScope)
		Me.Controls.Add(cmdReplaceAll)
		Me.Controls.Add(cmdReplace)
		Me.Controls.Add(cmdFind)
		Me.Controls.Add(txtReplace)
		Me.Controls.Add(txtFind)
		Me.Controls.Add(Label1)
		Me.Controls.Add(lblFind)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class