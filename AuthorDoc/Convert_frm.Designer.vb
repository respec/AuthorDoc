<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmConvert
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
	Public WithEvents Text1 As System.Windows.Forms.TextBox
	Public WithEvents _cmdConvert_1 As System.Windows.Forms.Button
	Public WithEvents _optTargetFormat_0 As System.Windows.Forms.RadioButton
	Public WithEvents _optTargetFormat_4 As System.Windows.Forms.RadioButton
	Public WithEvents chkID As System.Windows.Forms.CheckBox
	Public WithEvents UpNextCheck As System.Windows.Forms.CheckBox
	Public WithEvents TimestampCheck As System.Windows.Forms.CheckBox
	Public WithEvents _cmdConvert_0 As System.Windows.Forms.Button
	Public WithEvents ContentsCheck As System.Windows.Forms.CheckBox
	Public WithEvents _optTargetFormat_3 As System.Windows.Forms.RadioButton
	Public WithEvents _optTargetFormat_2 As System.Windows.Forms.RadioButton
	Public WithEvents _optTargetFormat_1 As System.Windows.Forms.RadioButton
	Public WithEvents ProjectCheck As System.Windows.Forms.CheckBox
	Public WithEvents frameConvertTo As System.Windows.Forms.GroupBox
	Public CmDialog1Open As System.Windows.Forms.OpenFileDialog
	Public CmDialog1Save As System.Windows.Forms.SaveFileDialog
	Public CmDialog1Font As System.Windows.Forms.FontDialog
	Public CmDialog1Color As System.Windows.Forms.ColorDialog
	Public CmDialog1Print As System.Windows.Forms.PrintDialog
	Public WithEvents cmdConvert As Microsoft.VisualBasic.Compatibility.VB6.ButtonArray
	Public WithEvents optTargetFormat As Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmConvert))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.Text1 = New System.Windows.Forms.TextBox
		Me.frameConvertTo = New System.Windows.Forms.GroupBox
		Me._cmdConvert_1 = New System.Windows.Forms.Button
		Me._optTargetFormat_0 = New System.Windows.Forms.RadioButton
		Me._optTargetFormat_4 = New System.Windows.Forms.RadioButton
		Me.chkID = New System.Windows.Forms.CheckBox
		Me.UpNextCheck = New System.Windows.Forms.CheckBox
		Me.TimestampCheck = New System.Windows.Forms.CheckBox
		Me._cmdConvert_0 = New System.Windows.Forms.Button
		Me.ContentsCheck = New System.Windows.Forms.CheckBox
		Me._optTargetFormat_3 = New System.Windows.Forms.RadioButton
		Me._optTargetFormat_2 = New System.Windows.Forms.RadioButton
		Me._optTargetFormat_1 = New System.Windows.Forms.RadioButton
		Me.ProjectCheck = New System.Windows.Forms.CheckBox
		Me.CmDialog1Open = New System.Windows.Forms.OpenFileDialog
		Me.CmDialog1Save = New System.Windows.Forms.SaveFileDialog
		Me.CmDialog1Font = New System.Windows.Forms.FontDialog
		Me.CmDialog1Color = New System.Windows.Forms.ColorDialog
		Me.CmDialog1Print = New System.Windows.Forms.PrintDialog
		Me.cmdConvert = New Microsoft.VisualBasic.Compatibility.VB6.ButtonArray(components)
		Me.optTargetFormat = New Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray(components)
		Me.frameConvertTo.SuspendLayout()
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		CType(Me.cmdConvert, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.optTargetFormat, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
		Me.Text = "HSPF Documentation Hypertext Converter"
		Me.ClientSize = New System.Drawing.Size(238, 450)
		Me.Location = New System.Drawing.Point(204, 58)
		Me.Icon = CType(resources.GetObject("frmConvert.Icon"), System.Drawing.Icon)
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
		Me.Name = "frmConvert"
		Me.Text1.AutoSize = False
		Me.Text1.BackColor = System.Drawing.SystemColors.Control
		Me.Text1.Size = New System.Drawing.Size(212, 72)
		Me.Text1.Location = New System.Drawing.Point(10, 370)
		Me.Text1.MultiLine = True
		Me.Text1.TabIndex = 10
		Me.Text1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Text1.AcceptsReturn = True
		Me.Text1.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.Text1.CausesValidation = True
		Me.Text1.Enabled = True
		Me.Text1.ForeColor = System.Drawing.SystemColors.WindowText
		Me.Text1.HideSelection = True
		Me.Text1.ReadOnly = False
		Me.Text1.Maxlength = 0
		Me.Text1.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.Text1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Text1.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.Text1.TabStop = True
		Me.Text1.Visible = True
		Me.Text1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.Text1.Name = "Text1"
		Me.frameConvertTo.Text = "Convert to:"
		Me.frameConvertTo.Size = New System.Drawing.Size(212, 351)
		Me.frameConvertTo.Location = New System.Drawing.Point(10, 10)
		Me.frameConvertTo.TabIndex = 0
		Me.frameConvertTo.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.frameConvertTo.BackColor = System.Drawing.SystemColors.Control
		Me.frameConvertTo.Enabled = True
		Me.frameConvertTo.ForeColor = System.Drawing.SystemColors.ControlText
		Me.frameConvertTo.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.frameConvertTo.Visible = True
		Me.frameConvertTo.Name = "frameConvertTo"
		Me._cmdConvert_1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me._cmdConvert_1.Text = "Preview"
		Me._cmdConvert_1.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._cmdConvert_1.Size = New System.Drawing.Size(92, 32)
		Me._cmdConvert_1.Location = New System.Drawing.Point(110, 310)
		Me._cmdConvert_1.TabIndex = 13
		Me._cmdConvert_1.BackColor = System.Drawing.SystemColors.Control
		Me._cmdConvert_1.CausesValidation = True
		Me._cmdConvert_1.Enabled = True
		Me._cmdConvert_1.ForeColor = System.Drawing.SystemColors.ControlText
		Me._cmdConvert_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._cmdConvert_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._cmdConvert_1.TabStop = True
		Me._cmdConvert_1.Name = "_cmdConvert_1"
		Me._optTargetFormat_0.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._optTargetFormat_0.Text = "From HSPF Manual"
		Me._optTargetFormat_0.Size = New System.Drawing.Size(192, 22)
		Me._optTargetFormat_0.Location = New System.Drawing.Point(10, 120)
		Me._optTargetFormat_0.TabIndex = 12
		Me._optTargetFormat_0.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._optTargetFormat_0.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._optTargetFormat_0.BackColor = System.Drawing.SystemColors.Control
		Me._optTargetFormat_0.CausesValidation = True
		Me._optTargetFormat_0.Enabled = True
		Me._optTargetFormat_0.ForeColor = System.Drawing.SystemColors.ControlText
		Me._optTargetFormat_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._optTargetFormat_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._optTargetFormat_0.Appearance = System.Windows.Forms.Appearance.Normal
		Me._optTargetFormat_0.TabStop = True
		Me._optTargetFormat_0.Checked = False
		Me._optTargetFormat_0.Visible = True
		Me._optTargetFormat_0.Name = "_optTargetFormat_0"
		Me._optTargetFormat_4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._optTargetFormat_4.Text = "HTML Help"
		Me._optTargetFormat_4.Size = New System.Drawing.Size(192, 22)
		Me._optTargetFormat_4.Location = New System.Drawing.Point(10, 75)
		Me._optTargetFormat_4.TabIndex = 11
		Me._optTargetFormat_4.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._optTargetFormat_4.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._optTargetFormat_4.BackColor = System.Drawing.SystemColors.Control
		Me._optTargetFormat_4.CausesValidation = True
		Me._optTargetFormat_4.Enabled = True
		Me._optTargetFormat_4.ForeColor = System.Drawing.SystemColors.ControlText
		Me._optTargetFormat_4.Cursor = System.Windows.Forms.Cursors.Default
		Me._optTargetFormat_4.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._optTargetFormat_4.Appearance = System.Windows.Forms.Appearance.Normal
		Me._optTargetFormat_4.TabStop = True
		Me._optTargetFormat_4.Checked = False
		Me._optTargetFormat_4.Visible = True
		Me._optTargetFormat_4.Name = "_optTargetFormat_4"
		Me.chkID.Text = "HelpContextID File"
		Me.chkID.Size = New System.Drawing.Size(192, 22)
		Me.chkID.Location = New System.Drawing.Point(10, 160)
		Me.chkID.TabIndex = 9
		Me.chkID.CheckState = System.Windows.Forms.CheckState.Checked
		Me.chkID.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkID.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkID.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkID.BackColor = System.Drawing.SystemColors.Control
		Me.chkID.CausesValidation = True
		Me.chkID.Enabled = True
		Me.chkID.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkID.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkID.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkID.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkID.TabStop = True
		Me.chkID.Visible = True
		Me.chkID.Name = "chkID"
		Me.UpNextCheck.Text = "Up/Next Navigation"
		Me.UpNextCheck.Size = New System.Drawing.Size(192, 22)
		Me.UpNextCheck.Location = New System.Drawing.Point(10, 220)
		Me.UpNextCheck.TabIndex = 8
		Me.UpNextCheck.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.UpNextCheck.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.UpNextCheck.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.UpNextCheck.BackColor = System.Drawing.SystemColors.Control
		Me.UpNextCheck.CausesValidation = True
		Me.UpNextCheck.Enabled = True
		Me.UpNextCheck.ForeColor = System.Drawing.SystemColors.ControlText
		Me.UpNextCheck.Cursor = System.Windows.Forms.Cursors.Default
		Me.UpNextCheck.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.UpNextCheck.Appearance = System.Windows.Forms.Appearance.Normal
		Me.UpNextCheck.TabStop = True
		Me.UpNextCheck.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.UpNextCheck.Visible = True
		Me.UpNextCheck.Name = "UpNextCheck"
		Me.TimestampCheck.Text = "Footer Timestamps"
		Me.TimestampCheck.Size = New System.Drawing.Size(192, 22)
		Me.TimestampCheck.Location = New System.Drawing.Point(10, 280)
		Me.TimestampCheck.TabIndex = 7
		Me.TimestampCheck.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.TimestampCheck.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.TimestampCheck.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.TimestampCheck.BackColor = System.Drawing.SystemColors.Control
		Me.TimestampCheck.CausesValidation = True
		Me.TimestampCheck.Enabled = True
		Me.TimestampCheck.ForeColor = System.Drawing.SystemColors.ControlText
		Me.TimestampCheck.Cursor = System.Windows.Forms.Cursors.Default
		Me.TimestampCheck.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.TimestampCheck.Appearance = System.Windows.Forms.Appearance.Normal
		Me.TimestampCheck.TabStop = True
		Me.TimestampCheck.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.TimestampCheck.Visible = True
		Me.TimestampCheck.Name = "TimestampCheck"
		Me._cmdConvert_0.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me._cmdConvert_0.Text = "Convert"
		Me._cmdConvert_0.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._cmdConvert_0.Size = New System.Drawing.Size(92, 32)
		Me._cmdConvert_0.Location = New System.Drawing.Point(10, 310)
		Me._cmdConvert_0.TabIndex = 6
		Me._cmdConvert_0.BackColor = System.Drawing.SystemColors.Control
		Me._cmdConvert_0.CausesValidation = True
		Me._cmdConvert_0.Enabled = True
		Me._cmdConvert_0.ForeColor = System.Drawing.SystemColors.ControlText
		Me._cmdConvert_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._cmdConvert_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._cmdConvert_0.TabStop = True
		Me._cmdConvert_0.Name = "_cmdConvert_0"
		Me.ContentsCheck.Text = "Contents"
		Me.ContentsCheck.Size = New System.Drawing.Size(192, 22)
		Me.ContentsCheck.Location = New System.Drawing.Point(10, 250)
		Me.ContentsCheck.TabIndex = 4
		Me.ContentsCheck.CheckState = System.Windows.Forms.CheckState.Checked
		Me.ContentsCheck.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.ContentsCheck.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.ContentsCheck.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.ContentsCheck.BackColor = System.Drawing.SystemColors.Control
		Me.ContentsCheck.CausesValidation = True
		Me.ContentsCheck.Enabled = True
		Me.ContentsCheck.ForeColor = System.Drawing.SystemColors.ControlText
		Me.ContentsCheck.Cursor = System.Windows.Forms.Cursors.Default
		Me.ContentsCheck.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.ContentsCheck.Appearance = System.Windows.Forms.Appearance.Normal
		Me.ContentsCheck.TabStop = True
		Me.ContentsCheck.Visible = True
		Me.ContentsCheck.Name = "ContentsCheck"
		Me._optTargetFormat_3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._optTargetFormat_3.Text = "Windows Help"
		Me._optTargetFormat_3.Size = New System.Drawing.Size(192, 22)
		Me._optTargetFormat_3.Location = New System.Drawing.Point(10, 53)
		Me._optTargetFormat_3.TabIndex = 3
		Me._optTargetFormat_3.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._optTargetFormat_3.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._optTargetFormat_3.BackColor = System.Drawing.SystemColors.Control
		Me._optTargetFormat_3.CausesValidation = True
		Me._optTargetFormat_3.Enabled = True
		Me._optTargetFormat_3.ForeColor = System.Drawing.SystemColors.ControlText
		Me._optTargetFormat_3.Cursor = System.Windows.Forms.Cursors.Default
		Me._optTargetFormat_3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._optTargetFormat_3.Appearance = System.Windows.Forms.Appearance.Normal
		Me._optTargetFormat_3.TabStop = True
		Me._optTargetFormat_3.Checked = False
		Me._optTargetFormat_3.Visible = True
		Me._optTargetFormat_3.Name = "_optTargetFormat_3"
		Me._optTargetFormat_2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._optTargetFormat_2.Text = "Printable Document"
		Me._optTargetFormat_2.Size = New System.Drawing.Size(192, 22)
		Me._optTargetFormat_2.Location = New System.Drawing.Point(10, 30)
		Me._optTargetFormat_2.TabIndex = 2
		Me._optTargetFormat_2.Checked = True
		Me._optTargetFormat_2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._optTargetFormat_2.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._optTargetFormat_2.BackColor = System.Drawing.SystemColors.Control
		Me._optTargetFormat_2.CausesValidation = True
		Me._optTargetFormat_2.Enabled = True
		Me._optTargetFormat_2.ForeColor = System.Drawing.SystemColors.ControlText
		Me._optTargetFormat_2.Cursor = System.Windows.Forms.Cursors.Default
		Me._optTargetFormat_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._optTargetFormat_2.Appearance = System.Windows.Forms.Appearance.Normal
		Me._optTargetFormat_2.TabStop = True
		Me._optTargetFormat_2.Visible = True
		Me._optTargetFormat_2.Name = "_optTargetFormat_2"
		Me._optTargetFormat_1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._optTargetFormat_1.Text = "HTML Pages"
		Me._optTargetFormat_1.Size = New System.Drawing.Size(162, 22)
		Me._optTargetFormat_1.Location = New System.Drawing.Point(10, 98)
		Me._optTargetFormat_1.TabIndex = 1
		Me._optTargetFormat_1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._optTargetFormat_1.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._optTargetFormat_1.BackColor = System.Drawing.SystemColors.Control
		Me._optTargetFormat_1.CausesValidation = True
		Me._optTargetFormat_1.Enabled = True
		Me._optTargetFormat_1.ForeColor = System.Drawing.SystemColors.ControlText
		Me._optTargetFormat_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._optTargetFormat_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._optTargetFormat_1.Appearance = System.Windows.Forms.Appearance.Normal
		Me._optTargetFormat_1.TabStop = True
		Me._optTargetFormat_1.Checked = False
		Me._optTargetFormat_1.Visible = True
		Me._optTargetFormat_1.Name = "_optTargetFormat_1"
		Me.ProjectCheck.Text = "Project File"
		Me.ProjectCheck.Size = New System.Drawing.Size(192, 22)
		Me.ProjectCheck.Location = New System.Drawing.Point(10, 190)
		Me.ProjectCheck.TabIndex = 5
		Me.ProjectCheck.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.ProjectCheck.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.ProjectCheck.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.ProjectCheck.BackColor = System.Drawing.SystemColors.Control
		Me.ProjectCheck.CausesValidation = True
		Me.ProjectCheck.Enabled = True
		Me.ProjectCheck.ForeColor = System.Drawing.SystemColors.ControlText
		Me.ProjectCheck.Cursor = System.Windows.Forms.Cursors.Default
		Me.ProjectCheck.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.ProjectCheck.Appearance = System.Windows.Forms.Appearance.Normal
		Me.ProjectCheck.TabStop = True
		Me.ProjectCheck.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.ProjectCheck.Visible = True
		Me.ProjectCheck.Name = "ProjectCheck"
		Me.Controls.Add(Text1)
		Me.Controls.Add(frameConvertTo)
		Me.frameConvertTo.Controls.Add(_cmdConvert_1)
		Me.frameConvertTo.Controls.Add(_optTargetFormat_0)
		Me.frameConvertTo.Controls.Add(_optTargetFormat_4)
		Me.frameConvertTo.Controls.Add(chkID)
		Me.frameConvertTo.Controls.Add(UpNextCheck)
		Me.frameConvertTo.Controls.Add(TimestampCheck)
		Me.frameConvertTo.Controls.Add(_cmdConvert_0)
		Me.frameConvertTo.Controls.Add(ContentsCheck)
		Me.frameConvertTo.Controls.Add(_optTargetFormat_3)
		Me.frameConvertTo.Controls.Add(_optTargetFormat_2)
		Me.frameConvertTo.Controls.Add(_optTargetFormat_1)
		Me.frameConvertTo.Controls.Add(ProjectCheck)
		Me.cmdConvert.SetIndex(_cmdConvert_1, CType(1, Short))
		Me.cmdConvert.SetIndex(_cmdConvert_0, CType(0, Short))
		Me.optTargetFormat.SetIndex(_optTargetFormat_0, CType(0, Short))
		Me.optTargetFormat.SetIndex(_optTargetFormat_4, CType(4, Short))
		Me.optTargetFormat.SetIndex(_optTargetFormat_3, CType(3, Short))
		Me.optTargetFormat.SetIndex(_optTargetFormat_2, CType(2, Short))
		Me.optTargetFormat.SetIndex(_optTargetFormat_1, CType(1, Short))
		CType(Me.optTargetFormat, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.cmdConvert, System.ComponentModel.ISupportInitialize).EndInit()
		Me.frameConvertTo.ResumeLayout(False)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class