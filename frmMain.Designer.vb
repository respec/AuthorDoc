<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmMain
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
	Public WithEvents TimerSlowAction As System.Windows.Forms.Timer
	Public WithEvents txtFind As System.Windows.Forms.TextBox
	Public WithEvents txtReplace As System.Windows.Forms.TextBox
	Public WithEvents cmdFind As System.Windows.Forms.Button
	Public WithEvents cmdReplace As System.Windows.Forms.Button
	Public WithEvents fraFind As System.Windows.Forms.Panel
	Public WithEvents Timer1 As System.Windows.Forms.Timer
	Public WithEvents sash As System.Windows.Forms.Panel
	Public WithEvents txtMain As System.Windows.Forms.RichTextBox
	Public cdlgOpen As System.Windows.Forms.OpenFileDialog
	Public cdlgSave As System.Windows.Forms.SaveFileDialog
	Public WithEvents tree1 As AxComctlLib.AxTreeView
	Public cdlgImageOpen As System.Windows.Forms.OpenFileDialog
	Public WithEvents mnuContext As Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray
	Public WithEvents mnuRecent As Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray
	Public WithEvents mnuTop As Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray
	Public WithEvents mnuOpenProject As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuSaveProject As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuNewProject As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents sep1 As System.Windows.Forms.ToolStripSeparator
	Public WithEvents mnuNewSection As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuSaveFile As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuRevert As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuAutoSave As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents sep2 As System.Windows.Forms.ToolStripSeparator
	Public WithEvents mnuConvert As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents _mnuRecent_0 As System.Windows.Forms.ToolStripSeparator
	Public WithEvents sep3 As System.Windows.Forms.ToolStripSeparator
	Public WithEvents mnuExit As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents _mnuTop_0 As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuCut As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuCopy As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuPaste As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents sep5 As System.Windows.Forms.ToolStripSeparator
	Public WithEvents mnuFindSelection As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuFind As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents _mnuTop_1 As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuUnderline As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuBold As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuItalic As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuLink As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuLinkSection As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuImage As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuIndexword As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuKeyword As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuOL As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuUL As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuPRE As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuFigure As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents sep4 As System.Windows.Forms.ToolStripSeparator
	Public WithEvents mnuAutoParagraph As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents _mnuTop_2 As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuFormatting As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuFormatWhileTyping As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuOptions As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuTextImage As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents _mnuTop_3 As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents _mnuContext_0 As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuContextTop As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuHelpContents As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuHelpAbout As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuHelpWebsite As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuTopHelp As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents MainMenu1 As System.Windows.Forms.MenuStrip
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmMain))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.TimerSlowAction = New System.Windows.Forms.Timer(components)
		Me.fraFind = New System.Windows.Forms.Panel
		Me.txtFind = New System.Windows.Forms.TextBox
		Me.txtReplace = New System.Windows.Forms.TextBox
		Me.cmdFind = New System.Windows.Forms.Button
		Me.cmdReplace = New System.Windows.Forms.Button
		Me.Timer1 = New System.Windows.Forms.Timer(components)
		Me.sash = New System.Windows.Forms.Panel
		Me.txtMain = New System.Windows.Forms.RichTextBox
		Me.cdlgOpen = New System.Windows.Forms.OpenFileDialog
		Me.cdlgSave = New System.Windows.Forms.SaveFileDialog
		Me.tree1 = New AxComctlLib.AxTreeView
		Me.cdlgImageOpen = New System.Windows.Forms.OpenFileDialog
		Me.mnuContext = New Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray(components)
		Me.mnuRecent = New Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray(components)
		Me.mnuTop = New Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray(components)
		Me.MainMenu1 = New System.Windows.Forms.MenuStrip
		Me._mnuTop_0 = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuOpenProject = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuSaveProject = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuNewProject = New System.Windows.Forms.ToolStripMenuItem
		Me.sep1 = New System.Windows.Forms.ToolStripSeparator
		Me.mnuNewSection = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuSaveFile = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuRevert = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuAutoSave = New System.Windows.Forms.ToolStripMenuItem
		Me.sep2 = New System.Windows.Forms.ToolStripSeparator
		Me.mnuConvert = New System.Windows.Forms.ToolStripMenuItem
		Me._mnuRecent_0 = New System.Windows.Forms.ToolStripSeparator
		Me.sep3 = New System.Windows.Forms.ToolStripSeparator
		Me.mnuExit = New System.Windows.Forms.ToolStripMenuItem
		Me._mnuTop_1 = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuCut = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuCopy = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuPaste = New System.Windows.Forms.ToolStripMenuItem
		Me.sep5 = New System.Windows.Forms.ToolStripSeparator
		Me.mnuFindSelection = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuFind = New System.Windows.Forms.ToolStripMenuItem
		Me._mnuTop_2 = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuUnderline = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuBold = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuItalic = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuLink = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuLinkSection = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuImage = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuIndexword = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuKeyword = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuOL = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuUL = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuPRE = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuFigure = New System.Windows.Forms.ToolStripMenuItem
		Me.sep4 = New System.Windows.Forms.ToolStripSeparator
		Me.mnuAutoParagraph = New System.Windows.Forms.ToolStripMenuItem
		Me._mnuTop_3 = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuFormatting = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuFormatWhileTyping = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuOptions = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuTextImage = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuContextTop = New System.Windows.Forms.ToolStripMenuItem
		Me._mnuContext_0 = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuTopHelp = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuHelpContents = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuHelpAbout = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuHelpWebsite = New System.Windows.Forms.ToolStripMenuItem
		Me.fraFind.SuspendLayout()
		Me.MainMenu1.SuspendLayout()
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		CType(Me.tree1, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.mnuContext, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.mnuRecent, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.mnuTop, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.Text = "AuthorDoc"
		Me.ClientSize = New System.Drawing.Size(739, 475)
		Me.Location = New System.Drawing.Point(14, 62)
		Me.Icon = CType(resources.GetObject("frmMain.Icon"), System.Drawing.Icon)
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
		Me.Name = "frmMain"
		Me.TimerSlowAction.Enabled = False
		Me.TimerSlowAction.Interval = 1000
		Me.fraFind.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.fraFind.Text = "Frame1"
		Me.fraFind.Size = New System.Drawing.Size(391, 41)
		Me.fraFind.Location = New System.Drawing.Point(160, 0)
		Me.fraFind.TabIndex = 3
		Me.fraFind.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.fraFind.BackColor = System.Drawing.SystemColors.Control
		Me.fraFind.Enabled = True
		Me.fraFind.ForeColor = System.Drawing.SystemColors.ControlText
		Me.fraFind.Cursor = System.Windows.Forms.Cursors.Default
		Me.fraFind.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.fraFind.Visible = True
		Me.fraFind.Name = "fraFind"
		Me.txtFind.AutoSize = False
		Me.txtFind.Size = New System.Drawing.Size(111, 24)
		Me.txtFind.Location = New System.Drawing.Point(70, 10)
		Me.txtFind.TabIndex = 7
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
		Me.txtReplace.AutoSize = False
		Me.txtReplace.Size = New System.Drawing.Size(111, 24)
		Me.txtReplace.Location = New System.Drawing.Point(280, 10)
		Me.txtReplace.TabIndex = 6
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
		Me.cmdFind.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdFind.Text = "Find:"
		Me.cmdFind.Size = New System.Drawing.Size(61, 21)
		Me.cmdFind.Location = New System.Drawing.Point(0, 10)
		Me.cmdFind.TabIndex = 5
		Me.cmdFind.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdFind.BackColor = System.Drawing.SystemColors.Control
		Me.cmdFind.CausesValidation = True
		Me.cmdFind.Enabled = True
		Me.cmdFind.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdFind.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdFind.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdFind.TabStop = True
		Me.cmdFind.Name = "cmdFind"
		Me.cmdReplace.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdReplace.Text = "Replace:"
		Me.cmdReplace.Size = New System.Drawing.Size(81, 21)
		Me.cmdReplace.Location = New System.Drawing.Point(190, 10)
		Me.cmdReplace.TabIndex = 4
		Me.cmdReplace.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdReplace.BackColor = System.Drawing.SystemColors.Control
		Me.cmdReplace.CausesValidation = True
		Me.cmdReplace.Enabled = True
		Me.cmdReplace.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdReplace.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdReplace.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdReplace.TabStop = True
		Me.cmdReplace.Name = "cmdReplace"
		Me.Timer1.Enabled = False
		Me.Timer1.Interval = 100
		Me.sash.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.sash.Text = "Frame1"
		Me.sash.Size = New System.Drawing.Size(6, 462)
		Me.sash.Location = New System.Drawing.Point(150, 0)
		Me.sash.Cursor = System.Windows.Forms.Cursors.SizeWE
		Me.sash.TabIndex = 0
		Me.sash.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.sash.BackColor = System.Drawing.SystemColors.Control
		Me.sash.Enabled = True
		Me.sash.ForeColor = System.Drawing.SystemColors.ControlText
		Me.sash.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.sash.Visible = True
		Me.sash.Name = "sash"
		Me.txtMain.Size = New System.Drawing.Size(391, 421)
		Me.txtMain.Location = New System.Drawing.Point(160, 40)
		Me.txtMain.TabIndex = 2
		Me.txtMain.Enabled = True
		Me.txtMain.HideSelection = False
		Me.txtMain.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.Vertical
		Me.txtMain.RTF = resources.GetString("txtMain.TextRTF")
		Me.txtMain.Font = New System.Drawing.Font("Courier New", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtMain.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtMain.Name = "txtMain"
		tree1.OcxState = CType(resources.GetObject("tree1.OcxState"), System.Windows.Forms.AxHost.State)
		Me.tree1.Size = New System.Drawing.Size(152, 462)
		Me.tree1.Location = New System.Drawing.Point(0, 0)
		Me.tree1.TabIndex = 1
		Me.tree1.Name = "tree1"
		Me._mnuTop_0.Name = "_mnuTop_0"
		Me._mnuTop_0.Text = "&File"
		Me._mnuTop_0.Checked = False
		Me._mnuTop_0.Enabled = True
		Me._mnuTop_0.Visible = True
		Me.mnuOpenProject.Name = "mnuOpenProject"
		Me.mnuOpenProject.Text = "&Open Project"
		Me.mnuOpenProject.Checked = False
		Me.mnuOpenProject.Enabled = True
		Me.mnuOpenProject.Visible = True
		Me.mnuSaveProject.Name = "mnuSaveProject"
		Me.mnuSaveProject.Text = "Save Project As"
		Me.mnuSaveProject.Checked = False
		Me.mnuSaveProject.Enabled = True
		Me.mnuSaveProject.Visible = True
		Me.mnuNewProject.Name = "mnuNewProject"
		Me.mnuNewProject.Text = "New Project"
		Me.mnuNewProject.Checked = False
		Me.mnuNewProject.Enabled = True
		Me.mnuNewProject.Visible = True
		Me.sep1.Enabled = True
		Me.sep1.Visible = True
		Me.sep1.Name = "sep1"
		Me.mnuNewSection.Name = "mnuNewSection"
		Me.mnuNewSection.Text = "&New Section"
		Me.mnuNewSection.Checked = False
		Me.mnuNewSection.Enabled = True
		Me.mnuNewSection.Visible = True
		Me.mnuSaveFile.Name = "mnuSaveFile"
		Me.mnuSaveFile.Text = "&Save Section"
		Me.mnuSaveFile.Enabled = False
		Me.mnuSaveFile.Checked = False
		Me.mnuSaveFile.Visible = True
		Me.mnuRevert.Name = "mnuRevert"
		Me.mnuRevert.Text = "&Revert to Saved"
		Me.mnuRevert.Checked = False
		Me.mnuRevert.Enabled = True
		Me.mnuRevert.Visible = True
		Me.mnuAutoSave.Name = "mnuAutoSave"
		Me.mnuAutoSave.Text = "&Auto-Save"
		Me.mnuAutoSave.Checked = False
		Me.mnuAutoSave.Enabled = True
		Me.mnuAutoSave.Visible = True
		Me.sep2.Enabled = True
		Me.sep2.Visible = True
		Me.sep2.Name = "sep2"
		Me.mnuConvert.Name = "mnuConvert"
		Me.mnuConvert.Text = "&Convert"
		Me.mnuConvert.Checked = False
		Me.mnuConvert.Enabled = True
		Me.mnuConvert.Visible = True
		Me._mnuRecent_0.Visible = False
		Me._mnuRecent_0.Enabled = True
		Me._mnuRecent_0.Name = "_mnuRecent_0"
		Me.sep3.Enabled = True
		Me.sep3.Visible = True
		Me.sep3.Name = "sep3"
		Me.mnuExit.Name = "mnuExit"
		Me.mnuExit.Text = "E&xit"
		Me.mnuExit.Checked = False
		Me.mnuExit.Enabled = True
		Me.mnuExit.Visible = True
		Me._mnuTop_1.Name = "_mnuTop_1"
		Me._mnuTop_1.Text = "&Edit"
		Me._mnuTop_1.Checked = False
		Me._mnuTop_1.Enabled = True
		Me._mnuTop_1.Visible = True
		Me.mnuCut.Name = "mnuCut"
		Me.mnuCut.Text = "Cut"
		Me.mnuCut.Checked = False
		Me.mnuCut.Enabled = True
		Me.mnuCut.Visible = True
		Me.mnuCopy.Name = "mnuCopy"
		Me.mnuCopy.Text = "Copy"
		Me.mnuCopy.Checked = False
		Me.mnuCopy.Enabled = True
		Me.mnuCopy.Visible = True
		Me.mnuPaste.Name = "mnuPaste"
		Me.mnuPaste.Text = "Paste"
		Me.mnuPaste.Checked = False
		Me.mnuPaste.Enabled = True
		Me.mnuPaste.Visible = True
		Me.sep5.Enabled = True
		Me.sep5.Visible = True
		Me.sep5.Name = "sep5"
		Me.mnuFindSelection.Name = "mnuFindSelection"
		Me.mnuFindSelection.Text = "Find Selection"
		Me.mnuFindSelection.ShortcutKeys = CType(System.Windows.Forms.Keys.Control or System.Windows.Forms.Keys.F, System.Windows.Forms.Keys)
		Me.mnuFindSelection.Checked = False
		Me.mnuFindSelection.Enabled = True
		Me.mnuFindSelection.Visible = True
		Me.mnuFind.Name = "mnuFind"
		Me.mnuFind.Text = "Find"
		Me.mnuFind.ShortcutKeys = CType(System.Windows.Forms.Keys.F3, System.Windows.Forms.Keys)
		Me.mnuFind.Checked = False
		Me.mnuFind.Enabled = True
		Me.mnuFind.Visible = True
		Me._mnuTop_2.Name = "_mnuTop_2"
		Me._mnuTop_2.Text = "&Tags"
		Me._mnuTop_2.Checked = False
		Me._mnuTop_2.Enabled = True
		Me._mnuTop_2.Visible = True
		Me.mnuUnderline.Name = "mnuUnderline"
		Me.mnuUnderline.Text = "&Underline <u>...</u>"
		Me.mnuUnderline.Checked = False
		Me.mnuUnderline.Enabled = True
		Me.mnuUnderline.Visible = True
		Me.mnuBold.Name = "mnuBold"
		Me.mnuBold.Text = "&Bold <b>...</b>"
		Me.mnuBold.Checked = False
		Me.mnuBold.Enabled = True
		Me.mnuBold.Visible = True
		Me.mnuItalic.Name = "mnuItalic"
		Me.mnuItalic.Text = "&Italic <i>...</i>"
		Me.mnuItalic.ShortcutKeys = CType(System.Windows.Forms.Keys.Control or System.Windows.Forms.Keys.I, System.Windows.Forms.Keys)
		Me.mnuItalic.Checked = False
		Me.mnuItalic.Enabled = True
		Me.mnuItalic.Visible = True
		Me.mnuLink.Name = "mnuLink"
		Me.mnuLink.Text = "&Link <a href=""..."">...</a>"
		Me.mnuLink.ShortcutKeys = CType(System.Windows.Forms.Keys.Control or System.Windows.Forms.Keys.L, System.Windows.Forms.Keys)
		Me.mnuLink.Checked = False
		Me.mnuLink.Enabled = True
		Me.mnuLink.Visible = True
		Me.mnuLinkSection.Name = "mnuLinkSection"
		Me.mnuLinkSection.Text = "Link &Section"
		Me.mnuLinkSection.Checked = False
		Me.mnuLinkSection.Enabled = True
		Me.mnuLinkSection.Visible = True
		Me.mnuImage.Name = "mnuImage"
		Me.mnuImage.Text = "I&mage <img src=""..."">"
		Me.mnuImage.ShortcutKeys = CType(System.Windows.Forms.Keys.Control or System.Windows.Forms.Keys.M, System.Windows.Forms.Keys)
		Me.mnuImage.Checked = False
		Me.mnuImage.Enabled = True
		Me.mnuImage.Visible = True
		Me.mnuIndexword.Name = "mnuIndexword"
		Me.mnuIndexword.Text = "Inde&x word <indexword=...>"
		Me.mnuIndexword.Checked = False
		Me.mnuIndexword.Enabled = True
		Me.mnuIndexword.Visible = True
		Me.mnuKeyword.Name = "mnuKeyword"
		Me.mnuKeyword.Text = "&Keyword <keyword=...>"
		Me.mnuKeyword.Checked = False
		Me.mnuKeyword.Enabled = True
		Me.mnuKeyword.Visible = True
		Me.mnuOL.Name = "mnuOL"
		Me.mnuOL.Text = "&Numbered List <ol><li>...</ol>"
		Me.mnuOL.Checked = False
		Me.mnuOL.Enabled = True
		Me.mnuOL.Visible = True
		Me.mnuUL.Name = "mnuUL"
		Me.mnuUL.Text = "Bulle&ts <ul><li>...</ul>"
		Me.mnuUL.Checked = False
		Me.mnuUL.Enabled = True
		Me.mnuUL.Visible = True
		Me.mnuPRE.Name = "mnuPRE"
		Me.mnuPRE.Text = "&Preformatted <pre>...</pre>"
		Me.mnuPRE.ShortcutKeys = CType(System.Windows.Forms.Keys.Control or System.Windows.Forms.Keys.P, System.Windows.Forms.Keys)
		Me.mnuPRE.Checked = False
		Me.mnuPRE.Enabled = True
		Me.mnuPRE.Visible = True
		Me.mnuFigure.Name = "mnuFigure"
		Me.mnuFigure.Text = "&Figure <figure>...</figure>"
		Me.mnuFigure.Checked = False
		Me.mnuFigure.Enabled = True
		Me.mnuFigure.Visible = True
		Me.sep4.Enabled = True
		Me.sep4.Visible = True
		Me.sep4.Name = "sep4"
		Me.mnuAutoParagraph.Name = "mnuAutoParagraph"
		Me.mnuAutoParagraph.Text = "Automatic Paragraphs <p>"
		Me.mnuAutoParagraph.Checked = True
		Me.mnuAutoParagraph.Enabled = True
		Me.mnuAutoParagraph.Visible = True
		Me._mnuTop_3.Name = "_mnuTop_3"
		Me._mnuTop_3.Text = "&View"
		Me._mnuTop_3.Checked = False
		Me._mnuTop_3.Enabled = True
		Me._mnuTop_3.Visible = True
		Me.mnuFormatting.Name = "mnuFormatting"
		Me.mnuFormatting.Text = "&Formatting"
		Me.mnuFormatting.Checked = True
		Me.mnuFormatting.Enabled = True
		Me.mnuFormatting.Visible = True
		Me.mnuFormatWhileTyping.Name = "mnuFormatWhileTyping"
		Me.mnuFormatWhileTyping.Text = "Format While Typing"
		Me.mnuFormatWhileTyping.Checked = False
		Me.mnuFormatWhileTyping.Enabled = True
		Me.mnuFormatWhileTyping.Visible = True
		Me.mnuOptions.Name = "mnuOptions"
		Me.mnuOptions.Text = "&Options"
		Me.mnuOptions.Checked = False
		Me.mnuOptions.Enabled = True
		Me.mnuOptions.Visible = True
		Me.mnuTextImage.Name = "mnuTextImage"
		Me.mnuTextImage.Text = "Test TextImage"
		Me.mnuTextImage.Checked = False
		Me.mnuTextImage.Enabled = True
		Me.mnuTextImage.Visible = True
		Me.mnuContextTop.Name = "mnuContextTop"
		Me.mnuContextTop.Text = "Context"
		Me.mnuContextTop.Checked = False
		Me.mnuContextTop.Enabled = True
		Me.mnuContextTop.Visible = True
		Me._mnuContext_0.Name = "_mnuContext_0"
		Me._mnuContext_0.Text = "Delete"
		Me._mnuContext_0.Checked = False
		Me._mnuContext_0.Enabled = True
		Me._mnuContext_0.Visible = True
		Me.mnuTopHelp.Name = "mnuTopHelp"
		Me.mnuTopHelp.Text = "&Help"
		Me.mnuTopHelp.Checked = False
		Me.mnuTopHelp.Enabled = True
		Me.mnuTopHelp.Visible = True
		Me.mnuHelpContents.Name = "mnuHelpContents"
		Me.mnuHelpContents.Text = "&Contents"
		Me.mnuHelpContents.Checked = False
		Me.mnuHelpContents.Enabled = True
		Me.mnuHelpContents.Visible = True
		Me.mnuHelpAbout.Name = "mnuHelpAbout"
		Me.mnuHelpAbout.Text = "&About"
		Me.mnuHelpAbout.Checked = False
		Me.mnuHelpAbout.Enabled = True
		Me.mnuHelpAbout.Visible = True
		Me.mnuHelpWebsite.Name = "mnuHelpWebsite"
		Me.mnuHelpWebsite.Text = "&Web Site"
		Me.mnuHelpWebsite.Checked = False
		Me.mnuHelpWebsite.Enabled = True
		Me.mnuHelpWebsite.Visible = True
		Me.Controls.Add(fraFind)
		Me.Controls.Add(sash)
		Me.Controls.Add(txtMain)
		Me.Controls.Add(tree1)
		Me.fraFind.Controls.Add(txtFind)
		Me.fraFind.Controls.Add(txtReplace)
		Me.fraFind.Controls.Add(cmdFind)
		Me.fraFind.Controls.Add(cmdReplace)
		Me.mnuContext.SetIndex(_mnuContext_0, CType(0, Short))
		Me.mnuRecent.SetIndex(_mnuRecent_0, CType(0, Short))
		Me.mnuTop.SetIndex(_mnuTop_0, CType(0, Short))
		Me.mnuTop.SetIndex(_mnuTop_1, CType(1, Short))
		Me.mnuTop.SetIndex(_mnuTop_2, CType(2, Short))
		Me.mnuTop.SetIndex(_mnuTop_3, CType(3, Short))
		CType(Me.mnuTop, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.mnuRecent, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.mnuContext, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.tree1, System.ComponentModel.ISupportInitialize).EndInit()
		MainMenu1.Items.AddRange(New System.Windows.Forms.ToolStripItem(){Me._mnuTop_0, Me._mnuTop_1, Me._mnuTop_2, Me._mnuTop_3, Me.mnuContextTop, Me.mnuTopHelp})
		_mnuTop_0.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem(){Me.mnuOpenProject, Me.mnuSaveProject, Me.mnuNewProject, Me.sep1, Me.mnuNewSection, Me.mnuSaveFile, Me.mnuRevert, Me.mnuAutoSave, Me.sep2, Me.mnuConvert, Me._mnuRecent_0, Me.sep3, Me.mnuExit})
		_mnuTop_1.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem(){Me.mnuCut, Me.mnuCopy, Me.mnuPaste, Me.sep5, Me.mnuFindSelection, Me.mnuFind})
		_mnuTop_2.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem(){Me.mnuUnderline, Me.mnuBold, Me.mnuItalic, Me.mnuLink, Me.mnuLinkSection, Me.mnuImage, Me.mnuIndexword, Me.mnuKeyword, Me.mnuOL, Me.mnuUL, Me.mnuPRE, Me.mnuFigure, Me.sep4, Me.mnuAutoParagraph})
		_mnuTop_3.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem(){Me.mnuFormatting, Me.mnuFormatWhileTyping, Me.mnuOptions, Me.mnuTextImage})
		mnuContextTop.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem(){Me._mnuContext_0})
		mnuTopHelp.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem(){Me.mnuHelpContents, Me.mnuHelpAbout, Me.mnuHelpWebsite})
		Me.Controls.Add(MainMenu1)
		Me.fraFind.ResumeLayout(False)
		Me.MainMenu1.ResumeLayout(False)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class