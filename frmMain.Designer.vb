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
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMain))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.TimerSlowAction = New System.Windows.Forms.Timer(Me.components)
        Me.fraFind = New System.Windows.Forms.Panel
        Me.txtFind = New System.Windows.Forms.TextBox
        Me.txtReplace = New System.Windows.Forms.TextBox
        Me.cmdFind = New System.Windows.Forms.Button
        Me.cmdReplace = New System.Windows.Forms.Button
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.sash = New System.Windows.Forms.Panel
        Me.txtMain = New System.Windows.Forms.RichTextBox
        Me.cdlgOpen = New System.Windows.Forms.OpenFileDialog
        Me.cdlgSave = New System.Windows.Forms.SaveFileDialog
        Me.tree1 = New AxComctlLib.AxTreeView
        Me.cdlgImageOpen = New System.Windows.Forms.OpenFileDialog
        Me.mnuContext = New Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray(Me.components)
        Me._mnuContext_0 = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuRecent = New Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray(Me.components)
        Me.mnuTop = New Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray(Me.components)
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
        Me.MainMenu1 = New System.Windows.Forms.MenuStrip
        Me.mnuContextTop = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuTopHelp = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuHelpContents = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuHelpAbout = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuHelpWebsite = New System.Windows.Forms.ToolStripMenuItem
        Me.fraFind.SuspendLayout()
        CType(Me.tree1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.mnuContext, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.mnuRecent, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.mnuTop, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.MainMenu1.SuspendLayout()
        Me.SuspendLayout()
        '
        'TimerSlowAction
        '
        Me.TimerSlowAction.Interval = 1000
        '
        'fraFind
        '
        Me.fraFind.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.fraFind.BackColor = System.Drawing.SystemColors.Control
        Me.fraFind.Controls.Add(Me.txtFind)
        Me.fraFind.Controls.Add(Me.txtReplace)
        Me.fraFind.Controls.Add(Me.cmdFind)
        Me.fraFind.Controls.Add(Me.cmdReplace)
        Me.fraFind.Cursor = System.Windows.Forms.Cursors.Default
        Me.fraFind.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.fraFind.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraFind.Location = New System.Drawing.Point(348, 0)
        Me.fraFind.Name = "fraFind"
        Me.fraFind.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraFind.Size = New System.Drawing.Size(391, 27)
        Me.fraFind.TabIndex = 3
        Me.fraFind.Text = "Frame1"
        '
        'txtFind
        '
        Me.txtFind.AcceptsReturn = True
        Me.txtFind.BackColor = System.Drawing.SystemColors.Window
        Me.txtFind.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtFind.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtFind.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtFind.Location = New System.Drawing.Point(67, 3)
        Me.txtFind.MaxLength = 0
        Me.txtFind.Name = "txtFind"
        Me.txtFind.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtFind.Size = New System.Drawing.Size(111, 20)
        Me.txtFind.TabIndex = 7
        '
        'txtReplace
        '
        Me.txtReplace.AcceptsReturn = True
        Me.txtReplace.BackColor = System.Drawing.SystemColors.Window
        Me.txtReplace.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtReplace.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtReplace.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtReplace.Location = New System.Drawing.Point(271, 3)
        Me.txtReplace.MaxLength = 0
        Me.txtReplace.Name = "txtReplace"
        Me.txtReplace.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtReplace.Size = New System.Drawing.Size(111, 20)
        Me.txtReplace.TabIndex = 6
        '
        'cmdFind
        '
        Me.cmdFind.BackColor = System.Drawing.SystemColors.Control
        Me.cmdFind.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdFind.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdFind.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdFind.Location = New System.Drawing.Point(0, 3)
        Me.cmdFind.Name = "cmdFind"
        Me.cmdFind.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdFind.Size = New System.Drawing.Size(61, 21)
        Me.cmdFind.TabIndex = 5
        Me.cmdFind.Text = "Find:"
        Me.cmdFind.UseVisualStyleBackColor = False
        '
        'cmdReplace
        '
        Me.cmdReplace.BackColor = System.Drawing.SystemColors.Control
        Me.cmdReplace.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdReplace.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdReplace.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdReplace.Location = New System.Drawing.Point(184, 3)
        Me.cmdReplace.Name = "cmdReplace"
        Me.cmdReplace.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdReplace.Size = New System.Drawing.Size(81, 21)
        Me.cmdReplace.TabIndex = 4
        Me.cmdReplace.Text = "Replace:"
        Me.cmdReplace.UseVisualStyleBackColor = False
        '
        'Timer1
        '
        '
        'sash
        '
        Me.sash.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.sash.BackColor = System.Drawing.SystemColors.Control
        Me.sash.Cursor = System.Windows.Forms.Cursors.SizeWE
        Me.sash.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.sash.ForeColor = System.Drawing.SystemColors.ControlText
        Me.sash.Location = New System.Drawing.Point(150, 34)
        Me.sash.Name = "sash"
        Me.sash.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.sash.Size = New System.Drawing.Size(10, 440)
        Me.sash.TabIndex = 0
        Me.sash.Text = "Frame1"
        '
        'txtMain
        '
        Me.txtMain.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtMain.Font = New System.Drawing.Font("Courier New", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtMain.HideSelection = False
        Me.txtMain.Location = New System.Drawing.Point(160, 30)
        Me.txtMain.Name = "txtMain"
        Me.txtMain.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.Vertical
        Me.txtMain.Size = New System.Drawing.Size(579, 445)
        Me.txtMain.TabIndex = 2
        Me.txtMain.Text = "txtMain"
        '
        'tree1
        '
        Me.tree1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.tree1.Location = New System.Drawing.Point(0, 30)
        Me.tree1.Name = "tree1"
        Me.tree1.OcxState = CType(resources.GetObject("tree1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.tree1.Size = New System.Drawing.Size(152, 444)
        Me.tree1.TabIndex = 1
        '
        'mnuContext
        '
        '
        '_mnuContext_0
        '
        Me.mnuContext.SetIndex(Me._mnuContext_0, CType(0, Short))
        Me._mnuContext_0.Name = "_mnuContext_0"
        Me._mnuContext_0.Size = New System.Drawing.Size(138, 24)
        Me._mnuContext_0.Text = "Delete"
        '
        'mnuRecent
        '
        '
        '_mnuTop_0
        '
        Me._mnuTop_0.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuOpenProject, Me.mnuSaveProject, Me.mnuNewProject, Me.sep1, Me.mnuNewSection, Me.mnuSaveFile, Me.mnuRevert, Me.mnuAutoSave, Me.sep2, Me.mnuConvert, Me._mnuRecent_0, Me.sep3, Me.mnuExit})
        Me.mnuTop.SetIndex(Me._mnuTop_0, CType(0, Short))
        Me._mnuTop_0.Name = "_mnuTop_0"
        Me._mnuTop_0.Size = New System.Drawing.Size(45, 23)
        Me._mnuTop_0.Text = "&File"
        '
        'mnuOpenProject
        '
        Me.mnuOpenProject.Name = "mnuOpenProject"
        Me.mnuOpenProject.Size = New System.Drawing.Size(205, 24)
        Me.mnuOpenProject.Text = "&Open Project"
        '
        'mnuSaveProject
        '
        Me.mnuSaveProject.Name = "mnuSaveProject"
        Me.mnuSaveProject.Size = New System.Drawing.Size(205, 24)
        Me.mnuSaveProject.Text = "Save Project As"
        '
        'mnuNewProject
        '
        Me.mnuNewProject.Name = "mnuNewProject"
        Me.mnuNewProject.Size = New System.Drawing.Size(205, 24)
        Me.mnuNewProject.Text = "New Project"
        '
        'sep1
        '
        Me.sep1.Name = "sep1"
        Me.sep1.Size = New System.Drawing.Size(202, 6)
        '
        'mnuNewSection
        '
        Me.mnuNewSection.Name = "mnuNewSection"
        Me.mnuNewSection.Size = New System.Drawing.Size(205, 24)
        Me.mnuNewSection.Text = "&New Section"
        '
        'mnuSaveFile
        '
        Me.mnuSaveFile.Enabled = False
        Me.mnuSaveFile.Name = "mnuSaveFile"
        Me.mnuSaveFile.Size = New System.Drawing.Size(205, 24)
        Me.mnuSaveFile.Text = "&Save Section"
        '
        'mnuRevert
        '
        Me.mnuRevert.Name = "mnuRevert"
        Me.mnuRevert.Size = New System.Drawing.Size(205, 24)
        Me.mnuRevert.Text = "&Revert to Saved"
        '
        'mnuAutoSave
        '
        Me.mnuAutoSave.Name = "mnuAutoSave"
        Me.mnuAutoSave.Size = New System.Drawing.Size(205, 24)
        Me.mnuAutoSave.Text = "&Auto-Save"
        '
        'sep2
        '
        Me.sep2.Name = "sep2"
        Me.sep2.Size = New System.Drawing.Size(202, 6)
        '
        'mnuConvert
        '
        Me.mnuConvert.Name = "mnuConvert"
        Me.mnuConvert.Size = New System.Drawing.Size(205, 24)
        Me.mnuConvert.Text = "&Convert"
        '
        '_mnuRecent_0
        '
        Me._mnuRecent_0.Name = "_mnuRecent_0"
        Me._mnuRecent_0.Size = New System.Drawing.Size(202, 6)
        Me._mnuRecent_0.Visible = False
        '
        'sep3
        '
        Me.sep3.Name = "sep3"
        Me.sep3.Size = New System.Drawing.Size(202, 6)
        '
        'mnuExit
        '
        Me.mnuExit.Name = "mnuExit"
        Me.mnuExit.Size = New System.Drawing.Size(205, 24)
        Me.mnuExit.Text = "E&xit"
        '
        '_mnuTop_1
        '
        Me._mnuTop_1.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuCut, Me.mnuCopy, Me.mnuPaste, Me.sep5, Me.mnuFindSelection, Me.mnuFind})
        Me.mnuTop.SetIndex(Me._mnuTop_1, CType(1, Short))
        Me._mnuTop_1.Name = "_mnuTop_1"
        Me._mnuTop_1.Size = New System.Drawing.Size(48, 23)
        Me._mnuTop_1.Text = "&Edit"
        '
        'mnuCut
        '
        Me.mnuCut.Name = "mnuCut"
        Me.mnuCut.Size = New System.Drawing.Size(246, 24)
        Me.mnuCut.Text = "Cut"
        '
        'mnuCopy
        '
        Me.mnuCopy.Name = "mnuCopy"
        Me.mnuCopy.Size = New System.Drawing.Size(246, 24)
        Me.mnuCopy.Text = "Copy"
        '
        'mnuPaste
        '
        Me.mnuPaste.Name = "mnuPaste"
        Me.mnuPaste.Size = New System.Drawing.Size(246, 24)
        Me.mnuPaste.Text = "Paste"
        '
        'sep5
        '
        Me.sep5.Name = "sep5"
        Me.sep5.Size = New System.Drawing.Size(243, 6)
        '
        'mnuFindSelection
        '
        Me.mnuFindSelection.Name = "mnuFindSelection"
        Me.mnuFindSelection.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.F), System.Windows.Forms.Keys)
        Me.mnuFindSelection.Size = New System.Drawing.Size(246, 24)
        Me.mnuFindSelection.Text = "Find Selection"
        '
        'mnuFind
        '
        Me.mnuFind.Name = "mnuFind"
        Me.mnuFind.ShortcutKeys = System.Windows.Forms.Keys.F3
        Me.mnuFind.Size = New System.Drawing.Size(246, 24)
        Me.mnuFind.Text = "Find"
        '
        '_mnuTop_2
        '
        Me._mnuTop_2.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuUnderline, Me.mnuBold, Me.mnuItalic, Me.mnuLink, Me.mnuLinkSection, Me.mnuImage, Me.mnuIndexword, Me.mnuKeyword, Me.mnuOL, Me.mnuUL, Me.mnuPRE, Me.mnuFigure, Me.sep4, Me.mnuAutoParagraph})
        Me.mnuTop.SetIndex(Me._mnuTop_2, CType(2, Short))
        Me._mnuTop_2.Name = "_mnuTop_2"
        Me._mnuTop_2.Size = New System.Drawing.Size(55, 23)
        Me._mnuTop_2.Text = "&Tags"
        '
        'mnuUnderline
        '
        Me.mnuUnderline.Name = "mnuUnderline"
        Me.mnuUnderline.Size = New System.Drawing.Size(361, 24)
        Me.mnuUnderline.Text = "&Underline <u>...</u>"
        '
        'mnuBold
        '
        Me.mnuBold.Name = "mnuBold"
        Me.mnuBold.Size = New System.Drawing.Size(361, 24)
        Me.mnuBold.Text = "&Bold <b>...</b>"
        '
        'mnuItalic
        '
        Me.mnuItalic.Name = "mnuItalic"
        Me.mnuItalic.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.I), System.Windows.Forms.Keys)
        Me.mnuItalic.Size = New System.Drawing.Size(361, 24)
        Me.mnuItalic.Text = "&Italic <i>...</i>"
        '
        'mnuLink
        '
        Me.mnuLink.Name = "mnuLink"
        Me.mnuLink.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.L), System.Windows.Forms.Keys)
        Me.mnuLink.Size = New System.Drawing.Size(361, 24)
        Me.mnuLink.Text = "&Link <a href=""..."">...</a>"
        '
        'mnuLinkSection
        '
        Me.mnuLinkSection.Name = "mnuLinkSection"
        Me.mnuLinkSection.Size = New System.Drawing.Size(361, 24)
        Me.mnuLinkSection.Text = "Link &Section"
        '
        'mnuImage
        '
        Me.mnuImage.Name = "mnuImage"
        Me.mnuImage.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.M), System.Windows.Forms.Keys)
        Me.mnuImage.Size = New System.Drawing.Size(361, 24)
        Me.mnuImage.Text = "I&mage <img src=""..."">"
        '
        'mnuIndexword
        '
        Me.mnuIndexword.Name = "mnuIndexword"
        Me.mnuIndexword.Size = New System.Drawing.Size(361, 24)
        Me.mnuIndexword.Text = "Inde&x word <indexword=...>"
        '
        'mnuKeyword
        '
        Me.mnuKeyword.Name = "mnuKeyword"
        Me.mnuKeyword.Size = New System.Drawing.Size(361, 24)
        Me.mnuKeyword.Text = "&Keyword <keyword=...>"
        '
        'mnuOL
        '
        Me.mnuOL.Name = "mnuOL"
        Me.mnuOL.Size = New System.Drawing.Size(361, 24)
        Me.mnuOL.Text = "&Numbered List <ol><li>...</ol>"
        '
        'mnuUL
        '
        Me.mnuUL.Name = "mnuUL"
        Me.mnuUL.Size = New System.Drawing.Size(361, 24)
        Me.mnuUL.Text = "Bulle&ts <ul><li>...</ul>"
        '
        'mnuPRE
        '
        Me.mnuPRE.Name = "mnuPRE"
        Me.mnuPRE.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.P), System.Windows.Forms.Keys)
        Me.mnuPRE.Size = New System.Drawing.Size(361, 24)
        Me.mnuPRE.Text = "&Preformatted <pre>...</pre>"
        '
        'mnuFigure
        '
        Me.mnuFigure.Name = "mnuFigure"
        Me.mnuFigure.Size = New System.Drawing.Size(361, 24)
        Me.mnuFigure.Text = "&Figure <figure>...</figure>"
        '
        'sep4
        '
        Me.sep4.Name = "sep4"
        Me.sep4.Size = New System.Drawing.Size(358, 6)
        '
        'mnuAutoParagraph
        '
        Me.mnuAutoParagraph.Checked = True
        Me.mnuAutoParagraph.CheckState = System.Windows.Forms.CheckState.Checked
        Me.mnuAutoParagraph.Name = "mnuAutoParagraph"
        Me.mnuAutoParagraph.Size = New System.Drawing.Size(361, 24)
        Me.mnuAutoParagraph.Text = "Automatic Paragraphs <p>"
        '
        '_mnuTop_3
        '
        Me._mnuTop_3.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuFormatting, Me.mnuFormatWhileTyping, Me.mnuOptions, Me.mnuTextImage})
        Me.mnuTop.SetIndex(Me._mnuTop_3, CType(3, Short))
        Me._mnuTop_3.Name = "_mnuTop_3"
        Me._mnuTop_3.Size = New System.Drawing.Size(55, 23)
        Me._mnuTop_3.Text = "&View"
        '
        'mnuFormatting
        '
        Me.mnuFormatting.Checked = True
        Me.mnuFormatting.CheckState = System.Windows.Forms.CheckState.Checked
        Me.mnuFormatting.Name = "mnuFormatting"
        Me.mnuFormatting.Size = New System.Drawing.Size(242, 24)
        Me.mnuFormatting.Text = "&Formatting"
        '
        'mnuFormatWhileTyping
        '
        Me.mnuFormatWhileTyping.Name = "mnuFormatWhileTyping"
        Me.mnuFormatWhileTyping.Size = New System.Drawing.Size(242, 24)
        Me.mnuFormatWhileTyping.Text = "Format While Typing"
        '
        'mnuOptions
        '
        Me.mnuOptions.Name = "mnuOptions"
        Me.mnuOptions.Size = New System.Drawing.Size(242, 24)
        Me.mnuOptions.Text = "&Options"
        '
        'mnuTextImage
        '
        Me.mnuTextImage.Name = "mnuTextImage"
        Me.mnuTextImage.Size = New System.Drawing.Size(242, 24)
        Me.mnuTextImage.Text = "Test TextImage"
        '
        'MainMenu1
        '
        Me.MainMenu1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me._mnuTop_0, Me._mnuTop_1, Me._mnuTop_2, Me._mnuTop_3, Me.mnuContextTop, Me.mnuTopHelp})
        Me.MainMenu1.Location = New System.Drawing.Point(0, 0)
        Me.MainMenu1.Name = "MainMenu1"
        Me.MainMenu1.Size = New System.Drawing.Size(739, 27)
        Me.MainMenu1.TabIndex = 4
        '
        'mnuContextTop
        '
        Me.mnuContextTop.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me._mnuContext_0})
        Me.mnuContextTop.Name = "mnuContextTop"
        Me.mnuContextTop.Size = New System.Drawing.Size(75, 23)
        Me.mnuContextTop.Text = "Context"
        '
        'mnuTopHelp
        '
        Me.mnuTopHelp.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuHelpContents, Me.mnuHelpAbout, Me.mnuHelpWebsite})
        Me.mnuTopHelp.Name = "mnuTopHelp"
        Me.mnuTopHelp.Size = New System.Drawing.Size(53, 23)
        Me.mnuTopHelp.Text = "&Help"
        '
        'mnuHelpContents
        '
        Me.mnuHelpContents.Name = "mnuHelpContents"
        Me.mnuHelpContents.Size = New System.Drawing.Size(156, 24)
        Me.mnuHelpContents.Text = "&Contents"
        '
        'mnuHelpAbout
        '
        Me.mnuHelpAbout.Name = "mnuHelpAbout"
        Me.mnuHelpAbout.Size = New System.Drawing.Size(156, 24)
        Me.mnuHelpAbout.Text = "&About"
        '
        'mnuHelpWebsite
        '
        Me.mnuHelpWebsite.Name = "mnuHelpWebsite"
        Me.mnuHelpWebsite.Size = New System.Drawing.Size(156, 24)
        Me.mnuHelpWebsite.Text = "&Web Site"
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(739, 475)
        Me.Controls.Add(Me.fraFind)
        Me.Controls.Add(Me.sash)
        Me.Controls.Add(Me.txtMain)
        Me.Controls.Add(Me.tree1)
        Me.Controls.Add(Me.MainMenu1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = New System.Drawing.Point(14, 62)
        Me.Name = "frmMain"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text = "AuthorDoc"
        Me.fraFind.ResumeLayout(False)
        Me.fraFind.PerformLayout()
        CType(Me.tree1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.mnuContext, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.mnuRecent, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.mnuTop, System.ComponentModel.ISupportInitialize).EndInit()
        Me.MainMenu1.ResumeLayout(False)
        Me.MainMenu1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
#End Region 
End Class