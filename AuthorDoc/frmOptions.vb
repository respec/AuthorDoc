Option Strict Off
Option Explicit On
Friend Class frmOptions
	Inherits System.Windows.Forms.Form
	'Copyright 2000 by AQUA TERRA Consultants
	
	Private Sub Command1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command1.Click
		If IsNumeric(txtTreeIndent.Text) Then frmMain.tree1.Indentation = CSng(txtTreeIndent.Text)
		If IsNumeric(txtFindTimeout.Text) Then FindTimeout = CSng(txtFindTimeout.Text)
		CopyFont2RichText(txtFont, (frmMain.txtMain))
		Me.Close()
	End Sub
	
	Private Sub frmOptions_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		txtTreeIndent.Text = CStr(frmMain.tree1.Indentation)
		txtFindTimeout.Text = CStr(FindTimeout)
		
		CopyFontFromRichText((frmMain.txtMain), txtFont)
		With txtFont
			.Text = .Font.Name & .Font.SizeInPoints
			If .Font.Bold Then .Text = .Text & "Bold"
			If .Font.Italic Then .Text = .Text & "Italic"
			If .Font.Underline Then .Text = .Text & "Underline"
		End With
		
	End Sub
	
	Private Sub txtFont_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFont.Click
		'UPGRADE_ISSUE: MSComDlg.CommonDialog control cdlg was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E047632-2D91-44D6-B2A3-0801707AF686"'
		CopyFont(txtFont, (frmMain.cdlg))
		On Error GoTo ExitSub
		'UPGRADE_WARNING: The CommonDialog CancelError property is not supported in Visual Basic .NET. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="8B377936-3DF7-4745-AA26-DD00FA5B9BE1"'
		frmMain.cdlg.CancelError = True
		'UPGRADE_ISSUE: Constant cdlCFBoth was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
		'UPGRADE_ISSUE: MSComDlg.CommonDialog property cdlg.flags was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		frmMain.cdlg.Flags = MSComDlg.FontsConstants.cdlCFBoth
		'UPGRADE_ISSUE: Constant cdlCFScalableOnly was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
		'UPGRADE_ISSUE: MSComDlg.CommonDialog property cdlg.flags was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		frmMain.cdlg.Flags = MSComDlg.FontsConstants.cdlCFScalableOnly
		'UPGRADE_ISSUE: Constant cdlCFWYSIWYG was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
		'UPGRADE_ISSUE: MSComDlg.CommonDialog property cdlg.flags was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		frmMain.cdlg.Flags = MSComDlg.FontsConstants.cdlCFWYSIWYG
		frmMain.cdlgFont.ShowDialog()
		
		'UPGRADE_ISSUE: MSComDlg.CommonDialog control cdlg was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E047632-2D91-44D6-B2A3-0801707AF686"'
		CopyFont((frmMain.cdlg), txtFont)
		If frmMain.cdlgFont.Font.Name = txtFont.Font.Name Then
			txtFont.Text = frmMain.cdlgFont.Font.Name
		End If
		
ExitSub: 
		'UPGRADE_WARNING: The CommonDialog CancelError property is not supported in Visual Basic .NET. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="8B377936-3DF7-4745-AA26-DD00FA5B9BE1"'
		frmMain.cdlg.CancelError = False
	End Sub
	
	Private Sub CopyFont(ByRef src As Object, ByRef dst As Object)
		On Error Resume Next 'Some objects have only some of the font attributes
		'UPGRADE_WARNING: Couldn't resolve default property of object dst.FontBold. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object src.FontBold. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		dst.FontBold = src.FontBold
		'UPGRADE_WARNING: Couldn't resolve default property of object dst.FontItalic. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object src.FontItalic. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		dst.FontItalic = src.FontItalic
		'UPGRADE_WARNING: Couldn't resolve default property of object dst.FontName. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object src.FontName. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		dst.FontName = src.FontName
		'UPGRADE_WARNING: Couldn't resolve default property of object dst.FontSize. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object src.FontSize. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		dst.FontSize = src.FontSize
		'UPGRADE_WARNING: Couldn't resolve default property of object dst.FontStrikethru. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object src.FontStrikethru. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		dst.FontStrikethru = src.FontStrikethru
		'UPGRADE_WARNING: Couldn't resolve default property of object dst.FontUnderline. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object src.FontUnderline. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		dst.FontUnderline = src.FontUnderline
		'UPGRADE_WARNING: Couldn't resolve default property of object dst.FontTransparent. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object src.FontTransparent. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		dst.FontTransparent = src.FontTransparent
	End Sub
	
	Private Sub CopyFont2RichText(ByRef src As System.Windows.Forms.TextBox, ByRef dst As System.Windows.Forms.RichTextBox)
		Dim lSelStart, lSelLength As Integer
		'On Error Resume Next 'Some objects have only some of the font attributes
		'  With dst.Font
		'    .Bold = src.FontBold
		'    .Italic = src.FontItalic
		'    .Name = src.FontName
		'    .Size = src.FontSize
		'    .Underline = src.FontUnderline
		'  End With
		lSelStart = dst.SelectionStart
		lSelLength = dst.SelectionLength
		dst.SelectionStart = 0
		dst.SelectionLength = Len(dst.RTF)
		dst.Font = VB6.FontChangeBold(dst.SelectionFont, src.Font.Bold)
		dst.SelectionFont = VB6.FontChangeItalic(dst.SelectionFont, src.Font.Italic)
		dst.SelectionFont = VB6.FontChangeName(dst.SelectionFont, src.Font.Name)
		dst.SelectionFont = VB6.FontChangeSize(dst.SelectionFont, src.Font.SizeInPoints)
		dst.SelectionFont = VB6.FontChangeStrikeOut(dst.SelectionFont, src.Font.StrikeOut)
		dst.SelectionFont = VB6.FontChangeUnderline(dst.SelectionFont, src.Font.Underline)
		dst.SelectionStart = lSelStart
		dst.SelectionLength = lSelLength
	End Sub
	
	Private Sub CopyFontFromRichText(ByRef src As System.Windows.Forms.RichTextBox, ByRef dst As System.Windows.Forms.TextBox)
		dst.Font = VB6.FontChangeBold(dst.Font, src.SelectionFont.Bold)
		dst.Font = VB6.FontChangeItalic(dst.Font, src.SelectionFont.Italic)
		dst.Font = VB6.FontChangeName(dst.Font, src.SelectionFont.Name)
		dst.Font = VB6.FontChangeSize(dst.Font, src.SelectionFont.SizeInPoints)
		dst.Font = VB6.FontChangeUnderline(dst.Font, src.SelectionFont.Underline)
	End Sub
End Class