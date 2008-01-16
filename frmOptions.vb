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
        Dim lFontDialog As New FontDialog
        With lFontDialog
            .Font = txtFont.Font
            .AllowVectorFonts = True
            .AllowVerticalFonts = False
            .FontMustExist = True

            If .ShowDialog = Windows.Forms.DialogResult.OK Then
                txtFont.Font = .Font
                txtFont.Text = .Font.Name
            End If
        End With
    End Sub
	
    'Private Sub CopyFont(ByRef src As Object, ByRef dst As Object)
    '	On Error Resume Next 'Some objects have only some of the font attributes
    '       dst.FontBold = src.FontBold
    '	dst.FontItalic = src.FontItalic
    '	dst.FontName = src.FontName
    '	dst.FontSize = src.FontSize
    '	dst.FontStrikethru = src.FontStrikethru
    '	dst.FontUnderline = src.FontUnderline
    '	dst.FontTransparent = src.FontTransparent
    'End Sub
	
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