Option Strict Off
Option Explicit On
Friend Class frmOptions
	Inherits System.Windows.Forms.Form
    'Copyright 2000-2008 by AQUA TERRA Consultants
	
    Private Sub Command1_Click(ByVal aEventSender As System.Object, ByVal aEventArgs As System.EventArgs) Handles Command1.Click
        If IsNumeric(txtTreeIndent.Text) Then frmMain.tree1.Indent = CInt(txtTreeIndent.Text)
        If IsNumeric(txtFindTimeout.Text) Then pFindTimeout = CInt(txtFindTimeout.Text)
        CopyFont2RichText(txtFont, (frmMain.txtMain))
        Me.Close()
    End Sub

    Private Sub frmOptions_Load(ByVal aEventSender As System.Object, ByVal aEventArgs As System.EventArgs) Handles MyBase.Load
        txtTreeIndent.Text = CStr(frmMain.tree1.Indent)
        txtFindTimeout.Text = CStr(pFindTimeout)

        CopyFontFromRichText((frmMain.txtMain), txtFont)
        With txtFont
            .Text = .Font.Name & .Font.SizeInPoints
            If .Font.Bold Then .Text &= "Bold"
            If .Font.Italic Then .Text &= "Italic"
            If .Font.Underline Then .Text &= "Underline"
        End With

    End Sub
	
    Private Sub txtFont_Click(ByVal aEventSender As System.Object, ByVal aEventArgs As System.EventArgs) Handles txtFont.Click
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
	
    Private Sub CopyFont2RichText(ByRef aTextBox As System.Windows.Forms.TextBox, ByRef aRichTextBox As System.Windows.Forms.RichTextBox)
        'On Error Resume Next 'Some objects have only some of the font attributes
        '  With dst.Font
        '    .Bold = src.FontBold
        '    .Italic = src.FontItalic
        '    .Name = src.FontName
        '    .Size = src.FontSize
        '    .Underline = src.FontUnderline
        '  End With
        Dim lSelectionStart As Integer = aRichTextBox.SelectionStart
        Dim lSelectionLength As Integer = aRichTextBox.SelectionLength
        aRichTextBox.SelectionStart = 0
        aRichTextBox.SelectionLength = aRichTextBox.Rtf.Length
        aRichTextBox.Font = aTextBox.Font
        aRichTextBox.SelectionFont = aTextBox.Font
        aRichTextBox.SelectionStart = lSelectionStart
        aRichTextBox.SelectionLength = lSelectionLength
    End Sub
	
    Private Sub CopyFontFromRichText(ByRef aRichTextBox As System.Windows.Forms.RichTextBox, ByRef aTextBox As System.Windows.Forms.TextBox)
        aTextBox.Font = aRichTextBox.SelectionFont
    End Sub
End Class