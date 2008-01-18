Option Strict Off
Option Explicit On

Imports atcUtility
Imports MapWinUtility
Imports Microsoft.Office.Interop.Word

Module modConvert
    'Copyright 2000-2008 by AQUA TERRA Consultants
	
    'pBaseName (~) is the name of program being documented.
    'File ~.txt contains list of source files (pProjectFileName)
	'~.hlp will be created if converting to help (also optionally ~.cnt, ~.hpj)
	'~.doc will be created if converting to printable
	'~.hhp, ~.hhc, ~.ID -> ~.chm
	
    Public OutputFormat As outputType
	
    Public Enum OutputType
        tASCII = 0
        tHTML = 1
        tPRINT = 2
        tHELP = 3
        tHTMLHELP = 4
        NONE = -999
    End Enum
	
    Private pWordBasic As Word.WordBasic
    Private pWordApp As Microsoft.Office.Interop.Word.Application

    Private Const maxLevels As Integer = 9 ' Do you really want sections nested deeper than this?

    Private ProjectFile As Integer

    Private SourceWin As String
    Private ContentsWin As String
    Private TargetWin As String

    Private mTargetText As String

    Private mSourceFilename As String

    Private mSourceBaseDirectory As String
    Private mSaveDirectory As String

    Private HelpSourceRTFName As String
    Private Directory As String

    Private mProjectFileEntrys As New Collection
    Private mNextProjectFileEntry As Integer

    Private HeadingWord(8) As String
    Private HeadingText(maxLevels) As String
    Private HeadingFile(maxLevels) As String

    Private BeforeHTML As String

    Private ContentsEntries(maxLevels) As Integer
    Private HeaderStyle(maxLevels) As String
    Private FooterStyle(maxLevels) As String
    Private BodyStyle(maxLevels) As String
    Private WordStyle(maxLevels) As Collection
    Private BodyTag As String
    Private StyleFile(maxLevels) As String
    Private PromptForFiles As Boolean

    Private FirstHeaderInFile As Boolean
    Private NotFirstPrintFooter As Boolean
    Private NotFirstPrintHeader As Boolean

    Private TablePrintFormat As Integer
    Private TablePrintApply As Integer
    Private TableLines As Boolean

    Private InsertParagraphsAroundImages As Boolean
    Private BuildContents As Boolean
    Private BuildProject As Boolean
    Private FooterTimestamps As Boolean
    Private UpNext As Boolean
    Private BuildID As Boolean
    Private IDfile As Integer
    Private IDnum As Integer
    Private AliasSection As String
    Private HTMLContentsfile As Integer
    Private HTMLHelpProjectfile As Integer
    Private HTMLIndexfile As Integer

    Private SaveFilename As String
    Private InPre As Boolean
    Private AlreadyInitialized As Boolean
    Private LastHeadingLevel As Integer
    Private HeadingLevel As Integer
    Private BookLevel As Integer
    Private StyleLevel As Integer ', IconLevel%
    Private SectionLevelName(99) As String

    Private Const CuteButtons As Boolean = False
    Private Const MoveHeadings As Integer = 0
    Private Const MakeBoxyHeaders As Boolean = False
    Private LinkToImageFiles As Integer '0=store data in document, 1=link+store in doc 2=soft links, -1=do not process images (assigned in Init())

    Public Const Asterisks80 As String = "********************************************************************************"
    Private Const SixSplats As String = "******"
    Private Const SevenSplats As String = "*******"
    Private Const TensPlace As String = "         1         2         3         4         5         6         7         8"
    Private Const OnesPlace As String = "12345678901234567890123456789012345678901234567890123456789012345678901234567890"
    Private Const MaxRowLength As Integer = 80
    Private Const MaxSectionNameLen As Integer = 53
    Private Const TableType As String = "Table-type "
    Private Const lenTableType As Integer = 11
    Private WholeCardHeader As String
    Private lenWholeCardHeader As Integer

    Private TotalTruncated As Integer
    Private TotalRepeated As Integer

    Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Integer, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Integer) As Integer

    'Returns position of first character from chars in str
    'Returns len(str) + 1 if none were found (0 if none found and reverse=true)
    Private Function FirstCharPos(ByRef start As Integer, ByRef aString As String, ByRef chars As String, Optional ByRef reverse As Boolean = False) As Integer
        Dim CharPos, retval, curval, LenChars As Integer
        If reverse Then retval = 0 Else retval = Len(aString) + 1
        LenChars = Len(chars)
        For CharPos = 1 To LenChars
            If reverse Then
                curval = InStrRev(aString, Mid(chars, CharPos, 1), start)
                If curval > retval Then retval = curval
            Else
                curval = InStr(start, aString, Mid(chars, CharPos, 1))
                If curval > 0 And curval < retval Then retval = curval
            End If
        Next CharPos
        FirstCharPos = retval
    End Function

    Public Sub CreateHelpProject(ByRef IDfileExists As Boolean)
        Dim outf As Integer
        outf = FreeFile
        FileOpen(outf, mSaveDirectory & pBaseName & ".hpj", OpenMode.Output)
        PrintLine(outf, "[OPTIONS]" & vbCrLf)
        PrintLine(outf, "LCID=0x409 0x0 0x0 ; English (United States)" & vbCrLf)
        PrintLine(outf, "REPORT=Yes" & vbCrLf)
        PrintLine(outf, "CNT=" & pBaseName & ".cnt" & vbCrLf & vbCrLf)
        PrintLine(outf, "HLP=" & pBaseName & ".hlp" & vbCrLf & vbCrLf)

        PrintLine(outf, "[FILES]" & vbCrLf)
        PrintLine(outf, HelpSourceRTFName & vbCrLf & vbCrLf)

        If IDfileExists Then
            PrintLine(outf, "[MAP]" & vbCrLf)
            PrintLine(outf, "#include <" & pBaseName & ".ID>" & vbCrLf & vbCrLf)
        End If

        PrintLine(outf, "[WINDOWS]" & vbCrLf)
        PrintLine(outf, "Main=" & Chr(34) & pBaseName & " Manual" & Chr(34) & ", , 60672, (r14876671), (r12632256), f2; " & vbCrLf & vbCrLf & "")

        PrintLine(outf, "[CONFIG]" & vbCrLf)
        PrintLine(outf, "BrowseButtons()" & vbCrLf)
        FileClose(outf)
    End Sub

    Public Function HTMLRelativeFilename(ByRef WinFilename As String, ByRef WinStartPath As String) As String
        HTMLRelativeFilename = ReplaceString(RelativeFilename(WinFilename, WinStartPath), "\", "/")
    End Function

    Private Sub OpenHTMLHelpProjectfile()
        'If OutputFormat = tHTMLHELP Then
        HTMLHelpProjectfile = FreeFile
        FileOpen(HTMLHelpProjectfile, mSaveDirectory & pBaseName & ".hhp", OpenMode.Output)
        Print(HTMLHelpProjectfile, "[OPTIONS]" & vbLf)
        Print(HTMLHelpProjectfile, "Auto Index=Yes" & vbLf)
        Print(HTMLHelpProjectfile, "Compatibility=1.1 Or later" & vbLf)
        Print(HTMLHelpProjectfile, "Compiled file=" & pBaseName & ".chm" & vbLf)
        Print(HTMLHelpProjectfile, "Contents file=" & pBaseName & ".hhc" & vbLf)
        'Print #HTMLHelpProjectfile, "Default topic=Introduction.html"
        Print(HTMLHelpProjectfile, "Display compile progress=Yes" & vbLf)
        Print(HTMLHelpProjectfile, "Enhanced decompilation=Yes" & vbLf)
        Print(HTMLHelpProjectfile, "Full-text search=Yes" & vbLf)
        Print(HTMLHelpProjectfile, "Index file = " & pBaseName & ".hhk" & vbLf)
        Print(HTMLHelpProjectfile, "Language=0x409 English (United States)" & vbLf)
        Print(HTMLHelpProjectfile, "Title=" & pBaseName & " Manual" & vbLf & vbLf)
        'Print #HTMLHelpProjectfile, ""
        Print(HTMLHelpProjectfile, "[Files]" & vbLf)
        AliasSection = vbLf & "[ALIAS]"
    End Sub

    Private Sub CheckStyle()
        Dim startTag, closeTag As Integer
        startTag = InStr(LCase(mTargetText), "<style")
        If startTag > 0 Then
            closeTag = InStr(startTag, mTargetText, ">")
            If closeTag < startTag Then
                Logger.Msg("Style tag not terminated in " & mSourceFilename)
            Else
                ReadStyleFile(Mid(mTargetText, startTag + 7, closeTag - startTag - 7), HeadingLevel)
            End If
        ElseIf HeadingLevel <= StyleLevel Then
            StyleLevel = StyleLevel - 1
            While Len(StyleFile(StyleLevel)) = 0
                StyleLevel = StyleLevel - 1
            End While
            ReadStyleFile("", StyleLevel)
        End If
    End Sub

    Private Sub ReadStyleFile(ByRef StyleFilename As String, ByRef HeadingLevel As Integer)
        Dim CurrSection As String = ""
        Dim level As Integer

        BeforeHTML = ""

        For level = 1 To maxLevels
            WordStyle(level) = Nothing
            WordStyle(level) = New Collection
        Next

        If Len(StyleFilename) = 0 Then
            StyleFilename = StyleFile(HeadingLevel)
        Else
            If Not IO.File.Exists(StyleFilename) Then
                If IO.File.Exists(StyleFilename & ".sty") Then
                    StyleFilename = StyleFilename & ".sty"
                End If
            End If
            StyleFilename = CurDir() & "\" & StyleFilename
        End If

        If IO.File.Exists(StyleFilename) Then
            StyleFile(HeadingLevel) = StyleFilename
            StyleLevel = HeadingLevel
            For Each lLine As String In LinesInFile(StyleFilename)
                lLine = lLine.Trim
                Dim FirstChar As String = Left(lLine, 1)
                Select Case FirstChar
                    Case "#", ""
                        'skip comments and blank lines
                    Case "["
                        CurrSection = LCase(Mid(lLine, 2, Len(lLine) - 2))
                        level = 0
                    Case Else
                        If IsNumeric(FirstChar) Then
                            level = CShort(FirstChar)
                            lLine = Mid(lLine, 2)
                            While IsNumeric(Left(lLine, 1))
                                level = level * 10 + CShort(Left(lLine, 1))
                                lLine = Mid(lLine, 2)
                            End While
                            While Left(lLine, 1) = " " Or Left(lLine, 1) = "="
                                lLine = Mid(lLine, 2)
                            End While
                            'UPGRADE_ISSUE: GoSub statement is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C5A1A479-AB8B-4D40-AAF4-DB19A2E5E77F"'
                            Select Case CurrSection
                                Case "beforehtml"
                                    If level = 0 Then BeforeHTML = BeforeHTML & lLine & vbCrLf
                                Case "printsection" : WordStyle(level).Add(lLine)
                                Case "top" : HeaderStyle(level) = lLine
                                Case "bottom" : FooterStyle(level) = lLine
                                Case "body"
                                    If lLine.Length > 0 Then
                                        BodyStyle(level) = "<body " & lLine & ">"
                                    Else
                                        BodyStyle(level) = "<body>"
                                    End If
                            End Select
                        ElseIf CurrSection = "printstart" Then
                            If OutputFormat = OutputType.tPRINT Then WordCommand(lLine, 0)
                        Else
                            For level = 0 To maxLevels
                                'UPGRADE_ISSUE: GoSub statement is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C5A1A479-AB8B-4D40-AAF4-DB19A2E5E77F"'
                                Select Case CurrSection
                                    Case "beforehtml"
                                        If level = 0 Then BeforeHTML = BeforeHTML & lLine & vbCrLf
                                    Case "printsection" : WordStyle(level).Add(lLine)
                                    Case "top" : HeaderStyle(level) = lLine
                                    Case "bottom" : FooterStyle(level) = lLine
                                    Case "body"
                                        If Len(lLine) > 0 Then
                                            BodyStyle(level) = "<body " & lLine & ">"
                                        Else
                                            BodyStyle(level) = "<body>"
                                        End If
                                End Select
                            Next level
                        End If
                End Select
            Next
        End If
    End Sub

    Private Sub WordCommand(ByVal cmdline As String, ByVal localHeadingLevel As Integer)
        Dim arg, cmd, lValue As String
        Dim isnum As Boolean
        Dim consuming As String
        Dim intval As Integer

        System.Windows.Forms.Application.DoEvents()
        On Error GoTo WordCommandErr
        consuming = cmdline
        cmd = StrSplit(consuming, " ", """")
        Dim posVal, typeVal, firstVal As Integer
        With pWordBasic
            Select Case LCase(cmd)
                '      Case "applystyle":
                '        If IsNumeric(consuming) Then localHeadingLevel = CLng(consuming)
                '        .EditBookmark "Hstart" & localHeadingLevel
                '        .EditGoTo "Hend" & localHeadingLevel
                '        .Insert vbCr & vbCr
                '        .EditGoTo "Hstart" & localHeadingLevel
                '        .ExtendSelection
                '        .EditGoTo "Hend" & localHeadingLevel
                '        .Style "ADheading" & localHeadingLevel
                '        .Cancel
                '        .CharRight
                Case "borderbottom" : If IsNumeric(consuming) Then .BorderBottom(CShort(consuming))
                Case "borderinside" : If IsNumeric(consuming) Then .BorderInside(CShort(consuming))
                Case "borderleft" : If IsNumeric(consuming) Then .BorderLeft(CShort(consuming))
                Case "borderlinestyle" '0=none, 1 to 6 increasing thickness, 7,8,9 double, 10 gray, 11 dashed
                    Select Case LCase(Trim(consuming))
                        Case "0", "1", "2", "3", "4", "5", "6"
                            .BorderLineStyle(CShort(consuming))
                        Case "none" : .BorderLineStyle(0)
                        Case "thin" : .BorderLineStyle(1)
                        Case "thick" : .BorderLineStyle(6)
                        Case "double" : .BorderLineStyle(7)
                        Case "doublethick" : .BorderLineStyle(9)
                        Case "dashed" : .BorderLineStyle(11)
                        Case Else : Logger.Msg("Unknown BorderLineStyle: " & consuming, MsgBoxStyle.OkOnly, "AuthorDoc:WordCommand")
                    End Select
                Case "bordernone" : If IsNumeric(consuming) Then .BorderNone(CShort(consuming))
                Case "borderoutside" : If IsNumeric(consuming) Then .BorderOutside(CShort(consuming))
                Case "borderright" : If IsNumeric(consuming) Then .BorderRight(CShort(consuming))
                Case "bordertop" : If IsNumeric(consuming) Then .BorderTop(CShort(consuming))
                Case "charleft" : .CharLeft()
                Case "charright" : .CharRight()
                Case "centerpara" : .CenterPara()
                Case "editclear"
                    If IsNumeric(consuming) Then
                        .EditClear(CInt(consuming))
                    Else
                        .EditClear()
                    End If
                Case "editselectall" : .EditSelectAll()
                Case "filepagesetup"
                    While Len(consuming) > 0
                        lValue = StrSplit(consuming, ",", """")
                        arg = StrSplit(lValue, ":=", """")
                        If IsNumeric(lValue) Then intval = CShort(lValue) Else intval = 0
                        Select Case LCase(arg)
                            Case "topmargin" : pWordBasic.FilePageSetup(TopMargin:=lValue)
                            Case "bottommargin" : pWordBasic.FilePageSetup(BottomMargin:=lValue)
                            Case "leftmargin" : pWordBasic.FilePageSetup(LeftMargin:=lValue)
                            Case "rightmargin" : pWordBasic.FilePageSetup(RightMargin:=lValue)
                            Case "headerdistance" : pWordBasic.FilePageSetup(HeaderDistance:=lValue)
                            Case "facingpages" : pWordBasic.FilePageSetup(FacingPages:=intval)
                            Case "oddandevenpages" : pWordBasic.FilePageSetup(OddAndEvenPages:=intval)
                        End Select
                    End While
                    '      Case "formatdefinestyleborders"
                    '        While Len(consuming) > 0
                    '          val = StrSplit(consuming, ",", """")
                    '          arg = StrSplit(val, ":=", """")
                    '          If IsNumeric(val) Then
                    '            intval = CInt(val)
                    '            Select Case LCase(arg)
                    '              Case "topborder":      .FormatDefineStyleBorders TopBorder:=intval
                    '              Case "leftborder":     .FormatDefineStyleBorders LeftBorder:=intval
                    '              Case "bottomborder":   .FormatDefineStyleBorders BottomBorder:=intval
                    '              Case "rightborder":    .FormatDefineStyleBorders RightBorder:=intval
                    '              Case "horizborder":    .FormatDefineStyleBorders HorizBorder:=intval
                    '              Case "vertborder":     .FormatDefineStyleBorders VertBorder:=intval
                    '              Case "topcolor":       .FormatDefineStyleBorders TopColor:=intval
                    '              Case "leftcolor":      .FormatDefineStyleBorders LeftColor:=intval
                    '              Case "bottomcolor":    .FormatDefineStyleBorders BottomColor:=intval
                    '              Case "rightcolor":     .FormatDefineStyleBorders RightColor:=intval
                    '              Case "horizcolor":     .FormatDefineStyleBorders HorizColor:=intval
                    '              Case "vertcolor":      .FormatDefineStyleBorders VertColor:=intval
                    '              Case "foreground":     .FormatDefineStyleBorders Foreground:=intval
                    '              Case "background":     .FormatDefineStyleBorders Background:=intval
                    '              Case "shading":        .FormatDefineStyleBorders Shading:=intval
                    '              Case "fineshading":    .FormatDefineStyleBorders FineShading:=intval
                    '            End Select
                    '          Else
                    '            logger.msg "non-numeric value for " & arg & " in " & cmd, vbOKOnly, "AuthorDoc:WordCommand"
                    '          End If
                    '        Wend
                Case "formatdefinestylefont"
                    While Len(consuming) > 0
                        lValue = StrSplit(consuming, ",", """")
                        arg = StrSplit(lValue, ":=", """")
                        If LCase(arg) = "font" Then
                            .FormatDefineStyleFont(Font:=lValue)
                        ElseIf IsNumeric(lValue) Then
                            intval = CShort(lValue)
                            Select Case LCase(arg)
                                Case "points" : .FormatDefineStyleFont(Points:=intval)
                                Case "underline" : .FormatDefineStyleFont(Underline:=intval)
                                Case "allcaps" : .FormatDefineStyleFont(AllCaps:=intval)
                                Case "kerning" : .FormatDefineStyleFont(Kerning:=intval)
                                Case "kerningmin" : .FormatDefineStyleFont(KerningMin:=intval)
                                Case "bold" : .FormatDefineStyleFont(Bold:=intval)
                                Case "italic" : .FormatDefineStyleFont(Italic:=intval)
                                Case "outline" : .FormatDefineStyleFont(Outline:=intval)
                                Case "shadow" : .FormatDefineStyleFont(Shadow:=intval)
                                Case "font"
                            End Select
                        End If
                    End While
                    '      Case "formatdefinestylepara"
                    '        While Len(consuming) > 0
                    '          val = StrSplit(consuming, ",", """")
                    '          arg = StrSplit(val, ":=", """")
                    '          isnum = IsNumeric(val)
                    '          If isnum Then
                    '            intval = CInt(val)
                    '            Select Case LCase(arg)
                    '              Case "before":       .FormatDefineStylePara Before:=intval
                    '              Case "after":        .FormatDefineStylePara After:=intval
                    '              Case "keepwithnext": .FormatDefineStylePara KeepWithNext:=intval
                    '              Case "alignment":    .FormatDefineStylePara Alignment:=intval
                    '            End Select
                    '          Else
                    '            logger.msg "non-numeric value for " & arg & " in " & cmd, vbOKOnly, "AuthorDoc:WordCommand"
                    '          End If
                    '        Wend
                Case "formatfont"
                    While Len(consuming) > 0
                        lValue = StrSplit(consuming, ",", """")
                        arg = StrSplit(lValue, ":=", """")
                        If LCase(arg) = "font" Then
                            .FormatFont(Font:=lValue)
                        ElseIf Len(lValue) = 0 Then
                            If IsNumeric(arg) Then .FormatFont(Points:=arg)
                        ElseIf IsNumeric(lValue) Then
                            intval = CShort(lValue)
                            Select Case LCase(arg)
                                Case "points" : .FormatFont(Points:=intval)
                                Case "underline" : .FormatFont(Underline:=intval)
                                Case "allcaps" : .FormatFont(AllCaps:=intval)
                                Case "kerning" : .FormatFont(Kerning:=intval)
                                Case "kerningmin" : .FormatFont(KerningMin:=intval)
                                Case "bold" : .FormatFont(Bold:=intval)
                                Case "italic" : .FormatFont(Italic:=intval)
                                Case "outline" : .FormatFont(Outline:=intval)
                                Case "shadow" : .FormatFont(Shadow:=intval)
                            End Select
                        Else
                            Logger.Msg("non-numeric value for " & arg & " in " & cmd, MsgBoxStyle.OkOnly, "AuthorDoc:WordCommand")
                        End If
                    End While
                Case "formatheaderfooterlink" : .FormatHeaderFooterLink()
                Case "formatpagenumber"
                    While Len(consuming) > 0
                        lValue = StrSplit(consuming, ",", """")
                        arg = StrSplit(lValue, ":=", """")
                        If LCase(arg) = "font" Then
                            .FormatFont(Font:=lValue)
                        ElseIf Len(lValue) = 0 Then
                            If IsNumeric(arg) Then .FormatFont(Points:=arg)
                        ElseIf IsNumeric(lValue) Then
                            intval = CShort(lValue)
                            Select Case LCase(arg)
                                Case "chapternumber" : .FormatPageNumber(ChapterNumber:=intval)
                                Case "numrestart" : .FormatPageNumber(NumRestart:=intval)
                                Case "numformat" : .FormatPageNumber(NumFormat:=intval)
                                Case "startingnum" : .FormatPageNumber(StartingNum:=intval)
                                Case "level" : .FormatPageNumber(Level:=intval)
                                Case "separator" : .FormatPageNumber(Separator:=intval)
                            End Select
                        Else
                            Logger.Msg("non-numeric value for " & arg & " in " & cmd, MsgBoxStyle.OkOnly, "AuthorDoc:WordCommand")
                        End If
                    End While
                Case "formatpara", "formatparagraph"
                    While Len(consuming) > 0
                        lValue = StrSplit(consuming, ",", """")
                        arg = StrSplit(lValue, ":=", """")
                        isnum = IsNumeric(lValue)
                        If isnum Then
                            intval = CShort(lValue)
                            Select Case LCase(arg)
                                Case "before" : .FormatParagraph(Before:=intval)
                                Case "after" : .FormatParagraph(After:=intval)
                                Case "keepwithnext" : .FormatParagraph(KeepWithNext:=intval)
                                Case "alignment" : .FormatParagraph(Alignment:=intval)
                            End Select
                        Else
                            Logger.Msg("non-numeric value for " & arg & " in " & cmd, MsgBoxStyle.OkOnly, "AuthorDoc:WordCommand")
                        End If
                    End While
                    '      Case "formatstyle"
                    '        arg = StrSplit(consuming, ",", """")
                    '        If arg = "Normal" Then 'Provide some good defaults in case they aren't explicit in the style file
                    '          .FormatStyle arg, AddToTemplate:=1, Define:=1
                    '          .FormatDefineStyleFont 11, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 10, 1, 1, "Times New Roman", 0, 0, 0, 0
                    '          .FormatDefineStylePara Chr$(34), Chr$(34), 0, 0, 0, 0, 0, 1, 0, 0, 0, 0, 0, 1, Chr$(34)
                    '          .FormatDefineStyleLang "English (US)", 1
                    '          .FormatDefineStyleBorders 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, -1
                    '        Else
                    '          .FormatStyle arg, BasedOn:="Normal", AddToTemplate:=0, Define:=1
                    '          .FormatStyle arg, Delete:=1
                    '          .FormatStyle arg, BasedOn:="Normal", AddToTemplate:=0, Define:=1
                    '        End If
                Case "formattabs"
                    arg = StrSplit(consuming, ",", """")
                    lValue = StrSplit(consuming, ",", """") '1=left, 2=right
                    If Not IsNumeric(arg) Then
                        Logger.Msg("non-numeric value for tab position in " & cmd, MsgBoxStyle.OkOnly, "AuthorDoc:WordCommand")
                    ElseIf Not IsNumeric(lValue) Then
                        Logger.Msg("non-numeric value for alignment in " & cmd, MsgBoxStyle.OkOnly, "AuthorDoc:WordCommand")
                    Else
                        intval = CShort(lValue)
                        .FormatTabs(arg & """", Align:=intval, Set:=1)
                    End If
                Case "formattabsclear" : .FormatTabs(ClearAll:=1)
                Case "gotoheaderfooter" : .GoToHeaderFooter()
                Case "insert" : .Insert(ReplaceStyleString(consuming, localHeadingLevel))
                Case "insertbreak"
                    Select Case LCase(Trim(consuming))
                        '0 (zero) Page break, 1 Column break, 2 Next Page section break, 3 Continuous section break, 4 Even Page section break, 5 Odd Page section break, 6 Line break (newline character)
                        Case "0", "1", "2", "3", "4", "5", "6"
                            .InsertBreak(CShort(consuming))
                        Case "page" : .InsertBreak(0)
                        Case "column" : .InsertBreak(1)
                        Case "pagesection" : .InsertBreak(2)
                        Case "contsection" : .InsertBreak(3)
                        Case "evenpagesection" : .InsertBreak(4)
                        Case "oddpagesection" : .InsertBreak(5)
                        Case "line" : .InsertBreak(6)
                        Case Else : Logger.Msg("Unknown argument to InsertBreak: " & consuming)
                    End Select
                Case "insertdatetime"
                    If Len(Trim(consuming)) > 0 Then
                        .InsertDateTime(consuming, 0)
                    Else
                        .InsertDateTime("   hh:mm MMMM d, yyyy", 0)
                    End If
                Case "insertfield" : .InsertField(consuming)
                Case "insertpagenumbers"
                    typeVal = 1
                    posVal = 1
                    firstVal = 0
                    While Len(consuming) > 0
                        lValue = StrSplit(consuming, ",", """")
                        arg = StrSplit(lValue, ":=", """")
                        isnum = IsNumeric(lValue)
                        If isnum Then
                            intval = CShort(lValue)
                            Select Case LCase(arg)
                                Case "type" : typeVal = intval
                                Case "position" : posVal = intval
                                Case "firstpage" : firstVal = intval
                            End Select
                        Else
                            Logger.Msg("non-numeric value for " & arg & " in " & cmd, MsgBoxStyle.OkOnly, "AuthorDoc:WordCommand")
                        End If
                    End While
                    .InsertPageNumbers(Type:=typeVal, Position:=posVal, FirstPage:=firstVal)
                Case "InsertParagraphsAroundImages"
                    Select Case LCase(consuming)
                        Case "0", "false" : InsertParagraphsAroundImages = False
                        Case "1", "true" : InsertParagraphsAroundImages = True
                    End Select
                Case "shownextheaderfooter" : .ShowNextHeaderFooter()
                Case "startofdocument" : .StartOfDocument()
                Case "tableprintapply" : If IsNumeric(consuming) Then TablePrintApply = CInt(consuming)
                Case "tableprintformat" : If IsNumeric(consuming) Then TablePrintFormat = CInt(consuming)
                Case "toggleheaderfooterlink" : .ToggleHeaderFooterLink()
                Case "viewfooter" : pWordApp.ActiveWindow.ActivePane.View.SeekView = WdSeekView.wdSeekCurrentPageFooter ' .ViewFooter()
                Case "viewheader" : pWordApp.ActiveWindow.ActivePane.View.SeekView = WdSeekView.wdSeekCurrentPageHeader ' .ViewHeader()
                Case "viewfooterandset"
                    ViewFooterAndSet(ReplaceStyleString(consuming, localHeadingLevel))
                Case "viewheaderandset"
                    ViewHeaderAndSet(ReplaceStyleString(consuming, localHeadingLevel))
                Case "viewnormal" : pWordApp.ActiveWindow.ActivePane.View.SeekView = WdSeekView.wdSeekMainDocument '.ViewNormal()
                Case "viewpage" : .ViewPage()
                Case Else : Logger.Msg("WordCommand not recognized: " & cmd)
            End Select
        End With
        Exit Sub
WordCommandErr:
        'logger.msg "Error with Word command '" & cmdline & "'" & vbCr & Err.Description
        Logger.Dbg("Error with Word command '" & cmdline & "'" & vbCr & Err.Description)
    End Sub

    Private Function ReplaceStyleString(ByRef aString As String, ByRef localHeadingLevel As Integer) As String
        Dim retval, wordstr As String
        Dim endwordpos, level, wordpos, wordnum As Integer
        retval = aString
        retval = ReplaceString(retval, "<sectionname>", HeadingText(localHeadingLevel))
        For level = 1 To localHeadingLevel
            retval = ReplaceString(retval, "<sectionname " & level & ">", HeadingText(level))
        Next
        retval = ReplaceString(retval, "vbTab", vbTab)
        retval = ReplaceString(retval, "vbCr", vbCr)
        retval = ReplaceString(retval, "vbLf", vbLf)
        retval = ReplaceString(retval, "vbCrLf", vbCrLf)
        wordpos = InStr(retval, "<sectionword")
        While wordpos > 0
            endwordpos = InStr(wordpos + 12, retval, ">")
            If endwordpos = 0 Then
                wordpos = 0
            Else
                wordstr = Trim(Mid(retval, wordpos + 12, endwordpos - wordpos - 12))
                If IsNumeric(wordstr) Then
                    wordnum = CShort(wordstr)
                    wordstr = HeadingText(localHeadingLevel)
                    While wordnum > 1
                        StrSplit(wordstr, " ", "")
                    End While
                    wordstr = StrSplit(wordstr, " ", "")
                    retval = Left(retval, wordpos - 1) & wordstr & Mid(retval, endwordpos + 1)
                Else
                    retval = Left(retval, wordpos - 1) & HeadingText(localHeadingLevel) & Mid(retval, endwordpos + 1)
                End If
                wordpos = InStr(wordpos + 1, retval, "<sectionword")
            End If
        End While
        ReplaceStyleString = retval
    End Function

    Public Sub Convert(ByRef aOutputAs As OutputType, ByRef makeContents As Boolean, ByRef timestamps As Boolean, ByRef makeUpNext As Boolean, ByRef makeID As Boolean, ByRef makeProject As Boolean)
        Dim keyword As Object
        Dim replaceSelectionOption As Integer

        Logger.StartToFile(CurDir() & "\log\authordoc.log", False, True)
        Logger.Dbg("StartConvert " & aOutputAs)

        pKeywords = New Collection
        Init()
        OutputFormat = aOutputAs
        BuildContents = makeContents
        BuildProject = makeProject
        FooterTimestamps = timestamps
        UpNext = makeUpNext
        BuildID = makeID
        frmConvert.CmDialog1Open.DefaultExt = "txt"
        If IO.File.Exists(pProjectFileName) Then
            PromptForFiles = False

            For Each lLine As String In LinesInFile(pProjectFileName)
                If lLine.Trim.Length > 0 Then
                    mProjectFileEntrys.Add(lLine)
                End If
            Next
            mNextProjectFileEntry = 1
        Else
            Logger.Msg("Could not open project file " & pProjectFileName)
            Exit Sub
        End If

        mSourceBaseDirectory = IO.Path.GetDirectoryName(pProjectFileName) & "\"
        ChDriveDir(mSourceBaseDirectory)
        mSaveDirectory = mSourceBaseDirectory & "Out\"
        If Not IO.Directory.Exists(mSaveDirectory) Then
            IO.Directory.CreateDirectory(mSaveDirectory)
        End If
        If BuildProject Then
            If OutputFormat = OutputType.tHELP Then
                CreateHelpProject(True)
            ElseIf OutputFormat = OutputType.tHTMLHELP Then
                OpenHTMLHelpProjectfile()
            ElseIf OutputFormat = OutputType.tASCII Then
                HTMLHelpProjectfile = FreeFile()
                FileOpen(HTMLHelpProjectfile, mSaveDirectory & pBaseName & ".txt", OpenMode.Output)
            End If
        End If

        If BuildID Then
            IDfile = FreeFile()
            FileOpen(IDfile, mSaveDirectory & pBaseName & ".ID", OpenMode.Output)
            IDnum = 2
        End If

        InitContents()
        PromptForFiles = False
        Dim lastSourceFilename As String = ""
        mSourceFilename = NextSourceFilename()
        If OutputFormat = OutputType.tPRINT Or OutputFormat = OutputType.tHELP Then
            'pWordApp = New Microsoft.Office.Interop.Word.Application
            'pWordBasic = pWordApp.WordBasic 
            pWordBasic = CreateObject("Word.Basic")
            pWordApp = GetObject(, "Word.Application")

            With pWordBasic
                .AppShow()
                '.ToolsOptionsView PicturePlaceHolders:=1
                .ChDir(mSaveDirectory)
                If OutputFormat = OutputType.tPRINT Then
                    .FileNewDefault()
                    DefinePrintStyles()
                    .FileSaveAs(mSaveDirectory & pBaseName & ".doc", 0)
                    TargetWin = .WindowName
                ElseIf OutputFormat = OutputType.tHELP Then
                    .FileNewDefault()
                    .FilePageSetup(PageWidth:="12 in")
                    .FileSaveAs(mSaveDirectory & HelpSourceRTFName, 6)
                    TargetWin = .WindowName
                End If
                .ChDir(mSourceBaseDirectory)
            End With
        End If
        ReadStyleFile(pBaseName & ".sty", 0)
        LastHeadingLevel = 0
        While mSourceFilename.Length > 0 AndAlso mSourceFilename <> lastSourceFilename
OpeningFile:
            Status("Opening " & mSourceFilename)
            lastSourceFilename = mSourceFilename
            FirstHeaderInFile = True
            System.Windows.Forms.Application.DoEvents()
            If OutputFormat = OutputType.tPRINT Or OutputFormat = OutputType.tHELP Then
                With pWordBasic
                    .Activate(TargetWin)
                    .ScreenUpdating(0) 'comment out to debug (show lots of updates)
                    .EditBookmark("CurrentFileStart")
                    Try
                        .Insert(WholeFileString(Directory & mSourceFilename))
                    Catch ex As Exception
                        GoTo FileNotFound
                    End Try
                    NumberHeaderTagsWithWord()
                    If LinkToImageFiles >= 0 Then
                        .EditGoTo("CurrentFileStart")
                        TranslateIMGtags(Directory & mSourceFilename)
                    End If
                    .EndOfDocument()
                    .ScreenUpdating(1)
                End With
            ElseIf OutputFormat = OutputType.tASCII Then
                Dim i As Integer = FreeFile()
                Try
                    FileOpen(i, Directory & mSourceFilename, OpenMode.Input) 'SourceBaseDirectory &
                Catch ex As Exception
                    GoTo FileNotFound
                End Try
                While Not EOF(i) ' Loop until end of file.
                    ParseHSPFsourceLine(i)
                End While
                If BuildProject And pKeywords.Count() > 0 Then
                    Print(HTMLHelpProjectfile, vbCrLf & "[Keywords]" & vbCrLf)
                    For Each keyword In pKeywords 'TODO: where do keywords come from?
                        'UPGRADE_WARNING: Couldn't resolve default property of object keyword. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        Print(HTMLHelpProjectfile, keyword & vbCrLf)
                    Next keyword
                End If
                FileClose(i)
            ElseIf OutputFormat = OutputType.tHTML Or OutputFormat = OutputType.tHTMLHELP Then
                mTargetText = WholeFileString(mSourceBaseDirectory & mSourceFilename).Trim
TrimTargetText:
                Select Case Left(mTargetText, 1)
                    Case vbCr, vbLf, vbTab, " "
                        mTargetText = Mid(mTargetText, 2)
                        GoTo TrimTargetText
                End Select
                If mTargetText.Length = 0 Then mTargetText = "<toc>"
                Dim lDotPos As Integer = InStrRev(mSourceFilename, ".")
                If lDotPos > 1 Then
                    SaveFilename = Left(mSourceFilename, lDotPos - 1)
                Else
                    SaveFilename = mSourceFilename
                End If
                SaveFilename = SaveFilename & ".html"
                If OutputFormat = OutputType.tHTMLHELP Then
                    FormatTag("b", OutputFormat)
                    FormatKeywordsHTMLHelp()
                    If BuildProject Then Print(HTMLHelpProjectfile, SaveFilename & vbLf)
                End If
                NumberHeaderTags()
                CheckStyle()
                FormatHeadings(OutputType.tHTML, SaveFilename)
                TranslateButtons(OutputFormat)
                MakeLocalTOCs()
                HREFsInsureExtension()
                AbsoluteToRelative()
                CopyImages()
                'FormatCardGraphic()
                SaveInNewDir(mSaveDirectory & SaveFilename)
            End If
            Status("Closing " & mSourceFilename)
OpenNextFile:
            mSourceFilename = NextSourceFilename()
        End While
        If OutputFormat = OutputType.tHTMLHELP And makeProject Then
            Print(HTMLHelpProjectfile, AliasSection & vbLf)
            Print(HTMLHelpProjectfile, "[MAP]" & vbLf & "#include " & pBaseName & ".ID" & vbLf)
            FileClose(HTMLHelpProjectfile)
        ElseIf OutputFormat = OutputType.tASCII Then
            FileClose(IDfile)
            FileClose(HTMLHelpProjectfile)
        End If
        If (OutputFormat = OutputType.tPRINT Or OutputFormat = OutputType.tHELP) Then
            With pWordBasic
                .ToolsOptionsEdit((replaceSelectionOption)) 'save current value of this option
                .ToolsOptionsEdit((1)) 'be sure option is on
                .ScreenUpdating(0) 'comment out to debug (show lots of updates)
                .Activate(TargetWin)
                ConvertTablesToWord()
                ConvertTagsToWord()
                If makeContents Then
                    If OutputFormat = OutputType.tHTMLHELP Or OutputFormat = OutputType.tHTML Then
                        FinishHTMLHelpContents()
                        'ElseIf OutputFormat = tHTML Then
                        '  .Activate ContentsWin
                        '  .FileSaveAs Directory & "Contents.html", 2
                        '  .FileClose 2
                    ElseIf OutputFormat = OutputType.tPRINT Then
                        .Activate(TargetWin)
                        .StartOfDocument()
                        .Insert("Contents" & vbCr & vbCr)
                        .InsertTableOfContents(0, 0, AddedStyles:="ADheading1,1,ADheading2,2,ADheading3,3,ADheading4,4,ADheading5,5,ADheading6,6", RightAlignPageNumbers:=1)
                    ElseIf Len(ContentsWin) > 0 Then
                        .Activate(ContentsWin)
                        .FileSave()
                        .FileClose(2)
                    End If
                End If
                .ToolsOptionsEdit((replaceSelectionOption))
                If Len(TargetWin) > 0 Then
                    .Activate(TargetWin)
                    Status("Saving file: " & TargetWin)
                    .FileSave()
                    .FileClose(2)
                End If
                .ScreenUpdating(1)
                .AppClose()
            End With
        ElseIf OutputFormat = OutputType.tHTMLHELP Or OutputFormat = OutputType.tHTML Then
            FinishHTMLHelpContents()
        End If
        pWordBasic = Nothing
        If IDfile > -1 Then FileClose(IDfile)
        If TotalTruncated > 0 Or TotalRepeated > 0 Then
            Logger.Msg("Total Truncated = " & TotalTruncated & vbCr & "Total Repeated = " & TotalRepeated)
        End If
        Status("Conversion Finished")
        If OutputFormat = OutputType.tHELP Then
            ShellExecute(frmConvert.Handle.ToInt32, "Open", mSaveDirectory & pBaseName & ".hpj", vbNullString, vbNullString, 1) 'SW_SHOWNORMAL"
        ElseIf OutputFormat = OutputType.tHTMLHELP Then
            ShellExecute(frmConvert.Handle.ToInt32, "Open", mSaveDirectory & pBaseName & ".hhp", vbNullString, vbNullString, 1) 'SW_SHOWNORMAL"
        End If
        Logger.Flush()
        Exit Sub

FileNotFound:
        If Logger.Msg("Error opening " & Directory & mSourceFilename & " (" & Err.Description & ")", MsgBoxStyle.RetryCancel, "Help Convert") = MsgBoxResult.Retry Then
            GoTo OpeningFile
        Else
            GoTo OpenNextFile
        End If
    End Sub

    Private Sub ParseHSPFsourceLine(ByRef inFile As Integer)
        Dim buf As String
        Dim SectionDir, SectionNum, SectionDirName, SectionName As String
        Dim parsePos As Integer
        Dim Lenbuf As Integer
        Dim keyword As String
        Dim a2, a, AllCapsStart As Integer
        Static DirectoryLevels As Integer
        Dim InHeader As Boolean
        Dim v As Object
        Static CurrentOutputDirectory As String
        Static CurrentOutputFilename, ImageFilename As String
        Dim dummy As String
        Dim FileRepeat As Integer

        DirectoryLevels = 0
        buf = LineInput(inFile)
        buf = ReplaceString(buf, "<", "&lt;")
        buf = ReplaceString(buf, ">", "&gt;")
        buf = ReplaceString(buf, "&lt;pre&gt;", "<pre>")
        buf = ReplaceString(buf, "&lt;/pre&gt;", "</pre>")
        buf = ReplaceString(buf, "&lt;ol&gt;", "<ol>")
        buf = ReplaceString(buf, "&lt;/ol&gt;", "</ol>")
        buf = ReplaceString(buf, "&lt;li&gt;", "<li>")
        If IsNumeric(Left(buf, 1)) And Left(buf, 10) <> "1234567890" Then 'And Mid(buf, 3, 1) <> vbTab Then
            If IsNumeric(Trim(buf)) Then 'This indicates an image rather than a section header
                If Len(buf) > 3 Then GoTo NormalLine
                keyword = CStr(CInt(Left(buf, 2)) * 2)
                keyword = New String("0", 3 - Len(keyword)) & keyword
                keyword = mSourceFilename & "_files\image" & keyword & ".png"
                For parsePos = 1 To DirectoryLevels
                    keyword = "../" & keyword
                Next
                buf = "<p>" & "<img src=""" & keyword & """>"
            Else
                If Mid(buf, 2, 1) <> "." Then GoTo NormalLine
                If Not IsNumeric(Mid(buf, 3, 1)) Then GoTo NormalLine
                InHeader = True
                If IDfile > 0 Then
                    If InPre Then Print(IDfile, vbCrLf & "</pre>" & vbCrLf)
                    If FileKeywords.Count() > 0 Then
                        For Each v In FileKeywords
                            'UPGRADE_WARNING: Couldn't resolve default property of object v. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            PrintLine(IDfile, "<keyword=" & v & ">" & vbCrLf)
                        Next v
                    End If
                    FileClose(IDfile)
                End If
                InPre = False
                FileKeywords = Nothing
                FileKeywords = New Collection
                parsePos = InStr(buf, " ")
                SectionNum = Left(buf, parsePos - 1)
                SectionName = Trim(Mid(buf, parsePos + 1))

                buf = LineInput(inFile)
                SectionName = SectionName & " " & Trim(buf)
                buf = ""

                parsePos = InStr(SectionName, " -- ")
                If parsePos > 0 Then
                    buf = Trim(Mid(SectionName, parsePos + 4))
                    SectionName = Trim(Left(SectionName, parsePos - 1))
                    If InStr(UCase(SectionName), "BLOCK") > 0 Then
                        SectionName = buf
                        buf = ""
                    End If
                Else
                    parsePos = InStr(SectionName, "(")
                    If parsePos > 0 Then
                        buf = Trim(Mid(SectionName, parsePos))
                        SectionName = Trim(Left(SectionName, parsePos - 1))
                    End If
                End If
                If UCase(Left(SectionName, 7)) = "SECTION" Then SectionName = Mid(SectionName, 9)
                If UCase(Left(SectionName, lenTableType)) = UCase(TableType) Then SectionName = Mid(SectionName, lenTableType + 1)
                buf = "<SecNum " & SectionNum & "> " & "<h>" & SectionName & "</h>" & vbCrLf & buf & vbCrLf


                If Right(SectionNum, 2) = ".0" Then SectionNum = Left(SectionNum, Len(SectionNum) - 2)
                SectionDir = ""
                SectionDirName = ""
                parsePos = InStr(SectionNum, ".")
                If parsePos > 0 Then
                    DirectoryLevels = 1
                    parsePos = InStrRev(SectionNum, ".")
                    SectionDir = Left(SectionNum, parsePos - 1)
                    SectionDirName = SectionLevelName(DirectoryLevels)
                    parsePos = InStr(SectionDir, ".")
                    While parsePos > 0
                        DirectoryLevels = DirectoryLevels + 1
                        SectionDir = Left(SectionDir, parsePos - 1) & "\" & Mid(SectionDir, parsePos + 1)
                        SectionDirName = SectionDirName & "\" & SectionLevelName(DirectoryLevels)
                        parsePos = InStr(parsePos + 1, SectionDir, ".")
                    End While
                    SectionDir = SectionDir & "\"
                    SectionDirName = SectionDirName & "\"
                End If
                IDfile = FreeFile
                If Not IO.Directory.Exists(mSaveDirectory & SectionDirName) Then IO.Directory.CreateDirectory(mSaveDirectory & SectionDirName)
                'Debug.Print
                'Debug.Print SectionDir & SectionNum & ":" & CurrentOutputDirectory & CurrentOutputFilename
                CurrentOutputDirectory = mSaveDirectory & SectionDirName 'SectionDir
                dummy = MakeValidFilename(SectionName)
                If Len(dummy) <= MaxSectionNameLen Then
                    SectionLevelName(DirectoryLevels + 1) = dummy
                Else
                    TotalTruncated = TotalTruncated + 1
                    SectionLevelName(DirectoryLevels + 1) = Trim(Left(dummy, 34) & Right(dummy, 1)) 'MakeValidFilename(buf)
                    Debug.Print("Truncated " & dummy & vbLf & "Shorter = " & SectionLevelName(DirectoryLevels + 1))
                End If
                FileRepeat = 1
SetFilenameHere:
                CurrentOutputFilename = SectionLevelName(DirectoryLevels + 1) & ".txt" 'Mid(SectionNum, Len(SectionDir) + 1) & ".txt"
                If Len(CurrentOutputDirectory & CurrentOutputFilename) > 255 Then
                    Logger.Msg("Path longer than 255 characters detected:" & vbCr & CurrentOutputDirectory & vbCr & CurrentOutputFilename)
                End If
                If IO.File.Exists(CurrentOutputDirectory & CurrentOutputFilename) Then
                    FileRepeat = FileRepeat + 1
                    SectionLevelName(DirectoryLevels + 1) = SectionLevelName(DirectoryLevels + 1) & FileRepeat
                    GoTo SetFilenameHere
                End If
                'Debug.Print Space(2 * DirectoryLevels) & "<li><a href=""Functional Description" & Mid(CurrentOutputDirectory, 21) & SectionLevelName(DirectoryLevels + 1) & """>" & dummy & "</a>"
                If FileRepeat > 1 Then TotalRepeated = TotalRepeated + 1
                FileOpen(IDfile, CurrentOutputDirectory & CurrentOutputFilename, OpenMode.Output)
                If BuildProject Then PrintLine(HTMLHelpProjectfile, Space(2 * DirectoryLevels) & SectionLevelName(DirectoryLevels + 1)) 'Trim(Mid(buf, Len(SectionNum) + 1))  'Mid(SectionNum, Len(SectionDir) + 1)
            End If
        Else
NormalLine:
            If Trim(buf) <> "" Then
                If Left(buf, 3) = "{{{" Then
                    ImageFilename = Mid(buf, 4, InStr(buf, "}}}") - 4) & ".png"
                    buf = "<p><img src=""" & ImageFilename & """>"
RetryImage:
                    Try
                        FileCopy(mSourceBaseDirectory & "png\" & ImageFilename, CurrentOutputDirectory & ImageFilename)
                    Catch
                        Select Case Logger.Msg("Missing Image: " & vbCr & ImageFilename, MsgBoxStyle.AbortRetryIgnore, "Missing")
                            Case MsgBoxResult.Retry : GoTo RetryImage
                            Case MsgBoxResult.Ignore
                            Case MsgBoxResult.Abort : Exit Sub
                        End Select
                    End Try
                End If
                buf = ReplaceString(buf, "[[[", "<br><figure>")
                buf = ReplaceString(buf, "]]]", "</figure>")
                If InPre Then
                    If InStr(buf, "Explanation") > 0 Then
                        InPre = False
                        buf = "</pre>" & vbCrLf & buf
                    End If
                Else
                    If InStr(buf, "****************************************") > 0 Or InStr(buf, "----------------------------------------") > 0 Then
                        InPre = True
                        buf = "<pre>" & vbCrLf & buf
                    End If
                End If
                If Not InPre Then
                    If Left(buf, 4) <> "<li>" Then buf = "<p>" & buf
                End If
            End If
        End If
        AllCapsStart = 0
        Lenbuf = Len(buf)
        Dim buf2 As String = ""
        For parsePos = 1 To Lenbuf
            a = Asc(Mid(buf, parsePos, 1))
            If a > 64 And a < 91 Then 'If capital letter, set AllCapsStart
                If AllCapsStart = 0 Then AllCapsStart = parsePos
            Else
                If AllCapsStart > 0 Then
                    Select Case a
                        Case 48 To 57 'Allow numbers as in PWAT-PARM1 but not F10
                            If AllCapsStart < parsePos - 1 Then GoTo NextChar
                        Case 32, 45, 95 'Allow spaces, dashes, underscores in keywords
                            If parsePos + 2 <= Lenbuf Then
                                a2 = Asc(Mid(buf, parsePos + 1, 1))
                                If a2 > 64 And a2 < 91 Then 'If the next char after is capital
                                    a2 = Asc(Mid(buf, parsePos + 2, 1)) 'And the next one is not lowercase
                                    If Not (a2 > 96 And a2 < 123) Then GoTo NextChar
                                End If
                            End If
                    End Select

                    If parsePos - AllCapsStart > 2 Then
                        keyword = Mid(buf, AllCapsStart, parsePos - AllCapsStart)
                        Try
                            'UPGRADE_WARNING: Couldn't resolve default property of object FileKeywords(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            dummy = FileKeywords.Item(keyword) 'Debug.Print "[" & FileKeywords(keyword) & "]";
                        Catch
                        End Try
                        'Try
                        '    Debug.Print("(" & Keywords(keyword) & ")")
                        'Catch
                        '    Keywords.Add(keyword, keyword)
                        '    Debug.Print("+" & keyword & ";")
                        'End Try
                        'If InHeader Then
                        buf2 &= keyword & Chr(a)
                        'Else
                        '  buf2 = buf2 & vbcrlf _
                        ''      & "<Object id=hhctrl type=""application/x-oleobject""" & vbcrlf _
                        ''      & "classid = ""clsid:adb880a6-d8ff-11cf-9377-00aa003b7a11""" & vbcrlf _
                        ''      & "codebase = ""hhctrl.ocx#Version=4,74,8702,0"" Width = 100 Height = 100>" & vbcrlf _
                        ''      & "<param name=""Command"" value=""KLink"">" & vbcrlf _
                        ''      & "<param name=""Button"" value=""Text:" & keyword & """>" & vbcrlf _
                        ''      & "<param name=""Item1"" value="""">" & vbcrlf _
                        ''      & "<param name=""Item2"" value=""" & keyword & """>" & vbcrlf & "</OBJECT>" & vbcrlf & Chr(a)
                        'End If
                    Else
                        buf2 = buf2 & Mid(buf, AllCapsStart, parsePos - AllCapsStart + 1)
                    End If
                    AllCapsStart = 0
                Else
                    buf2 = buf2 & Chr(a)
                End If
            End If
NextChar:
        Next
        If AllCapsStart > 0 Then buf2 = buf2 & Mid(buf, AllCapsStart)
        Print(IDfile, buf2 & vbCrLf)
        Exit Sub

    End Sub

    Private Sub ConvertTablesToWord()
        With pWordBasic
            .StartOfDocument()
            .EditFindClearFormatting()
            While FindAndDeleteTableStart
                ConvertTableToWord(1)
            End While
        End With
    End Sub

    Private Function FindAndDeleteTableStart() As Boolean
        With pWordBasic
            .EditFind("<table", "", 0)
            If Not .EditFindFound Then Exit Function
            .ExtendSelection()
            .EditFind(">")
            If .EditFindFound Then
                If InStr(LCase(.Selection), "border=0") > 0 Then
                    TableLines = False
                Else
                    TableLines = True
                End If
                .EditClear() 'delete <table...>
            End If
            .Cancel()
            FindAndDeleteTableStart = .EditFindFound
        End With
    End Function

    Private Sub ConvertTableToWord(ByRef RecursionLevel As Integer)
        Dim TableText As String
        Dim TableCols As Integer
        Dim TableLen As Integer
        Dim ColPos As Integer
        Dim RowEnd As Integer
        Dim HeaderCell(,) As Boolean
        Dim MergeCells As Integer
        With pWordBasic
            .EditBookmark("TableStart" & RecursionLevel)
FindEnd:
            .EditFind("</table>")
            If Not .EditFindFound Then Exit Sub

            .EditBookmark("TableEnd" & RecursionLevel)
            .ExtendSelection()
            .EditGoTo("TableStart" & RecursionLevel)
            .Cancel()

            If InStr(LCase(.Selection), "<table") > 0 Then
                If FindAndDeleteTableStart() Then
                    ConvertTableToWord(RecursionLevel + 1)
                    .EditGoTo("TableStart" & RecursionLevel)
                    GoTo FindEnd
                Else
                    .EditGoTo("TableEnd" & RecursionLevel)
                End If
            Else
                .EditGoTo("TableEnd" & RecursionLevel)
            End If

            .EditClear() 'delete </table>
            .EditBookmark("TableEnd" & RecursionLevel)
            .ExtendSelection()
            .EditGoTo("TableStart" & RecursionLevel)

            TableLen = Len(.Selection)
            'skip leading blanks and newlines
SkipBlanks:
            Select Case Asc(.Selection)
                Case 10, 13, 32
                    .CharRight()
                    GoTo SkipBlanks
            End Select
            .Cancel() 'stop extending selection
            .EditBookmark("TableAll")

            'Count columns = # <th> + # <td> in first row
            TableText = LCase(.Selection)

            If InStr(TableText, "<table") > 0 Then
                .EditGoTo("TableEnd" & RecursionLevel)
                .EditBookmark("TableEnd" & RecursionLevel)
            End If

            TableCols = 0
            ColPos = InStr(TableText, "<tr")
            RowEnd = InStr(ColPos + 2, TableText, "tr>")
            If RowEnd = 0 Then RowEnd = Len(TableText)
            While ColPos > 0 And ColPos < RowEnd
                TableCols = TableCols + 1
                ColPos = InStr(ColPos + 1, TableText, "<th")
                If Mid(TableText, ColPos + 3, 8) = " colspan" Then
                    ColPos = ColPos + 12
                    While Not IsNumeric(Mid(TableText, ColPos, 1))
                        ColPos = ColPos + 1
                    End While
                    TableCols = TableCols + CShort(Mid(TableText, ColPos, 1)) - 1
                End If
            End While
            ColPos = InStr(TableText, "<tr")
            ColPos = InStr(ColPos + 1, TableText, "<td")
            While ColPos > 0 And ColPos < RowEnd
                TableCols = TableCols + 1
                ColPos = InStr(ColPos + 1, TableText, "<td")
            End While
            If TableCols > 1 Then TableCols = TableCols - 1
            ReDim HeaderCell(500, TableCols)

            .EditGoTo("TableAll")
            .EditFind("<tr ", "", 0)
            While .EditFindFound
                .CharRight()
                .CharLeft()
                .ExtendSelection()
                .EditFind(">")
                .CharLeft()
                .EditClear()
                .EditGoTo("TableAll")
                .EditFind("<tr ", "", 0)
            End While

            .EditGoTo("TableAll")
            .EditFind("<td ", "", 0)
            While .EditFindFound
                .CharRight()
                .CharLeft()
                .ExtendSelection()
                .EditFind(">")
                .CharLeft()
                .EditClear()
                .EditGoTo("TableAll")
                .EditFind("<td ", "", 0)
            End While

            .EditGoTo("TableAll")
            .EditFind("<th ", "", 0)
            While .EditFindFound
                .CharRight()
                .CharLeft()
                .ExtendSelection()
                .EditFind(">")
                .CharLeft()
                .EditClear()
                .EditGoTo("TableAll")
                .EditFind("<th ", "", 0)
            End While

            .EditGoTo("TableAll")
            .EditReplace("^p^w", "^p", ReplaceAll:=True)
            .EditReplace("^w^p", "^p", ReplaceAll:=True)
            .EditReplace("^p", " ", ReplaceAll:=True)
            .FormatTabs("2""", Align:=0, Set:=1)
            .EditReplace("<tr><th>", "^p", 0, 0, 0, 0, 0, 0, 0, 1, 0, 0, 0)
            .EditReplace("<tr><td>", "^p", 0, 0, 0, 0, 0, 0, 0, 1, 0, 0, 0)
            .EditGoTo("TableAll")
            .EditReplace("<tr>^w<th>", "^p", 0, 0, 0, 0, 0, 0, 0, 1, 0, 0, 0)
            .EditReplace("<tr>^w<td>", "^p", 0, 0, 0, 0, 0, 0, 0, 1, 0, 0, 0)
            .EditReplace("<tr>", "^p", 0, 0, 0, 0, 0, 0, 0, 1, 0, 0, 0)
            .EditReplace("<td>", vbTab, 0, 0, 0, 0, 0, 0, 0, 1, 0, 0, 0)
            .EditReplace("<th>", vbTab, 0, 0, 0, 0, 0, 0, 0, 1, 0, 0, 0)
            .EditReplace("<td ", vbTab, 0, 0, 0, 0, 0, 0, 0, 1, 0, 0, 0)
            .EditReplace("<th ", vbTab, 0, 0, 0, 0, 0, 0, 0, 1, 0, 0, 0)
            .EditReplace("<p>", "", 0, 0, 0, 0, 0, 0, 0, 1, 0, 0, 0)
            .EditReplace("</tr>", "", 0, 0, 0, 0, 0, 0, 0, 1, 0, 0, 0)
            .EditReplace("</td>", "", 0, 0, 0, 0, 0, 0, 0, 1, 0, 0, 0)
            .EditReplace("</thead>", "", 0, 0, 0, 0, 0, 0, 0, 1, 0, 0, 0)
            .EditReplace("<thead>", "", 0, 0, 0, 0, 0, 0, 0, 1, 0, 0, 0)

            .EditGoTo("TableEnd" & RecursionLevel)
            .ExtendSelection()
            .EditGoTo("TableStart" & RecursionLevel)
SkipBlanks2:
            Select Case Asc(.Selection)
                Case 10, 13, 32
                    .CharRight()
                    GoTo SkipBlanks2
            End Select

            .TableInsertTable(ConvertFrom:=1, NumColumns:=TableCols, Format:=TablePrintFormat, Apply:=TablePrintApply)
            .EditBookmark("TableAll")
            .TableColumnWidth(AutoFit:=1)

            If Not TableLines Then
                '.TableGridlines 0
                .BorderBottom(0)
                .BorderInside(0)
                .BorderLeft(0)
                .BorderOutside(0)
                .BorderRight(0)
                .BorderTop(0)
            End If
            .EditFind("colspan=")
            While .EditFindFound
                .EditClear()
                .ExtendSelection()
                .CharRight() 'Want to merge more than 9 columns? probably not.
                MergeCells = CInt(.Selection)
                .CharRight() '>
                .EditClear()
                .NextCell()
                .CharLeft()
                .ExtendSelection()
                'For MergeCount = 2 To MergeCells
                .CharRight(MergeCells)
                'Next
                .TableMergeCells()
                .EditGoTo("TableAll")
                .EditFind("colspan=")
            End While
            .CharRight()
        End With
    End Sub

    Private Sub ConvertTagsToWord()
        With pWordBasic
            Status("Removing HTML Headers")
            RemoveStuffOutsideBody()

            Status("Translating Paragraph Marks")
            InsertParagraphsInPRE(OutputFormat)
            .StartOfDocument()
            Status("Removing Whitespace After Paragraph Marks")
            .EditReplace("^p^w", "^p", ReplaceAll:=True)
            Status("Removing Whitespace Before Paragraph Marks")
            .EditReplace("^w^p", "^p", ReplaceAll:=True)
            Status("Removing Non-HTML Paragraphs")
            .EditReplace("^p", " ", ReplaceAll:=True)

            TranslateLists("ul", 1)
            TranslateLists("ol", 7)

            Status("Replacing HTML Paragraphs")
            .EditReplace("<p>", "^p", ReplaceAll:=True)
            Status("Replacing HTML Line Breaks")
            .EditReplace("<br>", "^l", ReplaceAll:=True)
            Status("Removing Whitespace After Paragraph Marks")
            .EditReplace("^p^w", "^p", ReplaceAll:=True)
            Status("Removing Whitespace Before Paragraph Marks")
            .EditReplace("^w^p", "^p", ReplaceAll:=True)
            Status("Replacing HTML Page Breaks")
            .EditReplace("<page>", "^m", ReplaceAll:=True)

            .EditSelectAll()
            .FormatParagraph(After:=10, LineSpacingRule:=3, LineSpacing:=32)
            If OutputFormat = outputType.tHELP Then .FormatFont(12)
            .Cancel()
            Status("Translating Buttons")
            TranslateButtons(OutputFormat)
            Status("Formatting Headings")
            FormatHeadings(OutputFormat, TargetWin)
            FormatTag("div", OutputFormat)
            FormatTag("pre", OutputFormat)
            FormatTag("figure", OutputFormat)
            FormatTag("u", OutputFormat)
            FormatTag("b", OutputFormat)
            FormatTag("i", OutputFormat)
            FormatTag("sub", OutputFormat)
            FormatTag("sup", OutputFormat)

            If OutputFormat = outputType.tHELP Then HREFsToHelpHyperlinks()
            Status("Removing Remaining HTML Tags")
            HTMLQuotedCharsToPrint()

            'Once more, replace the few places where unwanted whitespace has crept back in
            .EditReplace("^p^w", "^p", ReplaceAll:=True)
            .EditReplace("^w^p", "^p", ReplaceAll:=True)
            .EditReplace("^p^p^p", "^p^p", ReplaceAll:=True)

            .EditReplace("^pXyZ", "^p", ReplaceAll:=True) 'Finally, unhide whitespace at start of lines in <pre>

            'Replace stupid quotes with smart quotes
            .ToolsOptionsAutoFormat(1, 0, 0, 0, 0, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 1)
            .FormatAutoFormat()
        End With
    End Sub

    Private Sub SaveInNewDir(ByRef newFilePath As String)
        Dim fname, path As String
        path = IO.Path.GetDirectoryName(newFilePath)
        fname = Mid(newFilePath, Len(path) + 2)
        If Not IO.Directory.Exists(path) Then IO.Directory.CreateDirectory(path)
        Dim OutFile As Integer
        Dim oldpath As String
        If OutputFormat = outputType.tHTML Or OutputFormat = outputType.tHTMLHELP Then
            OutFile = FreeFile
            FileOpen(OutFile, newFilePath, OpenMode.Output)
            PrintLine(OutFile, mTargetText)
            FileClose(OutFile)
        Else
            With pWordBasic
                oldpath = .DefaultDir(14)
                .ChDir(path)
                .FileSaveAs(fname, 2, AddToMRU:=False)
                .ChDir(oldpath)
            End With
        End If
    End Sub

    Private Function NextSourceFilename() As String
        Dim pos, lvl, ach As Integer
        Dim ch, fn As String 'FileName, will be return value
        Dim buf As String

        If mNextProjectFileEntry >= mProjectFileEntrys.Count Then
            NextSourceFilename = ""
        Else
            buf = mProjectFileEntrys(mNextProjectFileEntry)
            mNextProjectFileEntry += +1
            'insert levels of hierarchy for subsections indented two spaces
            fn = ""
            lvl = 1
            While buf.StartsWith("  ")
                buf = Mid(buf, 3)
                fn &= HeadingWord(lvl) & "\"
                lvl += 1
            End While
            buf = Trim(buf)
            HeadingWord(lvl) = fn
            pos = 1
            ch = Mid(buf, pos, 1)
            ach = Asc(ch)
            While ach > 31 And ach < 127 ' > 47 And ach < 58 Or ach > 64 And ach < 91 Or ach > 96 And ach < 123 Or ach = 92 'alphanumeric or \
                If ach = 34 Or ach = 42 Or ach = 47 Or ach = 58 Or ach = 60 Or ach = 62 Or ach = 63 Or ach = 124 Then 'illegal for file names
                    fn = fn & "_"
                Else
                    fn = fn & ch
                End If
                pos = pos + 1
                If pos <= Len(buf) Then
                    ch = Mid(buf, pos, 1)
                    ach = Asc(ch)
                Else
                    ach = 0
                End If
            End While
            If Len(fn) > Len(HeadingWord(lvl)) Then
                HeadingWord(lvl) = Mid(fn, 1 + Len(HeadingWord(lvl)))
                HeadingLevel = lvl
                NextSourceFilename = fn & pSourceExtension
            Else
                NextSourceFilename = ""
            End If
        End If
    End Function

    'Private Function oldNextSourceFilename() As String
    '  Dim lvl&, ch$, pos&, ach&, fn$ 'FileName, will be return value
    '  Dim ListEntry$
    'Beginning:
    '  If EOF(ProjectFile) Then
    '    NextSourceFilename = ""
    '  Else
    '    Line Input #ProjectFile, ListEntry
    '    If Len(Trim(ListEntry)) = 0 Then GoTo Beginning 'skip blank lines
    '    'insert levels of hierarchy for subsections indented two spaces
    '    fn = ""
    '    lvl = 1
    '    While Left(ListEntry, 2) = "  "
    '      ListEntry = Mid(ListEntry, 3)
    '      fn = fn & HeadingWord(lvl) & "\"
    '      lvl = lvl + 1
    '    Wend
    '    ListEntry = Trim(ListEntry)
    '    HeadingWord(lvl) = fn
    '    pos = 1
    '    ch = Mid(ListEntry, pos, 1)
    '    ach = Asc(ch)
    '    While ach > 31 And ach < 127 ' > 47 And ach < 58 Or ach > 64 And ach < 91 Or ach > 96 And ach < 123 Or ach = 92 'alphanumeric or \
    '      If ach = 34 Or ach = 42 Or ach = 47 Or ach = 58 Or ach = 60 Or ach = 62 Or ach = 63 Or ach = 124 Then 'illegal for file names
    '        fn = fn & "_"
    '      Else
    '        fn = fn & ch
    '      End If
    '      pos = pos + 1
    '      If pos <= Len(ListEntry) Then
    '        ch = Mid(ListEntry, pos, 1)
    '        ach = Asc(ch)
    '      Else
    '        ach = 0
    '      End If
    '    Wend
    '    If Len(fn) > Len(HeadingWord(lvl)) Then
    '      HeadingWord(lvl) = Mid(fn, 1 + Len(HeadingWord(lvl)))
    '      HeadingLevel = lvl
    '      NextSourceFilename = fn & pSourceExtension
    '    Else
    '      NextSourceFilename = ""
    '    End If
    '  End If
    'End Function

    Private Sub InitContents()
        If BuildContents Then
            If OutputFormat = outputType.tHTML Then
                HTMLContentsfile = FreeFile
                FileOpen(HTMLContentsfile, mSaveDirectory & "Contents.html", OpenMode.Output)
                PrintLine(HTMLContentsfile, "<html><head><title>" & pBaseName & " Help Contents</title></head>")
                PrintLine(HTMLContentsfile, "<body>")
                PrintLine(HTMLContentsfile, "<h1>Contents</h1>")
            ElseIf OutputFormat = outputType.tHTMLHELP Then
                HTMLContentsfile = FreeFile
                FileOpen(HTMLContentsfile, mSaveDirectory & pBaseName & ".hhc", OpenMode.Output)
                PrintLine(HTMLContentsfile, "<html><head><!-- Sitemap 1.0 --></head>")
                PrintLine(HTMLContentsfile, "<body>")
                PrintLine(HTMLContentsfile, "<OBJECT type=""text/site properties"">")
                PrintLine(HTMLContentsfile, "<param name=""ImageType"" value=""Folder"">")
                PrintLine(HTMLContentsfile, "</OBJECT>")

                HTMLIndexfile = FreeFile
                FileOpen(HTMLIndexfile, mSaveDirectory & pBaseName & ".hhk", OpenMode.Output)
                PrintLine(HTMLIndexfile, "<html><head></head>")
                PrintLine(HTMLIndexfile, "<body>")
                PrintLine(HTMLIndexfile, "<ul>")

            ElseIf OutputFormat = outputType.tHELP Then
                With pWordBasic
                    'Header of contents file
                    .FileNewDefault()
                    .Insert(":Title " & pBaseName & " Help" & vbCr)
                    .Insert(":Base " & pBaseName & ".hlp" & vbCr)
                    .ChDir(mSaveDirectory)
                    .FileSaveAs(pBaseName & ".cnt", 2)
                    .ChDir(mSourceBaseDirectory)
                    ContentsWin = .WindowName()
                End With
            End If
        End If
    End Sub

    Sub Init()
        Dim HeaderLevel As Integer

        TotalTruncated = 0
        TotalRepeated = 0
        BodyTag = "<body>"
        PromptForFiles = True
        NotFirstPrintHeader = False
        NotFirstPrintFooter = False
        InsertParagraphsAroundImages = False
        HelpSourceRTFName = pBaseName & ".rtf"
        TablePrintFormat = 0
        TablePrintApply = 511
        IDfile = -1
        HTMLContentsfile = -1
        HTMLIndexfile = -1
        If AlreadyInitialized Then Exit Sub
        AlreadyInitialized = True
        PromptForFiles = True
        LinkToImageFiles = 0 '2 ' make soft links with word95 and large document
        frmConvert.CmDialog1Open.DefaultExt = "doc"
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        frmConvert.Show()
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        'IconLevel = 999

        WholeCardHeader = Asterisks80 & vbCrLf & TensPlace & vbCrLf & OnesPlace
        lenWholeCardHeader = Len(WholeCardHeader)

        'set default HTML styles
        For HeaderLevel = 0 To maxLevels
            HeaderStyle(HeaderLevel) = "<hr size=7><h2><sectionname></h2><hr size=7>"
            FooterStyle(HeaderLevel) = ""
            BodyStyle(HeaderLevel) = "<body>"
            WordStyle(HeaderLevel) = New Collection
        Next
    End Sub

    Sub SetUnInitialized()
        AlreadyInitialized = False
    End Sub

    'Function OpenFile(title, Optional filename) As Boolean
    '  Dim InputFile%, buf$, lbuf$
    '  Dim pos&
    '  OpenFile = False
    '  If IsMissing(filename) Then filename = ""
    '
    '  If Not PromptForFiles And filename <> "" And Len(Dir(Directory & filename)) > 0 Then
    '    On Error GoTo nofile
    '    Word.FileNew
    '    InputFile = FreeFile
    '    Open filename For Input As InputFile
    '    lbuf = Input(LOF(InputFile), InputFile)
    '    lbuf = ReplaceString(lbuf, "<html>", "      ")
    '    lbuf = ReplaceString(lbuf, "<head>", "      ")
    '    lbuf = ReplaceString(lbuf, "<body>", "      ")
    '    lbuf = ReplaceString(lbuf, "<form>", "      ")
    '    lbuf = ReplaceString(lbuf, "</html>", "       ")
    '    lbuf = ReplaceString(lbuf, "</head>", "       ")
    '    lbuf = ReplaceString(lbuf, "</body>", "       ")
    '    lbuf = ReplaceString(lbuf, "</form>", "       ")
    '    Word.Insert lbuf
    ''    While Not EOF(InputFile)  ' Loop until end of file.
    ''      Line Input #InputFile, buf
    ''      lbuf = LCase(buf)
    ''
    ''      'blank out confusing tags
    ''      pos = InStr(lbuf, "<html>")
    ''      If pos > 0 Then buf = Left(buf, pos - 1) & "      " & Mid(buf, pos + 6)
    ''      pos = InStr(lbuf, "<head>")
    ''      If pos > 0 Then buf = Left(buf, pos - 1) & "      " & Mid(buf, pos + 6)
    ''      pos = InStr(lbuf, "<body>")
    ''      If pos > 0 Then buf = Left(buf, pos - 1) & "      " & Mid(buf, pos + 6)
    ''      pos = InStr(lbuf, "<form>")
    ''      If pos > 0 Then buf = Left(buf, pos - 1) & "      " & Mid(buf, pos + 6)
    ''
    ''      pos = InStr(lbuf, "</html>")
    ''      If pos > 0 Then buf = Left(buf, pos - 1) & "       " & Mid(buf, pos + 7)
    ''      pos = InStr(lbuf, "</head>")
    ''      If pos > 0 Then buf = Left(buf, pos - 1) & "       " & Mid(buf, pos + 7)
    ''      pos = InStr(lbuf, "</body>")
    ''      If pos > 0 Then buf = Left(buf, pos - 1) & "       " & Mid(buf, pos + 7)
    ''      pos = InStr(lbuf, "</form>")
    ''      If pos > 0 Then buf = Left(buf, pos - 1) & "       " & Mid(buf, pos + 7)
    ''
    ''      Word.Insert buf & vbLf
    ''    Wend
    '    Close InputFile
    '    'Word.FileOpen Directory & filename, 0
    '    OpenFile = True
    '    Word.ViewNormal
    '    Exit Function
    '  End If
    '  With frmConvert.CmDialog1
    '    If Len(filename) > 0 Then .filename = filename
    '
    '    On Error GoTo nofile
    '    .CancelError = True
    '    .DialogTitle = title
    '    .ShowOpen
    '
    '    If IsNull(Dir(.filename)) Then GoTo nofile
    '
    '    Directory = Left(.filename, Len(.filename) - Len(.FileTitle))
    '    Word.FileOpen .filename
    '  End With
    '  Screen.MousePointer = vbHourglass
    '  With Word
    '    .ViewNormal
    '    .ToolsOptionsView Hidden:=1
    '    .ToolsOptionsEdit AutoWordSelection:=0, SmartCutPaste:=0
    '  End With
    '
    '  Screen.MousePointer = vbDefault
    '  OpenFile = True
    '
    '  Exit Function
    'nofile:
    '  frmConvert.CmDialog1.filename = ""
    'End Function

    Private Sub InsertParagraphsInPRE(ByRef OutputFormat As outputType)
        With pWordBasic
            If OutputFormat = outputType.tPRINT Or OutputFormat = outputType.tHELP Then
                .StartOfDocument()
                .EditFindClearFormatting()
                .EditFind("<pre>", "", 0)
                While .EditFindFound
                    .CharRight()
                    .EditBookmark("Hstart")
                    .EditFind("</pre>")
                    If Not .EditFindFound Then Exit Sub
                    .EditBookmark("Hend")
                    .ExtendSelection()
                    .EditGoTo("Hstart")
                    .Cancel()
                    .EditReplace("^p ", "<P>XyZ ", ReplaceAll:=True, Wrap:=0) 'Hide spaces at start of line in <pre>
                    .EditReplace("^p", "<P>", ReplaceAll:=True, Wrap:=0)
                    .EditGoTo("Hend")
                    .CharRight()
                    .EditFind("<pre>", "", 0)
                End While
            End If
        End With
    End Sub

    Private Sub ApplyWordFormat(ByRef f As String, ByRef OutputFormat As outputType, ByRef divArgs As String)
        Dim caption As String
        With pWordBasic
            Select Case LCase(f)
                Case "sub" : .Subscript()
                Case "sup" : .Superscript()
                Case "b" : .FormatFont(Bold:=1)
                Case "i" : .FormatFont(Italic:=1)
                Case "u"
                    If OutputFormat = outputType.tHELP Then .FormatFont(Bold:=1) Else .FormatFont(Underline:=1)
                Case "figure"
                    caption = .Selection
                    .EditClear()
                    .InsertCaption("Figure", "", ": " & caption, Position:=1)
                    .Insert(vbCr)
                Case "pre"
                    .FormatParagraph(After:=0)
                    .FormatFont(Font:="Courier New", Points:=9.5)
                Case "div"
                    If InStr(divArgs, "left") > 0 Then .FormatParagraph(Alignment:=0)
                    If InStr(divArgs, "center") > 0 Then .FormatParagraph(Alignment:=1)
                    If InStr(divArgs, "right") > 0 Then .FormatParagraph(Alignment:=2)
                    If InStr(divArgs, "justify") > 0 Then .FormatParagraph(Alignment:=3)
            End Select
        End With
    End Sub


    Private Sub FormatTag(ByRef tag As String, ByRef OutputFormat As outputType)
        Dim taggedText, begintag, endtag As String
        Dim closeTag, startTag, lenBeginTag As Integer
        Dim divArgs As String

        begintag = "<" & tag & ">"
        endtag = "</" & tag & ">"
        If tag = "div" Then begintag = "<" & tag & " "
        lenBeginTag = Len(begintag)

        Status("Formatting HTML " & begintag)

        Select Case OutputFormat
            Case outputType.tPRINT, outputType.tHELP
                With pWordBasic
                    .StartOfDocument()
                    .EditFindClearFormatting()
                    .EditFind(begintag, "", 0)
                    While .EditFindFound
                        .EditClear() 'delete beginTag
                        .EditBookmark("Hstart")
                        If tag = "div" Then
                            .EditFind(">")
                            If Not .EditFindFound Then Exit Sub
                            .EditClear()
                            .ExtendSelection()
                            .EditGoTo("Hstart")
                            divArgs = LCase(Trim(.Selection))
                            .EditClear()
                            .Insert(vbCr)
                            .EditBookmark("Hstart")
                        Else
                            divArgs = ""
                        End If
                        .EditFind(endtag)
                        If Not .EditFindFound Then Exit Sub
                        .EditClear() 'delete endTag
                        .EditBookmark("Hend")
                        .ExtendSelection()
                        .EditGoTo("Hstart")
                        .Cancel()
                        ApplyWordFormat(tag, OutputFormat, divArgs)
                        .CharRight()
                        .EditFind(begintag, "", 0)
                    End While
                End With
            Case outputType.tHTML, outputType.tHTMLHELP
                startTag = InStr(LCase(mTargetText), begintag)
                While startTag > 0
                    closeTag = InStr(startTag + 2, LCase(mTargetText), endtag)
                    Dim insertText As String = ""
                    If closeTag > 0 Then
                        taggedText = Mid(mTargetText, startTag + lenBeginTag, closeTag - (startTag + lenBeginTag))
                        Select Case LCase(tag)
                            Case "b"
                                If InStr(taggedText, "<") > 0 Then
                                    insertText = ""
                                ElseIf InStr(taggedText, ">") > 0 Then
                                    insertText = ""
                                Else
                                    insertText = "<a name=""" & taggedText & """>" 'Insert link target for bold text
                                    If OutputFormat = outputType.tHTMLHELP Then 'Insert bold text in index
                                        insertText = insertText & "<indexword=""" & taggedText & """>"
                                    End If
                                    mTargetText = Left(mTargetText, startTag - 1) & insertText & Mid(mTargetText, startTag)
                                End If
                        End Select
                    End If
                    startTag = InStr(startTag + Len(insertText) + 2, LCase(mTargetText), begintag)
                End While
        End Select
    End Sub

    Private Sub FormatHeadings(ByRef OutputFormat As outputType, ByRef targetFilename As String)
        Dim localHeadingLevel As Integer
        Dim direction As Integer
        Dim Selection As String
        Dim startTag As Integer
        Dim startNumber As Integer
        Dim endtag As Integer
        Dim closeTag As Integer
        Dim CloseTagEnd As Integer
        If OutputFormat = outputType.tPRINT Or OutputFormat = outputType.tHELP Then
            FormatHeadingsWithWord(OutputFormat, targetFilename)
        Else

            BodyTag = BodyStyle(HeadingLevel)
            startTag = InStr(LCase(mTargetText), "<body")
            If startTag > 0 Then
                endtag = InStr(startTag, mTargetText, ">")
                If endtag > startTag Then
                    BodyTag = Mid(mTargetText, startTag, endtag - startTag + 1)
                    mTargetText = Left(mTargetText, startTag - 1) & Mid(mTargetText, endtag + 1)
                End If
            End If

            startTag = InStr(LCase(mTargetText), "<h")
            While startTag > 0
                If localHeadingLevel = 0 Then localHeadingLevel = HeadingLevel
                direction = 1
                Select Case Mid(mTargetText, startTag + 2, 1)
                    Case ">" : endtag = startTag + 2 : GoTo FindingHend
                        '        Case "+": startNumber = startTag + 3
                        '        Case "-": startNumber = startTag + 3: direction = -1
                    Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
                        startNumber = startTag + 2
                        'endTag = InStr(startTag, TargetText, ">")
                        localHeadingLevel = 0
                    Case Else : GoTo NextHeader
                End Select
                endtag = InStr(startTag, mTargetText, ">")
                If endtag = 0 Then Exit Sub

                'now we have found the header number (startNumber..endtag-1)
                Selection = Mid(mTargetText, startNumber, endtag - startNumber)
                If Len(Selection) > 1 Then
                    If Logger.Msg("Warning: suspicious header tag '<h" & Selection & ">' " & vbCr & "Found in '" & mSourceFilename & "'" & vbCr & "Continue processing this section?", MsgBoxStyle.YesNo) = MsgBoxResult.No Then Exit Sub
                End If
                If IsNumeric(Selection) Then
                    localHeadingLevel = localHeadingLevel + direction * CShort(Selection)
                Else
                    localHeadingLevel = localHeadingLevel + direction
                End If
FindingHend:
                closeTag = InStr(startTag, LCase(mTargetText), "</h")
                If closeTag = 0 Then Exit Sub
                CloseTagEnd = InStr(closeTag, mTargetText, ">")

                HeadingText(localHeadingLevel) = Trim(Mid(mTargetText, endtag + 1, closeTag - endtag - 1))
                'If HeadingText(localHeadingLevel) = "Duration" Then Stop
                HeadingFile(localHeadingLevel) = SaveFilename

                mTargetText = Left(mTargetText, startTag - 1) & Mid(mTargetText, CloseTagEnd + 1)
                '& localHeadingLevel & ">" & HeadingText(localHeadingLevel) & "</h" & localHeadingLevel & ">"
                'look for icon
                '      If IconLevel >= localHeadingLevel Then IconLevel = 999
                '      If OutputFormat = tHTML Or OutputFormat = tHTMLHELP Then
                '        IconStart = InStr(LCase(Mid(TargetText, startTag, 10)), "<img src=")
                '        If IconStart > 0 Then
                '          IconStart = IconStart + startTag - 1
                '          IconEnd = InStr(IconStart, TargetText, ">")
                '          IconFilenameStart = InStr(IconStart, TargetText, """") + 1
                '          If IconFilenameStart > IconStart And IconFilenameStart < IconEnd Then
                '            IconFilenameEnd = InStr(IconFilenameStart + 1, TargetText, """") - 1
                '            If IconFilenameEnd > IconFilenameStart And IconFilenameEnd < IconEnd Then
                '              IconFilename = Mid(TargetText, IconFilenameStart, IconFilenameEnd - IconFilenameStart + 1)
                '              IconLevel = localHeadingLevel
                '              Debug.Print "Icon " & IconFilename & " - level " & IconLevel
                '              TargetText = Left(TargetText, IconStart - 1) & Mid(TargetText, IconEnd + 1)
                '            End If
                '          End If
                '        End If
                '      End If
                startTag = startTag - 1
                FormatHeadingHTML(localHeadingLevel, targetFilename, startTag) 'Insert header and adjust startTag to end of header
                LastHeadingLevel = localHeadingLevel
NextHeader:
                startTag = InStr(startTag + 2, LCase(mTargetText), "<h")
            End While
        End If
    End Sub

    Private Sub FormatHeadingsWithWord(ByRef OutputFormat As outputType, ByRef targetFilename As String)
        Dim localHeadingLevel, direction As Integer
        With pWordBasic
            .StartOfDocument()
            .EditFindClearFormatting()

            .EditFind("<h", "", 0)
            While .EditFindFound
                .EditClear()
                .EditBookmark("Hstart")
                .ExtendSelection()
                .CharRight()
                If localHeadingLevel = 0 Then localHeadingLevel = HeadingLevel
                direction = 1
                Select Case .Selection
                    Case ">"
                        .EditClear()
                        .EditBookmark("Hstart" & localHeadingLevel)
                        .EditFind("</h>")
                        GoTo FindingHend
                    Case "+" : .EditClear() : .EditBookmark("Hstart")
                    Case "-" : .EditClear() : .EditBookmark("Hstart") : direction = -1
                    Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
                        .Cancel() : .CharLeft() : localHeadingLevel = 0
                    Case Else : .Cancel() : .CharLeft() : .Insert("<h") : GoTo NextHeader
                End Select
                .EditFind(">", "", 0)
                If Not .EditFindFound Then .Insert("<h") : Exit Sub
                .EditClear()
                .ExtendSelection()
                .EditGoTo("Hstart")

                'now we have selected the header number and deleted the <h> from around it
                If Len(.Selection) > 1 Then
                    If Logger.Msg("Warning: suspicious header tag '<h" & .Selection & ">' " & vbCr & "Found in '" & .WindowName & "'." & vbCr & "Continue processing this section?", MsgBoxStyle.YesNo) = MsgBoxResult.No Then
                        .Insert("</h" & .Selection & ">")
                        Exit Sub
                    End If
                End If
                If IsNumeric(.Selection) Then
                    localHeadingLevel = localHeadingLevel + direction * CShort(.Selection)
                Else
                    localHeadingLevel = localHeadingLevel + direction
                End If
                .EditClear()
                .EditBookmark("Hstart" & localHeadingLevel)
                .EditFind("</h" & localHeadingLevel & ">") 'find end tag of this header
FindingHend:
                If Not .EditFindFound Then Exit Sub
                .EditClear() 'delete </hx> tag
                .EditBookmark("Hend" & localHeadingLevel)
                .ExtendSelection()
                .EditGoTo("Hstart" & localHeadingLevel)
                .Cancel()
                HeadingText(localHeadingLevel) = Trim(.Selection)
                'If HeadingText(localHeadingLevel) = "Duration" Then Stop
                HeadingFile(localHeadingLevel) = SaveFilename
                .CharRight()
                'look for icon
                '      If IconLevel >= localHeadingLevel Then IconLevel = 999
                '      .CharRight 4, 1
                '      .EditFind "^g"
                '      If .EditFindFound Then 'we may want to move icon before heading
                '        Dim sizex$, sizey$
                '        .FormatPicture
                '        sizex = Word.CurValues.FormatPicture.sizex
                '        sizey = Word.CurValues.FormatPicture.sizey
                '        'if image is less than 1" square, then move it.
                '        If val(Left(sizey, Len(sizey) - 1)) < 1 And _
                ''           val(Left(sizex, Len(sizex) - 1)) < 1 Then
                '          .EditCut
                '          IconLevel = localHeadingLevel
                '
                '          'look for a second icon
                '          .CharRight 4, 1
                '          .EditFind "^g"
                '          If .EditFindFound Then 'we may want to move this one, too
                '            .FormatPicture
                '            sizex = Word.CurValues.FormatPicture.sizex
                '            sizey = Word.CurValues.FormatPicture.sizey
                '            'if image is less than 1" square, then move it.
                '            If val(Left(sizey, Len(sizey) - 1)) < 1 And _
                ''               val(Left(sizex, Len(sizex) - 1)) < 1 Then
                '               .CharLeft  'paste first icon found earlier
                '               .EditPaste
                '               .CharLeft
                '               .CharRight 2, 1 'select both icons and cut them together
                '               .EditCut
                '            End If
                '          Else
                '            .CharLeft
                '          End If
                '        End If
                '      Else
                '        .CharLeft
                '      End If

                .EditGoTo("Hend" & localHeadingLevel)
                .ExtendSelection()
                .EditGoTo("Hstart" & localHeadingLevel)
                .Cancel()

                If OutputFormat = outputType.tPRINT Then
                    FormatHeadingPrint(localHeadingLevel)
                ElseIf OutputFormat = outputType.tHELP Then
                    FormatHeadingHelp(localHeadingLevel)
                Else
                    .Cancel()
                End If
                LastHeadingLevel = localHeadingLevel
NextHeader:
                .CharRight()
                .EditFind("<h", "", 0)
            End While
        End With
    End Sub

    Sub RemoveStuffOutsideBody()
        With pWordBasic
            .StartOfDocument()
            .EditReplace("<html>", "", ReplaceAll:=True)
            .StartOfDocument()
            .EditReplace("</html>", "", ReplaceAll:=True)
            .StartOfDocument()
            .EditFind("<head>", "", 0)
            While .EditFindFound
                .ExtendSelection()
                .EditFind("</head>")
                If .EditFindFound Then .EditClear()
                .EditFind("<head>")
            End While
            .StartOfDocument()
            .EditReplace("<body>", "", ReplaceAll:=True)
            .StartOfDocument()
            .EditReplace("</body>", "", ReplaceAll:=True)
        End With
    End Sub

    Sub TranslateButtons(ByRef OutputFormat As outputType)

        If Not CuteButtons Then Exit Sub

        Dim label As String
        If OutputFormat = outputType.tHELP Or OutputFormat = outputType.tHTML Or OutputFormat = outputType.tHTMLHELP Then
            With pWordBasic
                .StartOfDocument()
                .EditFind("' button")
                While .EditFindFound
                    .CharLeft()
                    .EditBookmark("LabelEnd")
                    .EditFind("'", direction:=1)
                    .CharRight()
                    .ExtendSelection()
                    .EditGoTo("LabelEnd")
                    label = .Selection
                    If Len(label) < 20 Then
                        .EditClear(2)
                        .CharLeft()
                        .EditClear()
                        If OutputFormat = outputType.tHELP Then
                            .Insert("{button " & label & ",}")
                        ElseIf OutputFormat = outputType.tHTML Or OutputFormat = outputType.tHTMLHELP Then
                            .Insert("<input type=submit value=""" & label & """>")
                        Else 'should not get here
                            .Insert("'" & label & "'")
                        End If
                    Else
                        Status("false alarm, not a button")
                    End If
                    .EditFind("' button", direction:=0)
                End While
                .EditFind(PatternMatch:=0)
            End With
        End If
    End Sub

    Sub TranslateLists(ByRef tag As String, ByRef MarkerType As Integer)
        Dim begintag, endtag As String
        Dim bulletNumber As Integer

        begintag = "<" & tag & ">"
        endtag = "</" & tag & ">"
        With pWordBasic
            .StartOfDocument()
            Status("Translating HTML <" & tag & ">")
            .EditFind(endtag, Direction:=0)
            While .EditFindFound
                .EditClear()
                .Insert(vbCr)
                .CharLeft()
                .EditBookmark("ListEnd")
                .EditFind(begintag, Direction:=1)
                If .EditFindFound Then
                    '.Insert vbCr
                    .EditClear()
                    WordRemoveTrailingWhitespace()
                    .EditBookmark("ListStart")
                    .ExtendSelection()
                    .EditGoTo("ListEnd")
                    .EditBookmark("WholeList")
                    .Cancel()
                    bulletNumber = 1
                    '        .FormatBulletsAndNumbering Hang:=1, preset:=MarkerType
                    '        If tag = "ol" Then .FormatNumber StartAt:=bulletNumber: bulletNumber = bulletNumber + 1
                    '        .FormatParagraph LeftIndent:="0.5"""

                    .EditFind("<li>", Direction:=0)
                    '        If .EditFindFound Then
                    '          .EditClear
                    '          WordRemoveTrailingWhitespace
                    '          .EditFind "<li>", direction:=0
                    '        End If
                    While .EditFindFound
                        .Insert(vbCr)
                        WordRemoveTrailingWhitespace()
                        .FormatBulletsAndNumbering(Hang:=1, Preset:=MarkerType)
                        If tag = "ol" Then
                            .FormatNumber(StartAt:=bulletNumber)
                            bulletNumber = bulletNumber + 1
                            .FormatParagraph(LeftIndent:="0.5""")
                        End If
                        .EditGoTo("WholeList")
                        .EditFind("<li>")
                    End While
                    '.ScreenUpdating 1
                    '        .EditGoTo "WholeList"
                    .EditReplace("<p>", "<br><br>", ReplaceAll:=True)
                    '        .EditFind "<p>"
                    '        While .EditFindFound
                    '          .CharLeft
                    '          .Insert vbCr
                    '          .SkipNumbering
                    '          .EditGoTo "WholeList"
                    '          .EditFind "<p>"
                    '        Wend


                    .EditGoTo("ListEnd")
                End If
                .EditFind(endtag)
            End While
        End With
    End Sub

    Private Sub WordRemoveTrailingWhitespace()
        Dim asci As Integer
        With pWordBasic
            .ExtendSelection()
            .CharRight()
            If Len(.Selection) > 0 Then asci = Asc(.Selection) Else asci = 33
            While asci < 33 'skip leading blanks and newlines
                .EditClear()
                .CharRight()
                If Len(.Selection) > 0 Then asci = Asc(.Selection) Else asci = 33
            End While
            .CharLeft()
        End With
    End Sub

    'Private Sub InsertHTMLskeleton(headText$)
    '  '<form> allows buttons to be displayed as buttons in browser
    '  With Word
    '    .StartOfDocument
    '    .Insert "<html>" & vbLf & "<head>" & headText & "</head>" & vbLf
    '    .Insert "<body>" & vbLf
    '    If CuteButtons Then .Insert "<form>" & vbLf
    '    .EndOfDocument 'in case there is already some text in the file
    '    If CuteButtons Then .Insert vbLf & "</form>"
    '    .Insert vbLf & "</body></html>"
    '    .StartOfDocument
    '    If CuteButtons Then .LineDown 3 Else .LineDown 2
    '  End With
    'End Sub

    Private Function AlinkAnchor(ByRef keyword As String) As String
        AlinkAnchor = "<Object type=""application/x-oleobject"" classid=""clsid:1e2a7bd0-dab9-11d0-b93a-00c04fc99f9e"">" & vbCrLf & "    <param name=""ALink Name"" value=""" & keyword & """>" & vbCrLf & "</Object>" & vbCrLf
    End Function

    Private Function KeywordAnchor(ByRef keyword As String) As String
        KeywordAnchor = "<Object type=""application/x-oleobject"" classid=""clsid:1e2a7bd0-dab9-11d0-b93a-00c04fc99f9e"">" & vbCrLf & "    <param name=""Keyword"" value=""" & keyword & """>" & vbCrLf & "</Object>" & vbCrLf
    End Function

    Private Function KeywordButton(ByRef keyword As String) As String
        KeywordButton = "<Object id=hhctrl type=""application/x-oleobject""" & vbCrLf & "classid=""clsid:adb880a6-d8ff-11cf-9377-00aa003b7a11""" & vbCrLf & "codebase=""hhctrl.ocx#Version=4,74,8702,0"" Width = 100 Height = 100>" & vbCrLf & "<param name=""Command"" value=""KLink"">" & vbCrLf & "<param name=""Button"" value=""Text:" & keyword & """>" & vbCrLf & "<param name=""Item1"" value="""">" & vbCrLf & "<param name=""Item2"" value=""" & keyword & """>" & vbCrLf & "</OBJECT>" & vbCrLf

    End Function

    Private Sub FormatKeywordsHTMLHelp()
        Dim startPos, endPos As Integer
        Dim keyword, KeywordPreamble As String
        startPos = InStr(LCase(mTargetText), "<keyword=")
        If startPos > 0 Then
            KeywordPreamble = "<p>" & vbCrLf & "Keywords:"
            mTargetText = Left(mTargetText, startPos - 1) & KeywordPreamble & Mid(mTargetText, startPos)
            startPos = startPos + Len(KeywordPreamble)
        End If
        While startPos > 0
            endPos = InStr(startPos, mTargetText, ">")
            If endPos > 0 Then
                keyword = TrimQuotes(Trim(Mid(mTargetText, startPos + 9, endPos - startPos - 9)))
                mTargetText = Left(mTargetText, startPos - 1) & KeywordAnchor(keyword) & "&nbsp;" & KeywordButton(keyword) & Mid(mTargetText, endPos + 1)
            End If
            startPos = InStr(endPos, LCase(mTargetText), "<keyword=")
        End While

        startPos = InStr(LCase(mTargetText), "<indexword=")
        While startPos > 0
            endPos = InStr(startPos, mTargetText, ">")
            If endPos > 0 Then
                keyword = TrimQuotes(Trim(Mid(mTargetText, startPos + 11, endPos - startPos - 11)))
                mTargetText = Left(mTargetText, startPos - 1) & KeywordAnchor(keyword) & Mid(mTargetText, endPos + 1)
            End If
            startPos = InStr(endPos, LCase(mTargetText), "<indexword=")
        End While

    End Sub

    'Private Sub FormatKeywordsHTMLHelpUsingWord()
    '  Dim keyword$
    '  With Word
    '    .StartOfDocument
    '    .EditFind "<Keyword=", "", 0
    '    While .EditFindFound
    '      .EditClear
    '      .EditBookmark "LinkStart"
    '      .EditFind ">"
    '      If Not .EditFindFound Then Exit Sub
    '      .EditClear
    '      .ExtendSelection
    '      .EditGoTo "LinkStart"
    '      keyword = Trim(.Selection)
    '      .EditClear
    '      .Cancel
    '      .Insert "<Object type=""application/x-oleobject"" classid=""clsid:1e2a7bd0-dab9-11d0-b93a-00c04fc99f9e"">" & vbCrLf
    '      .Insert "    <param name=""Keyword"" value=""" & keyword & """>" & vbCrLf
    '      .Insert "</Object>"
    '      'Print #IDfile, "<p>"
    '      .Insert vbCrLf _
    ''            & "<Object id=hhctrl type=""application/x-oleobject""" & vbCrLf _
    ''            & "classid=""clsid:adb880a6-d8ff-11cf-9377-00aa003b7a11""" & vbCrLf _
    ''            & "codebase=""hhctrl.ocx#Version=4,74,8702,0"" Width = 100 Height = 100>" & vbCrLf _
    ''            & "<param name=""Command"" value=""KLink"">" & vbCrLf _
    ''            & "<param name=""Button"" value=""Text:" & keyword & """>" & vbCrLf _
    ''            & "<param name=""Item1"" value="""">" & vbCrLf _
    ''            & "<param name=""Item2"" value=""" & keyword & """>" & vbCrLf & "</OBJECT>" & vbCrLf
    '
    '      .Insert "<p>"
    '      .EditFind "<Keyword="
    '    Wend
    '  End With
    'End Sub

    Private Sub NumberHeaderTags()
        Dim curTag As String
        Dim selStr As String
        Dim direction, localHeadingLevel As Integer
        Dim IgnoreUnknown As Boolean
        Dim startPos, endPos As Integer
        If OutputFormat = outputType.tPRINT Or OutputFormat = outputType.tHELP Then
            NumberHeaderTagsWithWord()
        Else

            startPos = InStr(LCase(mTargetText), "<h")
            If startPos = 0 Then 'need to insert section header
                mTargetText = "<h" & HeadingLevel & ">" & HeadingWord(HeadingLevel) & "</h" & HeadingLevel & "> " & vbLf & mTargetText
            End If
            curTag = "<h"
            startPos = InStr(LCase(mTargetText), curTag)
            localHeadingLevel = HeadingLevel
            While startPos > 0
                endPos = InStr(startPos, mTargetText, ">")
                If endPos > 0 Then
                    selStr = Mid(mTargetText, startPos + Len(curTag), endPos - startPos - Len(curTag))
                    direction = 0
                    If selStr = "" Then
                        mTargetText = Left(mTargetText, startPos - 1) & curTag & localHeadingLevel & Mid(mTargetText, endPos)
                    Else
                        Select Case Left(selStr, 1)
                            Case "+", "-"
                                If Left(selStr, 1) = "+" Then direction = 1 Else direction = -1
                                selStr = Mid(selStr, 2)
                                If IsNumeric(selStr) Then
                                    localHeadingLevel = HeadingLevel + direction * CShort(selStr)
                                Else
                                    localHeadingLevel = HeadingLevel + direction
                                End If
                                mTargetText = Left(mTargetText, startPos + 1) & localHeadingLevel & Mid(mTargetText, endPos)
                            Case Else
                                If IsNumeric(selStr) Then
                                    localHeadingLevel = CShort(selStr)
                                ElseIf UCase(selStr) = "EAD" Or UCase(selStr) = "TML" Then
                                    'ignore <html> and <head> even though these should not be in source files
                                ElseIf UCase(Left(selStr, 1)) = "R" Then
                                    'ignore <hr> <hr size=7> etc.
                                Else
                                    If Not IgnoreUnknown Then
                                        If Logger.Msg("Unknown heading tag '<h" & selStr & ">'" & vbCr & "In file " & mSourceFilename & vbCr & "Warn about future unknown headers?", MsgBoxStyle.YesNo, "Number Header Tags") = MsgBoxResult.No Then
                                            IgnoreUnknown = True
                                        End If
                                    End If
                                End If
                        End Select
                    End If
                    If curTag = "<h" Then curTag = "</h" Else curTag = "<h"
                    startPos = InStr(endPos, LCase(mTargetText), curTag)
                End If
            End While
        End If
    End Sub

    Private Sub NumberHeaderTagsWithWord()
        Dim curTag As String
        Dim selStr As String
        Dim direction, localHeadingLevel As Integer
        With pWordBasic
            .EditGoTo("CurrentFileStart")
            .EditFind("<h", "", 0)
            If Not .EditFindFound Then 'need to insert section header
                .EditGoTo("CurrentFileStart")
                .Insert("<h" & HeadingLevel & ">" & HeadingWord(HeadingLevel) & "</h" & HeadingLevel & "> " & vbLf)
            End If
            curTag = "<h"
            .EditGoTo("CurrentFileStart")
            .EditFind(curTag, "", 0)
            While .EditFindFound
                .CharRight()
                .ExtendSelection()
                .EditFind(">", "", 0)
                localHeadingLevel = HeadingLevel
                If .EditFindFound Then
                    selStr = Left(.Selection, Len(.Selection) - 1)
                    .CharLeft()
                    direction = 0
                    If selStr = "" Then
                        .Cancel()
                        .Insert(CStr(localHeadingLevel))
                    Else
                        Select Case Left(selStr, 1)
                            Case "+", "-"
                                .EditClear()
                                If Left(selStr, 1) = "+" Then direction = 1 Else direction = -1
                                selStr = Mid(selStr, 2)
                                If IsNumeric(selStr) Then
                                    localHeadingLevel = HeadingLevel + direction * CShort(selStr)
                                Else
                                    localHeadingLevel = HeadingLevel + direction
                                End If
                                .Insert(CStr(localHeadingLevel))
                            Case "r", "R" : curTag = "</h" 'ignore <hr>
                            Case Else
                                If IsNumeric(selStr) Then
                                    localHeadingLevel = CShort(selStr)
                                Else
                                    .StartOfLine()
                                    .ExtendSelection()
                                    .EndOfLine()
                                    Logger.Msg("Suspicious heading " & selStr)
                                    Stop
                                End If
                        End Select
                        .Cancel()
                    End If
                    If curTag = "<h" Then curTag = "</h" Else curTag = "<h"
                    .EditFind(curTag, "", 0)
                End If
            End While
        End With
    End Sub

    Private Sub TranslateIMGtags(ByRef sourcefile As String)
        Dim LinkFilename, path As String
        path = sourcefile
        While path <> "" And Right(path, 1) <> "\"
            path = Left(path, Len(path) - 1)
        End While
        Dim InsertParagraphs As Boolean
        Dim curfilename As String
        Dim LinkToThisImageFile As Integer
        With pWordBasic
            .EditFind("<IMG ", "", 0)
            While .EditFindFound
                InsertParagraphs = InsertParagraphsAroundImages
                LinkToThisImageFile = LinkToImageFiles
                .EditClear()
                .EditBookmark("ImgStart")
                .EditFind("SRC=""", "", 0)
                .EditClear()
                .EditBookmark("LinkStart")
                .EditFind("""")
                If Not .EditFindFound Then Exit Sub
                .EditClear()
                .ExtendSelection()
                .EditGoTo("LinkStart")
                'curpath = path
                curfilename = .Selection
                .EditClear()
                .EditGoTo("ImgStart")
                LinkFilename = IO.Path.GetDirectoryName(pProjectFileName) & "\" & path & curfilename
                'While Left(curfilename, 2) = ".."
                '  curfilename = Right(curfilename, Len(curfilename) - 3)
                '  curpath = Left(curpath, Len(curpath) - 1)
                '  While Len(curpath) > 0 And Right(curpath, 1) <> "\"
                '    curpath = Left(curpath, Len(curpath) - 1)
                '  Wend
                'Wend
                'LinkFilename = curpath & curfilename
                .ExtendSelection()
                .EditFind(">")
                If InStr(1, LinkFilename, "icon", 1) > 0 Then
                    InsertParagraphs = False
                    'LinkToThisImageFile = 0
                End If
                If Len(.Selection) > 1 Then InsertParagraphs = False 'probably said ALIGN=LEFT in img tag
                .EditClear()
                .Cancel()
                If InsertParagraphs Then .Insert("<p>")
                .InsertPicture(LinkFilename, LinkToThisImageFile)

                '.Insert "{ INCLUDEPICTURE """ & LinkFilename & """ \* MERGEFORMAT \D }"
                '.CharLeft 1, 1

                'Dim sizexS$, sizeyS$, sizex!, sizey!
                '.FormatPicture
                '.FormatPicture ScaleX:="100%", ScaleY:="100%"
                'sizexS = Word.CurValues.FormatPicture.sizex
                'sizex = Val(Left(sizexS, Len(sizexS) - 1))
                'sizeyS = Word.CurValues.FormatPicture.sizey
                'sizey = Val(Left(sizeyS, Len(sizeyS) - 1))

                'If sizex > 6.5 Then
                '  Dim percent$
                '  percent = (100# * 6.5 / sizex) & "%"
                '  .FormatPicture ScaleX:=percent, ScaleY:=percent
                'End If
                '.CharRight

                If InStr(1, LinkFilename, "icon", 1) > 0 Then
                    .CharLeft(1, 1)
                    .FormatFont(Position:=-12) 'half-point units
                    .CharRight()
                End If

                If InsertParagraphs Then .Insert("<p>")
                .EditFind("<IMG ")
            End While
        End With
    End Sub

    Public Sub FormatHeadingHTML(ByRef thisHeadingLevel As Integer, ByRef targetFilename As String, ByRef thisHeadingStart As Integer)
        Static LinkToFirstHeader As String
        Dim hn As String
        Dim ht As String
        Dim TextToInsert As String
        Dim TextToPrepend As String
        Dim TextToAppend As String
        Dim RuleInTable As Integer
        Dim TableStart As Integer
        Dim RuleEnd As Integer
        Dim TableEnd As Integer
        Dim h As Integer
        Dim ParentHT As String 'HeadingText
        If MoveHeadings <> 0 Then
            hn = "h" & (thisHeadingLevel + MoveHeadings) & ">"
            mTargetText = Left(mTargetText, thisHeadingStart) & "<" & hn & HeadingText(thisHeadingLevel) & "</" & hn & Mid(mTargetText, thisHeadingStart + 1)
        Else
            'Dim IconPath$

            ht = HeadingText(thisHeadingLevel)
            TextToInsert = ""
            TextToPrepend = ""
            TextToAppend = ReplaceString(FooterStyle(thisHeadingLevel), "<sectionname>", ht)
            '    IconPath = ""
            '    For i = IconLevel + 1 To thisHeadingLevel
            '      IconPath = "../" & IconPath
            '    Next i

            'Insert name anchor around heading
            TextToInsert = TextToInsert & "<a name=""" & ht & """>"
            TextToInsert = TextToInsert & ReplaceString(HeaderStyle(thisHeadingLevel), "<sectionname>", ht) & "</a>" & vbLf
            '    If IconLevel <= thisHeadingLevel Then
            '      TextToInsert = TextToInsert & "<img src=""" & IconPath & IconFilename & """ align=right>"
            '    End If

            If FirstHeaderInFile Then 'Insert navigation to parents in hierarchy
                FirstHeaderInFile = False
                If UpNext Then
                    LinkToFirstHeader = "Up to: <a href=""#" & ht & """>" & ht & "</a>" & vbLf
                    If thisHeadingLevel > 1 Or (OutputFormat = outputType.tHTML And BuildContents) Then
                        TextToInsert = TextToInsert & "Up to: "
                        For h = thisHeadingLevel - 1 To 1 Step -1
                            ParentHT = HeadingText(h)
                            TextToInsert = TextToInsert & "<a href=""\" & HeadingFile(h) & "#" & ParentHT & """>" & ParentHT & "</a>, "
                        Next h
                        If OutputFormat = outputType.tHTML And BuildContents Then
                            TextToInsert = TextToInsert & "<a href=""\Contents.html#" & ht & """>Contents</a>"
                        Else 'remove last ", "
                            TextToInsert = Left(TextToInsert, Len(TextToInsert) - 2)
                        End If
                        TextToInsert = TextToInsert & "<p>" & vbLf
                    End If
                End If
                TextToPrepend = BeforeHTML & "<html><head><title>" & ht & "</title></head>" & vbCrLf & BodyTag & vbCrLf '& "<form>" & vbCrLf
                TextToAppend = vbCrLf & "</body>" & vbCrLf & "</html>" & vbCrLf

                If OutputFormat = outputType.tHTMLHELP Then
                    If InStr(mTargetText, "<param name=""Keyword"" value=""" & ht & """>") = 0 Then
                        TextToAppend = KeywordAnchor(ht) & TextToAppend
                    End If
                    If thisHeadingLevel = 5 Then
                        If Right(HeadingText(3), 5) = "Block" Then
                            TextToAppend = AlinkAnchor(ht & Left(HeadingText(3), Len(HeadingText(3)) - 6)) & TextToAppend
                        End If
                    End If
                End If
            End If
            mTargetText = TextToPrepend & Left(mTargetText, thisHeadingStart) & TextToInsert & Mid(mTargetText, thisHeadingStart + 1) & TextToAppend
            thisHeadingStart = thisHeadingStart + Len(TextToPrepend) + Len(TextToInsert)

            'Move <hr> inserted in a table out of the table
            RuleInTable = FindWithinTag(mTargetText, "table", "<hr ")
            If RuleInTable > 0 Then
                RuleEnd = InStr(RuleInTable, mTargetText, ">")
                TableStart = InStrRev(mTargetText, "<table", RuleInTable) 'move first <hr> above table
                mTargetText = Left(mTargetText, TableStart - 1) & Mid(mTargetText, RuleInTable, RuleEnd - RuleInTable + 1) & Mid(mTargetText, TableStart, RuleInTable - TableStart) & Mid(mTargetText, RuleEnd + 1)
                RuleInTable = FindWithinTag(mTargetText, "table", "<hr ", TableStart)
                If RuleInTable > 0 Then 'move second <hr> below table
                    RuleEnd = InStr(RuleInTable, mTargetText, ">")
                    TableEnd = InStr(RuleInTable, mTargetText, "</table") + 8
                    mTargetText = Left(mTargetText, RuleInTable - 1) & Mid(mTargetText, RuleEnd + 1, TableEnd - RuleEnd - 1) & Mid(mTargetText, RuleInTable, RuleEnd - RuleInTable + 1) & Mid(mTargetText, TableEnd + 1)
                End If

            End If

            If BuildContents Then HTMLContentsEntry(thisHeadingLevel, targetFilename, ht)

        End If
    End Sub

    Private Function FindWithinTag(ByVal allText As String, ByRef tag As String, ByVal Find As String, Optional ByVal startAt As Integer = 1) As Integer
        Dim tagStarts As Integer
        Dim tagEnds As Integer
        Dim findPos As Integer
        Dim searchThrough As String

        allText = LCase(allText)
        tag = LCase(tag)
        Find = LCase(Find)

        findPos = InStr(startAt, allText, Find)
        While findPos > 0
            searchThrough = Left(allText, findPos)
            tagStarts = CountString(searchThrough, "<" & tag)
            tagEnds = CountString(searchThrough, "</" & tag)
            If tagStarts > tagEnds Then
                FindWithinTag = findPos
                Exit Function
            End If
            findPos = InStr(findPos + 1, allText, Find)
        End While
    End Function

    Private Sub HTMLContentsEntry(ByRef thisHeadingLevel As Integer, ByRef targetFilename As String, ByRef headerText As String)
        Dim lvl As Integer
        Dim id, SafeFilename As String
        '.Activate ContentsWin
        For lvl = LastHeadingLevel + 1 To thisHeadingLevel
            PrintLine(HTMLContentsfile, Space((lvl - 1) * 4) & "<ul>")
        Next lvl
        For lvl = thisHeadingLevel + 1 To LastHeadingLevel
            PrintLine(HTMLContentsfile, Space((lvl - 1) * 4) & "</ul>")
        Next lvl
        PrintLine(HTMLContentsfile, "<li>")

        Dim objdef As String
        If OutputFormat = outputType.tHTML Then
            SafeFilename = ReplaceString(targetFilename, "\", "/")
            SafeFilename = ReplaceString(SafeFilename, " ", "%20")
            PrintLine(HTMLContentsfile, Space((thisHeadingLevel - 1) * 4) & "<a name=""" & headerText & """>")
            PrintLine(HTMLContentsfile, Space((thisHeadingLevel - 1) * 4) & "<a href=""" & SafeFilename & "#" & headerText & """>" & headerText & "</a></a>")
        ElseIf OutputFormat = outputType.tHTMLHELP Then
            objdef = Space((thisHeadingLevel - 1) * 4) & "<li><OBJECT type=""text/sitemap"">" & vbCr
            objdef = objdef & Space((thisHeadingLevel) * 4) & "<param name=""Name"" value=""" & headerText & """>" & vbCr
            objdef = objdef & Space((thisHeadingLevel) * 4) & "<param name=""Local"" value=""" & targetFilename & """>" & vbCr
            objdef = objdef & Space((thisHeadingLevel) * 4) & "</OBJECT>" & vbCr
            PrintLine(HTMLContentsfile, objdef)
            PrintLine(HTMLIndexfile, objdef)

            id = MakeValidHelpID(headerText)

            If BuildID Then
                PrintLine(IDfile, "#define " & id & vbTab & IDnum)
                IDnum = IDnum + 1
            End If
            If BuildProject Then
                AliasSection = AliasSection & vbLf & id & " = " & targetFilename
            End If
        End If
        'If IconLevel = thisHeadingLevel Then
        '  Dim slashpos%, IconPath$, i&
        '  IconPath = ""
        '  slashpos = 0
        '  For i = 1 To IconLevel - 1
        '    slashpos = InStr(slashpos + 1, targetFilename, "\")
        '  Next i
        '  If slashpos > 0 Then IconPath = Mid(targetFilename, 1, slashpos)
        '
        '  Print #HTMLContentsfile, Space((thisHeadingLevel) * 4) & "<img src=""" & IconPath & IconFilename & """>"
        'End If
        '.Activate SourceWin
    End Sub

    Public Sub FormatHeadingPrint(ByRef thisHeadingLevel As Integer)
        '  With Word
        '  Dim styleName$, ht$
        Dim cmd As Object
        '  styleName = "ADheading" & thisHeadingLevel
        '  ht = HeadingText(thisHeadingLevel)
        With pWordBasic
            .EditBookmark("Hall" & thisHeadingLevel)
            .CharRight()
            .EditBookmark("Hend" & thisHeadingLevel)
            .Insert(vbCr & vbCr)
            .EditGoTo("Hall" & thisHeadingLevel)
            .CharLeft()
            .Insert(vbCr)
            .ExtendSelection()
            .EditGoTo("Hend" & thisHeadingLevel)
            .Style("ADheading" & thisHeadingLevel)
            .Cancel()
        End With
        For Each cmd In WordStyle(thisHeadingLevel)
            'UPGRADE_WARNING: Couldn't resolve default property of object cmd. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            WordCommand(cmd, thisHeadingLevel)
            'If Left(LCase(cmd), 11) = "insertbreak" Then GoSub ApplyStyle
        Next cmd

        '  If thisHeadingLevel = 4 Then
        '    .InsertBreak 2
        '    GoSub ApplyStyle
        '    If MakeBoxyHeaders Then
        '      ViewHeaderAndSet ht & vbTab & vbTab & ht
        '      .FormatTabs ClearAll:=1
        '      .FormatTabs "3.25""", Align:=1, Set:=1 'Center text in center of pg
        '      .FormatTabs "6.5""", Align:=2, Set:=1 'flush right at right margin
        '      .BorderLineStyle 8    'make double-line border around header
        '      .BorderOutside 1
        '      'If IconLevel <= thisHeadingLevel Then
        '      '  .StartOfLine
        '      '  .EditPaste
        '      '  .Insert "   "
        '      'End If
        '      .ViewNormal
        '    End If
        '  ElseIf thisHeadingLevel = 5 Then
        '    .InsertBreak 2
        '    GoSub ApplyStyle
        '    If MakeBoxyHeaders Then
        '      ViewHeaderAndSet HeadingText(4) & vbTab & ht & vbTab & HeadingText(4)
        '      'If IconLevel <= thisHeadingLevel Then
        '      '  .StartOfLine
        '      '  .EditPaste
        '      '  .Insert "   "
        '      'End If
        '      .BorderLineStyle 8    'make double-line border around header
        '      .BorderOutside 1
        '      .ViewNormal
        '    End If
        '  ElseIf thisHeadingLevel > 5 Then
        '    .InsertBreak 3
        '    GoSub ApplyStyle
        '    'If IconLevel = thisHeadingLevel Then
        '    '  .EditGoTo "Hstart" & thisHeadingLevel
        '    '  .EditPaste
        '    '  .Insert "   "
        '    'End If
        '  Else 'If thisHeadingLevel < 4 Then
        '    .InsertBreak 2
        '    GoSub ApplyStyle
        '    'If IconLevel <= thisHeadingLevel Then
        '    '  .EditGoTo "Hstart" & thisHeadingLevel
        '    '  .EditPaste
        '    '  .Insert "   "
        '    'End If
        '    ViewHeaderAndSet ""
        '    .BorderNone 1
        '    .GoToHeaderFooter
        '    GoSub PageNumber
        '  End If
        ' End With
        Exit Sub

        'ApplyStyle:
        '  With Word
        '    .EditBookmark "Hstart" & thisHeadingLevel
        '    .EditGoTo "Hend" & thisHeadingLevel
        '    .Insert vbCr & vbCr
        '    .EditGoTo "Hstart" & thisHeadingLevel
        '    .ExtendSelection
        '    .EditGoTo "Hend" & thisHeadingLevel
        '    .Style StyleName
        '    .Cancel
        '    .CharRight
        '    Return
        '  End With

        'PageNumber:
        ' With Word
        '  On Error Resume Next
        '  .ToggleHeaderFooterLink
        '  On Error GoTo 0
        '  .EditSelectAll
        '  .CenterPara
        '  .Insert ht
        '  If FooterTimestamps Then .InsertDateTime "   hh:mm MMMM d, yyyy", InsertAsField:=0
        '  .ViewNormal
        '  .InsertPageNumbers 1, 4, 1
        '  .ViewFooter
        '  .EditSelectAll
        '  .FormatFont 9, Bold:=1
        '  .ViewNormal
        '  Return
        ' End With

    End Sub

    Private Sub ViewHeaderAndSet(ByRef s As String)
        With pWordBasic
            If Len(s) > 0 Then
                .FilePageSetup(TopMargin:="1.5""", HeaderDistance:="1""", ApplyPropsTo:=0)
            Else
                .FilePageSetup(TopMargin:="1""", HeaderDistance:="0.5""", ApplyPropsTo:=0)
            End If
            pWordApp.ActiveWindow.ActivePane.View.SeekView = WdSeekView.wdSeekCurrentPageHeader
            '.ViewHeader()
            If NotFirstPrintHeader Then
                Try
                    .FormatHeaderFooterLink()
                Catch
                    pWordBasic.ToggleHeaderFooterLink()
                End Try
            Else
                NotFirstPrintHeader = True
            End If
            .EditSelectAll()
            .EditClear()
            .Insert(s)
        End With
    End Sub

    Private Sub ViewFooterAndSet(ByRef s As String)
        With pWordBasic
            '.ViewFooter()
            pWordApp.ActiveWindow.ActivePane.View.SeekView = WdSeekView.wdSeekCurrentPageFooter
            If NotFirstPrintFooter Then
                On Error GoTo toggle
                .FormatHeaderFooterLink()
                On Error GoTo 0
            Else
                NotFirstPrintFooter = True
            End If
            .EditSelectAll()
            .EditClear()
            .Insert(s)
        End With
        Exit Sub
toggle:
        On Error GoTo foo
        pWordBasic.ToggleHeaderFooterLink()
        Resume Next
foo:
        Resume Next
    End Sub

    Sub DefinePrintStyles()
        Dim level As Integer
        With pWordBasic
            For level = 1 To maxLevels
                .FormatStyle("ADheading" & level)
            Next
            .FormatStyle("Normal")
            '    .ViewHeader
            '    .Insert " "
            '    .ViewNormal
            '    'Stop
            '    .ToolsOptionsGeneral Units:=0
            '    .FilePageSetup TopMargin:="1""", BottomMargin:="1""", LeftMargin:="1""", RightMargin:="1""", HeaderDistance:="0.5""", ApplyPropsTo:=4
            '
            '    .FormatStyle "Normal", AddToTemplate:=1, Define:=1
            '    .FormatDefineStyleFont 11, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 10, 1, 1, "Times New Roman", 0, 0, 0, 0
            '    .FormatDefineStylePara Chr$(34), Chr$(34), 0, 0, 0, 0, 0, 1, 0, 0, 0, 0, 0, 1, Chr$(34)
            '    .FormatDefineStyleLang "English (US)", 1
            '    .FormatDefineStyleBorders 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, -1
            '
            '    .FormatStyle "heading1", BasedOn:="Normal", AddToTemplate:=1, Define:=1
            '    .FormatDefineStyleFont 18, Kerning:=1, KerningMin:=14, Bold:=1
            '    .FormatDefineStylePara Before:=12, After:=9, KeepWithNext:=1
            '    .FormatDefineStyleBorders TopBorder:=2, BottomBorder:=2, HorizColor:=1
            '
            '    .FormatStyle "heading2", BasedOn:="Normal", AddToTemplate:=1, Define:=1
            '    .FormatDefineStyleFont 16, Bold:=1
            '    .FormatDefineStylePara Before:=12, After:=3, KeepWithNext:=1
            '
            '    .FormatStyle "heading3", BasedOn:="Normal", AddToTemplate:=1, Define:=1
            '    .FormatDefineStyleFont 14, Underline:=1, Bold:=1
            '    .FormatDefineStylePara Before:=12, After:=3, KeepWithNext:=1
            '
            '    .FormatStyle "heading4", BasedOn:="Normal", AddToTemplate:=1, Define:=1
            '    .FormatDefineStyleFont 12
            '    .FormatDefineStylePara Before:=12, After:=3, KeepWithNext:=1
            '
            '    .FormatStyle "heading5", BasedOn:="Normal", AddToTemplate:=1, Define:=1
            '    .FormatDefineStyleFont 12
            '    .FormatDefineStylePara Before:=12, After:=3, KeepWithNext:=1
            '
            '    .FormatStyle "heading6", BasedOn:="Normal", AddToTemplate:=1, Define:=1
            '    .FormatDefineStyleFont 12
            '    .FormatDefineStylePara Before:=12, After:=3, KeepWithNext:=1
            '
            '    .FormatStyle "heading7", BasedOn:="Normal", AddToTemplate:=1, Define:=1
            '    .FormatDefineStyleFont 12
            '    .FormatDefineStylePara Before:=12, After:=3, KeepWithNext:=1
        End With
    End Sub

    Private Function TrimQuotes(ByRef aString As String) As String
        Dim retval As String = aString.Trim
        If Left(retval, 1) = """" Then retval = Mid(retval, 2)
        If Right(retval, 1) = """" Then retval = Left(retval, Len(retval) - 1)
        Return retval
    End Function

    Sub HTMLQuotedCharsToPrint()
        With pWordBasic
            .EditReplace("<p>", "^p", ReplaceAll:=True)
            .StartOfDocument()
            .EditFind("<", "", 0)
            While .EditFindFound
                .EditClear()
                .EditBookmark("TagStart")
                .EditFind(">")
                If Not .EditFindFound Then Exit Sub
                .EditClear()
                .ExtendSelection()
                .EditGoTo("TagStart")
                .EditClear()
                .Cancel()
                .EditFind("<")
            End While
            .StartOfDocument()
            .EditReplace("&lt;", "<", ReplaceAll:=True)
            .StartOfDocument()
            .EditReplace("&gt;", ">", ReplaceAll:=True)
            .StartOfDocument()
            .EditReplace("&amp;", "&", ReplaceAll:=True)
            .StartOfDocument()
            .EditReplace("&quot;", """", ReplaceAll:=True)
        End With
    End Sub

    Sub HelpFootnotes(ByRef topic As String, ByRef id As String, ByRef Keywords As String)
        With pWordBasic
            .InsertFootnote("#", 1)
            .Insert(id)
            .OtherPane()
            .InsertFootnote("$", 1)
            .Insert(topic)
            .OtherPane()
            .InsertFootnote("K", 1)
            .Insert(Keywords)
            .OtherPane()
            .InsertFootnote("+", 1)
            .Insert("auto")
            .ClosePane()
        End With
    End Sub

    Public Sub FormatHeadingHelp(ByRef thisHeadingLevel As Integer)
        Dim topic As String
        Dim id As String
        Dim h As Integer
        topic = HeadingText(thisHeadingLevel)
        ' If topic = "Graph" Then Stop
        id = MakeValidHelpID(topic)
        Dim parentTopic As String
        With pWordBasic
            .Bold()
            .FontSize(.FontSize + 4)
            .Cancel()
            .CharLeft()
            .Insert(vbCr)
            .InsertBreak(0)
            HelpFootnotes(topic, id, topic)
            .EditBookmark("Hstart" & thisHeadingLevel)
            'If IconLevel <= thisHeadingLevel Then
            '  .EditPaste
            '  .Insert "   "
            'End If
            .EditGoTo("Hend" & thisHeadingLevel)
            .Insert(vbCr & vbCr)
            .CharLeft(2)
            MakeNonscrollingHereBackToPageBreak()
            .CharRight()
            .EditClear()
            .CharRight()
            If BuildContents Then HelpContentsEntry(topic, id, thisHeadingLevel)
            If BuildID Then
                PrintLine(IDfile, "#define " & id & vbTab & IDnum)
                IDnum = IDnum + 1
            End If
            On Error GoTo NoPrevSection
            .EditBookmark("temp")
            If UpNext Then
                If LastHeadingLevel >= thisHeadingLevel Then
                    .EditGoTo("UpFrom" & LastHeadingLevel)
                    .Insert(vbTab & "Next: ")
                    HelpHyperlink(topic, id)
                End If
                For h = thisHeadingLevel To LastHeadingLevel - 1
                    .EditGoTo("UpFrom" & h)
                    .Insert(vbTab & "Next: ")
                    HelpHyperlink(topic, id)
                Next h
            End If
NoPrevSection:
            .EditGoTo("temp")
            On Error GoTo 0

            If thisHeadingLevel > 1 Then 'Insert navigation to/from parents in hierarchy
                If UpNext Then
                    .Insert("Up to: ")
                    For h = thisHeadingLevel - 1 To 1 Step -1
                        parentTopic = HeadingText(h)
                        HelpHyperlink(parentTopic, MakeValidHelpID(parentTopic))
                        If h > 1 Then .Insert(", ")
                    Next h
                    .Insert(vbCr)
                    .EditBookmark("SectionContents" & thisHeadingLevel)
                    .CharLeft()
                    .EditBookmark("UpFrom" & thisHeadingLevel)

                    'insert entry in section contents of parent topic
                    If HeadingText(thisHeadingLevel - 1) <> "Tutorial" Then
                        ContentsEntries(thisHeadingLevel - 1) = ContentsEntries(thisHeadingLevel - 1) + 1
                        .EditGoTo("SectionContents" & thisHeadingLevel - 1)
                        If ContentsEntries(thisHeadingLevel - 1) = 1 Then
                            .Insert("In This Section:" & vbCr)
                            .CharLeft()
                        End If
                        .Insert(Chr(11) & vbTab)
                        HelpHyperlink(topic, id)
                        .EditBookmark("SectionContents" & thisHeadingLevel - 1)
                        .EditGoTo("SectionContents" & thisHeadingLevel)
                    End If
                End If
            Else
                .EditBookmark("UpFrom" & thisHeadingLevel)
                .Insert(vbCr)
                .EditBookmark("SectionContents" & thisHeadingLevel)
            End If
            ContentsEntries(thisHeadingLevel) = 0
        End With
    End Sub

    Sub FinishHTMLHelpContents()
        Dim lvl As Integer
        For lvl = HeadingLevel To 1 Step -1
            PrintLine(HTMLContentsfile, Space((lvl - 1) * 4) & "</ul>")
        Next lvl
        PrintLine(HTMLContentsfile, "</body>")
        PrintLine(HTMLContentsfile, "</html>")
        FileClose(HTMLContentsfile)

        If HTMLIndexfile >= 0 Then
            PrintLine(HTMLIndexfile, "</ul>")
            PrintLine(HTMLIndexfile, "</body>")
            PrintLine(HTMLIndexfile, "</html>")
            FileClose(HTMLIndexfile)
        End If
    End Sub

    Sub MakeNonscrollingHereBackToPageBreak()
        'Make non-scrolling region at top of help topic
        With pWordBasic
            '        .EditBookmark "NonscrollEnd"
            .ExtendSelection()
            .EditFind("^m", "", 1)
            .LineDown()
            .StartOfLine()
            .ParaKeepWithNext()
            .Cancel()
            .CharRight()
        End With
    End Sub

    Sub HelpHyperlink(ByRef label As Object, ByRef target As Object)
        With pWordBasic
            'DebugMsg "label: " & label & ", target: " & target
            .Insert("   ")
            .CharLeft(2)
            .EditBookmark("link1")
            .DoubleUnderline(1)
            'UPGRADE_WARNING: Couldn't resolve default property of object label. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            .Insert(label)
            .CharRight()
            .EditBookmark("link2")
            .DoubleUnderline(0)
            .Hidden(1)
            'UPGRADE_WARNING: Couldn't resolve default property of object target. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            .Insert(target)
            .CharRight()
            .Hidden(0)
            .EditBookmark("endlink")
            .EditGoTo("link1")
            .EditClear(-1)
            .EditGoTo("link2")
            .EditClear(-1)
            .EditGoTo("endlink")
        End With
    End Sub

    Public Function MakeValidHelpID(ByRef id As String) As String
        Dim a, i As Integer
        Dim retval, ch As String
        Dim lastReplaced As Boolean
        retval = ""
        lastReplaced = True
        For i = 1 To Len(id)
            ch = Mid(id, i, 1)
            a = Asc(ch)
            Select Case a
                Case 65 To 90, 97 To 122, 47 'in range A-Z a-z or /
                    retval &= ch
                    lastReplaced = False
                Case Else
                    If lastReplaced = False Then
                        retval &= "_"
                        lastReplaced = True
                    End If
            End Select
        Next i
        MakeValidHelpID = retval
    End Function

    Public Function MakeValidFilename(ByVal id As String) As String
        Dim a, i As Integer
        Dim retval, ch As String
        Dim lastReplaced As Boolean
        retval = ""
        id = Trim(id)
        lastReplaced = True ' Don't replace multiple illegal chars in a row with underscore
        For i = 1 To Len(id)
            ch = Mid(id, i, 1)
            a = Asc(ch)
            Select Case a
                Case 32, 33, 35 To 41, 43 To 46, 48 To 57, 65 To 90, 94, 95, 97 To 122
                    retval &= ch
                    lastReplaced = False
                Case 0 - 31
                    'Omit control characters and don't insert underscores for them
                Case Else
                    If lastReplaced = False Then
                        retval &= "_"
                        lastReplaced = True
                    End If
            End Select
        Next i
        MakeValidFilename = retval
    End Function

    Sub HelpContentsEntry(ByRef topic As String, ByRef id As String, ByRef thisHeadingLevel As Integer)
        Dim numlines As Integer
        Dim tmpstr As String
        With pWordBasic
            .Activate(ContentsWin)
            If LastHeadingLevel = 0 Then LastHeadingLevel = 1
            .Insert(thisHeadingLevel & " " & topic & "=" & id & vbCr)
            If thisHeadingLevel < LastHeadingLevel Then BookLevel = thisHeadingLevel
            If thisHeadingLevel > LastHeadingLevel And BookLevel < LastHeadingLevel Or BookLevel = thisHeadingLevel Then
                If thisHeadingLevel > LastHeadingLevel Then
                    numlines = 2
                    BookLevel = LastHeadingLevel
                Else
                    numlines = 1
                End If
                .LineUp(numlines)
                .EditFind("=", "", 0)
                .CharLeft()
                .StartOfLine(1)
                .CharRight(1, 1)

                tmpstr = .Selection
                .Cancel()
                .CharLeft()
                .Insert(tmpstr)
                'Before copy buffer was used for icons, the last 4 lines were:
                '.EditCopy
                '.Cancel
                '.CharLeft
                '.EditPaste

                .Insert(vbCr & thisHeadingLevel)
                .LineDown(numlines)
            End If
            'Dim lvl%
            'For lvl = LastHeadingLevel + 1 To thisHeadingLevel - 1
            '  .Insert lvl & " missing level" & vbCr
            'Next lvl
            .Activate(TargetWin)
        End With
    End Sub

    Private Sub HREFsToHelpHyperlinks()
        Dim topic, LinkRef, LinkLabel, id As String
        With pWordBasic
            .StartOfDocument()
            Status("Translating HTML links to Help Hyperlinks")
            .EditFind("<A HREF=""", "", 0)
            While .EditFindFound
                'set LinkRef$
                .EditClear()
                .EditBookmark("LinkStart")
                .EditFind(""">")
                If Not .EditFindFound Then Exit Sub
                .EditClear()
                .ExtendSelection()
                .EditGoTo("LinkStart")
                LinkRef = .Selection
                .EditClear()
                .Cancel()

                'set LinkLabel$
                .EditBookmark("LinkStart")
                .EditFind("</A>")
                If Not .EditFindFound Then Exit Sub
                .EditClear()
                .ExtendSelection()
                .EditGoTo("LinkStart")
                LinkLabel = .Selection
                .EditClear()
                .Cancel()

                'UPGRADE_WARNING: Couldn't resolve default property of object pos. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Dim pos As Object = InStr(1, LinkRef, "#")
                'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
                If IsDBNull(pos) Then
                    Status("Warning: Link '" & LinkLabel & " -> " & LinkRef & "' does not contain valid help topic.")
                    'UPGRADE_WARNING: Couldn't resolve default property of object pos. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                ElseIf pos = 0 Then
                    Status("Warning: Link '" & LinkLabel & " -> " & LinkRef & "' does not contain valid help topic.")
                Else
                    'UPGRADE_WARNING: Couldn't resolve default property of object pos. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    topic = Mid(LinkRef, pos + 1)
                    id = MakeValidHelpID(topic)
                    HelpHyperlink(LinkLabel, id)
                End If

                .EditFind("<A HREF=""")
            End While
        End With
    End Sub

    Private Sub AbsoluteToRelative()
        Dim startTag, LevelCount As Integer
        Dim FilePath As String
        startTag = InStr(mTargetText, "=""/")
        While startTag > 0
            FilePath = ""
            For LevelCount = 1 To HeadingLevel - 1
                FilePath = FilePath & "../"
            Next
            mTargetText = Left(mTargetText, startTag + 1) & FilePath & Mid(mTargetText, startTag + 3)
            startTag = InStr(mTargetText, "=""/")
        End While
    End Sub

    Private Sub MakeLocalTOCs()
        Dim startTag As Integer
        startTag = InStr(LCase(mTargetText), "<toc>")
        If startTag > 0 Then
            mTargetText = Left(mTargetText, startTag - 1) & SectionContents & Mid(mTargetText, startTag + 5)
            If InStr(LCase(mTargetText), "<toc>") Then Logger.Msg("More than one <toc> in '" & mSourceFilename & "'")
        End If
    End Sub

    Private Function SectionContents() As String
        Dim retval As String = ""
        Dim localNextEntry As Integer
        Dim nextLevel, lvl, prevLevel As Integer
        Dim nextName As String = ""
        Dim nextHref As String = ""
        Dim localHeadingWord(10) As String

        localNextEntry = mNextProjectFileEntry
        localHeadingWord(HeadingLevel) = HeadingWord(HeadingLevel)
        prevLevel = HeadingLevel
        GetNextEntryLevel(localNextEntry, nextLevel, nextName, nextHref, localHeadingWord)
        If nextLevel > HeadingLevel Then
            While nextLevel > HeadingLevel

                If nextLevel > prevLevel Then
                    For lvl = prevLevel To (nextLevel - 1)
                        retval &= "<ul>" & vbCr
                    Next
                ElseIf nextLevel < prevLevel Then
                    For lvl = nextLevel To (prevLevel - 1)
                        retval &= "</ul>" & vbCr
                    Next
                End If

                retval &= "<li><a href=""" & nextHref & """>" & nextName & "</a>" & vbCr
                prevLevel = nextLevel
                GetNextEntryLevel(localNextEntry, nextLevel, nextName, nextHref, localHeadingWord)
            End While
            retval &= "</ul>" & vbCr
        End If
        SectionContents = retval
    End Function

    Private Sub GetNextEntryLevel(ByRef localNextEntry As Integer, _
                                  ByRef nextLevel As Integer, _
                                  ByRef nextName As String, _
                                  ByRef nextHref As String, _
                                  ByRef localHeadingWord() As String)
        If localNextEntry >= mProjectFileEntrys.Count Then
            nextLevel = 0
        Else
            nextName = LTrim(mProjectFileEntrys(localNextEntry))
            nextLevel = (Len(mProjectFileEntrys(localNextEntry)) - Len(nextName)) / 2 + 1
            nextName = RTrim(nextName)
            localHeadingWord(nextLevel) = nextName
            nextHref = ""
            For lvl As Integer = HeadingLevel To nextLevel - 1
                nextHref = nextHref & localHeadingWord(lvl) & "\"
            Next
            nextHref = nextHref & nextName & ".html"
            localNextEntry = localNextEntry + 1
        End If
    End Sub

    Private Sub CopyImages()
        Dim endPos, lStartPos As Integer
        Dim SrcPath, ImageFilename, DstPath As String
        Dim HTMLsafeFilename As String

        If OutputFormat = OutputType.tHELP OrElse OutputFormat = OutputType.tPRINT Then
            Exit Sub
        End If
        Status("Copying Images")
        SrcPath = IO.Path.GetDirectoryName(mSourceBaseDirectory & mSourceFilename) & "\"
        DstPath = IO.Path.GetDirectoryName(mSaveDirectory & SaveFilename) & "\"
        Dim lIgnoreAll As Boolean = False
        While Assign(lStartPos, mTargetText.IndexOf(" src=""", StringComparison.OrdinalIgnoreCase)) > 0
            endPos = mTargetText.IndexOf("""", lStartPos + 6)
            If endPos = 0 Then Exit Sub
            ImageFilename = Mid(mTargetText, lStartPos + 6, endPos - lStartPos - 6)
CheckForImage:
            If IO.File.Exists(SrcPath & ImageFilename) Then
                MkDirPath(IO.Path.GetDirectoryName(AbsolutePath(ReplaceString(ImageFilename, "/", "\"), DstPath)))
                FileCopy(SrcPath & ImageFilename, DstPath & ImageFilename)
            ElseIf Not lIgnoreAll Then
                Select Case Logger.Msg("Missing image: " & vbCr & SrcPath & ImageFilename, MsgBoxStyle.AbortRetryIgnore, "AuthorDoc")
                    Case MsgBoxResult.Abort : Exit Sub
                    Case MsgBoxResult.Retry : GoTo CheckForImage
                    Case MsgBoxResult.Ignore
                        If Logger.Msg("Ignore all missing images?", MsgBoxStyle.YesNo, "AuthorDoc") = MsgBoxResult.Yes Then
                            lIgnoreAll = True
                        End If
                End Select
            End If

            If OutputFormat = OutputType.tHTML Then
                HTMLsafeFilename = ReplaceString(ImageFilename, "\", "/")
                HTMLsafeFilename = ReplaceString(HTMLsafeFilename, " ", "%20")
                If HTMLsafeFilename <> ImageFilename Then
                    mTargetText = Left(mTargetText, lStartPos + 5) & HTMLsafeFilename & Mid(mTargetText, endPos)
                End If
            End If
        End While
    End Sub

    Private Sub HREFsInsureExtension()
        Dim LinkFile, LinkRef, LinkTopic As String
        Dim endPos, startPos, pos As Integer

        If OutputFormat = OutputType.tHELP Then
            HREFsInsureExtensionWithWord()
        ElseIf OutputFormat = OutputType.tPRINT Then
            'We don't preserve links in printable, so skip this step
        Else
            Status("HREFsInsureExtension")
            startPos = InStr(LCase(mTargetText), "<a href=""")
            While startPos > 0
                endPos = InStr(startPos + 9, mTargetText, """")
                If endPos = 0 Then Exit Sub
                LinkRef = Mid(mTargetText, startPos + 9, endPos - startPos - 9)
                pos = InStr(LinkRef, "#")
                If pos = 0 Then
                    LinkFile = LinkRef
                    LinkTopic = ""
                Else
                    LinkFile = Left(LinkRef, pos - 1)
                    LinkTopic = Mid(LinkRef, pos)
                End If
                If Len(LinkFile) > 0 Then
                    If InStr(LCase(LinkFile), ".html") < 1 And InStr(LinkFile, ":") < 1 Then
                        If LCase(Right(LinkFile, 4)) = ".txt" Then
                            LinkFile = Left(LinkFile, Len(LinkFile) - 3) & "html"
                        Else
                            LinkFile = LinkFile & ".html"
                        End If
                    End If
                    If OutputFormat = OutputType.tHTML Then
                        LinkFile = ReplaceString(LinkFile, "\", "/")
                        LinkFile = ReplaceString(LinkFile, " ", "%20")
                    End If
                End If
                If LinkFile & LinkTopic <> LinkRef Then
                    mTargetText = Left(mTargetText, startPos + 8) & LinkFile & LinkTopic & Mid(mTargetText, endPos)
                End If
                startPos = InStr(endPos, LCase(mTargetText), "<a href=""")
            End While
        End If
    End Sub

    Private Sub HREFsInsureExtensionWithWord()
        Dim LinkRef As String
        Dim pos As Integer
        With pWordBasic
            .StartOfDocument()
            .Cancel()
            Status("HREFsInsureExtension")
            .EditFind("<A HREF=""", "", 0)
            While .EditFindFound
                'set LinkRef$
                .CharRight() '.EditClear
                .EditBookmark("LinkStart")
                .EditFind(""">")
                If Not .EditFindFound Then Exit Sub
                .CharLeft() '.EditClear
                .ExtendSelection()
                .EditGoTo("LinkStart")
                LinkRef = .Selection
                .Cancel()
                .CharRight()
                If InStr(LCase(LinkRef), ".html") < 1 Then
                    pos = InStr(1, LinkRef, "#")
                    If pos > 0 Then .CharLeft(Len(LinkRef) - pos + 1)
                    .ExtendSelection()
                    .CharLeft(4)
                    If LCase(.Selection) = ".txt" Then
                        .EditClear()
                    Else
                        .Cancel()
                        .CharRight()
                    End If
                    .Insert(".html")
                End If

                .EditFind("<A HREF=""")
            End While
        End With
    End Sub

    Private Sub Status(ByRef aMessage As String)
        Logger.Dbg(aMessage)
        frmConvert.Text1.Text = aMessage
        System.Windows.Forms.Application.DoEvents()
    End Sub

    'Public Sub FormatCardGraphic()
    '    Dim startPos, endPos As Integer
    '    Dim ImageFilename, TableText, ImageDirectory As String
    '    Dim ImageMap As String
    '    startPos = InStr(TargetText, WholeCardHeader)
    '    If startPos > 0 Then
    '        endPos = InStrRev(TargetText, Asterisks80)
    '        If endPos > 0 Then
    '            ImageDirectory = IO.Path.GetDirectoryName(SaveDirectory & SaveFilename)
    '            ImageFilename = FilenameOnly(SaveFilename) & ".bmp"
    '            TableText = Mid(TargetText, startPos + lenWholeCardHeader, endPos - startPos - lenWholeCardHeader)
    '            ImageMap = CardImage(TableText)
    '            If Len(ImageMap) > 0 Then
    '                TargetText = "<map name=""CardImageMap"">" & vbCrLf & ImageMap & "</map>" & vbCrLf & Left(TargetText, startPos - 1) & "<p>" & vbCrLf & "<img src=""" & ImageFilename & """ usemap=""#CardImageMap"" border=0>" & vbCrLf & "<p>" & Mid(TargetText, endPos + 81)
    '            Else
    '                TargetText = Left(TargetText, startPos - 1) & "<p>" & vbCrLf & "<img src=""" & ImageFilename & """>" & vbCrLf & "<p>" & Mid(TargetText, endPos + 81)
    '            End If
    '            'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
    '            If Len(Dir(ImageDirectory, FileAttribute.Directory)) = 0 Then MkDir(ImageDirectory)
    '            'UPGRADE_WARNING: SavePicture was upgraded to System.Drawing.Image.Save and has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
    '            frmSample.img.Image.Save(ImageDirectory & "\" & ImageFilename)
    '        End If
    '    End If
    'End Sub

    'Creates image on frmSample.img and returns HTML map for links
    '    Public Function CardImage(ByRef TableText As String) As String
    '        Dim TextRow(255) As String
    '        Dim RowY(255) As Integer
    '        Dim TensY As Integer
    '        Dim OnesY As Integer
    '        Dim Row, Rows As Integer
    '        Dim col, lentxt As Integer
    '        Dim lastCR, thisCR As Integer
    '        Dim GrayColor As System.Drawing.Color
    '        Dim curChar As String
    '        Dim CharWidth As Integer
    '        Dim XMargin As Integer
    '        Dim CharHeight As Integer
    '        Dim txt As String
    '        Dim GrayLevel As Integer
    '        Dim RangeExists As Boolean
    '        Dim RowStopChecking As Integer
    '        Dim SubsectionName As String
    '        Dim retval As String
    '        Dim StartLinkCol, parsePos, StopLinkCol As Integer
    '        Dim SaveFileNameOnly As String
    '        Dim DoLinks As Boolean

    '        DoLinks = True

    '        SaveFileNameOnly = FilenameOnly(SaveFilename)
    '        retval = ""
    '        GrayLevel = 170
    '        txt = ReplaceString(TableText, "&gt;", ">")
    '        txt = ReplaceString(txt, "&lt;", "<")
    '        GrayColor = System.Drawing.ColorTranslator.FromOle(RGB(GrayLevel, GrayLevel, GrayLevel))

    '        lentxt = Len(txt)
    '        thisCR = 0
    '        Rows = 0
    '        RowStopChecking = 256
    '        FindCR(lastCR, thisCR, lentxt, txt)
    '        While lastCR <= lentxt
    '            Rows = Rows + 1
    '            TextRow(Rows) = Mid(txt, lastCR + 1, thisCR - lastCR - 1)
    '            If Right(TextRow(Rows), 1) = vbCr Then TextRow(Rows) = Left(TextRow(Rows), Len(TextRow(Rows)) - 1)
    '            Select Case TextRow(Rows)
    '                Case SixSplats, SevenSplats, Asterisks80, TensPlace, OnesPlace 'Skip some rows
    '                    Rows = Rows - 1
    '                Case "Example"
    '                    If RowStopChecking > 0 Then RowStopChecking = Rows
    '                Case "<otyp>"
    '                    DoLinks = False
    '                Case "SPEC-ACTIONS"
    '                    DoLinks = False
    '                    RowStopChecking = 0
    '                Case Else 'Split long rows
    '                    While Len(TextRow(Rows)) > MaxRowLength
    '                        Rows = Rows + 1
    '                        TextRow(Rows) = ""
    '                        col = 1
    '                        While Mid(TextRow(Rows - 1), col, 1) = " "
    '                            TextRow(Rows) = TextRow(Rows) & " "
    '                            col = col + 1
    '                        End While
    '                        TextRow(Rows) = TextRow(Rows) & Mid(TextRow(Rows - 1), MaxRowLength + 1)
    '                        TextRow(Rows - 1) = Left(TextRow(Rows - 1), MaxRowLength)
    '                    End While
    '            End Select
    '            FindCR(lastCR, thisCR, lentxt, txt)
    '        End While

    '        frmSample.Visible = True
    '        With frmSample.img
    '            'UPGRADE_ISSUE: PictureBox method img.TextWidth was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
    '            CharWidth = .TextWidth("X")
    '            XMargin = CharWidth / 2
    '            'UPGRADE_ISSUE: PictureBox method img.TextHeight was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
    '            CharHeight = .TextHeight("X")
    '            .Width = CharWidth * MaxRowLength + XMargin * 2
    '            .Height = CharHeight * 10 'Start with enough height for header, adjust again after header
    '            'UPGRADE_ISSUE: PictureBox method img.Cls was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
    '            .Cls()
    '            'frmSample.img.Line (0, 0)-(.Width, 0), vbBlack

    '            'Print tens places in gray
    '            .ForeColor = GrayColor
    '            'UPGRADE_ISSUE: PictureBox property img.CurrentY was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
    '            .CurrentY = CharHeight / 2
    '            'UPGRADE_ISSUE: PictureBox property img.CurrentY was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
    '            TensY = .CurrentY
    '            For col = 1 To 8
    '                curChar = CStr(col)
    '                'UPGRADE_ISSUE: PictureBox method img.TextWidth was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
    '                'UPGRADE_ISSUE: PictureBox property img.CurrentX was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
    '                .CurrentX = XMargin + (col * 10 - 1) * CharWidth + (CharWidth - .TextWidth(curChar)) / 2
    '                'UPGRADE_ISSUE: PictureBox method img.Print was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
    '                frmSample.img.Print(curChar)
    '            Next
    '            'UPGRADE_ISSUE: PictureBox method img.Print was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
    '            frmSample.img.Print()
    '            'Print Ones Place in gray
    '            'UPGRADE_ISSUE: PictureBox property img.CurrentY was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
    '            OnesY = .CurrentY
    '            For col = 1 To 80
    '                curChar = CStr(col Mod 10)
    '                'UPGRADE_ISSUE: PictureBox method img.TextWidth was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
    '                'UPGRADE_ISSUE: PictureBox property img.CurrentX was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
    '                .CurrentX = XMargin + (col - 1) * CharWidth + (CharWidth - .TextWidth(curChar)) / 2
    '                'UPGRADE_ISSUE: PictureBox method img.Print was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
    '                frmSample.img.Print(curChar)
    '            Next
    '            'UPGRADE_ISSUE: PictureBox method img.Print was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
    '            frmSample.img.Print()
    '            'UPGRADE_ISSUE: PictureBox property img.CurrentY was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
    '            RowY(0) = .CurrentY
    '            .Height = RowY(0) + (OnesY - TensY) * Rows + TensY
    '            .ForeColor = System.Drawing.Color.Black
    '            If InStr(txt, "<-range>") > 0 Then
    '                RangeExists = True
    '                col = 5
    '                'UPGRADE_ISSUE: PictureBox method img.Line was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
    '				frmSample.img.Line (XMargin + col * CharWidth, 0) - (0, .Height), GrayColor
    '                'UPGRADE_ISSUE: GoSub statement is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C5A1A479-AB8B-4D40-AAF4-DB19A2E5E77F"'
    '				GoSub NumberCol
    '                col = 10
    '                'UPGRADE_ISSUE: PictureBox method img.Line was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
    '				frmSample.img.Line (XMargin + col * CharWidth, 0) - (0, .Height), GrayColor
    '                'UPGRADE_ISSUE: GoSub statement is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C5A1A479-AB8B-4D40-AAF4-DB19A2E5E77F"'
    '				GoSub NumberCol
    '            Else
    '                RangeExists = False
    '            End If
    '            For Row = 1 To Rows
    '                RowY(Row) = RowY(Row - 1) + OnesY - TensY

    '                StartLinkCol = 0
    '                StopLinkCol = 0

    '                If LCase(Mid(TextRow(Row), 3, lenTableType)) = LCase(TableType) Then
    '                    Mid(TextRow(Row), 3, lenTableType) = TableType
    '                    StartLinkCol = lenTableType + 3
    '                    'UPGRADE_ISSUE: GoSub statement is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C5A1A479-AB8B-4D40-AAF4-DB19A2E5E77F"'
    '					GoSub StartArea
    '                End If

    '                If LCase(Mid(TextRow(Row), 3, 13)) = "general input" Then
    '                    Mid(TextRow(Row), 3, 7) = "General input"
    '                    StartLinkCol = 3
    '                    'UPGRADE_ISSUE: GoSub statement is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C5A1A479-AB8B-4D40-AAF4-DB19A2E5E77F"'
    '					GoSub StartArea
    '                End If

    '                If LCase(Mid(TextRow(Row), 3, 7)) = "section" Then
    '                    Mid(TextRow(Row), 3, 7) = "Section"
    '                    StartLinkCol = 11
    '                    'UPGRADE_ISSUE: GoSub statement is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C5A1A479-AB8B-4D40-AAF4-DB19A2E5E77F"'
    '					GoSub StartArea
    '                End If

    '                If Not DoLinks Then StartLinkCol = 0

    '                For col = 1 To Len(TextRow(Row))
    '                    curChar = Mid(TextRow(Row), col, 1)
    '                    'UPGRADE_ISSUE: PictureBox method img.TextWidth was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
    '                    'UPGRADE_ISSUE: PictureBox property img.CurrentX was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
    '                    .CurrentX = XMargin + (col - 1) * CharWidth + (CharWidth - .TextWidth(curChar)) / 2
    '                    'UPGRADE_ISSUE: PictureBox property img.CurrentY was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
    '                    .CurrentY = RowY(Row)

    '                    If col = StartLinkCol Then .ForeColor = System.Drawing.Color.Blue
    '                    'UPGRADE_ISSUE: GoSub statement is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C5A1A479-AB8B-4D40-AAF4-DB19A2E5E77F"'
    '					If col = StopLinkCol Then GoSub EndTableLink

    '                    'UPGRADE_ISSUE: PictureBox method img.Print was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
    '                    frmSample.img.Print(curChar)

    '                    If (Not RangeExists Or col > 10) And Row < RowStopChecking Then
    '                        Select Case curChar
    '                            Case "<"
    '                                'UPGRADE_ISSUE: PictureBox method img.Line was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
    '								frmSample.img.Line (XMargin + (col - 1) * CharWidth, 0) - (0, .Height), GrayColor
    '                                'GoSub NumberCol
    '                            Case ">"
    '                                'UPGRADE_ISSUE: PictureBox method img.Line was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
    '								frmSample.img.Line (XMargin + col * CharWidth, 0) - (0, .Height), GrayColor
    '                                'UPGRADE_ISSUE: GoSub statement is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C5A1A479-AB8B-4D40-AAF4-DB19A2E5E77F"'
    '								GoSub NumberCol
    '                        End Select
    '                    End If
    '                Next
    '                'UPGRADE_ISSUE: GoSub statement is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C5A1A479-AB8B-4D40-AAF4-DB19A2E5E77F"'
    '				If System.Drawing.ColorTranslator.ToOle(.ForeColor) <> System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black) Then GoSub EndTableLink
    '                'UPGRADE_ISSUE: PictureBox method img.Print was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
    '                frmSample.img.Print()
    '                'frmSample.img.Line (col * CharWidth, 0)-((col + 1) * CharWidth, .Height), RGB(222, 222, 222), BF
    '            Next
    '            'Clipboard.SetData .Image
    '            If frmSample.WindowState <> System.Windows.Forms.FormWindowState.Normal Then frmSample.WindowState = System.Windows.Forms.FormWindowState.Normal
    '            frmSample.SetBounds(frmSample.Left, frmSample.Top, VB6.TwipsToPixelsX((.Width * VB6.TwipsPerPixelX) + 108), VB6.TwipsToPixelsY((.Height * VB6.TwipsPerPixelY) + 372))
    '            CardImage = retval

    '            Exit Function

    'StartArea:
    '            If DoLinks Then
    '                'UPGRADE_ISSUE: PictureBox property img.CurrentY was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
    '                'UPGRADE_ISSUE: PictureBox property img.CurrentX was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
    '                retval &= "<area coords=""" & .CurrentX & "," & .CurrentY
    '                StopLinkCol = InStr(StartLinkCol, TextRow(Row), "]")
    '                If StopLinkCol > 0 Then 'Ignore tables containing the string "Tables in brackets  are ..."
    '                    If Mid(TextRow(Row), StopLinkCol - 1, 1) = "[" Then StopLinkCol = 0
    '                End If
    '                If StopLinkCol = 0 Then
    '                    StopLinkCol = InStr(StartLinkCol, TextRow(Row), "  ")
    '                    If StopLinkCol = 0 Then
    '                        StopLinkCol = 999
    '                    End If
    '                End If
    '            End If
    '            'UPGRADE_WARNING: Return has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
    '            Return

    'EndTableLink:
    '            If DoLinks Then
    '                .ForeColor = System.Drawing.Color.Black
    '                If curChar = "]" Or curChar = " " Then
    '                    SubsectionName = Trim(Mid(TextRow(Row), 3, col - 3))
    '                Else
    '                    SubsectionName = Trim(Mid(TextRow(Row), 3))
    '                End If
    '                If Left(SubsectionName, 7) = "Section" Then SubsectionName = Mid(SubsectionName, 9)
    '                If Left(SubsectionName, lenTableType) = TableType Then SubsectionName = Mid(SubsectionName, lenTableType + 1)
    '                'UPGRADE_ISSUE: PictureBox property img.CurrentY was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
    '                'UPGRADE_ISSUE: PictureBox property img.CurrentX was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
    '                retval &= "," & .CurrentX & "," & .CurrentY + CharHeight & """ href="""

    '                'Several special cases for sections that are not where they are expected
    '                Select Case Left(SubsectionName, 9)
    '                    Case "SOIL-DATA", "CROP-DATE"
    '                        retval &= "PWATER Input"
    '                    Case "OXRX inpu", "NUTRX inp", "PLANK inp", "PHCARB in"
    '                        retval &= SaveFileNameOnly
    '                        If SaveFileNameOnly <> "Input for RQUAL sections" Then
    '                            retval &= "/Input for RQUAL sections"
    '                        End If
    '                    Case "SURF-EXPO"
    '                        If SaveFileNameOnly = "GQUAL input" Then
    '                            retval &= "Input for RQUAL sections/PLANK input"
    '                        Else
    '                            retval &= SaveFileNameOnly
    '                        End If
    '                    Case Else
    '                        If SaveFileNameOnly <> "OXRX input" And (SubsectionName = "ELEV" Or Left(SubsectionName, 3) = "OX-") Then
    '                            retval &= "Input for RQUAL sections/OXRX input"
    '                        Else 'Most links do not need tweaking and fall through to here
    '                            retval &= SaveFileNameOnly
    '                        End If
    '                End Select
    '                retval &= "/" & SubsectionName & ".html"">" & vbCrLf
    '            End If
    '            'UPGRADE_WARNING: Return has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
    '            Return

    'NumberCol:
    '            If col > 9 Then 'And curChar <> "<" Then
    '                curChar = CStr(Int(col / 10))
    '                'UPGRADE_ISSUE: PictureBox method img.TextWidth was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
    '                'UPGRADE_ISSUE: PictureBox property img.CurrentX was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
    '                .CurrentX = XMargin + (col - 1) * CharWidth + (CharWidth - .TextWidth(curChar)) / 2
    '                'UPGRADE_ISSUE: PictureBox property img.CurrentY was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
    '                .CurrentY = TensY
    '                'UPGRADE_ISSUE: PictureBox method img.Print was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
    '                frmSample.img.Print(curChar)
    '            End If
    '            curChar = CStr(col Mod 10)
    '            'UPGRADE_ISSUE: PictureBox method img.TextWidth was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
    '            'UPGRADE_ISSUE: PictureBox property img.CurrentX was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
    '            .CurrentX = XMargin + (col - 1) * CharWidth + (CharWidth - .TextWidth(curChar)) / 2
    '            'UPGRADE_ISSUE: PictureBox property img.CurrentY was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
    '            .CurrentY = OnesY
    '            'UPGRADE_ISSUE: PictureBox method img.Print was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
    '            frmSample.img.Print(curChar)
    '            'UPGRADE_WARNING: Return has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
    '            Return
    '        End With
    '    End Function

    Private Sub FindCR(ByRef lastCR As Integer, ByRef thisCR As Integer, ByVal lentxt As Integer, ByVal txt As String)
        lastCR = thisCR
        If lastCR < lentxt Then
            If Mid(txt, lastCR + 1, 1) = vbLf Then lastCR = lastCR + 1
        End If

        thisCR = InStr(thisCR + 1, txt, vbLf)
        If thisCR = 0 Then thisCR = lentxt + 1
    End Sub

    'Public Sub PictureString(buf As String)
    '  Dim col As Long, maxcol As Long, curChar As String
    '  maxcol = Len(buf)
    '  If maxcol > Cols Then maxcol = Cols
    '  With frmSample.img
    '    .CurrentY = Row * CharHeight
    '    For col = 1 To maxcol
    '      curChar = Mid(buf, col, 1)
    '      .CurrentX = (col - 1) * CharWidth + (CharWidth - .TextWidth(curChar)) / 2
    '      frmSample.img.Print curChar;
    '    Next
    '  End With
    '  Row = Row + 1
    '  If Len(buf) > CharWidth Then PictureString Mid(buf, CharWidth + 1)
    'End Sub
End Module