Option Strict Off
Option Explicit On

Imports atcUtility
Imports MapWinUtility
Imports Microsoft.Office.Interop.Word
Imports System.Text

Module modConvert
    'Copyright 2000-2008 by AQUA TERRA Consultants
	
    'pBaseName (~) is the name of program being documented.
    'File ~.txt contains list of source files (pProjectFileName)
	'~.hlp will be created if converting to help (also optionally ~.cnt, ~.hpj)
	'~.doc will be created if converting to printable
	'~.hhp, ~.hhc, ~.ID -> ~.chm
	
    Public pOutputFormat As outputType

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

    Private Const mMaxLevels As Integer = 9 ' Do you really want sections nested deeper than this?

    Private mContentsWindowName As String
    Private mTargetWindowName As String

    Private mTargetText As String

    Private mSourceFilename As String

    Private mSourceBaseDirectory As String
    Private mSaveDirectory As String

    Private mHelpSourceRTFName As String

    Private mProjectFileEntrys As New Collection
    Private mNextProjectFileIndex As Integer

    Private mHeadingWord(8) As String
    Private mHeadingText(mMaxLevels) As String
    Private mHeadingFile(mMaxLevels) As String

    Private mBeforeHTML As String

    Private mContentsEntries(mMaxLevels) As Integer
    Private mHeaderStyle(mMaxLevels) As String
    Private mFooterStyle(mMaxLevels) As String
    Private mBodyStyle(mMaxLevels) As String
    Private mWordStyle(mMaxLevels) As Collection
    Private mBodyTag As String
    Private mStyleFile(mMaxLevels) As String
    'Private mPromptForFiles As Boolean

    Private mFirstHeaderInFile As Boolean
    Private mNotFirstPrintFooter As Boolean
    Private mNotFirstPrintHeader As Boolean

    Private mTablePrintFormat As Integer
    Private mTablePrintApply As Integer
    Private mTableLines As Boolean

    Private mInsertParagraphsAroundImages As Boolean
    Private mBuildContents As Boolean
    Private mBuildProject As Boolean
    Private mFooterTimestamps As Boolean
    Private mUpNext As Boolean
    Private mBuildID As Boolean
    Private mIDfile As Integer
    Private mIDnum As Integer
    Private mAliasSection As String
    Private mHTMLContentsfile As Integer
    Private mHTMLHelpProjectfile As Integer
    Private mHTMLIndexfile As Integer

    Private mSaveFilename As String
    Private mInPre As Boolean
    Private mAlreadyInitialized As Boolean
    Private mLastHeadingLevel As Integer
    Private mHeadingLevel As Integer
    Private mBookLevel As Integer
    Private mStyleLevel As Integer ', IconLevel%
    Private mSectionLevelName(99) As String

    Private Const mCuteButtons As Boolean = False
    Private Const mMoveHeadings As Integer = 0
    Private Const mMakeBoxyHeaders As Boolean = False
    Private mLinkToImageFiles As Integer '0=store data in document, 1=link+store in doc 2=soft links, -1=do not process images (assigned in Init())

    Public Const mAsterisks80 As String = "********************************************************************************"
    Private Const mTensPlace As String = "         1         2         3         4         5         6         7         8"
    Private Const mOnesPlace As String = "12345678901234567890123456789012345678901234567890123456789012345678901234567890"
    Private Const mMaxSectionNameLen As Integer = 53
    Private Const mTableType As String = "Table-type "
    Private Const mTableTypeLength As Integer = 11
    Private mWholeCardHeader As String
    Private mWholeCardHeaderLength As Integer

    Private mTotalTruncated As Integer
    Private mTotalRepeated As Integer

    Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Integer, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Integer) As Integer

    Public Sub CreateHelpProject(ByRef aIDfileExists As Boolean)
        Dim lSB As New StringBuilder
        lSB.AppendLine("[OPTIONS]")
        lSB.AppendLine("LCID=0x409 0x0 0x0 ; English (United States)")
        lSB.AppendLine("REPORT=Yes")
        lSB.AppendLine("CNT=" & pBaseName & ".cnt")
        lSB.AppendLine("")
        lSB.AppendLine("HLP=" & pBaseName & ".hlp")
        lSB.AppendLine("")

        lSB.AppendLine("[FILES]")
        lSB.AppendLine(mHelpSourceRTFName)
        lSB.AppendLine("")

        If aIDfileExists Then
            lSB.AppendLine("[MAP]")
            lSB.AppendLine("#include <" & pBaseName & ".ID>")
            lSB.AppendLine("")
        End If

        lSB.AppendLine("[WINDOWS]")
        lSB.AppendLine("Main=" & Chr(34) & pBaseName & " Manual" & Chr(34) & ", , 60672, (r14876671), (r12632256), f2; ")
        lSB.AppendLine("")

        lSB.AppendLine("[CONFIG]")
        lSB.AppendLine("BrowseButtons()")

        SaveFileString(mSaveDirectory & pBaseName & ".hpj", lSB.ToString)
    End Sub

    Public Function HTMLRelativeFilename(ByRef aWinFilename As String, ByRef aWinStartPath As String) As String
        HTMLRelativeFilename = ReplaceString(RelativeFilename(aWinFilename, aWinStartPath), "\", "/")
    End Function

    Private Sub OpenHTMLHelpProjectfile()
        mHTMLHelpProjectfile = FreeFile()
        FileOpen(mHTMLHelpProjectfile, mSaveDirectory & pBaseName & ".hhp", OpenMode.Output)
        Print(mHTMLHelpProjectfile, "[OPTIONS]" & vbLf)
        Print(mHTMLHelpProjectfile, "Auto Index=Yes" & vbLf)
        Print(mHTMLHelpProjectfile, "Compatibility=1.1 Or later" & vbLf)
        Print(mHTMLHelpProjectfile, "Compiled file=" & pBaseName & ".chm" & vbLf)
        Print(mHTMLHelpProjectfile, "Contents file=" & pBaseName & ".hhc" & vbLf)
        'Print #HTMLHelpProjectfile, "Default topic=Introduction.html"
        Print(mHTMLHelpProjectfile, "Display compile progress=Yes" & vbLf)
        Print(mHTMLHelpProjectfile, "Enhanced decompilation=Yes" & vbLf)
        Print(mHTMLHelpProjectfile, "Full-text search=Yes" & vbLf)
        Print(mHTMLHelpProjectfile, "Index file = " & pBaseName & ".hhk" & vbLf)
        Print(mHTMLHelpProjectfile, "Language=0x409 English (United States)" & vbLf)
        Print(mHTMLHelpProjectfile, "Title=" & pBaseName & " Manual" & vbLf & vbLf)
        'Print #HTMLHelpProjectfile, ""
        Print(mHTMLHelpProjectfile, "[Files]" & vbLf)
        mAliasSection = vbLf & "[ALIAS]"
    End Sub

    Private Sub CheckStyle()
        Dim lStartTagPos As Integer = mTargetText.ToLower.IndexOf("<style")
        If lStartTagPos >= 0 Then
            Dim lCloseTagPos As Integer = mTargetText.IndexOf(">", lStartTagPos)
            If lCloseTagPos < lStartTagPos Then
                Logger.Msg("Style tag not terminated in " & mSourceFilename)
            Else
                ReadStyleFile(Mid(mTargetText, lStartTagPos + 7, lCloseTagPos - lStartTagPos - 7), mHeadingLevel)
            End If
        ElseIf mHeadingLevel <= mStyleLevel Then
            mStyleLevel -= 1
            While mStyleFile(mStyleLevel).Length = 0
                mStyleLevel -= 1
            End While
            ReadStyleFile("", mStyleLevel)
        End If
    End Sub

    Private Sub ReadStyleFile(ByRef aStyleFilename As String, ByRef aHeadingLevel As Integer)
        mBeforeHTML = ""

        Dim lLevel As Integer
        For lLevel = 1 To mMaxLevels
            mWordStyle(lLevel) = Nothing
            mWordStyle(lLevel) = New Collection
        Next

        If aStyleFilename.Length = 0 Then
            aStyleFilename = mStyleFile(aHeadingLevel)
        Else
            If Not IO.File.Exists(aStyleFilename) Then
                If IO.File.Exists(aStyleFilename & ".sty") Then
                    aStyleFilename = aStyleFilename & ".sty"
                End If
            End If
            aStyleFilename = CurDir() & "\" & aStyleFilename
        End If

        If IO.File.Exists(aStyleFilename) Then
            mStyleFile(aHeadingLevel) = aStyleFilename
            mStyleLevel = aHeadingLevel
            Dim lCurrSection As String = ""
            For Each lLine As String In LinesInFile(aStyleFilename)
                lLine = lLine.Trim
                Dim lFirstChar As String = Left(lLine, 1)
                Select Case lFirstChar
                    Case "#", ""
                        'skip comments and blank lines
                    Case "["
                        lCurrSection = LCase(Mid(lLine, 2, Len(lLine) - 2))
                        lLevel = 0
                    Case Else
                        If IsNumeric(lFirstChar) Then
                            lLevel = CShort(lFirstChar)
                            lLine = Mid(lLine, 2)
                            While IsNumeric(Left(lLine, 1))
                                lLevel = lLevel * 10 + CShort(Left(lLine, 1))
                                lLine = Mid(lLine, 2)
                            End While
                            While Left(lLine, 1) = " " Or Left(lLine, 1) = "="
                                lLine = Mid(lLine, 2)
                            End While
                            'UPGRADE_ISSUE: GoSub statement is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C5A1A479-AB8B-4D40-AAF4-DB19A2E5E77F"'
                            Select Case lCurrSection
                                Case "beforehtml"
                                    If lLevel = 0 Then
                                        mBeforeHTML &= lLine & vbCrLf
                                    End If
                                Case "printsection" : mWordStyle(lLevel).Add(lLine)
                                Case "top" : mHeaderStyle(lLevel) = lLine
                                Case "bottom" : mFooterStyle(lLevel) = lLine
                                Case "body"
                                    If lLine.Length > 0 Then
                                        mBodyStyle(lLevel) = "<body " & lLine & ">"
                                    Else
                                        mBodyStyle(lLevel) = "<body>"
                                    End If
                            End Select
                        ElseIf lCurrSection = "printstart" Then
                            If pOutputFormat = OutputType.tPRINT Then WordCommand(lLine, 0)
                        Else
                            For lLevel = 0 To mMaxLevels
                                'UPGRADE_ISSUE: GoSub statement is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C5A1A479-AB8B-4D40-AAF4-DB19A2E5E77F"'
                                Select Case lCurrSection
                                    Case "beforehtml"
                                        If lLevel = 0 Then
                                            mBeforeHTML &= lLine & vbCrLf
                                        End If
                                    Case "printsection" : mWordStyle(lLevel).Add(lLine)
                                    Case "top" : mHeaderStyle(lLevel) = lLine
                                    Case "bottom" : mFooterStyle(lLevel) = lLine
                                    Case "body"
                                        If Len(lLine) > 0 Then
                                            mBodyStyle(lLevel) = "<body " & lLine & ">"
                                        Else
                                            mBodyStyle(lLevel) = "<body>"
                                        End If
                                End Select
                            Next lLevel
                        End If
                End Select
            Next
        End If
    End Sub

    Private Sub WordCommand(ByVal aCommandToProcess As String, ByVal aLocalHeadingLevel As Integer)
        System.Windows.Forms.Application.DoEvents()
        Dim lArgument As String
        Dim lValueString As String
        Dim lValueInteger As Integer
        Try
            Dim lCommandToProcess As String = aCommandToProcess
            Dim lCommand As String = StrSplit(lCommandToProcess, " ", """")
            With pWordBasic
                Select Case lCommand.ToLower
                    '      Case "applystyle":
                    '        If IsNumeric(lCommandToProcess) Then localHeadingLevel = CLng(lCommandToProcess)
                    '        .EditBookmark "Hstart" & localHeadingLevel
                    '        .EditGoTo "Hend" & localHeadingLevel
                    '        .Insert vbCr & vbCr
                    '        .EditGoTo "Hstart" & localHeadingLevel
                    '        .ExtendSelection
                    '        .EditGoTo "Hend" & localHeadingLevel
                    '        .Style "ADheading" & localHeadingLevel
                    '        .Cancel
                    '        .CharRight
                    Case "borderbottom" : If IsNumeric(lCommandToProcess) Then .BorderBottom(CInt(lCommandToProcess))
                    Case "borderinside" : If IsNumeric(lCommandToProcess) Then .BorderInside(CInt(lCommandToProcess))
                    Case "borderleft" : If IsNumeric(lCommandToProcess) Then .BorderLeft(CInt(lCommandToProcess))
                    Case "borderlinestyle" '0=none, 1 to 6 increasing thickness, 7,8,9 double, 10 gray, 11 dashed
                        Select Case lCommandToProcess.Trim.ToLower
                            Case "0", "1", "2", "3", "4", "5", "6"
                                .BorderLineStyle(CInt(lCommandToProcess))
                            Case "none" : .BorderLineStyle(0)
                            Case "thin" : .BorderLineStyle(1)
                            Case "thick" : .BorderLineStyle(6)
                            Case "double" : .BorderLineStyle(7)
                            Case "doublethick" : .BorderLineStyle(9)
                            Case "dashed" : .BorderLineStyle(11)
                            Case Else
                                Logger.Msg("Unknown BorderLineStyle: " & lCommandToProcess, MsgBoxStyle.OkOnly, "AuthorDoc:WordCommand")
                        End Select
                    Case "bordernone" : If IsNumeric(lCommandToProcess) Then .BorderNone(CInt(lCommandToProcess))
                    Case "borderoutside" : If IsNumeric(lCommandToProcess) Then .BorderOutside(CInt(lCommandToProcess))
                    Case "borderright" : If IsNumeric(lCommandToProcess) Then .BorderRight(CInt(lCommandToProcess))
                    Case "bordertop" : If IsNumeric(lCommandToProcess) Then .BorderTop(CInt(lCommandToProcess))
                    Case "charleft" : .CharLeft()
                    Case "charright" : .CharRight()
                    Case "centerpara" : .CenterPara()
                    Case "editclear"
                        If IsNumeric(lCommandToProcess) Then
                            .EditClear(CInt(lCommandToProcess))
                        Else
                            .EditClear()
                        End If
                    Case "editselectall" : .EditSelectAll()
                    Case "filepagesetup"
                        While lCommandToProcess.Length > 0
                            lValueString = StrSplit(lCommandToProcess, ",", """")
                            lArgument = StrSplit(lValueString, ":=", """")
                            If IsNumeric(lValueString) Then
                                lValueInteger = CInt(lValueString)
                            Else
                                lValueInteger = 0
                            End If
                            Select Case lArgument.ToLower
                                Case "topmargin" : pWordBasic.FilePageSetup(TopMargin:=lValueString)
                                Case "bottommargin" : pWordBasic.FilePageSetup(BottomMargin:=lValueString)
                                Case "leftmargin" : pWordBasic.FilePageSetup(LeftMargin:=lValueString)
                                Case "rightmargin" : pWordBasic.FilePageSetup(RightMargin:=lValueString)
                                Case "headerdistance" : pWordBasic.FilePageSetup(HeaderDistance:=lValueString)
                                Case "facingpages" : pWordBasic.FilePageSetup(FacingPages:=lValueInteger)
                                Case "oddandevenpages" : pWordBasic.FilePageSetup(OddAndEvenPages:=lValueInteger)
                            End Select
                        End While
                        '      Case "formatdefinestyleborders"
                        '        While Len(lCommandToProcess) > 0
                        '          lValueString = StrSplit(lCommandToProcess, ",", """")
                        '          lArgument = StrSplit(lValueString, ":=", """")
                        '          If IsNumeric(lValueString) Then
                        '            lValueInteger = CInt(lValueString)
                        '            Select Case LCase(lArgument)
                        '              Case "topborder":      .FormatDefineStyleBorders TopBorder:=lValueInteger
                        '              Case "leftborder":     .FormatDefineStyleBorders LeftBorder:=lValueInteger
                        '              Case "bottomborder":   .FormatDefineStyleBorders BottomBorder:=lValueInteger
                        '              Case "rightborder":    .FormatDefineStyleBorders RightBorder:=lValueInteger
                        '              Case "horizborder":    .FormatDefineStyleBorders HorizBorder:=lValueInteger
                        '              Case "vertborder":     .FormatDefineStyleBorders VertBorder:=lValueInteger
                        '              Case "topcolor":       .FormatDefineStyleBorders TopColor:=lValueInteger
                        '              Case "leftcolor":      .FormatDefineStyleBorders LeftColor:=lValueInteger
                        '              Case "bottomcolor":    .FormatDefineStyleBorders BottomColor:=lValueInteger
                        '              Case "rightcolor":     .FormatDefineStyleBorders RightColor:=lValueInteger
                        '              Case "horizcolor":     .FormatDefineStyleBorders HorizColor:=lValueInteger
                        '              Case "vertcolor":      .FormatDefineStyleBorders VertColor:=lValueInteger
                        '              Case "foreground":     .FormatDefineStyleBorders Foreground:=lValueInteger
                        '              Case "background":     .FormatDefineStyleBorders Background:=lValueInteger
                        '              Case "shading":        .FormatDefineStyleBorders Shading:=lValueInteger
                        '              Case "fineshading":    .FormatDefineStyleBorders FineShading:=lValueInteger
                        '            End Select
                        '          Else
                        '            logger.msg "non-numeric value for " & lArgument & " in " & cmd, vbOKOnly, "AuthorDoc:WordCommand"
                        '          End If
                        '        Wend
                    Case "formatdefinestylefont"
                        While lCommandToProcess.Length > 0
                            lValueString = StrSplit(lCommandToProcess, ",", """")
                            lArgument = StrSplit(lValueString, ":=", """")
                            If lArgument.ToLower = "font" Then
                                .FormatDefineStyleFont(Font:=lValueString)
                            ElseIf IsNumeric(lValueString) Then
                                lValueInteger = CInt(lValueString)
                                Select Case lArgument.ToLower
                                    Case "points" : .FormatDefineStyleFont(Points:=lValueInteger)
                                    Case "underline" : .FormatDefineStyleFont(Underline:=lValueInteger)
                                    Case "allcaps" : .FormatDefineStyleFont(AllCaps:=lValueInteger)
                                    Case "kerning" : .FormatDefineStyleFont(Kerning:=lValueInteger)
                                    Case "kerningmin" : .FormatDefineStyleFont(KerningMin:=lValueInteger)
                                    Case "bold" : .FormatDefineStyleFont(Bold:=lValueInteger)
                                    Case "italic" : .FormatDefineStyleFont(Italic:=lValueInteger)
                                    Case "outline" : .FormatDefineStyleFont(Outline:=lValueInteger)
                                    Case "shadow" : .FormatDefineStyleFont(Shadow:=lValueInteger)
                                    Case "font"
                                End Select
                            End If
                        End While
                        '      Case "formatdefinestylepara"
                        '        While Len(lCommandToProcess) > 0
                        '          lValueString = StrSplit(lCommandToProcess, ",", """")
                        '          lArgument = StrSplit(lValueString, ":=", """")
                        '          isnum = IsNumeric(lValueString)
                        '          If isnum Then
                        '            lValueInteger = CInt(lValueString)
                        '            Select Case LCase(lArgument)
                        '              Case "before":       .FormatDefineStylePara Before:=lValueInteger
                        '              Case "after":        .FormatDefineStylePara After:=lValueInteger
                        '              Case "keepwithnext": .FormatDefineStylePara KeepWithNext:=lValueInteger
                        '              Case "alignment":    .FormatDefineStylePara Alignment:=lValueInteger
                        '            End Select
                        '          Else
                        '            logger.msg "non-numeric value for " & lArgument & " in " & cmd, vbOKOnly, "AuthorDoc:WordCommand"
                        '          End If
                        '        Wend
                    Case "formatfont"
                        While lCommandToProcess.Length > 0
                            lValueString = StrSplit(lCommandToProcess, ",", """")
                            lArgument = StrSplit(lValueString, ":=", """")
                            If LCase(lArgument) = "font" Then
                                .FormatFont(Font:=lValueString)
                            ElseIf Len(lValueString) = 0 Then
                                If IsNumeric(lArgument) Then .FormatFont(Points:=lArgument)
                            ElseIf IsNumeric(lValueString) Then
                                lValueInteger = CInt(lValueString)
                                Select Case LCase(lArgument)
                                    Case "points" : .FormatFont(Points:=lValueInteger)
                                    Case "underline" : .FormatFont(Underline:=lValueInteger)
                                    Case "allcaps" : .FormatFont(AllCaps:=lValueInteger)
                                    Case "kerning" : .FormatFont(Kerning:=lValueInteger)
                                    Case "kerningmin" : .FormatFont(KerningMin:=lValueInteger)
                                    Case "bold" : .FormatFont(Bold:=lValueInteger)
                                    Case "italic" : .FormatFont(Italic:=lValueInteger)
                                    Case "outline" : .FormatFont(Outline:=lValueInteger)
                                    Case "shadow" : .FormatFont(Shadow:=lValueInteger)
                                End Select
                            Else
                                Logger.Msg("non-numeric value for " & lArgument & " in " & lCommand, MsgBoxStyle.OkOnly, "AuthorDoc:WordCommand")
                            End If
                        End While
                    Case "formatheaderfooterlink" : .FormatHeaderFooterLink()
                    Case "formatpagenumber"
                        While lCommandToProcess.Length > 0
                            lValueString = StrSplit(lCommandToProcess, ",", """")
                            lArgument = StrSplit(lValueString, ":=", """")
                            If LCase(lArgument) = "font" Then
                                .FormatFont(Font:=lValueString)
                            ElseIf Len(lValueString) = 0 Then
                                If IsNumeric(lArgument) Then .FormatFont(Points:=lArgument)
                            ElseIf IsNumeric(lValueString) Then
                                lValueInteger = CInt(lValueString)
                                Select Case LCase(lArgument)
                                    Case "chapternumber" : .FormatPageNumber(ChapterNumber:=lValueInteger)
                                    Case "numrestart" : .FormatPageNumber(NumRestart:=lValueInteger)
                                    Case "numformat" : .FormatPageNumber(NumFormat:=lValueInteger)
                                    Case "startingnum" : .FormatPageNumber(StartingNum:=lValueInteger)
                                    Case "level" : .FormatPageNumber(Level:=lValueInteger)
                                    Case "separator" : .FormatPageNumber(Separator:=lValueInteger)
                                End Select
                            Else
                                Logger.Msg("non-numeric value for " & lArgument & " in " & lCommand, MsgBoxStyle.OkOnly, "AuthorDoc:WordCommand")
                            End If
                        End While
                    Case "formatpara", "formatparagraph"
                        While lCommandToProcess.Length > 0
                            lValueString = StrSplit(lCommandToProcess, ",", """")
                            lArgument = StrSplit(lValueString, ":=", """")
                            If IsNumeric(lValueString) Then
                                lValueInteger = CInt(lValueString)
                                Select Case LCase(lArgument)
                                    Case "before" : .FormatParagraph(Before:=lValueInteger)
                                    Case "after" : .FormatParagraph(After:=lValueInteger)
                                    Case "keepwithnext" : .FormatParagraph(KeepWithNext:=lValueInteger)
                                    Case "alignment" : .FormatParagraph(Alignment:=lValueInteger)
                                End Select
                            Else
                                Logger.Msg("non-numeric value for " & lArgument & " in " & lCommand, MsgBoxStyle.OkOnly, "AuthorDoc:WordCommand")
                            End If
                        End While
                        '      Case "formatstyle"
                        '        lArgument = StrSplit(lCommandToProcess, ",", """")
                        '        If lArgument = "Normal" Then 'Provide some good defaults in case they aren't explicit in the style file
                        '          .FormatStyle lArgument, AddToTemplate:=1, Define:=1
                        '          .FormatDefineStyleFont 11, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 10, 1, 1, "Times New Roman", 0, 0, 0, 0
                        '          .FormatDefineStylePara Chr$(34), Chr$(34), 0, 0, 0, 0, 0, 1, 0, 0, 0, 0, 0, 1, Chr$(34)
                        '          .FormatDefineStyleLang "English (US)", 1
                        '          .FormatDefineStyleBorders 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, -1
                        '        Else
                        '          .FormatStyle lArgument, BasedOn:="Normal", AddToTemplate:=0, Define:=1
                        '          .FormatStyle lArgument, Delete:=1
                        '          .FormatStyle lArgument, BasedOn:="Normal", AddToTemplate:=0, Define:=1
                        '        End If
                    Case "formattabs"
                        lArgument = StrSplit(lCommandToProcess, ",", """")
                        lValueString = StrSplit(lCommandToProcess, ",", """") '1=left, 2=right
                        If Not IsNumeric(lArgument) Then
                            Logger.Msg("non-numeric value for tab position in " & lCommand, MsgBoxStyle.OkOnly, "AuthorDoc:WordCommand")
                        ElseIf Not IsNumeric(lValueString) Then
                            Logger.Msg("non-numeric value for alignment in " & lCommand, MsgBoxStyle.OkOnly, "AuthorDoc:WordCommand")
                        Else
                            lValueInteger = CInt(lValueString)
                            .FormatTabs(lArgument & """", Align:=lValueInteger, Set:=1)
                        End If
                    Case "formattabsclear" : .FormatTabs(ClearAll:=1)
                    Case "gotoheaderfooter" : .GoToHeaderFooter()
                    Case "insert" : .Insert(ReplaceStyleString(lCommandToProcess, aLocalHeadingLevel))
                    Case "insertbreak"
                        Select Case LCase(Trim(lCommandToProcess))
                            '0 (zero) Page break, 1 Column break, 2 Next Page section break, 3 Continuous section break, 4 Even Page section break, 5 Odd Page section break, 6 Line break (newline character)
                            Case "0", "1", "2", "3", "4", "5", "6"
                                .InsertBreak(CInt(lCommandToProcess))
                            Case "page" : .InsertBreak(0)
                            Case "column" : .InsertBreak(1)
                            Case "pagesection" : .InsertBreak(2)
                            Case "contsection" : .InsertBreak(3)
                            Case "evenpagesection" : .InsertBreak(4)
                            Case "oddpagesection" : .InsertBreak(5)
                            Case "line" : .InsertBreak(6)
                            Case Else : Logger.Dbg("Unknown argument to InsertBreak: " & lCommandToProcess)
                        End Select
                    Case "insertdatetime"
                        If Len(Trim(lCommandToProcess)) > 0 Then
                            .InsertDateTime(lCommandToProcess, 0)
                        Else
                            .InsertDateTime("   hh:mm MMMM d, yyyy", 0)
                        End If
                    Case "insertfield" : .InsertField(lCommandToProcess)
                    Case "insertpagenumbers"
                        Dim lTypeVal As Integer = 1
                        Dim lPosVal As Integer = 1
                        Dim lFirstVal As Integer = 0
                        While lCommandToProcess.Length > 0
                            lValueString = StrSplit(lCommandToProcess, ",", """")
                            lArgument = StrSplit(lValueString, ":=", """")
                            If IsNumeric(lValueString) Then
                                lValueInteger = CInt(lValueString)
                                Select Case LCase(lArgument)
                                    Case "type" : lTypeVal = lValueInteger
                                    Case "position" : lPosVal = lValueInteger
                                    Case "firstpage" : lFirstVal = lValueInteger
                                End Select
                            Else
                                Logger.Msg("non-numeric value for " & lArgument & " in " & lCommand, MsgBoxStyle.OkOnly, "AuthorDoc:WordCommand")
                            End If
                        End While
                        .InsertPageNumbers(Type:=lTypeVal, Position:=lPosVal, FirstPage:=lFirstVal)
                    Case "InsertParagraphsAroundImages"
                        Select Case lCommandToProcess.ToLower
                            Case "0", "false" : mInsertParagraphsAroundImages = False
                            Case "1", "true" : mInsertParagraphsAroundImages = True
                        End Select
                    Case "shownextheaderfooter"
                        .ShowNextHeaderFooter()
                    Case "startofdocument"
                        .StartOfDocument()
                    Case "tableprintapply"
                        If IsNumeric(lCommandToProcess) Then mTablePrintApply = CInt(lCommandToProcess)
                    Case "tableprintformat"
                        If IsNumeric(lCommandToProcess) Then mTablePrintFormat = CInt(lCommandToProcess)
                    Case "toggleheaderfooterlink"
                        .ToggleHeaderFooterLink()
                    Case "viewfooter"
                        pWordApp.ActiveWindow.ActivePane.View.SeekView = WdSeekView.wdSeekCurrentPageFooter ' .ViewFooter()
                    Case "viewheader"
                        pWordApp.ActiveWindow.ActivePane.View.SeekView = WdSeekView.wdSeekCurrentPageHeader ' .ViewHeader()
                    Case "viewfooterandset"
                        ViewFooterAndSet(ReplaceStyleString(lCommandToProcess, aLocalHeadingLevel))
                    Case "viewheaderandset"
                        ViewHeaderAndSet(ReplaceStyleString(lCommandToProcess, aLocalHeadingLevel))
                    Case "viewnormal"
                        pWordApp.ActiveWindow.ActivePane.View.SeekView = WdSeekView.wdSeekMainDocument '.ViewNormal()
                    Case "viewpage"
                        .ViewPage()
                    Case Else
                        Logger.Dbg("WordCommand not recognized: " & lCommand)
                End Select
            End With
        Catch
            Logger.Dbg("Error with Word command '" & aCommandToProcess & "'" & vbCr & Err.Description)
        End Try
    End Sub

    Private Function ReplaceStyleString(ByRef aString As String, ByRef aLocalHeadingLevel As Integer) As String
        Dim lReturnString As String = aString
        lReturnString = lReturnString.Replace("<sectionname>", mHeadingText(aLocalHeadingLevel))
        For lHeadingLevel As Integer = 1 To aLocalHeadingLevel
            lReturnString = ReplaceString(lReturnString, "<sectionname " & lHeadingLevel & ">", mHeadingText(lHeadingLevel))
        Next
        lReturnString = ReplaceString(lReturnString, "vbTab", vbTab)
        lReturnString = ReplaceString(lReturnString, "vbCr", vbCr)
        lReturnString = ReplaceString(lReturnString, "vbLf", vbLf)
        lReturnString = ReplaceString(lReturnString, "vbCrLf", vbCrLf)
        Dim lSectionWord As String
        Dim lSectionWordPos As Integer = lReturnString.IndexOf("<sectionword")
        While lSectionWordPos > -1
            Dim lSectionWordEndPos As Integer = lReturnString.IndexOf(">", lSectionWordPos + 12)
            If lSectionWordEndPos = -1 Then
                lSectionWordPos = 0
            Else
                lSectionWord = Trim(Mid(lReturnString, lSectionWordPos + 12, lSectionWordEndPos - lSectionWordPos - 12))
                If IsNumeric(lSectionWord) Then
                    Dim wordnum As Integer = CInt(lSectionWord)
                    lSectionWord = mHeadingText(aLocalHeadingLevel)
                    While wordnum > 1
                        StrSplit(lSectionWord, " ", "")
                    End While
                    lSectionWord = StrSplit(lSectionWord, " ", "")
                    lReturnString = Left(lReturnString, lSectionWordPos - 1) & lSectionWord & Mid(lReturnString, lSectionWordEndPos + 1)
                Else
                    lReturnString = Left(lReturnString, lSectionWordPos - 1) & mHeadingText(aLocalHeadingLevel) & Mid(lReturnString, lSectionWordEndPos + 1)
                End If
                lSectionWordPos = InStr(lSectionWordPos + 1, lReturnString, "<sectionword")
            End If
        End While
        Return lReturnString
    End Function

    Public Sub Convert(ByRef aOutputAs As OutputType, _
                       ByRef aMakeContents As Boolean, ByRef aTimestamps As Boolean, _
                       ByRef aMakeUpNext As Boolean, ByRef aMakeID As Boolean, ByRef aMakeProject As Boolean)
        Logger.StartToFile(CurDir() & "\log\authordoc.log", False, True)
        Logger.Dbg("StartConvert " & aOutputAs)

        pKeywords = New Collection
        Init()
        pOutputFormat = aOutputAs
        mBuildContents = aMakeContents
        mBuildProject = aMakeProject
        mFooterTimestamps = aTimestamps
        mUpNext = aMakeUpNext
        mBuildID = aMakeID
        frmConvert.CmDialog1Open.DefaultExt = "txt"
        If IO.File.Exists(pProjectFileName) Then
            'mPromptForFiles = False
            mProjectFileEntrys.Clear()
            For Each lLine As String In LinesInFile(pProjectFileName)
                If lLine.Trim.Length > 0 Then
                    mProjectFileEntrys.Add(lLine)
                End If
            Next
            mNextProjectFileIndex = 1
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
        If mBuildProject Then
            If pOutputFormat = OutputType.tHELP Then
                CreateHelpProject(True)
            ElseIf pOutputFormat = OutputType.tHTMLHELP Then
                OpenHTMLHelpProjectfile()
            ElseIf pOutputFormat = OutputType.tASCII Then
                mHTMLHelpProjectfile = FreeFile()
                FileOpen(mHTMLHelpProjectfile, mSaveDirectory & pBaseName & ".txt", OpenMode.Output)
            End If
        End If

        If mBuildID Then
            mIDfile = FreeFile()
            FileOpen(mIDfile, mSaveDirectory & pBaseName & ".ID", OpenMode.Output)
            mIDnum = 2
        End If

        InitContents()
        'mPromptForFiles = False
        Dim lastSourceFilename As String = ""
        mSourceFilename = NextSourceFilename()
        If pOutputFormat = OutputType.tPRINT Or pOutputFormat = OutputType.tHELP Then
            'pWordApp = New Microsoft.Office.Interop.Word.Application
            'pWordBasic = pWordApp.WordBasic 
            pWordBasic = CreateObject("Word.Basic")
            pWordApp = GetObject(, "Word.Application")

            With pWordBasic
                .AppShow()
                pWordApp.ActiveWindow.View.Type = WdViewType.wdPrintView 'ensures all commands available
                '.ToolsOptionsView PicturePlaceHolders:=1
                .ChDir(mSaveDirectory)
                If pOutputFormat = OutputType.tPRINT Then
                    .FileNewDefault()
                    DefinePrintStyles()
                    .FileSaveAs(mSaveDirectory & pBaseName & ".doc", 0)
                    mTargetWindowName = .WindowName
                ElseIf pOutputFormat = OutputType.tHELP Then
                    .FileNewDefault()
                    .FilePageSetup(PageWidth:="12 in")
                    .FileSaveAs(mSaveDirectory & mHelpSourceRTFName, 6)
                    mTargetWindowName = .WindowName
                End If
                .ChDir(mSourceBaseDirectory)
            End With
        End If
        ReadStyleFile(pBaseName & ".sty", 0)
        mLastHeadingLevel = 0
        While mSourceFilename.Length > 0 AndAlso mSourceFilename <> lastSourceFilename
OpeningFile:
            Status("Opening " & mSourceFilename)
            lastSourceFilename = mSourceFilename
            mFirstHeaderInFile = True
            System.Windows.Forms.Application.DoEvents()
            If pOutputFormat = OutputType.tPRINT Or pOutputFormat = OutputType.tHELP Then
                With pWordBasic
                    .Activate(mTargetWindowName)
                    .ScreenUpdating(0) 'comment out to debug (show lots of updates)
                    .EditBookmark("CurrentFileStart")
                    Try
                        .Insert(WholeFileString(mSourceFilename))
                    Catch ex As Exception
                        GoTo FileNotFound
                    End Try
                    NumberHeaderTagsWithWord()
                    If mLinkToImageFiles >= 0 Then
                        .EditGoTo("CurrentFileStart")
                        TranslateIMGtags(PathNameOnly(AbsolutePath(mSourceFilename, CurDir)))
                    End If
                    .EndOfDocument()
                    .ScreenUpdating(1)
                End With
            ElseIf pOutputFormat = OutputType.tASCII Then
                Dim i As Integer = FreeFile()
                Try
                    FileOpen(i, mSourceFilename, OpenMode.Input) 'SourceBaseDirectory &
                Catch ex As Exception
                    GoTo FileNotFound
                End Try
                While Not EOF(i) ' Loop until end of file.
                    ParseHSPFsourceLine(i)
                End While
                If mBuildProject And pKeywords.Count() > 0 Then
                    Print(mHTMLHelpProjectfile, vbCrLf & "[Keywords]" & vbCrLf)
                    For Each lKeyword As Object In pKeywords 'TODO: where do keywords come from?
                        'UPGRADE_WARNING: Couldn't resolve default property of object keyword. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        Print(mHTMLHelpProjectfile, lKeyword & vbCrLf)
                    Next lKeyword
                End If
                FileClose(i)
            ElseIf pOutputFormat = OutputType.tHTML Or pOutputFormat = OutputType.tHTMLHELP Then
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
                    mSaveFilename = Left(mSourceFilename, lDotPos - 1)
                Else
                    mSaveFilename = mSourceFilename
                End If
                mSaveFilename = mSaveFilename & ".html"
                If pOutputFormat = OutputType.tHTMLHELP Then
                    FormatTag("b", pOutputFormat)
                    FormatKeywordsHTMLHelp()
                    If mBuildProject Then Print(mHTMLHelpProjectfile, mSaveFilename & vbLf)
                End If
                NumberHeaderTags()
                CheckStyle()
                FormatHeadings(OutputType.tHTML, mSaveFilename)
                TranslateButtons(pOutputFormat)
                MakeLocalTOCs()
                HREFsInsureExtension()
                AbsoluteToRelative()
                If pOutputFormat <> OutputType.tHELP AndAlso pOutputFormat <> OutputType.tPRINT Then
                    CopyImages()
                End If
                'FormatCardGraphic()
                SaveInNewDir(mSaveDirectory & mSaveFilename)
            End If
            Status("Closing " & mSourceFilename)
OpenNextFile:
            mSourceFilename = NextSourceFilename()
        End While
        If pOutputFormat = OutputType.tHTMLHELP And aMakeProject Then
            Print(mHTMLHelpProjectfile, mAliasSection & vbLf)
            Print(mHTMLHelpProjectfile, "[MAP]" & vbLf & "#include " & pBaseName & ".ID" & vbLf)
            FileClose(mHTMLHelpProjectfile)
        ElseIf pOutputFormat = OutputType.tASCII Then
            FileClose(mIDfile)
            FileClose(mHTMLHelpProjectfile)
        End If
        If (pOutputFormat = OutputType.tPRINT Or pOutputFormat = OutputType.tHELP) Then
            With pWordBasic
                Logger.Dbg("WordCount:" & pWordApp.ActiveDocument.Words.Count)
                Dim lReplaceSelectionOption As Integer
                .ToolsOptionsEdit((lReplaceSelectionOption)) 'save current value of this option
                .ToolsOptionsEdit((1)) 'be sure option is on
                .ScreenUpdating(0) 'comment out to debug (show lots of updates)
                .Activate(mTargetWindowName)
                ConvertTablesToWord()
                Logger.Dbg("WordCount:" & pWordApp.ActiveDocument.Words.Count)
                ConvertTagsToWord()
                If aMakeContents Then
                    If pOutputFormat = OutputType.tHTMLHELP Or pOutputFormat = OutputType.tHTML Then
                        FinishHTMLHelpContents()
                        'ElseIf OutputFormat = tHTML Then
                        '  .Activate ContentsWin
                        '  .FileSaveAs Directory & "Contents.html", 2
                        '  .FileClose 2
                    ElseIf pOutputFormat = OutputType.tPRINT Then
                        .Activate(mTargetWindowName)
                        .StartOfDocument()
                        .Insert("Contents" & vbCr & vbCr)
                        .InsertTableOfContents(0, 0, AddedStyles:="ADheading1,1,ADheading2,2,ADheading3,3,ADheading4,4,ADheading5,5,ADheading6,6", RightAlignPageNumbers:=1)
                    ElseIf mContentsWindowName.Length > 0 Then
                        .Activate(mContentsWindowName)
                        .FileSave()
                        .FileClose(2)
                    End If
                End If
                .ToolsOptionsEdit((lReplaceSelectionOption))
                If Len(mTargetWindowName) > 0 Then
                    .Activate(mTargetWindowName)
                    Status("Saving file: " & mTargetWindowName)
                    .FileSave()
                    .FileClose(2)
                End If
                .ScreenUpdating(1)
                .AppClose()
            End With
        ElseIf pOutputFormat = OutputType.tHTMLHELP Or pOutputFormat = OutputType.tHTML Then
            FinishHTMLHelpContents()
        End If
        pWordBasic = Nothing
        If mIDfile > -1 Then FileClose(mIDfile)
        If mTotalTruncated > 0 Or mTotalRepeated > 0 Then
            Logger.Msg("Total Truncated = " & mTotalTruncated & vbCr & "Total Repeated = " & mTotalRepeated)
        End If
        Status("Conversion Finished")
        If pOutputFormat = OutputType.tHELP Then
            ShellExecute(frmConvert.Handle.ToInt32, "Open", mSaveDirectory & pBaseName & ".hpj", vbNullString, vbNullString, 1) 'SW_SHOWNORMAL"
        ElseIf pOutputFormat = OutputType.tHTMLHELP Then
            ShellExecute(frmConvert.Handle.ToInt32, "Open", mSaveDirectory & pBaseName & ".hhp", vbNullString, vbNullString, 1) 'SW_SHOWNORMAL"
        End If
        Logger.Flush()
        Exit Sub

FileNotFound:
        If Logger.Msg("Error opening " & mSourceFilename & " (" & Err.Description & ")", MsgBoxStyle.RetryCancel, "Help Convert") = MsgBoxResult.Retry Then
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
                If mIDfile > 0 Then
                    If mInPre Then Print(mIDfile, vbCrLf & "</pre>" & vbCrLf)
                    If FileKeywords.Count() > 0 Then
                        For Each v In FileKeywords
                            'UPGRADE_WARNING: Couldn't resolve default property of object v. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            PrintLine(mIDfile, "<keyword=" & v & ">" & vbCrLf)
                        Next v
                    End If
                    FileClose(mIDfile)
                End If
                mInPre = False
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
                If UCase(Left(SectionName, mTableTypeLength)) = UCase(mTableType) Then SectionName = Mid(SectionName, mTableTypeLength + 1)
                buf = "<SecNum " & SectionNum & "> " & "<h>" & SectionName & "</h>" & vbCrLf & buf & vbCrLf


                If Right(SectionNum, 2) = ".0" Then SectionNum = Left(SectionNum, Len(SectionNum) - 2)
                SectionDir = ""
                SectionDirName = ""
                parsePos = InStr(SectionNum, ".")
                If parsePos > 0 Then
                    DirectoryLevels = 1
                    parsePos = InStrRev(SectionNum, ".")
                    SectionDir = Left(SectionNum, parsePos - 1)
                    SectionDirName = mSectionLevelName(DirectoryLevels)
                    parsePos = InStr(SectionDir, ".")
                    While parsePos > 0
                        DirectoryLevels = DirectoryLevels + 1
                        SectionDir = Left(SectionDir, parsePos - 1) & "\" & Mid(SectionDir, parsePos + 1)
                        SectionDirName = SectionDirName & "\" & mSectionLevelName(DirectoryLevels)
                        parsePos = InStr(parsePos + 1, SectionDir, ".")
                    End While
                    SectionDir = SectionDir & "\"
                    SectionDirName = SectionDirName & "\"
                End If
                mIDfile = FreeFile
                If Not IO.Directory.Exists(mSaveDirectory & SectionDirName) Then IO.Directory.CreateDirectory(mSaveDirectory & SectionDirName)
                'Debug.Print
                'Debug.Print SectionDir & SectionNum & ":" & CurrentOutputDirectory & CurrentOutputFilename
                CurrentOutputDirectory = mSaveDirectory & SectionDirName 'SectionDir
                dummy = MakeValidFilename(SectionName)
                If Len(dummy) <= mMaxSectionNameLen Then
                    mSectionLevelName(DirectoryLevels + 1) = dummy
                Else
                    mTotalTruncated = mTotalTruncated + 1
                    mSectionLevelName(DirectoryLevels + 1) = Trim(Left(dummy, 34) & Right(dummy, 1)) 'MakeValidFilename(buf)
                    Debug.Print("Truncated " & dummy & vbLf & "Shorter = " & mSectionLevelName(DirectoryLevels + 1))
                End If
                FileRepeat = 1
SetFilenameHere:
                CurrentOutputFilename = mSectionLevelName(DirectoryLevels + 1) & ".txt" 'Mid(SectionNum, Len(SectionDir) + 1) & ".txt"
                If Len(CurrentOutputDirectory & CurrentOutputFilename) > 255 Then
                    Logger.Msg("Path longer than 255 characters detected:" & vbCr & CurrentOutputDirectory & vbCr & CurrentOutputFilename)
                End If
                If IO.File.Exists(CurrentOutputDirectory & CurrentOutputFilename) Then
                    FileRepeat = FileRepeat + 1
                    mSectionLevelName(DirectoryLevels + 1) = mSectionLevelName(DirectoryLevels + 1) & FileRepeat
                    GoTo SetFilenameHere
                End If
                'Debug.Print Space(2 * DirectoryLevels) & "<li><a href=""Functional Description" & Mid(CurrentOutputDirectory, 21) & SectionLevelName(DirectoryLevels + 1) & """>" & dummy & "</a>"
                If FileRepeat > 1 Then mTotalRepeated = mTotalRepeated + 1
                FileOpen(mIDfile, CurrentOutputDirectory & CurrentOutputFilename, OpenMode.Output)
                If mBuildProject Then PrintLine(mHTMLHelpProjectfile, Space(2 * DirectoryLevels) & mSectionLevelName(DirectoryLevels + 1)) 'Trim(Mid(buf, Len(SectionNum) + 1))  'Mid(SectionNum, Len(SectionDir) + 1)
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
                If mInPre Then
                    If InStr(buf, "Explanation") > 0 Then
                        mInPre = False
                        buf = "</pre>" & vbCrLf & buf
                    End If
                Else
                    If InStr(buf, "****************************************") > 0 Or InStr(buf, "----------------------------------------") > 0 Then
                        mInPre = True
                        buf = "<pre>" & vbCrLf & buf
                    End If
                End If
                If Not mInPre Then
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
        Print(mIDfile, buf2 & vbCrLf)
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
            If Not .EditFindFound Then
                Exit Function
            End If
            .ExtendSelection()
            .EditFind(">")
            If .EditFindFound Then
                If .Selection.ToLower.IndexOf("border=0") > -1 Then
                    mTableLines = False
                Else
                    mTableLines = True
                End If
                .EditClear() 'delete <table...>
            End If
            .Cancel()
            Return .EditFindFound
        End With
    End Function

    Private Sub ConvertTableToWord(ByRef aRecursionLevel As Integer)
        With pWordBasic
            .EditBookmark("TableStart" & aRecursionLevel)
FindEnd:
            .EditFind("</table>")
            If Not .EditFindFound Then
                Logger.Msg("  EndTableNotFound")
                Exit Sub
            End If

            .EditBookmark("TableEnd" & aRecursionLevel)
            .ExtendSelection()
            .EditGoTo("TableStart" & aRecursionLevel)
            .Cancel()

            If .Selection.ToLower.IndexOf("<table") > -1 Then
                If FindAndDeleteTableStart() Then
                    Logger.Dbg("RecursiveTable " & aRecursionLevel + 1)
                    ConvertTableToWord(aRecursionLevel + 1)
                    .EditGoTo("TableStart" & aRecursionLevel)
                    GoTo FindEnd
                Else
                    .EditGoTo("TableEnd" & aRecursionLevel)
                End If
            Else
                .EditGoTo("TableEnd" & aRecursionLevel)
            End If

            .EditClear() 'delete </table>
            .EditBookmark("TableEnd" & aRecursionLevel)
            .ExtendSelection()
            .EditGoTo("TableStart" & aRecursionLevel)

            Dim lTableLen As Integer = .Selection.Length
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
            Dim lTableText As String = .Selection.ToLower

            If lTableText.IndexOf("<table") > -1 Then
                .EditGoTo("TableEnd" & aRecursionLevel)
                .EditBookmark("TableEnd" & aRecursionLevel)
            End If

            Dim lTableCols As Integer = 0
            Dim lColPos As Integer = lTableText.IndexOf("<tr")
            Dim lRowEnd As Integer = lTableText.IndexOf("tr>", lColPos + 2)
            If lRowEnd = -1 Then
                lRowEnd = lTableText.Length
            End If
            While lColPos >= 0 And lColPos < lRowEnd
                lTableCols += 1
                lColPos = lTableText.IndexOf("<th", lColPos + 1)
                If Mid(lTableText, lColPos + 3, 8) = " colspan" Then
                    lColPos += 12
                    While Not IsNumeric(Mid(lTableText, lColPos, 1))
                        lColPos += 1
                    End While
                    lTableCols += CShort(Mid(lTableText, lColPos, 1)) - 1
                End If
            End While
            lColPos = lTableText.IndexOf("<tr")
            lColPos = InStr(lColPos + 1, lTableText, "<td")
            While lColPos > 0 And lColPos < lRowEnd
                lTableCols += 1
                lColPos = InStr(lColPos + 1, lTableText, "<td")
            End While
            If lTableCols > 1 Then
                lTableCols -= 1
            End If
            'Dim lHeaderCell(500, lTableCols) As Boolean
            Logger.Dbg("Table " & aRecursionLevel & " Length:" & lTableLen & " Columns:" & lTableCols)

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
            .EditGoTo("TableAll")
            .EditReplace("^w^p", "^p", ReplaceAll:=True)
            .EditGoTo("TableAll")
            .EditReplace("^p", " ", ReplaceAll:=True)
            .EditGoTo("TableAll")
            .FormatTabs("2""", Align:=0, Set:=1)
            .EditReplace("<tr><th>", "^p", 0, 0, 0, 0, 0, 0, 0, 1, 0, 0, 0)
            .EditGoTo("TableAll")
            .EditReplace("<tr><td>", "^p", 0, 0, 0, 0, 0, 0, 0, 1, 0, 0, 0)
            .EditGoTo("TableAll")
            .EditReplace("<tr>^w<th>", "^p", 0, 0, 0, 0, 0, 0, 0, 1, 0, 0, 0)
            .EditGoTo("TableAll")
            .EditReplace("<tr>^w<td>", "^p", 0, 0, 0, 0, 0, 0, 0, 1, 0, 0, 0)
            .EditGoTo("TableAll")
            .EditReplace("<tr>", "^p", 0, 0, 0, 0, 0, 0, 0, 1, 0, 0, 0)
            .EditGoTo("TableAll")
            .EditReplace("<td>", vbTab, 0, 0, 0, 0, 0, 0, 0, 1, 0, 0, 0)
            .EditGoTo("TableAll")
            .EditReplace("<th>", vbTab, 0, 0, 0, 0, 0, 0, 0, 1, 0, 0, 0)
            .EditGoTo("TableAll")
            .EditReplace("<td ", vbTab, 0, 0, 0, 0, 0, 0, 0, 1, 0, 0, 0)
            .EditGoTo("TableAll")
            .EditReplace("<th ", vbTab, 0, 0, 0, 0, 0, 0, 0, 1, 0, 0, 0)
            .EditGoTo("TableAll")
            .EditReplace("<p>", "", 0, 0, 0, 0, 0, 0, 0, 1, 0, 0, 0)
            .EditGoTo("TableAll")
            .EditReplace("</tr>", "", 0, 0, 0, 0, 0, 0, 0, 1, 0, 0, 0)
            .EditGoTo("TableAll")
            .EditReplace("</td>", "", 0, 0, 0, 0, 0, 0, 0, 1, 0, 0, 0)
            .EditGoTo("TableAll")
            .EditReplace("</thead>", "", 0, 0, 0, 0, 0, 0, 0, 1, 0, 0, 0)
            .EditGoTo("TableAll")
            .EditReplace("<thead>", "", 0, 0, 0, 0, 0, 0, 0, 1, 0, 0, 0)

            .EditGoTo("TableEnd" & aRecursionLevel)
            .ExtendSelection()
            .EditGoTo("TableStart" & aRecursionLevel)
SkipBlanks2:
            Select Case Asc(.Selection)
                Case 10, 13, 32
                    .CharRight()
                    GoTo SkipBlanks2
            End Select

            If lTableCols = 0 Then
                Dim lPrintLength As Integer = Math.Min(.Selection.Length, 500)
                Logger.Dbg("--------TableProblem:" & .Selection.Substring(0, lPrintLength - 1))
            Else
                .TableInsertTable(ConvertFrom:=1, NumColumns:=lTableCols, Format:=mTablePrintFormat, Apply:=mTablePrintApply)
                .EditBookmark("TableAll")
                .TableColumnWidth(AutoFit:=1)

                If Not mTableLines Then
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
                    Dim lMergeCells As Integer = CInt(.Selection)
                    .CharRight() '>
                    .EditClear()
                    .NextCell()
                    .CharLeft()
                    .ExtendSelection()
                    'For MergeCount = 2 To MergeCells
                    .CharRight(lMergeCells)
                    'Next
                    .TableMergeCells()
                    .EditGoTo("TableAll")
                    .EditFind("colspan=")
                End While
                .CharRight()
            End If
        End With
    End Sub

    Private Sub ConvertTagsToWord()
        With pWordBasic
            Status("Removing HTML Headers")
            RemoveStuffOutsideBody()
            Logger.Dbg("WordCount:" & pWordApp.ActiveDocument.Words.Count)

            Status("Translating Paragraph Marks")
            InsertParagraphsInPRE(pOutputFormat)
            .StartOfDocument()
            Logger.Dbg("WordCount:" & pWordApp.ActiveDocument.Words.Count)

            Status("Removing Whitespace After Paragraph Marks")
            .EditReplace("^p^w", "^p", ReplaceAll:=True)
            Logger.Dbg("WordCount:" & pWordApp.ActiveDocument.Words.Count)

            Status("Removing Whitespace Before Paragraph Marks")
            .EditReplace("^w^p", "^p", ReplaceAll:=True)
            Logger.Dbg("WordCount:" & pWordApp.ActiveDocument.Words.Count)

            Status("Removing Non-HTML Paragraphs")
            .EditReplace("^p", " ", ReplaceAll:=True)
            Logger.Dbg("WordCount:" & pWordApp.ActiveDocument.Words.Count)

            TranslateLists("ul", 1)
            TranslateLists("ol", 7)

            Status("Replacing HTML Paragraphs")
            .EditReplace("<p>", "^p", ReplaceAll:=True)
            Logger.Dbg("WordCount:" & pWordApp.ActiveDocument.Words.Count)

            Status("Replacing HTML Line Breaks")
            .EditReplace("<br>", "^l", ReplaceAll:=True)
            Logger.Dbg("WordCount:" & pWordApp.ActiveDocument.Words.Count)

            Status("Removing Whitespace After Paragraph Marks")
            .EditReplace("^p^w", "^p", ReplaceAll:=True)
            Logger.Dbg("WordCount:" & pWordApp.ActiveDocument.Words.Count)

            Status("Removing Whitespace Before Paragraph Marks")
            .EditReplace("^w^p", "^p", ReplaceAll:=True)
            Logger.Dbg("WordCount:" & pWordApp.ActiveDocument.Words.Count)

            Status("Replacing HTML Page Breaks")
            .EditReplace("<page>", "^m", ReplaceAll:=True)

            .EditSelectAll()
            .FormatParagraph(After:=10, LineSpacingRule:=3, LineSpacing:=32)
            If pOutputFormat = outputType.tHELP Then .FormatFont(12)
            .Cancel()
            Status("Translating Buttons")
            TranslateButtons(pOutputFormat)
            Logger.Dbg("WordCount:" & pWordApp.ActiveDocument.Words.Count)

            Status("Formatting Headings")
            FormatHeadings(pOutputFormat, mTargetWindowName)
            FormatTag("div", pOutputFormat)
            FormatTag("pre", pOutputFormat)
            FormatTag("figure", pOutputFormat)
            FormatTag("u", pOutputFormat)
            FormatTag("b", pOutputFormat)
            FormatTag("i", pOutputFormat)
            FormatTag("sub", pOutputFormat)
            FormatTag("sup", pOutputFormat)
            Logger.Dbg("WordCount:" & pWordApp.ActiveDocument.Words.Count)

            If pOutputFormat = outputType.tHELP Then HREFsToHelpHyperlinks()
            Status("Removing Remaining HTML Tags")
            HTMLQuotedCharsToPrint()
            Logger.Dbg("WordCount:" & pWordApp.ActiveDocument.Words.Count)

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

    Private Sub SaveInNewDir(ByRef aNewFilePath As String)
        Dim lPath As String = IO.Path.GetDirectoryName(aNewFilePath)
        MkDirPath(lPath)

        If pOutputFormat = OutputType.tHTML OrElse pOutputFormat = OutputType.tHTMLHELP Then
            SaveFileString(aNewFilePath, mTargetText)
        Else
            With pWordBasic
                Dim lOldPath As String = .DefaultDir(14)
                .ChDir(lPath)
                .FileSaveAs(FilenameNoPath(aNewFilePath), 2, AddToMru:=False)
                .ChDir(lOldPath)
            End With
        End If
    End Sub

    Private Function NextSourceFilename() As String
        If mNextProjectFileIndex > mProjectFileEntrys.Count Then 'all done
            Return ""
        Else
            Dim lNextProjectFileEntry As String = mProjectFileEntrys(mNextProjectFileIndex)
            mNextProjectFileIndex += +1
            'insert levels of hierarchy for subsections indented two spaces
            Dim lFileName As String = ""
            Dim lLevel As Integer = 1
            While lNextProjectFileEntry.StartsWith("  ")
                lNextProjectFileEntry = lNextProjectFileEntry.Substring(2)
                lFileName &= mHeadingWord(lLevel) & "\"
                lLevel += 1
            End While
            lNextProjectFileEntry = lNextProjectFileEntry.Trim
            mHeadingWord(lLevel) = lFileName
            Dim lPos As Integer = 0
            Dim lChar As String = lNextProjectFileEntry.Substring(lPos, 1)
            Dim lCharAsInt As Integer = Asc(lChar)
            While lCharAsInt > 31 And lCharAsInt < 127 ' > 47 And ach < 58 Or ach > 64 And ach < 91 Or ach > 96 And ach < 123 Or ach = 92 'alphanumeric or \
                If lCharAsInt = 34 Or lCharAsInt = 42 Or lCharAsInt = 47 Or lCharAsInt = 58 Or lCharAsInt = 60 Or lCharAsInt = 62 Or lCharAsInt = 63 Or lCharAsInt = 124 Then 'illegal for file names
                    lFileName &= "_"
                Else
                    lFileName &= lChar
                End If
                lPos += +1
                If lPos < lNextProjectFileEntry.Length Then
                    lChar = lNextProjectFileEntry.Substring(lPos, 1)
                    lCharAsInt = Asc(lChar)
                Else
                    lCharAsInt = 0
                End If
            End While
            If lFileName.Length > mHeadingWord(lLevel).Length Then
                mHeadingWord(lLevel) = Mid(lFileName, 1 + Len(mHeadingWord(lLevel)))
                mHeadingLevel = lLevel
                Return lFileName & pSourceExtension
            Else
                Return ""
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
        If mBuildContents Then
            If pOutputFormat = outputType.tHTML Then
                mHTMLContentsfile = FreeFile
                FileOpen(mHTMLContentsfile, mSaveDirectory & "Contents.html", OpenMode.Output)
                PrintLine(mHTMLContentsfile, "<html><head><title>" & pBaseName & " Help Contents</title></head>")
                PrintLine(mHTMLContentsfile, "<body>")
                PrintLine(mHTMLContentsfile, "<h1>Contents</h1>")
            ElseIf pOutputFormat = outputType.tHTMLHELP Then
                mHTMLContentsfile = FreeFile
                FileOpen(mHTMLContentsfile, mSaveDirectory & pBaseName & ".hhc", OpenMode.Output)
                PrintLine(mHTMLContentsfile, "<html><head><!-- Sitemap 1.0 --></head>")
                PrintLine(mHTMLContentsfile, "<body>")
                PrintLine(mHTMLContentsfile, "<OBJECT type=""text/site properties"">")
                PrintLine(mHTMLContentsfile, "<param name=""ImageType"" value=""Folder"">")
                PrintLine(mHTMLContentsfile, "</OBJECT>")

                mHTMLIndexfile = FreeFile
                FileOpen(mHTMLIndexfile, mSaveDirectory & pBaseName & ".hhk", OpenMode.Output)
                PrintLine(mHTMLIndexfile, "<html><head></head>")
                PrintLine(mHTMLIndexfile, "<body>")
                PrintLine(mHTMLIndexfile, "<ul>")

            ElseIf pOutputFormat = outputType.tHELP Then
                With pWordBasic
                    'Header of contents file
                    .FileNewDefault()
                    .Insert(":Title " & pBaseName & " Help" & vbCr)
                    .Insert(":Base " & pBaseName & ".hlp" & vbCr)
                    .ChDir(mSaveDirectory)
                    .FileSaveAs(pBaseName & ".cnt", 2)
                    .ChDir(mSourceBaseDirectory)
                    mContentsWindowName = .WindowName()
                End With
            End If
        End If
    End Sub

    Sub Init()
        mTotalTruncated = 0
        mTotalRepeated = 0
        mBodyTag = "<body>"
        'mPromptForFiles = True
        mNotFirstPrintHeader = False
        mNotFirstPrintFooter = False
        mInsertParagraphsAroundImages = False
        mHelpSourceRTFName = pBaseName & ".rtf"
        mTablePrintFormat = 0
        mTablePrintApply = 511
        mIDfile = -1
        mHTMLContentsfile = -1
        mHTMLIndexfile = -1
        If mAlreadyInitialized Then
            Exit Sub
        End If
        mAlreadyInitialized = True
        'mPromptForFiles = True
        mLinkToImageFiles = 0 '2 ' make soft links with word95 and large document
        frmConvert.CmDialog1Open.DefaultExt = "doc"
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        frmConvert.Show()
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        'IconLevel = 999

        mWholeCardHeader = mAsterisks80 & vbCrLf & mTensPlace & vbCrLf & mOnesPlace
        mWholeCardHeaderLength = mWholeCardHeader.Length

        'set default HTML styles
        For lHeaderLevel As Integer = 0 To mMaxLevels
            mHeaderStyle(lHeaderLevel) = "<hr size=7><h2><sectionname></h2><hr size=7>"
            mFooterStyle(lHeaderLevel) = ""
            mBodyStyle(lHeaderLevel) = "<body>"
            mWordStyle(lHeaderLevel) = New Collection
        Next
    End Sub

    Sub SetUnInitialized()
        mAlreadyInitialized = False
    End Sub

    Private Sub InsertParagraphsInPRE(ByRef aOutputFormat As OutputType)
        With pWordBasic
            If aOutputFormat = OutputType.tPRINT Or aOutputFormat = OutputType.tHELP Then
                .StartOfDocument()
                .EditFindClearFormatting()
                .EditFind("<pre>", "", 0)
                While .EditFindFound
                    .CharRight()
                    .EditBookmark("Hstart")
                    .EditFind("</pre>")
                    If Not .EditFindFound Then
                        Exit Sub
                    End If
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

    Private Sub ApplyWordFormat(ByRef aTag As String, ByRef aOutputFormat As OutputType, ByRef aDivTags As String)
        With pWordBasic
            Select Case aTag.ToLower
                Case "sub" : .Subscript()
                Case "sup" : .Superscript()
                Case "b" : .FormatFont(Bold:=1)
                Case "i" : .FormatFont(Italic:=1)
                Case "u"
                    If aOutputFormat = OutputType.tHELP Then
                        .FormatFont(Bold:=1)
                    Else
                        .FormatFont(Underline:=1)
                    End If
                Case "figure"
                    Dim lCaption As String = .Selection
                    .EditClear()
                    .InsertCaption("Figure", "", ": " & lCaption, Position:=1)
                    .Insert(vbCr)
                Case "pre"
                    .FormatParagraph(After:=0)
                    .FormatFont(Font:="Courier New", Points:=9.5)
                Case "div"
                    If InStr(aDivTags, "left") > 0 Then .FormatParagraph(Alignment:=0)
                    If InStr(aDivTags, "center") > 0 Then .FormatParagraph(Alignment:=1)
                    If InStr(aDivTags, "right") > 0 Then .FormatParagraph(Alignment:=2)
                    If InStr(aDivTags, "justify") > 0 Then .FormatParagraph(Alignment:=3)
            End Select
        End With
    End Sub

    Private Sub FormatTag(ByRef aTag As String, ByRef aOutputFormat As outputType)
        Dim lBeginTag As String
        If aTag = "div" Then
            lBeginTag = "<" & aTag & " "
        Else
            lBegintag = "<" & aTag & ">"
        End If
        Dim lEndtag As String = "</" & aTag & ">"

        Status("Formatting HTML " & lBegintag)

        Select Case aOutputFormat
            Case outputType.tPRINT, outputType.tHELP
                With pWordBasic
                    Logger.Dbg("WordCount:" & pWordApp.ActiveDocument.Words.Count)
                    .StartOfDocument()
                    .EditFindClearFormatting()
                    .EditFind(lBeginTag, "", 0)
                    While .EditFindFound
                        .EditClear() 'delete beginTag
                        .EditBookmark("Hstart")
                        Dim lDivArgs As String
                        If aTag = "div" Then
                            .EditFind(">")
                            If Not .EditFindFound Then Exit Sub
                            .EditClear()
                            .ExtendSelection()
                            .EditGoTo("Hstart")
                            lDivArgs = .Selection.Trim.ToLower
                            .EditClear()
                            .Insert(vbCr)
                            .EditBookmark("Hstart")
                        Else
                            lDivArgs = ""
                        End If
                        .EditFind(lEndtag)
                        If Not .EditFindFound Then Exit Sub
                        .EditClear() 'delete endTag
                        .EditBookmark("Hend")
                        .ExtendSelection()
                        .EditGoTo("Hstart")
                        .Cancel()
                        ApplyWordFormat(aTag, aOutputFormat, lDivArgs)
                        .CharRight()
                        .EditFind(lBeginTag, "", 0)
                    End While
                End With
            Case outputType.tHTML, outputType.tHTMLHELP
                Dim lStartTag As Integer = InStr(mTargetText.ToLower, lBeginTag)
                While lStartTag > 0
                    Dim lCloseTag As Integer = InStr(lStartTag + 2, mTargetText.ToLower, lEndtag)
                    Dim lInsertText As String = ""
                    If lCloseTag > 0 Then
                        Dim lBeginTagLength As Integer = lBegintag.Length
                        Dim lTaggedText As String = Mid(mTargetText, lStartTag + lBeginTagLength, lCloseTag - (lStartTag + lBeginTagLength))
                        Select Case LCase(aTag)
                            Case "b"
                                If InStr(lTaggedText, "<") > 0 Then
                                    lInsertText = ""
                                ElseIf InStr(lTaggedText, ">") > 0 Then
                                    lInsertText = ""
                                Else
                                    lInsertText = "<a name=""" & lTaggedText & """>" 'Insert link target for bold text
                                    If aOutputFormat = outputType.tHTMLHELP Then 'Insert bold text in index
                                        lInsertText = lInsertText & "<indexword=""" & lTaggedText & """>"
                                    End If
                                    mTargetText = Left(mTargetText, lStartTag - 1) & lInsertText & Mid(mTargetText, lStartTag)
                                End If
                        End Select
                    End If
                    lStartTag = InStr(lStartTag + lInsertText.Length + 2, mTargetText.ToLower, lBeginTag)
                End While
        End Select
    End Sub

    Private Sub FormatHeadings(ByRef aOutputFormat As outputType, ByRef aTargetFilename As String)
        Dim localHeadingLevel As Integer
        Dim direction As Integer
        Dim Selection As String
        Dim startTag As Integer
        Dim startNumber As Integer
        Dim endtag As Integer
        Dim closeTag As Integer
        Dim CloseTagEnd As Integer
        If aOutputFormat = outputType.tPRINT Or aOutputFormat = outputType.tHELP Then
            FormatHeadingsWithWord(aOutputFormat, aTargetFilename)
        Else

            mBodyTag = mBodyStyle(mHeadingLevel)
            startTag = InStr(LCase(mTargetText), "<body")
            If startTag > 0 Then
                endtag = InStr(startTag, mTargetText, ">")
                If endtag > startTag Then
                    mBodyTag = Mid(mTargetText, startTag, endtag - startTag + 1)
                    mTargetText = Left(mTargetText, startTag - 1) & Mid(mTargetText, endtag + 1)
                End If
            End If

            startTag = InStr(LCase(mTargetText), "<h")
            While startTag > 0
                If localHeadingLevel = 0 Then localHeadingLevel = mHeadingLevel
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

                mHeadingText(localHeadingLevel) = Trim(Mid(mTargetText, endtag + 1, closeTag - endtag - 1))
                'If HeadingText(localHeadingLevel) = "Duration" Then Stop
                mHeadingFile(localHeadingLevel) = mSaveFilename

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
                FormatHeadingHTML(localHeadingLevel, aTargetFilename, startTag) 'Insert header and adjust startTag to end of header
                mLastHeadingLevel = localHeadingLevel
NextHeader:
                startTag = InStr(startTag + 2, LCase(mTargetText), "<h")
            End While
        End If
    End Sub

    Private Sub FormatHeadingsWithWord(ByRef aOutputFormat As outputType, ByRef aTargetFilename As String)
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
                If localHeadingLevel = 0 Then localHeadingLevel = mHeadingLevel
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
                mHeadingText(localHeadingLevel) = Trim(.Selection)
                'If HeadingText(localHeadingLevel) = "Duration" Then Stop
                mHeadingFile(localHeadingLevel) = mSaveFilename
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

                If aOutputFormat = outputType.tPRINT Then
                    FormatHeadingPrint(localHeadingLevel)
                ElseIf aOutputFormat = outputType.tHELP Then
                    FormatHeadingHelp(localHeadingLevel)
                Else
                    .Cancel()
                End If
                mLastHeadingLevel = localHeadingLevel
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

    Sub TranslateButtons(ByRef aOutputFormat As outputType)
        If mCuteButtons AndAlso _
           (aOutputFormat = OutputType.tHELP Or _
            aOutputFormat = OutputType.tHTML Or _
            aOutputFormat = OutputType.tHTMLHELP) Then
            With pWordBasic
                .StartOfDocument()
                .EditFind("' button")
                While .EditFindFound
                    .CharLeft()
                    .EditBookmark("LabelEnd")
                    .EditFind("'", Direction:=1)
                    .CharRight()
                    .ExtendSelection()
                    .EditGoTo("LabelEnd")
                    Dim lLabel As String = .Selection
                    If Len(lLabel) < 20 Then
                        .EditClear(2)
                        .CharLeft()
                        .EditClear()
                        If aOutputFormat = OutputType.tHELP Then
                            .Insert("{button " & lLabel & ",}")
                        ElseIf aOutputFormat = OutputType.tHTML Or aOutputFormat = OutputType.tHTMLHELP Then
                            .Insert("<input type=submit value=""" & lLabel & """>")
                        Else 'should not get here
                            .Insert("'" & lLabel & "'")
                        End If
                    Else
                        Status("false alarm, not a button")
                    End If
                    .EditFind("' button", Direction:=0)
                End While
                .EditFind(PatternMatch:=0)
            End With
        End If
    End Sub

    Sub TranslateLists(ByRef aTag As String, ByRef aMarkerType As Integer)
        Dim begintag, endtag As String
        Dim bulletNumber As Integer

        begintag = "<" & aTag & ">"
        endtag = "</" & aTag & ">"
        With pWordBasic
            .StartOfDocument()
            Status("Translating HTML <" & aTag & ">")
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
                        .FormatBulletsAndNumbering(Hang:=1, Preset:=aMarkerType)
                        If aTag = "ol" Then
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
        Dim lIgnoreUnknown As Boolean
        If pOutputFormat = OutputType.tPRINT Or pOutputFormat = OutputType.tHELP Then
            NumberHeaderTagsWithWord()
        Else
            Dim lStartPos As Integer = InStr(LCase(mTargetText), "<h")
            If lStartPos = 0 Then 'need to insert section header
                mTargetText = "<h" & mHeadingLevel & ">" & mHeadingWord(mHeadingLevel) & "</h" & mHeadingLevel & "> " & vbLf & mTargetText
            End If
            Dim lCurTag As String = "<h"
            lStartPos = InStr(LCase(mTargetText), lCurTag)
            Dim lLocalHeadingLevel As Integer = mHeadingLevel
            While lStartPos > 0
                Dim lEndPos As Integer = InStr(lStartPos, mTargetText, ">")
                If lEndPos > 0 Then
                    Dim lSelStr As String = Mid(mTargetText, lStartPos + Len(lCurTag), lEndPos - lStartPos - Len(lCurTag))
                    Dim lDirection As Integer = 0
                    If lSelStr = "" Then
                        mTargetText = Left(mTargetText, lStartPos - 1) & lCurTag & lLocalHeadingLevel & Mid(mTargetText, lEndPos)
                    Else
                        Select Case Left(lSelStr, 1)
                            Case "+", "-"
                                If Left(lSelStr, 1) = "+" Then lDirection = 1 Else lDirection = -1
                                lSelStr = Mid(lSelStr, 2)
                                If IsNumeric(lSelStr) Then
                                    lLocalHeadingLevel = mHeadingLevel + lDirection * CShort(lSelStr)
                                Else
                                    lLocalHeadingLevel = mHeadingLevel + lDirection
                                End If
                                mTargetText = Left(mTargetText, lStartPos + 1) & lLocalHeadingLevel & Mid(mTargetText, lEndPos)
                            Case Else
                                If IsNumeric(lSelStr) Then
                                    lLocalHeadingLevel = CShort(lSelStr)
                                ElseIf UCase(lSelStr) = "EAD" Or UCase(lSelStr) = "TML" Then
                                    'ignore <html> and <head> even though these should not be in source files
                                ElseIf UCase(Left(lSelStr, 1)) = "R" Then
                                    'ignore <hr> <hr size=7> etc.
                                Else
                                    If Not lIgnoreUnknown Then
                                        If Logger.Msg("Unknown heading tag '<h" & lSelStr & ">'" & vbCr & "In file " & mSourceFilename & vbCr & "Warn about future unknown headers?", MsgBoxStyle.YesNo, "Number Header Tags") = MsgBoxResult.No Then
                                            lIgnoreUnknown = True
                                        End If
                                    End If
                                End If
                        End Select
                    End If
                    If lCurTag = "<h" Then lCurTag = "</h" Else lCurTag = "<h"
                    lStartPos = InStr(lEndPos, LCase(mTargetText), lCurTag)
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
                .Insert("<h" & mHeadingLevel & ">" & mHeadingWord(mHeadingLevel) & "</h" & mHeadingLevel & "> " & vbLf)
            End If
            curTag = "<h"
            .EditGoTo("CurrentFileStart")
            .EditFind(curTag, "", 0)
            While .EditFindFound
                .CharRight()
                .ExtendSelection()
                .EditFind(">", "", 0)
                localHeadingLevel = mHeadingLevel
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
                                    localHeadingLevel = mHeadingLevel + direction * CShort(selStr)
                                Else
                                    localHeadingLevel = mHeadingLevel + direction
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

    Private Sub TranslateIMGtags(ByRef aPath As String)
        With pWordBasic
            .EditFind("<IMG ", "", 0)
            While .EditFindFound
                Dim lInsertParagraphs As Boolean = mInsertParagraphsAroundImages
                Dim lLinkToThisImageFile As Integer = mLinkToImageFiles
                .EditClear()
                '.EditBookmark("ImgStart")
                .EditFind("SRC=""", "", 0)
                .EditClear()
                .EditBookmark("LinkStart")
                .EditFind("""")
                If Not .EditFindFound Then
                    Exit Sub
                End If
                .EditClear()
                .ExtendSelection()
                .EditGoTo("LinkStart")
                'curpath = path
                Dim lCurrentFilename As String = .Selection
                .EditClear()
                'If .ExistingBookmark("ImgStart") <> 0 Then
                '    .EditGoTo("ImgStart")
                'Else
                '    Logger.Dbg("No ImgStart Bookmark for " & lCurrentFilename)
                'End If
                Dim lLinkFilename As String = aPath & "\" & lCurrentFilename
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
                If InStr(1, lLinkFilename, "icon", 1) > 0 Then
                    lInsertParagraphs = False
                    'LinkToThisImageFile = 0
                End If
                If .Selection.Length > 1 Then
                    lInsertParagraphs = False 'probably said ALIGN=LEFT in img tag
                End If
                .EditClear()
                .Cancel()
                If lInsertParagraphs Then .Insert("<p>")
                Try
                    .InsertPicture(lLinkFilename, lLinkToThisImageFile)
                Catch lEx As Exception
                    Logger.Dbg("InsertPictureProblem:File:" & lLinkFilename & ":Message:" & lEx.Message)
                End Try

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

                If InStr(1, lLinkFilename, "icon", 1) > 0 Then
                    .CharLeft(1, 1)
                    .FormatFont(Position:=-12) 'half-point units
                    .CharRight()
                End If

                If lInsertParagraphs Then
                    .Insert("<p>")
                End If
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
        If mMoveHeadings <> 0 Then
            hn = "h" & (thisHeadingLevel + mMoveHeadings) & ">"
            mTargetText = Left(mTargetText, thisHeadingStart) & "<" & hn & mHeadingText(thisHeadingLevel) & "</" & hn & Mid(mTargetText, thisHeadingStart + 1)
        Else
            'Dim IconPath$

            ht = mHeadingText(thisHeadingLevel)
            TextToInsert = ""
            TextToPrepend = ""
            TextToAppend = ReplaceString(mFooterStyle(thisHeadingLevel), "<sectionname>", ht)
            '    IconPath = ""
            '    For i = IconLevel + 1 To thisHeadingLevel
            '      IconPath = "../" & IconPath
            '    Next i

            'Insert name anchor around heading
            TextToInsert = TextToInsert & "<a name=""" & ht & """>"
            TextToInsert = TextToInsert & ReplaceString(mHeaderStyle(thisHeadingLevel), "<sectionname>", ht) & "</a>" & vbLf
            '    If IconLevel <= thisHeadingLevel Then
            '      TextToInsert = TextToInsert & "<img src=""" & IconPath & IconFilename & """ align=right>"
            '    End If

            If mFirstHeaderInFile Then 'Insert navigation to parents in hierarchy
                mFirstHeaderInFile = False
                If mUpNext Then
                    LinkToFirstHeader = "Up to: <a href=""#" & ht & """>" & ht & "</a>" & vbLf
                    If thisHeadingLevel > 1 Or (pOutputFormat = outputType.tHTML And mBuildContents) Then
                        TextToInsert = TextToInsert & "Up to: "
                        For h = thisHeadingLevel - 1 To 1 Step -1
                            ParentHT = mHeadingText(h)
                            TextToInsert = TextToInsert & "<a href=""\" & mHeadingFile(h) & "#" & ParentHT & """>" & ParentHT & "</a>, "
                        Next h
                        If pOutputFormat = outputType.tHTML And mBuildContents Then
                            TextToInsert = TextToInsert & "<a href=""\Contents.html#" & ht & """>Contents</a>"
                        Else 'remove last ", "
                            TextToInsert = Left(TextToInsert, Len(TextToInsert) - 2)
                        End If
                        TextToInsert = TextToInsert & "<p>" & vbLf
                    End If
                End If
                TextToPrepend = mBeforeHTML & "<html><head><title>" & ht & "</title></head>" & vbCrLf & mBodyTag & vbCrLf '& "<form>" & vbCrLf
                TextToAppend = vbCrLf & "</body>" & vbCrLf & "</html>" & vbCrLf

                If pOutputFormat = outputType.tHTMLHELP Then
                    If InStr(mTargetText, "<param name=""Keyword"" value=""" & ht & """>") = 0 Then
                        TextToAppend = KeywordAnchor(ht) & TextToAppend
                    End If
                    If thisHeadingLevel = 5 Then
                        If Right(mHeadingText(3), 5) = "Block" Then
                            TextToAppend = AlinkAnchor(ht & Left(mHeadingText(3), Len(mHeadingText(3)) - 6)) & TextToAppend
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

            If mBuildContents Then HTMLContentsEntry(thisHeadingLevel, targetFilename, ht)

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
        For lvl = mLastHeadingLevel + 1 To thisHeadingLevel
            PrintLine(mHTMLContentsfile, Space((lvl - 1) * 4) & "<ul>")
        Next lvl
        For lvl = thisHeadingLevel + 1 To mLastHeadingLevel
            PrintLine(mHTMLContentsfile, Space((lvl - 1) * 4) & "</ul>")
        Next lvl
        PrintLine(mHTMLContentsfile, "<li>")

        Dim objdef As String
        If pOutputFormat = outputType.tHTML Then
            SafeFilename = ReplaceString(targetFilename, "\", "/")
            SafeFilename = ReplaceString(SafeFilename, " ", "%20")
            PrintLine(mHTMLContentsfile, Space((thisHeadingLevel - 1) * 4) & "<a name=""" & headerText & """>")
            PrintLine(mHTMLContentsfile, Space((thisHeadingLevel - 1) * 4) & "<a href=""" & SafeFilename & "#" & headerText & """>" & headerText & "</a></a>")
        ElseIf pOutputFormat = outputType.tHTMLHELP Then
            objdef = Space((thisHeadingLevel - 1) * 4) & "<li><OBJECT type=""text/sitemap"">" & vbCr
            objdef = objdef & Space((thisHeadingLevel) * 4) & "<param name=""Name"" value=""" & headerText & """>" & vbCr
            objdef = objdef & Space((thisHeadingLevel) * 4) & "<param name=""Local"" value=""" & targetFilename & """>" & vbCr
            objdef = objdef & Space((thisHeadingLevel) * 4) & "</OBJECT>" & vbCr
            PrintLine(mHTMLContentsfile, objdef)
            PrintLine(mHTMLIndexfile, objdef)

            id = MakeValidHelpID(headerText)

            If mBuildID Then
                PrintLine(mIDfile, "#define " & id & vbTab & mIDnum)
                mIDnum = mIDnum + 1
            End If
            If mBuildProject Then
                mAliasSection = mAliasSection & vbLf & id & " = " & targetFilename
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
        For Each cmd In mWordStyle(thisHeadingLevel)
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
            If mNotFirstPrintHeader Then
                Try
                    .FormatHeaderFooterLink()
                Catch
                    pWordBasic.ToggleHeaderFooterLink()
                End Try
            Else
                mNotFirstPrintHeader = True
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
            If mNotFirstPrintFooter Then
                On Error GoTo toggle
                .FormatHeaderFooterLink()
                On Error GoTo 0
            Else
                mNotFirstPrintFooter = True
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
            For level = 1 To mMaxLevels
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
        topic = mHeadingText(thisHeadingLevel)
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
            If mBuildContents Then HelpContentsEntry(topic, id, thisHeadingLevel)
            If mBuildID Then
                PrintLine(mIDfile, "#define " & id & vbTab & mIDnum)
                mIDnum = mIDnum + 1
            End If
            On Error GoTo NoPrevSection
            .EditBookmark("temp")
            If mUpNext Then
                If mLastHeadingLevel >= thisHeadingLevel Then
                    .EditGoTo("UpFrom" & mLastHeadingLevel)
                    .Insert(vbTab & "Next: ")
                    HelpHyperlink(topic, id)
                End If
                For h = thisHeadingLevel To mLastHeadingLevel - 1
                    .EditGoTo("UpFrom" & h)
                    .Insert(vbTab & "Next: ")
                    HelpHyperlink(topic, id)
                Next h
            End If
NoPrevSection:
            .EditGoTo("temp")
            On Error GoTo 0

            If thisHeadingLevel > 1 Then 'Insert navigation to/from parents in hierarchy
                If mUpNext Then
                    .Insert("Up to: ")
                    For h = thisHeadingLevel - 1 To 1 Step -1
                        parentTopic = mHeadingText(h)
                        HelpHyperlink(parentTopic, MakeValidHelpID(parentTopic))
                        If h > 1 Then .Insert(", ")
                    Next h
                    .Insert(vbCr)
                    .EditBookmark("SectionContents" & thisHeadingLevel)
                    .CharLeft()
                    .EditBookmark("UpFrom" & thisHeadingLevel)

                    'insert entry in section contents of parent topic
                    If mHeadingText(thisHeadingLevel - 1) <> "Tutorial" Then
                        mContentsEntries(thisHeadingLevel - 1) = mContentsEntries(thisHeadingLevel - 1) + 1
                        .EditGoTo("SectionContents" & thisHeadingLevel - 1)
                        If mContentsEntries(thisHeadingLevel - 1) = 1 Then
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
            mContentsEntries(thisHeadingLevel) = 0
        End With
    End Sub

    Sub FinishHTMLHelpContents()
        Dim lvl As Integer
        For lvl = mHeadingLevel To 1 Step -1
            PrintLine(mHTMLContentsfile, Space((lvl - 1) * 4) & "</ul>")
        Next lvl
        PrintLine(mHTMLContentsfile, "</body>")
        PrintLine(mHTMLContentsfile, "</html>")
        FileClose(mHTMLContentsfile)

        If mHTMLIndexfile >= 0 Then
            PrintLine(mHTMLIndexfile, "</ul>")
            PrintLine(mHTMLIndexfile, "</body>")
            PrintLine(mHTMLIndexfile, "</html>")
            FileClose(mHTMLIndexfile)
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
            .Activate(mContentsWindowName)
            If mLastHeadingLevel = 0 Then mLastHeadingLevel = 1
            .Insert(thisHeadingLevel & " " & topic & "=" & id & vbCr)
            If thisHeadingLevel < mLastHeadingLevel Then mBookLevel = thisHeadingLevel
            If thisHeadingLevel > mLastHeadingLevel And mBookLevel < mLastHeadingLevel Or mBookLevel = thisHeadingLevel Then
                If thisHeadingLevel > mLastHeadingLevel Then
                    numlines = 2
                    mBookLevel = mLastHeadingLevel
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
            .Activate(mTargetWindowName)
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
            For LevelCount = 1 To mHeadingLevel - 1
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
        Dim lSectionContents As String = ""
        Dim lLocalNextIndex As Integer = mNextProjectFileIndex
        Dim nextLevel, lvl, prevLevel As Integer
        Dim nextName As String = ""
        Dim nextHref As String = ""
        Dim localHeadingWord(10) As String

        localHeadingWord(mHeadingLevel) = mHeadingWord(mHeadingLevel)
        prevLevel = mHeadingLevel
        GetNextEntryLevel(lLocalNextIndex, nextLevel, nextName, nextHref, localHeadingWord)
        If nextLevel > mHeadingLevel Then
            While nextLevel > mHeadingLevel

                If nextLevel > prevLevel Then
                    For lvl = prevLevel To (nextLevel - 1)
                        lSectionContents &= "<ul>" & vbCr
                    Next
                ElseIf nextLevel < prevLevel Then
                    For lvl = nextLevel To (prevLevel - 1)
                        lSectionContents &= "</ul>" & vbCr
                    Next
                End If

                lSectionContents &= "<li><a href=""" & nextHref & """>" & nextName & "</a>" & vbCr
                prevLevel = nextLevel
                GetNextEntryLevel(lLocalNextIndex, nextLevel, nextName, nextHref, localHeadingWord)
            End While
            lSectionContents &= "</ul>" & vbCr
        End If
        Return lSectionContents
    End Function

    Private Sub GetNextEntryLevel(ByRef aLocalNextEntry As Integer, _
                                  ByRef aNextLevel As Integer, _
                                  ByRef aNextName As String, _
                                  ByRef aNextHref As String, _
                                  ByRef aLocalHeadingWord() As String)
        If aLocalNextEntry >= mProjectFileEntrys.Count Then
            aNextLevel = 0
        Else
            aNextName = mProjectFileEntrys(aLocalNextEntry).ToString.TrimStart
            aNextLevel = (mProjectFileEntrys(aLocalNextEntry).ToString.Length - aNextName.ToString.Length) / 2 + 1
            aNextName = aNextName.TrimEnd
            aLocalHeadingWord(aNextLevel) = aNextName
            aNextHref = ""
            For lLevel As Integer = mHeadingLevel To aNextLevel - 1
                aNextHref = aNextHref & aLocalHeadingWord(lLevel) & "\"
            Next
            aNextHref &= aNextName & ".html"
            aLocalNextEntry += 1
        End If
    End Sub

    Private Sub CopyImages()
        Status("Copying Images")
        Dim lSrcPath As String = IO.Path.GetDirectoryName(mSourceBaseDirectory & mSourceFilename) & "\"
        Dim lDstPath As String = IO.Path.GetDirectoryName(mSaveDirectory & mSaveFilename) & "\"
        Dim lIgnoreAll As Boolean = False
        Dim lEndPos As Integer
        Dim lStartPos As Integer = -1
        While Assign(lStartPos, mTargetText.IndexOf(" src=""", lStartPos + 1, StringComparison.OrdinalIgnoreCase)) > 0
            lEndPos = mTargetText.IndexOf("""", lStartPos + 6)
            If lEndPos = 0 Then Exit Sub
            Dim ImageFilename As String = mTargetText.Substring(lStartPos + 6, lEndPos - lStartPos - 6)
CheckForImage:
            If IO.File.Exists(lSrcPath & ImageFilename) Then
                MkDirPath(IO.Path.GetDirectoryName(AbsolutePath(ReplaceString(ImageFilename, "/", "\"), lDstPath)))
                FileCopy(lSrcPath & ImageFilename, lDstPath & ImageFilename)
            ElseIf Not lIgnoreAll Then
                Select Case Logger.Msg("Missing image: " & vbCr & lSrcPath & ImageFilename, MsgBoxStyle.AbortRetryIgnore, "AuthorDoc")
                    Case MsgBoxResult.Abort : Exit Sub
                    Case MsgBoxResult.Retry : GoTo CheckForImage
                    Case MsgBoxResult.Ignore
                        If Logger.Msg("Ignore all missing images?", MsgBoxStyle.YesNo, "AuthorDoc") = MsgBoxResult.Yes Then
                            lIgnoreAll = True
                        End If
                End Select
            End If

            If pOutputFormat = OutputType.tHTML Then
                Dim lHTMLSafeFilename As String = ReplaceString(ImageFilename, "\", "/")
                lHTMLSafeFilename = ReplaceString(lHTMLSafeFilename, " ", "%20")
                If lHTMLSafeFilename <> lHTMLSafeFilename Then
                    mTargetText = Left(mTargetText, lStartPos + 5) & lHTMLSafeFilename & Mid(mTargetText, lEndPos)
                End If
            End If
        End While
    End Sub

    Private Sub HREFsInsureExtension()
        Dim LinkFile, LinkRef, LinkTopic As String
        Dim endPos, startPos, pos As Integer

        If pOutputFormat = OutputType.tHELP Then
            HREFsInsureExtensionWithWord()
        ElseIf pOutputFormat = OutputType.tPRINT Then
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
                    If pOutputFormat = OutputType.tHTML Then
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