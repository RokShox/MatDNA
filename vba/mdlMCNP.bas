Attribute VB_Name = "mdlMCNP"
'
' Copyright (c) 2017 Henkel Technical Computing, LLC. All rights reserved.
'
Option Explicit

Public mdSurfaces As Scripting.Dictionary

'
' writeMCNPRangeToFile
'
Public Sub writeMCNPRangeToFile(rng As Range, iFile As Integer, colPatSubst As Collection)
    Dim wkb As Workbook
    Dim sLine As String
    Dim iLen As Integer
    Dim lrow As Long
    Dim sToken As String
    Dim sDelim As String
    Dim iPos As Integer
    Dim rngInsert As Range
    Dim sInclude As String
    Dim colPatSubstInsert As Collection
    Dim wksThis As Worksheet
    Dim lCol As Long
    Dim nCol As Long
    Dim sOut As String
    Dim sFunc As String
    Const LINELEN As Integer = 132
    
    Set wksThis = rng.Worksheet
    sDelim = Chr(32)
    wksThis.Activate
    
    On Error Resume Next
    
    nCol = rng.Columns.Count
    For lrow = 1 To rng.Rows.Count
    
        sToken = Trim(rng.Cells(lrow, 1).Text)
        
        ' Check for not used
        If sToken = "n/u" Or sToken = "not used" Or sToken = "c not used" Then
            ' Skip this line
            
        ' Check for included range
        ElseIf Left(sToken, 1) = "#" Then
            ' Put the whole line together
            sLine = ""
            For lCol = 1 To nCol
                sLine = sLine & Trim(rng.Cells(lrow, lCol).Text)
                If lCol < nCol Then
                    sLine = sLine & sDelim
                End If
            Next lCol
            sLine = RTrim(sLine)
            parseInsert sLine, rngInsert, colPatSubstInsert, wksThis
            writeMCNPRangeToFile rngInsert, iFile, colPatSubstInsert
            wksThis.Activate
        ElseIf Left(sToken, 1) = "<" Then
            ' Check for function to run
            sFunc = Mid(sToken, 2)
'            If sFunc = "getSurfs" Then
'                getSurfs
'            ElseIf sFunc = "getCells" Then
'                getCells
'            End If
        
        Else
            ' Normal line
            ' Fix up first part of line
            If Left(sToken, 1) = "c" And Len(sToken) = 1 Then
                sToken = "c    "
            ElseIf Len(sToken) <> 0 Then
                iLen = Len(sToken)
                While iLen < 5
                    sToken = sToken & " "
                    iLen = Len(sToken)
                Wend
                ' if last char is not blank add a space
                If Right(sToken, 1) <> " " Then
                    sToken = sToken & sDelim
                End If
            Else
                sToken = "     "
            End If

            ' Put the rest of the line together
            sLine = ""
            For lCol = 2 To nCol
                sLine = sLine & Trim(rng.Cells(lrow, lCol).Text)
                If lCol < nCol Then
                    sLine = sLine & sDelim
                End If
            Next lCol
            sLine = RTrim(sLine)
            
            ' Prefix first part to line
            sLine = sToken & sLine
        
            doPatSubst sLine, sOut, colPatSubst
        
            ' If line too long, see if it can be shortened
            If Len(sOut) > LINELEN Then
                ' Check and see if a comment can be abbreviated
                If Left(sOut, 5) = "c    " Then
                    sOut = Mid(sOut, 1, LINELEN)
                    ' Replace last char with a vertical bar to indicate abbreviation
                    Mid(sOut, LINELEN, 1) = "|"
                Else
                    ' Try to shorten inline comment
                    iPos = InStr(1, sOut, "$")
                    If iPos = 0 Or iPos > LINELEN Then
                        MsgBox "Error: Line longer than " & LINELEN & " cols" & Chr(10) & sOut, vbCritical
                        Exit Sub
                    Else
                        sOut = Mid(sOut, 1, LINELEN)
                        ' Replace last char with a vertical bar to indicate abbreviation
                        Mid(sOut, LINELEN, 1) = "|"
                    End If
                End If
            End If
            Print #iFile, sOut
        End If
    Next lrow
End Sub

Public Sub getSurfs()
    Dim rngBlockRoom As Range
    Dim rngBlockPost As Range
    Dim rngPre As Range
    Dim rngTblRoom As Range
    Dim rngSurface As Range
    Dim colPatSubst As Collection
    Dim cRoom As clsBlock
    Dim cPost As clsBlock
    Dim iRoom As Integer
    Dim iFirst As Integer
    Dim rngStart As Range
    Dim rngFormat As Range
    Dim rngBgn As Range
    Dim rngInit As Range
    Dim nCol As Integer
    Dim nRow As Integer
    Dim nRoom As Integer
    
    Set rngPre = ActiveSheet.Names("!inpSurfacePre").RefersToRange
    Set rngBlockRoom = ActiveSheet.Names("blockSurfaceRoom").RefersToRange
    Set rngBlockPost = ActiveSheet.Names("blockSurfacePost").RefersToRange
    Set rngTblRoom = ActiveSheet.Names("tblRoom").RefersToRange
    Set rngSurface = ActiveSheet.Names("inpSurface").RefersToRange
    Set rngStart = rngSurface.Cells(1, 1)
    Set rngBgn = rngStart
    Set rngInit = rngStart
    nCol = rngPre.Columns.Count
    nRoom = rngTblRoom.Rows.Count
    
    rngSurface.Cells(1, 1).Resize(rngSurface.Rows.Count + 1, rngSurface.Columns.Count).Clear
    
    Application.StatusBar = "Pre surfaces"
    
    writeMCNPRangeToExcel rngPre, rngStart, Nothing, True
    
    Set cRoom = New clsBlock
    Set cRoom.BlockRange = rngBlockRoom
        
    iFirst = 1
    Set rngBgn = rngStart
    For iRoom = 1 To nRoom
    ' For iRoom = 1 To 2
        With cRoom
            .Parameter("row") = iRoom
            .Parameter("firstSurface") = iFirst
            .SetInputParameters
            .GetOutputParameters
        End With
        
        Application.StatusBar = "Row " & Format(iRoom, "#0") & "/" & nRoom & " Room " & cRoom.Parameter("room") & " surfaces"
        
        writeMCNPRangeToExcel cRoom.BlockRange, rngStart, Nothing, False
        
        iFirst = CInt(cRoom.Parameter("lastSurface")) + 1
    Next iRoom
    
    ' Copy formats separately
    Set rngFormat = rngBgn.Resize(cRoom.BlockRange.Rows.Count * nRoom, cRoom.BlockRange.Columns.Count)
    CopyRangeFormats cRoom.BlockRange, rngFormat

    Set cPost = New clsBlock
    With cPost
        Set .BlockRange = rngBlockPost
        .Parameter("firstSurface") = iFirst
        .SetInputParameters
        .GetOutputParameters
    End With
    
    Application.StatusBar = "Post surfaces"
    
    writeMCNPRangeToExcel cPost.BlockRange, rngStart, Nothing, True
        
    nRow = rngStart.Cells(1, 1).Row - rngInit.Cells(1, 1).Row
    ActiveSheet.Names("inpSurface").Delete
    ActiveSheet.Names.Add "inpSurface", rngInit.Resize(nRow, nCol)
    Set rngSurface = ActiveSheet.Names("inpSurface").RefersToRange
    rngSurface.Cells(nRow, nCol).Offset(1, 0).Value2 = "!inpSurface"
    InitSurfaces
    Application.Calculate
    
    Application.StatusBar = False

End Sub


Public Sub getCells()
    Dim rngBlockRoom As Range
    Dim rngBlockPost As Range
    Dim rngPre As Range
    Dim rngTblRoom As Range
    Dim rngCell As Range
    Dim colPatSubst As Collection
    Dim cRoom As clsBlock
    Dim cPost As clsBlock
    Dim iRoom As Integer
    Dim iFirst As Integer
    Dim rngStart As Range
    Dim rngBgn As Range
    Dim rngInit As Range
    Dim rngFormat As Range
    Dim nCol As Integer
    Dim nRow As Integer
    Dim nRoom As Integer
    Dim nAuxCols As Integer
    
    ' Don't include aux columns in final named range
    ' Currently 1 column is for ext parameters
    nAuxCols = 1
    
    Set rngPre = ActiveSheet.Names("!inpCellPre").RefersToRange
    Set rngBlockRoom = ActiveSheet.Names("blockCellRoom").RefersToRange
    Set rngBlockPost = ActiveSheet.Names("blockCellPost").RefersToRange
    Set rngTblRoom = ActiveSheet.Names("tblRoom").RefersToRange
    Set rngCell = ActiveSheet.Names("inpCell").RefersToRange
    Set rngStart = rngCell.Cells(1, 1)
    Set rngBgn = rngStart
    Set rngInit = rngStart
    nCol = rngPre.Columns.Count
    nRoom = rngTblRoom.Rows.Count
    
    rngCell.Cells(1, 1).Resize(rngCell.Rows.Count + 1, rngCell.Columns.Count).Clear
    
    Application.StatusBar = "Pre cells"
    
    writeMCNPRangeToExcel rngPre, rngStart, Nothing, True
    
    Set cRoom = New clsBlock
    Set cRoom.BlockRange = rngBlockRoom
    cRoom.ParameterOffset = -6
        
    Set rngBgn = rngStart
    iFirst = 1
    For iRoom = 1 To nRoom
    ' For iRoom = 1 To 2
        With cRoom
            .Parameter("row") = iRoom
            .Parameter("firstCell") = iFirst
            .SetInputParameters
            .GetOutputParameters
        End With
        
        Application.StatusBar = "Row " & Format(iRoom, "#0") & "/" & nRoom & " Room " & cRoom.Parameter("room") & " cells"
        
        writeMCNPRangeToExcel cRoom.BlockRange, rngStart, Nothing, False
        
        iFirst = CInt(cRoom.Parameter("lastCell")) + 1
    Next iRoom

    ' Copy formats separately
    Set rngFormat = rngBgn.Resize(cRoom.BlockRange.Rows.Count * nRoom, cRoom.BlockRange.Columns.Count)
    CopyRangeFormats cRoom.BlockRange, rngFormat

    Set cPost = New clsBlock
    With cPost
        Set .BlockRange = rngBlockPost
        .ParameterOffset = -6
        .Parameter("firstCell") = iFirst
        .SetInputParameters
        .GetOutputParameters
    End With
        
    Application.StatusBar = "Post cells"
    
    writeMCNPRangeToExcel cPost.BlockRange, rngStart, Nothing, True
        
    nRow = rngStart.Cells(1, 1).Row - rngInit.Cells(1, 1).Row
    ActiveSheet.Names("inpCell").Delete
    ActiveSheet.Names.Add "inpCell", rngInit.Resize(nRow, nCol - nAuxCols)
    Set rngCell = ActiveSheet.Names("inpCell").RefersToRange
    rngCell.Cells(nRow, nCol - nAuxCols).Offset(1, 0).Value2 = "!inpCell"
    Application.Calculate
    
    Application.StatusBar = False

End Sub

'
' writeMCNPRange
'
Public Sub writeMCNPRange()
    Dim sFile As String
    Dim rng As Range
    Dim iFile As Integer
    
    sFile = ThisWorkbook.Names("dirTmp").RefersToRange.Value & "\mcnp.txt"
    iFile = FreeFile
    Open sFile For Output As #iFile
    Set rng = Selection
    writeMCNPRangeToFile rng, iFile, Nothing
    Close #iFile

End Sub

'
' CopyRangeFormats
'
Public Sub CopyRangeFormats(rngSrc As Range, rngDest As Range)
    Dim blnScreen As Boolean

    blnScreen = Application.ScreenUpdating
    Application.ScreenUpdating = False
    
    rngSrc.Copy
    rngDest.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

    Application.ScreenUpdating = blnScreen

End Sub


'
' writeMCNPRangeToExcel
'
Public Sub writeMCNPRangeToExcel(rng As Range, rngStart As Range, colPatSubst As Collection, Optional blnFormats As Boolean = True)
    Dim wkb As Workbook
    Dim sLine As String
    Dim iLen As Integer
    Dim lrow As Long
    Dim sToken As String
    Dim sDelim As String
    Dim iPos As Integer
    Dim rngInsert As Range
    Dim sInclude As String
    Dim colPatSubstInsert As Collection
    Dim wksThis As Worksheet
    Dim lCol As Long
    Dim nCol As Long
    Dim sOut As String
    Dim blnScreen As Boolean
    Dim LastRow As Long
    Dim nRow As Long
    
    blnScreen = Application.ScreenUpdating
    Application.ScreenUpdating = False
    Set wksThis = rng.Worksheet
    sDelim = Chr(32)
    wksThis.Activate
    rng.Calculate
    LastRow = 0
    
    On Error Resume Next
    
    nCol = rng.Columns.Count
    For lrow = 1 To rng.Rows.Count
    
        sToken = Trim(rng.Cells(lrow, 1).Text)
        
        ' Include n/u lines
'        ' Check for not used
'        If sToken = "n/u" Or sToken = "not used" Or sToken = "c not used" Then
'            ' Skip this line
'
        ' Check for included range
        If Left(sToken, 1) = "#" Then
        
            ' Make pending copies. It takes too long to wtite out line by line
            ' So keep a marker pointing to the last copied row. Now that we hit an
            ' included range, can make the pending copies.
            If lrow - 1 > LastRow Then
                nRow = (lrow - 1) - LastRow
                rngStart.Resize(nRow, nCol).Value2 = rng.Cells(LastRow + 1, 1).Resize(nRow, nCol).Value2
                If blnFormats Then
                    rng.Cells(LastRow + 1, 1).Resize(nRow, nCol).Copy
                    rngStart.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
                End If
                Set rngStart = rngStart.Offset(nRow, 0)
                Application.StatusBar = rng.Name.Name & " " & nRow & " rows copied"
            End If
            LastRow = lrow
        
            ' Put the whole line together
            sLine = ""
            For lCol = 1 To nCol
                sLine = sLine & Trim(rng.Cells(lrow, lCol).Text)
                If lCol < nCol Then
                    sLine = sLine & sDelim
                End If
            Next lCol
            sLine = RTrim(sLine)
            parseInsert sLine, rngInsert, colPatSubstInsert, wksThis
            writeMCNPRangeToExcel rngInsert, rngStart, colPatSubstInsert
            wksThis.Activate
        ' Normal line
'        Else
'            rng.Cells(lrow, 1).Resize(1, nCol).Copy
'            'rngStart.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
'            rngStart.Resize(1, nCol).Value2 = rng.Cells(lrow, 1).Resize(1, nCol).Value2
'            rngStart.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
'            Application.StatusBar = rng.Name.Name & " " & lrow
'
'
'            'doPatSubstExcel rngStart, nCol, colPatSubst
'
'            Set rngStart = rngStart.Offset(1, 0)
         End If
    Next lrow
    
    If rng.Rows.Count > LastRow Then
        nRow = rng.Rows.Count - LastRow
        rng.Cells(LastRow + 1, 1).Resize(nRow, nCol).Copy
        rngStart.Resize(nRow, nCol).Value2 = rng.Cells(LastRow + 1, 1).Resize(nRow, nCol).Value2
        rngStart.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        Set rngStart = rngStart.Offset(nRow, 0)
        ' Application.StatusBar = rng.Name.Name & " " & nRow & " rows copied"
    End If
    
    Application.CutCopyMode = False
    Application.ScreenUpdating = blnScreen
    Application.StatusBar = False
End Sub


'
' MakeStudy
'
Public Sub MakeStudy()
    Dim sFile As String
    Dim iFile As Integer
    Dim rngInput As Range
    Dim sCase As String
    Dim sVersion As String
    Dim sDirCase As String
    Dim sDirRun As String
    Dim wks As Worksheet
       
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    Application.Calculate
    
    
    Set wks = ActiveSheet
    sCase = wks.Names("caseName").RefersToRange.Text
    sDirRun = ThisWorkbook.Names("dirRun").RefersToRange.Text
    sVersion = ThisWorkbook.Names("version").RefersToRange.Text
    sDirCase = sDirRun & "\" & sVersion
'
    SafeMkDir sDirCase
    Set rngInput = wks.Names("inpMCNP").RefersToRange
    
    sFile = sDirCase & "\" & sCase
    iFile = FreeFile
    Open sFile For Output As #iFile
    writeMCNPRangeToFile rngInput, iFile, Nothing
    Close #iFile
        
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
        
        
    WriteCaseEntry
End Sub


Public Sub WriteCaseEntry()
    Dim wksModel As Worksheet
    Dim i As Long
    Dim j As Long
    Dim rngStart As Range
    Dim rngResults As Range
    Dim rngTblCaseEntry As Range
    Dim sResults As String

    Set wksModel = ActiveSheet
    
    On Error GoTo NotFound
    Set rngTblCaseEntry = wksModel.Names("tblCaseEntry").RefersToRange
    
    sResults = rngTblCaseEntry.Cells(1, 3).Text
    
    If Left(sResults, 1) = "!" Then
        sResults = Mid(sResults, 2)
        Set rngResults = wksModel.Names(sResults).RefersToRange
    Else
        Set rngResults = ThisWorkbook.Names(sResults).RefersToRange
    End If
        
    ' Find first non-blank row
    Set rngStart = rngResults.Cells(1, 1)
    i = 0
    While rngStart.Offset(i, 0).Text <> "" And i < rngResults.Rows.Count
        i = i + 1
    Wend
    If i = rngResults.Rows.Count Then
        MsgBox "No more rows in " & sResults, vbCritical, "WriteCaseEntry Error"
        Exit Sub
    End If
    
    Set rngStart = rngStart.Offset(i, 0)
    
    For j = 2 To rngTblCaseEntry.Rows.Count
        If rngTblCaseEntry.Cells(j, 1).Text <> "" Then
            rngStart.Offset(0, j - 2).Value2 = rngTblCaseEntry.Cells(j, 3).Value2
        End If
    Next j
        
NotFound:
    
End Sub

Public Function MakeListParen(ParamArray vntVals() As Variant) As String
    Dim sList As String
    Dim i As Integer
    
    sList = "("
    For i = LBound(vntVals) To UBound(vntVals)
        sList = sList & Format(vntVals(i), "0.0000") & " "
    Next i
    sList = Left(sList, Len(sList) - 1)
    sList = sList & ")"
    MakeListParen = sList
End Function

Public Function MakeListBare(ParamArray vntVals() As Variant) As String
    Dim sList As String
    Dim i As Integer
    
    sList = ""
    For i = LBound(vntVals) To UBound(vntVals)
        sList = sList & Format(vntVals(i), "0.0000") & " "
    Next i
    sList = Left(sList, Len(sList) - 1)
    MakeListBare = sList
End Function

Public Function MakeUnion(ParamArray vntVals() As Variant) As String
    Dim sList As String
    Dim i As Integer
    
    sList = "("
    For i = LBound(vntVals) To UBound(vntVals)
        sList = sList & Format(vntVals(i), "0") & ":"
    Next i
    sList = Left(sList, Len(sList) - 1)
    sList = sList & ")"
    MakeUnion = sList
End Function

Public Function MakeInter(ParamArray vntVals() As Variant) As String
    Dim sList As String
    Dim i As Integer
    
    sList = "("
    For i = LBound(vntVals) To UBound(vntVals)
        sList = sList & Format(vntVals(i), "0") & " "
    Next i
    sList = Left(sList, Len(sList) - 1)
    sList = sList & ")"
    MakeInter = sList
End Function

Public Function GetTr(nTr As Long, lDim As Long) As Double
    Dim rngInpTr As Range
    Dim lrow As Long
    Dim str As String
    
    Application.Volatile True
    
    Set rngInpTr = ActiveSheet.Names("inpTranslation").RefersToRange
    str = "tr" & Format(nTr, "0")
    lrow = 1
    While rngInpTr.Cells(lrow, 1).Text <> str
        lrow = lrow + 1
        If lrow > rngInpTr.Rows.Count Then
            GetTr = CVErr(xlErrNA)
            Exit Function
        End If
    Wend
    
    GetTr = rngInpTr.Cells(lrow, lDim + 1).Value2
End Function

Public Function DoTr(nTr As Long, lDim As Long, dblLocal As Double) As Double
    Dim dblTr As Double
    
    Application.Volatile True
    
    dblTr = GetTr(nTr, lDim)
    
    If IsNumeric(dblTr) Then
        DoTr = dblLocal + dblTr
    Else
        DoTr = CVErr(xlErrNA)
    End If
End Function

Public Sub InitSurfaces()
    Dim cSurface As clsSurface
    
    Set mdSurfaces = New Scripting.Dictionary
    
    ' PNNL
    Set cSurface = New clsSurface
    Set cSurface.SurfaceRange = Worksheets("PNNL").Names("inpSurface").RefersToRange
    mdSurfaces.Add "PNNL", cSurface
    
    ' Elevator
    Set cSurface = New clsSurface
    Set cSurface.SurfaceRange = Worksheets("Elevator").Names("inpSurface").RefersToRange
    mdSurfaces.Add "Elevator", cSurface
    
    ' SAL
    Set cSurface = New clsSurface
    Set cSurface.SurfaceRange = Worksheets("SAL").Names("inpSurface").RefersToRange
    mdSurfaces.Add "SAL", cSurface
    
    ' HLRF
    Set cSurface = New clsSurface
    Set cSurface.SurfaceRange = Worksheets("HLRF").Names("inpSurface").RefersToRange
    mdSurfaces.Add "HLRF", cSurface
    
    
End Sub

Public Sub InitWorksheetSurface(sName As String)
    Dim cSurface As clsSurface
    
    If mdSurfaces Is Nothing Then
        Set mdSurfaces = New Scripting.Dictionary
    End If
    
    If mdSurfaces.Exists(sName) Then
        mdSurfaces.Remove sName
    End If
    
    Set cSurface = New clsSurface
    Set cSurface.SurfaceRange = Worksheets(sName).Names("inpSurface").RefersToRange
    mdSurfaces.Add sName, cSurface
    
    
End Sub
Public Function Surr(sSurf As String) As String
    Dim cSurface As clsSurface
    
    Application.Volatile True
    
    If mdSurfaces Is Nothing Then
        Set mdSurfaces = New Scripting.Dictionary
    End If
    
    If Not mdSurfaces.Exists(ActiveSheet.Name) Then
        InitWorksheetSurface ActiveSheet.Name
    End If
    
    Set cSurface = mdSurfaces.Item(ActiveSheet.Name)
    If cSurface Is Nothing Then
        InitWorksheetSurface ActiveSheet.Name
        Set cSurface = mdSurfaces.Item(ActiveSheet.Name)
    End If
    
    Surr = cSurface.Lookup(sSurf)
    
End Function

Public Function Unnn(rng As Range) As String

    Application.Volatile True

    Dim sSurf As String
    Dim lBgn As Long
    Dim lEnd As Long
    Dim l As Long
    
    If rng.Cells(1).Text = "(" Then
        lBgn = 2
    Else
        lBgn = 1
    End If
    
    If rng.Cells(rng.Cells.Count).Text = ")" Then
        lEnd = rng.Cells.Count - 1
    Else
        lEnd = rng.Cells.Count
    End If
    
        
    Unnn = "("
    For l = lBgn To lEnd
        Unnn = Unnn & Surr(rng.Cells(l).Text) & ":"
    Next
    Mid(Unnn, Len(Unnn), 1) = ")"
End Function

Public Function Sur(sSurf As String) As String
    Dim rngSurface As Range
    Dim nRow As Long
    Dim nCol As Long
    Dim lMatch As Long
    Dim sLabel As String
    Dim sSign As String
    Dim sID As String
    Dim i As Long
    Dim ioff As Long
    Dim blnFound As Boolean
    
    Application.Volatile True
        
    ' Check for initial minus or en dash
    sSign = Left(sSurf, 1)
    If sSign = "-" Or sSign = "–" Then
        sSurf = Mid(sSurf, 2)
        sSign = "-"
    ElseIf sSign = "+" Then
        sSurf = Mid(sSurf, 2)
        sSign = ""
    Else
        sSign = ""
    End If
    
    Set rngSurface = ActiveSheet.Names("inpSurface").RefersToRange
    nRow = rngSurface.Rows.Count
    nCol = rngSurface.Columns.Count
    
    blnFound = False
    For i = 1 To nRow
        sLabel = rngSurface.Cells(i, nCol).Text
        If sLabel = sSurf Then
            ioff = 0
            sID = rngSurface.Cells(i, 1).Text
            While sID = ""
                If ioff > 5 Then
                    blnFound = True
                    Exit For
                End If
                ioff = ioff + 1
                sID = rngSurface.Cells(i - ioff, 1).Text
            Wend
            If IsNumeric(sID) Then
                Sur = sSign & sID
                Exit Function
            Else
                blnFound = True
                Exit For
            End If
        End If
    Next i
    
    If blnFound Then
        Sur = "n/u"
    Else
        Sur = "!Error"
    End If
End Function

Public Function Unn(rng As Range) As String

    Application.Volatile True

    Dim sSurf As String
    Dim lBgn As Long
    Dim lEnd As Long
    Dim l As Long
    
    If rng.Cells(1).Text = "(" Then
        lBgn = 2
    Else
        lBgn = 1
    End If
    
    If rng.Cells(rng.Cells.Count).Text = ")" Then
        lEnd = rng.Cells.Count - 1
    Else
        lEnd = rng.Cells.Count
    End If
    
        
    Unn = "("
    For l = lBgn To lEnd
        Unn = Unn & Sur(rng.Cells(l).Text) & ":"
    Next
    Mid(Unn, Len(Unn), 1) = ")"
End Function

