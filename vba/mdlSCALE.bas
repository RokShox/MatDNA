Attribute VB_Name = "mdlSCALE"
'
' Copyright (c) 2017 Henkel Technical Computing, LLC. All rights reserved.
'
Option Explicit
'
' WriteSCALERange
'
Public Sub WriteSCALERange()
    Dim sFile As String
    Dim rng As Range
    Dim iFile As Integer
    
    sFile = ThisWorkbook.Names("dirWork").RefersToRange.Value & "\scale.inp"
    iFile = FreeFile
    Open sFile For Output As #iFile
    Set rng = Selection
    WriteSCALERangeToFile rng, iFile, Nothing, Chr(32)
    Close #iFile

End Sub

'
' MakeCSAS5Case
'
Public Sub MakeCSAS5Case()
    Dim sFile As String
    Dim iFile As Integer
    Dim rngInput As Range
    Dim sCase As String
    Dim sVersion As String
    Dim sDirCase As String
    Dim sDirRun As String
    
    Application.Calculate
    
    sDirRun = ThisWorkbook.Names("dirRun").RefersToRange.Value
    sVersion = ThisWorkbook.Names("version").RefersToRange.Text
    sDirCase = sDirRun & "\" & sVersion
    SafeMkDir sDirCase
    sCase = ActiveSheet.Names("caseName").RefersToRange.Text
    sFile = sDirCase & "\" & sCase & ".inp"
    iFile = FreeFile
    Open sFile For Output As #iFile
    On Error Resume Next
    Set rngInput = ActiveSheet.Names("inpCSAS5").RefersToRange
    
    WriteSCALERangeToFile rngInput, iFile, Nothing, Chr(32)
    Close #iFile
End Sub


'
' MakeCSAS6Case
'
Public Sub MakeCSAS6Case()
    Dim sFile As String
    Dim iFile As Integer
    Dim rngInput As Range
    Dim sCase As String
    Dim sVersion As String
    Dim sDirCase As String
    Dim sDirRun As String
    
    Application.Calculate
    
    sDirRun = ThisWorkbook.Names("dirRun").RefersToRange.Value
    sVersion = ThisWorkbook.Names("version").RefersToRange.Text
    sDirCase = sDirRun & "\" & sVersion
    SafeMkDir sDirCase
    sCase = ActiveSheet.Names("caseName").RefersToRange.Text
    sFile = sDirCase & "\" & sCase & ".inp"
    iFile = FreeFile
    Open sFile For Output As #iFile
    On Error Resume Next
    Set rngInput = ActiveSheet.Names("inpCSAS6").RefersToRange
    
    WriteSCALERangeToFile rngInput, iFile, Nothing, Chr(32)
    Close #iFile
End Sub


'
' WriteSCALERangeToFile
'
Public Sub WriteSCALERangeToFile(rng As Range, iFile As Integer, colPatSubst As Collection, sDelim As String)
    Dim lrow As Long
    Dim lCol As Long
    Dim nCol As Long
    Dim sLine As String
    Dim rngInsert As Range
    Dim colPatSubstInsert As Collection
    Dim sOut As String
    Dim wksThis As Worksheet
    
    If sDelim = "" Then
        sDelim = Chr(9)
    End If
    Set wksThis = rng.Worksheet
    
    nCol = rng.Columns.Count
    For lrow = 1 To rng.Rows.Count
        sLine = ""
        ' Put the whole line together
        For lCol = 1 To nCol
            sLine = sLine & rng.Cells(lrow, lCol).Text
            If lCol < nCol Then
                sLine = sLine & sDelim
            End If
        Next lCol
        sLine = RTrim(sLine)
        
        ' Check for included range
        If Left(sLine, 1) = "#" Then
            parseInsert sLine, rngInsert, colPatSubstInsert, wksThis
            WriteSCALERangeToFile rngInsert, iFile, colPatSubstInsert, sDelim
        
        ElseIf Left(sLine, 3) = "n/u" Then
            ' Skip this line
        
        Else
            doPatSubst sLine, sOut, colPatSubst
            Print #iFile, sOut
        End If
    Next lrow
End Sub


