Attribute VB_Name = "mdlWin"
'
' Copyright (c) 2017 Henkel Technical Computing, LLC. All rights reserved.
'
Option Explicit

Public Sub SafeMkDir(sDir As String, Optional blnChDir As Boolean = False)

    ' Only makes one level of directory
    On Error Resume Next
    MkDir sDir
    If blnChDir Then
        ChDir sDir
    End If

End Sub

Public Sub WriteRangeToFile(rng As Range, iFile As Integer, Optional sDelim As String)
    Dim lrow As Long
    Dim lCol As Long
    Dim nCol As Long
    Dim sLine As String
    
    If sDelim = "" Then
        sDelim = Chr(9)
    End If
    nCol = rng.Columns.Count
    For lrow = 1 To rng.Rows.Count
        sLine = ""
        For lCol = 1 To nCol
            sLine = sLine & rng.Cells(lrow, lCol).Text
            If lCol < nCol Then
                sLine = sLine & sDelim
            End If
        Next lCol
        Print #iFile, sLine
    Next lrow
End Sub

