VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLayout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' Copyright (c) 2017 Henkel Technical Computing, LLC. All rights reserved.
'
Option Explicit

' Represents a layout of assy, rods, control rods, etc
' Positions are indexed from 0
' LayoutRef is the upper left corner of the data in the layout (not including labels)
' ArraySize is the logical number of cells on each side of the square array
' CellSize is the number of Excel cells per logical cell
' The Excel cells in a logical cell may be merged


Dim mrngLayoutRef As Range
Dim miArraySize As Integer
Dim miCellSize As Integer


Private Sub Class_Initialize()
    miCellSize = 0
    miArraySize = 0
End Sub

Property Get ArraySize() As Integer
    ArraySize = miArraySize
End Property

Property Let ArraySize(i As Integer)
    If i <= 0 Then
        MsgBox "Invalid array size: " & i, vbCritical, "Invalid Data"
        End
    End If
    miArraySize = i
End Property

Property Get CellSize() As Integer
    CellSize = miCellSize
End Property

Property Let CellSize(i As Integer)
    If i <= 0 Then
        MsgBox "Invalid cell size: " & i, vbCritical, "Invalid Data"
        End
    End If
    miCellSize = i
End Property

Property Get LayoutRange() As Range
    Dim iSize As Integer
    iSize = miCellSize * (miArraySize - 1)
    If mrngLayoutRef Is Nothing Then
        Set LayoutRange = Nothing
    Else
        Set LayoutRange = Range(mrngLayoutRef, mrngLayoutRef.Offset(iSize, iSize))
    End If
End Property

Property Get LayoutRef() As Range
    Set LayoutRef = mrngLayoutRef
End Property

Property Set LayoutRef(rng As Range)
    Set mrngLayoutRef = rng.Cells(1, 1)
End Property

Property Get Comment() As String
    ' Something screwed up here with merged cells. Should be miArraySize * miCellSize + 1
    Comment = mrngLayoutRef.Offset(miArraySize * miCellSize, -1).Text
End Property

Property Get CellLabel(oRow As Integer, oCol As Integer) As String
    Dim sRow As String
    Dim sCol As String
    
    checkIJ oRow, oCol
    
    sRow = Chr(Asc("A") + oRow)
    sCol = Format(oCol + 1, "00")
    CellLabel = sRow & sCol
End Property

Public Sub CellIJ(sLabel As String, ByRef oRow As Integer, ByRef oCol As Integer)
    Dim sRow As String
    Dim sCol As String
    
    checkLabel sLabel
    
    sRow = UCase(Left(sLabel, 1))
    sCol = Right(sLabel, 2)
    
    oRow = Asc(sRow) - Asc("A")
    oCol = Int(val(sCol)) - 1

End Sub

Public Function HasText(oRow As Integer, oCol As Integer) As Boolean
    checkIJ oRow, oCol
    
    If Me.CellText(oRow, oCol) = "" Then
        HasText = False
    Else
        HasText = True
    End If
End Function

Public Function CellText(oRow As Integer, oCol As Integer, Optional oSubRow As Integer = 0, Optional oSubCol As Integer = 0) As String
    Dim oRowOff As Integer
    Dim oColOff As Integer
    
    checkIJ oRow, oCol
    
    If oSubRow >= miCellSize Then
        MsgBox "Invalid sub row index: " & oSubRow, vbCritical, "Invalid Data"
        End
    ElseIf oSubCol >= miCellSize Then
        MsgBox "Invalid sub column index: " & oSubCol, vbCritical, "Invalid Data"
        End
    Else
        oRowOff = oRow * miCellSize + oSubRow
        oColOff = oCol * miCellSize + oSubCol
        CellText = mrngLayoutRef.Offset(oRowOff, oColOff).MergeArea.Cells(1, 1).Text
    End If
End Function


Public Function CellValue(oRow As Integer, oCol As Integer, Optional oSubRow As Integer = 0, Optional oSubCol As Integer = 0) As Variant
    Dim oRowOff As Integer
    Dim oColOff As Integer
    
    checkIJ oRow, oCol
    
    If oSubRow >= miCellSize Then
        MsgBox "Invalid sub row index: " & oSubRow, vbCritical, "Invalid Data"
        End
    ElseIf oSubCol >= miCellSize Then
        MsgBox "Invalid sub column index: " & oSubCol, vbCritical, "Invalid Data"
        End
    Else
        oRowOff = oRow * miCellSize + oSubRow
        oColOff = oCol * miCellSize + oSubCol
        CellValue = mrngLayoutRef.Offset(oRowOff, oColOff).MergeArea.Cells(1, 1).Value2
    End If
End Function


Public Function Display() As String
    Dim i As Integer
    Dim j As Integer
    Dim iRowOff As Integer
    Dim iColOff As Integer
    Dim sText As String
    
    sText = ""
    
    For i = 0 To miArraySize - 1
        For j = 0 To miArraySize - 1
            sText = sText & Format(Me.CellText(i, j), "@@@@@;@@@@@")
        Next j
        sText = sText & vbCrLf
    Next i
    Display = sText
End Function

Public Sub UniqueValues(saUnique() As String, nUnique As Integer)
    Dim oRow As Integer
    Dim oCol As Integer
    Dim blnUnique As Boolean
    Dim oUnique As Integer
    Dim sText As String
    
    nUnique = 0
    For oRow = 0 To Me.ArraySize - 1
        For oCol = 0 To Me.ArraySize - 1
            sText = Me.CellText(oRow, oCol)
            
            If sText <> "" Then
                blnUnique = True
                For oUnique = 0 To nUnique - 1
                    If sText = saUnique(oUnique) Then
                        blnUnique = False
                        Exit For
                    End If
                Next oUnique
                
                If blnUnique Then
                    saUnique(nUnique) = sText
                    nUnique = nUnique + 1
                End If
            End If
        Next oCol
    Next oRow
End Sub

Public Sub SymmetricPositions(sLabel As String, ByRef saSym() As String, ByRef nSym As Integer)
    Dim oRow As Integer
    Dim oCol As Integer
    Dim oCenter As Integer
    Dim sNorm As String
    
    sNorm = Me.NormalizeLabel(sLabel)
    
    Me.CellIJ sNorm, oRow, oCol
    
    If Me.ArraySize Mod 2 <> 1 Then
        MsgBox "Even array sizes not yet implemented for symmetry", vbCritical
        End
    End If
    
    oCenter = Int((Me.ArraySize - 1) \ 2)
    
    If oRow = oCenter And oCol = oCenter Then
        nSym = 1
        saSym(0) = sNorm
    ElseIf oRow = oCenter Then
        nSym = 2
        saSym(0) = sNorm
        saSym(1) = Me.CellLabel(oRow, oCenter - (oCol - oCenter))
    ElseIf oCol = oCenter Then
        nSym = 2
        saSym(0) = sNorm
        saSym(1) = Me.CellLabel(oCenter - (oRow - oCenter), oCenter)
    Else
        nSym = 4
        ' SE
        saSym(0) = sNorm
        ' NE
        saSym(1) = Me.CellLabel(oCenter - (oRow - oCenter), oCol)
        ' NW
        saSym(2) = Me.CellLabel(oCenter - (oRow - oCenter), oCenter - (oCol - oCenter))
        ' SW
        saSym(3) = Me.CellLabel(oRow, oCenter - (oCol - oCenter))
    End If
End Sub

Public Function NormalizeLabel(sLabel As String) As String
    Dim oRow As Integer
    Dim oCol As Integer
    Dim oCenter As Integer

    If Me.ArraySize Mod 2 <> 1 Then
        MsgBox "Even array sizes not yet implemented for symmetry", vbCritical
        End
    End If
    
    checkLabel sLabel
    
    Me.CellIJ sLabel, oRow, oCol
    
    oCenter = Int((Me.ArraySize - 1) \ 2)
    
    If oRow < oCenter Then
        oRow = oCenter + (oCenter - oRow)
    End If
    
    If oCol < oCenter Then
        oCol = oCenter + (oCenter - oCol)
    End If
    
    NormalizeLabel = Me.CellLabel(oRow, oCol)
End Function

Public Function IsValidLabel(sLabel As String) As Boolean
    Dim sRow As String
    Dim sCol As String
    Dim oRow As Integer
    Dim oCol As Integer
    
    IsValidLabel = True
    If Len(sLabel) <> 3 Then
        IsValidLabel = False
    Else
        sRow = UCase(Left(sLabel, 1))
        oRow = Asc(sRow) - Asc("A")
        
        If oRow > Me.ArraySize - 1 Or oRow < 0 Then
            IsValidLabel = False
        End If
    
        sCol = Right(sLabel, 2)
        oCol = Int(val(sCol)) - 1
        
        If oCol > Me.ArraySize - 1 Or oCol < 0 Then
            IsValidLabel = False
        End If
    End If
End Function

Public Function IsValidIJ(oRow As Integer, oCol As Integer) As Boolean
    IsValidIJ = True
    If oRow < 0 Or oRow > Me.ArraySize - 1 Then
        IsValidIJ = False
    ElseIf oCol < 0 Or oCol > Me.ArraySize - 1 Then
        IsValidIJ = False
    End If
End Function

Public Function IsQuarterSymmetric() As Boolean
    Dim oRow As Integer
    Dim oCol As Integer
    Dim oCenter As Integer
    Dim oSym As Integer
    Dim saSym(3) As String
    Dim nSym As Integer
    Dim sText As String
    Dim sLabel As String
    Dim ooRow As Integer
    Dim ooCol As Integer

    If Me.ArraySize Mod 2 <> 1 Then
        MsgBox "Even array sizes not yet implemented for symmetry", vbCritical
        End
    End If
    
    oCenter = Int((Me.ArraySize - 1) \ 2)

    IsQuarterSymmetric = True
    
    For oRow = oCenter To Me.ArraySize - 1
        For oCol = oCenter To Me.ArraySize - 1
            sText = Me.CellText(oRow, oCol)
            sLabel = Me.CellLabel(oRow, oCol)
            Me.SymmetricPositions sLabel, saSym, nSym
            ' Skip first returned position since it's always in SE quadrant
            For oSym = 1 To nSym - 1
                Me.CellIJ saSym(oSym), ooRow, ooCol
                If Me.CellText(ooRow, ooCol) <> sText Then
                    IsQuarterSymmetric = False
                    Exit Function
                End If
            Next oSym
        Next oCol
    Next oRow
End Function

Private Sub checkLabel(sLabel As String)
    If Not Me.IsValidLabel(sLabel) Then
        MsgBox "Invalid cell label: " & sLabel, vbCritical
        End
    End If
End Sub

Private Sub checkIJ(oRow As Integer, oCol As Integer)
    If Not Me.IsValidIJ(oRow, oCol) Then
        MsgBox "Invalid cell offsets: " & oRow & "," & oCol, vbCritical
        End
    End If
End Sub
