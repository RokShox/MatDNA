Attribute VB_Name = "mdlIndex"
'
' Copyright (c) 2017 Henkel Technical Computing, LLC. All rights reserved.
'
Option Explicit

Public Sub PackIndex(i() As Integer, n() As Integer, k As Integer, Optional nDim As Integer = 2, Optional blnZeroBased As Boolean = True)
    Dim iDim As Integer
    Dim nTot As Integer
    
    If blnZeroBased Then
        k = 0
        nTot = 1
        For iDim = 0 To nDim - 1
            k = k + i(iDim) * nTot
            nTot = nTot * n(iDim)
        Next iDim
    Else
        k = 1
        nTot = 1
        For iDim = 0 To nDim - 1
            k = k + (i(iDim) - 1) * nTot
            nTot = nTot * n(iDim)
        Next iDim
    End If
End Sub

Public Sub UnpackIndex(i() As Integer, n() As Integer, k As Integer, Optional nDim As Integer = 2, Optional blnZeroBased As Boolean = True)
    Dim iDim As Integer
    Dim nTot As Integer
    Dim nLeft As Integer
    
    ' Total number of nodes
    nTot = 1
    For iDim = 0 To nDim - 1
        nTot = nTot * n(iDim)
    Next iDim
    
    If blnZeroBased Then
        nLeft = k
        For iDim = nDim - 1 To 0 Step -1
            nTot = nTot / n(iDim)
            i(iDim) = Int(nLeft / nTot)
            nLeft = nLeft - i(iDim) * nTot
        Next iDim
    Else
        nLeft = k
        For iDim = nDim - 1 To 0 Step -1
            nTot = nTot / n(iDim)
            i(iDim) = Int((nLeft - 1) / nTot) + 1
            nLeft = nLeft - (i(iDim) - 1) * nTot
        Next iDim
    End If
End Sub

Public Sub PackIndex2(i0 As Integer, i1 As Integer, k As Integer, n0 As Integer, Optional blnZeroBased As Boolean = True)
    Dim i(1) As Integer
    Dim n(1) As Integer
    
    i(0) = i0
    i(1) = i1
    n(0) = n0
    n(1) = 1
    PackIndex i, n, k, 2, blnZeroBased
End Sub

Public Sub UnpackIndex2(i0 As Integer, i1 As Integer, k As Integer, n0 As Integer, Optional blnZeroBased As Boolean = True)
    Dim i(1) As Integer
    Dim n(1) As Integer
    
    n(0) = n0
    n(1) = 1
    UnpackIndex i, n, k, 2, blnZeroBased
    i0 = i(0)
    i1 = i(1)
End Sub


Public Function CellIndexRow(cell As Range, rng As Range)

    CellIndexRow = cell.Cells(1, 1).Row - rng.Cells(1, 1).Row + 1

End Function

Public Function CellIndexColumn(cell As Range, rng As Range)

    CellIndexColumn = cell.Cells(1, 1).Column - rng.Cells(1, 1).Column + 1

End Function

Public Function CellIndexPosition(cell As Range, rng As Range)

    CellIndexPosition = (CellIndexRow(cell, rng) - 1) * rng.Columns.Count + CellIndexColumn(cell, rng)
    
End Function

Public Function NaturalIndexIJ(pos As Variant, nColumn As Variant, iDim As Variant, Optional blnZeroBased As Boolean = True) As Variant

    Dim iRow As Integer
    Dim iCol As Integer
    
    If blnZeroBased Then
        ' Index from 0
        iCol = pos Mod nColumn
        iRow = pos \ nColumn
    Else
        ' Index from 1
        iCol = ((pos - 1) Mod nColumn) + 1
        iRow = ((pos - 1) \ nColumn) + 1
    End If
    
    If iDim = 1 Then
        NaturalIndexIJ = iCol
    ElseIf iDim = 2 Then
        NaturalIndexIJ = iRow
    Else
        NaturalIndexIJ = 0
    End If

End Function

Public Function NaturalIndexPos(iCol As Variant, iRow As Variant, nColumn As Variant, Optional blnZeroBased As Boolean = True) As Variant
    Dim pos As Integer

    If blnZeroBased Then
        ' Index from 0
        pos = iRow * nColumn + iCol
    Else
        ' Index from 1
        pos = (iRow - 1) * nColumn + iCol
    End If
    
    NaturalIndexPos = pos

End Function

