VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsBin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' Copyright (c) 2017 Henkel Technical Computing, LLC. All rights reserved.
'
Option Explicit
' Recall dim a(n) means n+1 values indexed from 0 to n

' mlNumBin is the number of "true" bins. There are two catchall bins on either
' end of the defined bin boundaries. They are indexed by 0 and mlNumBin+1.
' The "true" bins are indexed from 1 to mlNumBin

' The upper limit of bin i is x(i).
' Bin i contains all values x(i-1) < x <= x(i)
' Bin 0 contains all values x <= x(0)
' Bin mlNumBin+1 is all values x(mlNumBin) < x
' Only the counts (and a weight for each count) are maintained, not the actual values
' Could store the values in separate arrays for each bin, or store an array of identifiers (e.g., indices)

Private Const MAXBIN As Long = 100
Private Const MAXCOUNT As Long = 10000

Private mdblaBinUpper(MAXBIN) As Double
Private mlaBinCount(MAXBIN) As Long
Private mlaIdent(MAXBIN, MAXCOUNT) As Long
Private mdblaBinWeight(MAXBIN) As Double
Private mdblaBinMoment(MAXBIN) As Double
Private mlNumBin As Long
Private mblnHasCatchallData

Private Sub Class_Initialize()
    Dim i As Long
    Dim j As Long
    
    mlNumBin = 0
    mblnHasCatchallData = False
    For i = 0 To MAXBIN
        mdblaBinUpper(i) = 0#
        mlaBinCount(i) = 0
        mdblaBinWeight(i) = 0#
        mdblaBinMoment(i) = 0#
        For j = 0 To MAXCOUNT
            mlaIdent(i, j) = 0
        Next j
    Next i
End Sub

Property Get NumBin() As Long
    NumBin = mlNumBin
End Property

Property Get BinUpper(i As Long) As Double
    ' Bin mlNumBin+1 exists, but it has no upper limit
    If i > mlNumBin Then
        MsgBox "Invalid bin", vbCritical
        End
    ElseIf i < 0 Then
        MsgBox "Invalid bin", vbCritical
        End
    End If
    
    BinUpper = mdblaBinUpper(i)
End Property

Property Get BinCount(i As Long) As Long
    If i > mlNumBin + 1 Then
        MsgBox "Invalid bin", vbCritical
        End
    ElseIf i < 0 Then
        MsgBox "Invalid bin", vbCritical
        End
    End If
    
    BinCount = mlaBinCount(i)
End Property

Property Get BinWeight(i As Long) As Double
    If i > mlNumBin + 1 Then
        MsgBox "Invalid bin", vbCritical
        End
    ElseIf i < 0 Then
        MsgBox "Invalid bin", vbCritical
        End
    End If
    
    BinWeight = mdblaBinWeight(i)
End Property

Property Get BinWeightedAverage(i As Long) As Double
    If i > mlNumBin + 1 Then
        MsgBox "Invalid bin", vbCritical
        End
    ElseIf i < 0 Then
        MsgBox "Invalid bin", vbCritical
        End
    End If
    
    If mdblaBinWeight(i) = 0 Then
        BinWeightedAverage = 0
    Else
        BinWeightedAverage = mdblaBinMoment(i) / CDbl(mdblaBinWeight(i))
    End If
    
End Property

Property Get TotalWeight() As Double
    Dim i As Long
    
    TotalWeight = 0#
    For i = 0 To mlNumBin + 1
        TotalWeight = TotalWeight + mdblaBinWeight(i)
    Next i
End Property

Property Get TotalCount() As Long
    Dim i As Long
    
    TotalCount = 0#
    For i = 0 To mlNumBin + 1
        TotalCount = TotalCount + mlaBinCount(i)
    Next i
End Property

Property Get TotalWeightedAverage() As Double
    Dim i As Long
    Dim dblSumMoment As Double
    Dim dblSumWeight As Double
    
    dblSumMoment = 0#
    dblSumWeight = 0#
    For i = 0 To mlNumBin + 1
        dblSumMoment = dblSumMoment + mdblaBinMoment(i)
        dblSumWeight = dblSumWeight + mdblaBinWeight(i)
    Next i
    If dblSumWeight = 0 Then
        TotalWeightedAverage = 0
    Else
         TotalWeightedAverage = dblSumMoment / dblSumWeight
    End If
End Property

Property Get TotalWeightFraction() As Double
    Dim i As Long
    Dim dblSum As Double
    
    dblSum = 0#
    For i = 0 To mlNumBin + 1
        dblSum = dblSum + Me.WeightFraction(i)
    Next i
    TotalWeightFraction = dblSum
End Property

Property Get WeightFraction(i As Long) As Double
    
    If i > mlNumBin + 1 Then
        MsgBox "Invalid bin", vbCritical
        End
    ElseIf i < 0 Then
        MsgBox "Invalid bin", vbCritical
        End
    End If
    
    WeightFraction = Me.BinWeight(i) / Me.TotalWeight
End Property

Property Get BinIdent(i As Long, j As Long) As Long

    If i > mlNumBin + 1 Then
        MsgBox "Invalid bin", vbCritical
        End
    ElseIf i < 0 Then
        MsgBox "Invalid bin", vbCritical
        End
    End If
    If j < 0 Then
        MsgBox "Invalid index", vbCritical
        End
    End If
    If j > mlaBinCount(i) - 1 Then
        MsgBox "Invalid index", vbCritical
        End
    End If

    BinIdent = mlaIdent(i, j)
End Property

Public Sub SetBinsbyLimits(nBin As Long, dblLimitLo As Double, dblLimitHi As Double)
    Dim dblBinSize As Double
    Dim i As Long
    
    ' nBin is the number of "true" bins. They are indexed from 1 to nBin
    
    If dblLimitLo >= dblLimitHi Then
        MsgBox "Invalid bin structure", vbCritical
        End
    End If
    
    If nBin < 1 Then
        MsgBox "Invalid bin structure", vbCritical
        End
    End If
    
    If nBin + 1 > MAXBIN Then
        MsgBox "Too many bins", vbCritical
        End
    End If
        
    ' Recall there are two "catchall bins" one on either end of the upper and lower limits
    mdblaBinUpper(0) = dblLimitLo
    
    ' There is no upper index on the last bin
    mdblaBinUpper(nBin) = dblLimitHi
    
    ' Just define the interior bin limits, since the extremes are known
    ' There will be nBin-2 increments
    dblBinSize = (dblLimitHi - dblLimitLo) / CDbl(nBin)
    For i = 1 To nBin - 1
        mdblaBinUpper(i) = mdblaBinUpper(0) + CDbl(i) * dblBinSize
    Next i
    
    mlNumBin = nBin

End Sub

Public Sub SetBinsbySize(nBin As Long, dblBinSize As Double, dblLimitLo As Double)
    Dim i As Long
    
    If nBin < 1 Then
        MsgBox "Invalid bin structure", vbCritical
        End
    End If
    
    If nBin + 1 > MAXBIN Then
        MsgBox "Too many bins", vbCritical
        End
    End If
        
    ' Recall there are two "catchall bins" one on either end of the upper and lower limits
    mdblaBinUpper(0) = dblLimitLo

    For i = 1 To nBin
        mdblaBinUpper(i) = mdblaBinUpper(0) + CDbl(i) * dblBinSize
    Next i

    mlNumBin = nBin
End Sub

Public Sub SetBinsbyBoundaries(nBin As Long, dblBinLimit() As Double)
    Dim i As Long
    
    If nBin < 1 Then
        MsgBox "Invalid bin structure", vbCritical
        End
    End If
    
    If nBin + 1 > MAXBIN Then
        MsgBox "Too many bins", vbCritical
        End
    End If
        
    For i = 0 To nBin
        mdblaBinUpper(i) = dblBinLimit(i)
        If i > 0 Then
            If dblBinLimit(i) <= dblBinLimit(i - 1) Then
                MsgBox "Invalid bin structure", vbCritical
                End
            End If
        End If
    Next i
    
    mlNumBin = nBin
End Sub

Public Sub Bin(dblVal As Double, dblWgt As Double, lIdent As Long)
    Dim i As Long
    
    If dblVal > mdblaBinUpper(mlNumBin) Then
        addCount dblVal, dblWgt, lIdent, mlNumBin + 1
        mblnHasCatchallData = True
    ElseIf dblVal <= mdblaBinUpper(0) Then
        addCount dblVal, dblWgt, lIdent, 0
        mblnHasCatchallData = True
    Else
        For i = 1 To mlNumBin
            If dblVal <= mdblaBinUpper(i) Then
                addCount dblVal, dblWgt, lIdent, i
                Exit For
            End If
        Next i
    End If
End Sub

Public Sub DisplayExcel(rng As Range)
    Dim i As Long
    Dim iLast As Long
    
    CalcSave
    CalcOff
    
    rng.Offset(-1, 0).Value = "Bin"
    rng.Offset(-1, 1).Value = "Lower"
    rng.Offset(-1, 2).Value = "Upper"
    rng.Offset(-1, 3).Value = "Count"
    rng.Offset(-1, 4).Value = "Weight"
    rng.Offset(-1, 5).Value = "Wgt Avg"
    rng.Offset(-1, 6).Value = "Wgt Frac"
    For i = 0 To mlNumBin + 1
        rng.Offset(i, 0).Value = i
        If i = 0 Then
            rng.Offset(i, 1).Value = "Lower"
        Else
            rng.Offset(i, 1).Value = mdblaBinUpper(i - 1)
        End If
        If i = mlNumBin + 1 Then
            rng.Offset(i, 2).Value = "Higher"
        Else
            rng.Offset(i, 2).Value = mdblaBinUpper(i)
        End If
        rng.Offset(i, 3).Value = mlaBinCount(i)
        rng.Offset(i, 4).Value = mdblaBinWeight(i)
        rng.Offset(i, 5).Value = Me.BinWeightedAverage(i)
        rng.Offset(i, 6).Value = Me.WeightFraction(i)
    Next i
    iLast = mlNumBin + 2
    rng.Offset(iLast, 0).Value = "Total"
    rng.Offset(iLast, 3).Value = Me.TotalCount
    rng.Offset(iLast, 4).Value = Me.TotalWeight
    rng.Offset(iLast, 5).Value = Me.TotalWeightedAverage
    rng.Offset(iLast, 6).Value = Me.TotalWeightFraction
    
    CalcRestore
End Sub

Private Sub addCount(dblVal As Double, dblWgt As Double, lIdent As Long, lBin As Long)
    
    mlaIdent(lBin, mlaBinCount(lBin)) = lIdent
    mdblaBinWeight(lBin) = mdblaBinWeight(lBin) + dblWgt
    mdblaBinMoment(lBin) = mdblaBinMoment(lBin) + dblVal * dblWgt
    mlaBinCount(lBin) = mlaBinCount(lBin) + 1

End Sub
