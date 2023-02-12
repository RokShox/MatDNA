Attribute VB_Name = "mdlStat"
'
' Copyright (c) 2017 Henkel Technical Computing, LLC. All rights reserved.
'
Option Explicit

'
' MaxIndex
'
' Return index of max value
'
Function MaxIndex(vntRange As Variant) As Variant
    Dim i As Integer
    Dim nRange As Integer
    Dim rng As Range
    Dim dblMax As Double
    
    Set rng = vntRange
    If rng.Rows.Count = 1 Then
        nRange = rng.Columns.Count
    ElseIf rng.Columns.Count Then
        nRange = rng.Rows.Count
    Else
        MsgBox "Must specify either a single row or column of values", vbCritical
        End
    End If
    
    dblMax = -10000000#
    For i = 1 To nRange
        If rng.Cells(i).Value > dblMax Then
            MaxIndex = i
            dblMax = rng.Cells(i).Value
        End If
    Next i
End Function
'
' MinIndex
'
' Return index of min value
'
Function MinIndex(vntRange As Variant) As Variant
    Dim i As Integer
    Dim nRange As Integer
    Dim rng As Range
    Dim dblMin As Double
    
    Set rng = vntRange
    If rng.Rows.Count = 1 Then
        nRange = rng.Columns.Count
    ElseIf rng.Columns.Count Then
        nRange = rng.Rows.Count
    Else
        MsgBox "Must specify either a single row or column of values", vbCritical
        End
    End If
    
    dblMin = 1E+28
    For i = 1 To nRange
        If rng.Cells(i).Value < dblMin Then
            MinIndex = i
            dblMin = rng.Cells(i).Value
        End If
    Next i
End Function

'
' combineMean
'
' Statistical combination of dose rate results
'
Sub combineMean( _
    rVals As Object, _
    dMean As Double, _
    dFracStdDev As Double, _
    Optional blnWeighted As Boolean = True)
Attribute combineMean.VB_ProcData.VB_Invoke_Func = " \n14"
    
    Dim dStdDev As Double
    Dim dVariance As Double
    Dim dSumMean As Double
    Dim dSumRecipVariance As Double
    Dim dSumVariance As Double
    Dim nSamps As Integer
    Dim dblSamps As Double
    Dim i As Integer
        
    nSamps = rVals.Rows.Count
        
    dblSamps = CDbl(nSamps)
    dSumMean = 0#
    dSumRecipVariance = 0#
    dSumVariance = 0#
    
    ' Optionally weight the computed mean by the standard deviations
    If blnWeighted Then
        For i = 1 To nSamps
            dMean = rVals.Cells(i, 1).Value
            dFracStdDev = rVals.Cells(i, 2).Value
            dStdDev = dFracStdDev * dMean
            dVariance = dStdDev ^ 2
            dSumMean = dSumMean + dMean / dVariance
            dSumRecipVariance = dSumRecipVariance + 1# / dVariance
        Next i
        dMean = dSumMean / dSumRecipVariance
        dVariance = 1# / dSumRecipVariance
        dFracStdDev = Sqr(dVariance) / dMean
    Else
        For i = 1 To nSamps
            dMean = rVals.Cells(i, 1).Value
            dFracStdDev = rVals.Cells(i, 2).Value
            dStdDev = dFracStdDev * dMean
            dVariance = dStdDev ^ 2
            dSumMean = dSumMean + dMean
            dSumVariance = dSumVariance + dVariance
        Next i
        dMean = dSumMean / dblSamps
        dVariance = dSumVariance / dblSamps
        dFracStdDev = Sqr(dVariance) / dMean
    End If
    
End Sub
    
'
' GetCombMean
'
' Worksheet accessor for combineMean sub
'
Function GetCombMean(vals As Object, Optional blnWeighted As Boolean = True) As Double
Attribute GetCombMean.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim dMean As Double
    Dim dFracStdDev As Double
    Call combineMean(vals, dMean, dFracStdDev, blnWeighted)
    GetCombMean = dMean
End Function
    
'
' GetCombFSD
'
' Worksheet accessor for combineMean sub
'
Function GetCombFSD(vals As Object, Optional blnWeighted As Boolean = True) As Double
Attribute GetCombFSD.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim dMean As Double
    Dim dFracStdDev As Double
    Call combineMean(vals, dMean, dFracStdDev, blnWeighted)
    GetCombFSD = dFracStdDev
End Function
   
'
' addVals
'
' Statistical summation of values
'
Sub addVals(rVals As Object, dSum As Double, dFracStdDev As Double)
Attribute addVals.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim dVal As Double
    Dim dStdDev As Double
    Dim dVariance As Double
    Dim dSumVariance As Double
    Dim cSamps As Single
    Dim blnRowData As Boolean
    Dim i As Integer
        
    If rVals.Rows.Count = 1 Then
        blnRowData = True
        If rVals.Columns.Count Mod 2 <> 0 Then
            MsgBox "Expected an even number of columns", vbCritical
            End
        End If
        cSamps = rVals.Columns.Count / 2
    Else
        blnRowData = False
        cSamps = rVals.Rows.Count
    End If
    dSum = 0
    dSumVariance = 0
    
    For i = 1 To cSamps
        If blnRowData Then
            dVal = rVals.Cells(1, 2 * i - 1)
            dFracStdDev = rVals.Cells(1, 2 * i)
        Else
            dVal = rVals.Cells(i, 1).Value
            dFracStdDev = rVals.Cells(i, 2).Value
        End If
        dStdDev = dFracStdDev * dVal
        dVariance = dStdDev ^ 2
        dSum = dSum + dVal
        dSumVariance = dSumVariance + dVariance
    Next i
    If dSum <> 0 Then
        dFracStdDev = (dSumVariance ^ 0.5) / dSum
    Else
        dFracStdDev = 0#
    End If
End Sub

'
' AddValsSum
'
' Worksheet accessor for addVals
'
Function AddValsSum(rVals As Object) As Double
Attribute AddValsSum.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim dSum As Double
    Dim dFracStdDev As Double
    Call addVals(rVals, dSum, dFracStdDev)
    AddValsSum = dSum
End Function
    
'
' AddValsFSD
'
' Worksheet accessor for addVals
'
Function AddValsFSD(rVals As Object) As Double
Attribute AddValsFSD.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim dSum As Double
    Dim dFracStdDev As Double
    Call addVals(rVals, dSum, dFracStdDev)
    AddValsFSD = dFracStdDev
End Function
'
' combineMeanRow
'
' Like combineMean, but data is row-wise
'
Sub combineMeanRow(rVals As Object, dMean As Double, dFracStdDev As Double, blnVarWeighted As Boolean)
    Dim dStdDev As Double
    Dim dVariance As Double
    Dim dSumMean As Double
    Dim dSumRecipVariance As Double
    Dim dSumObs As Double
    Dim dSumObsVar As Double
    Dim dSumVar As Double
    Dim nSamps As Integer
    Dim dSamps As Double
    Dim dObs As Double
    Dim dFSD As Double
    Dim ii As Integer
    Dim jj As Integer
    Dim i As Integer
        
    If rVals.Columns.Count Mod 2 Then
        MsgBox "Error! Must have even number of columns"
    End If
        
    nSamps = rVals.Columns.Count / 2
    dSamps = CDbl(nSamps)
    ' MsgBox "cSamps is " & cSamps
    dSumObs = 0
    dSumObsVar = 0
    dSumVar = 0
    dSumRecipVariance = 0
    
    For i = 1 To nSamps
        ii = 2 * i
        jj = ii - 1
        
        dObs = rVals.Cells(1, jj).Value
        dFSD = rVals.Cells(1, ii).Value
        dStdDev = dFSD * dObs
        dVariance = dStdDev ^ 2
        If dVariance = 0 Then
            dVariance = 1
        End If
        dSumObs = dSumObs + dObs
        dSumVar = dSumVar + dVariance
        dSumObsVar = dSumObsVar + dObs / dVariance
        dSumRecipVariance = dSumRecipVariance + 1 / dVariance
    Next i
  
    If blnVarWeighted Then
        dMean = dSumObsVar / dSumRecipVariance
        dVariance = 1 / dSumRecipVariance
        dFracStdDev = (dVariance ^ 0.5) / dMean
    Else
        dMean = dSumObs / dSamps
        dVariance = dSumVar / (dSamps ^ 2)
        dFracStdDev = (dVariance ^ 0.5) / dMean
    End If
        
End Sub
    
'
' GetCombMeanRow
'
' Worksheet accessor for combineMeanRow
'
Function GetCombMeanRow(vals As Object, Optional blnWeighted As Boolean = True) As Double
    Dim dMean As Double
    Dim dFracStdDev As Double
    Call combineMeanRow(vals, dMean, dFracStdDev, blnWeighted)
    GetCombMeanRow = dMean
End Function
    
'
' GetCombFSDRow
'
' Worksheet accessor for combineMeanRow
'
Function GetCombFSDRow(vals As Object, Optional blnWeighted As Boolean = True) As Double
    Dim dMean As Double
    Dim dFracStdDev As Double
    Call combineMeanRow(vals, dMean, dFracStdDev, blnWeighted)
    GetCombFSDRow = dFracStdDev
End Function
   
'
' GetAvgRow
'
' Worksheet accessor for combineMeanRow
'
Function GetAvgRow(vals As Object) As Double
    Dim dMean As Double
    Dim dFracStdDev As Double
    Call combineMeanRow(vals, dMean, dFracStdDev, False)
    GetAvgRow = dMean
End Function
    
'
' GetAvgFSDRow
'
' Worksheet accessor for combineMeanRow
'
Function GetAvgFSDRow(vals As Object) As Double
    Dim dMean As Double
    Dim dFracStdDev As Double
    Call combineMeanRow(vals, dMean, dFracStdDev, False)
    GetAvgFSDRow = dFracStdDev
End Function

'
' combineMeanArea
'
' Statistical combination of dose rate results
'
Sub combineMeanArea( _
    rngArea As Range, _
    dMean As Double, _
    dFracStdDev As Double, _
    Optional blnWeighted As Boolean = True)
    
    Dim dStdDev As Double
    Dim dVariance As Double
    Dim dSumMean As Double
    Dim dSumRecipVariance As Double
    Dim dSumVariance As Double
    Dim nSamps As Integer
    Dim nSampsTot As Integer
    Dim dblSampsTot As Double
    Dim rngVals As Range
    Dim i As Long
        
    nSampsTot = 0
    dSumMean = 0#
    dSumRecipVariance = 0#
    dSumVariance = 0#
    
    ' MsgBox "Area count is " & rngArea.Areas.Count
    
    For Each rngVals In rngArea.Areas
        
        nSamps = rngVals.Rows.Count
        nSampsTot = nSampsTot + nSamps
        
        ' Optionally weight the computed mean by the standard deviations
        If blnWeighted Then
            For i = 1 To nSamps
                dMean = rngVals.Cells(i, 1).Value
                dFracStdDev = rngVals.Cells(i, 2).Value
                dStdDev = dFracStdDev * dMean
                dVariance = dStdDev ^ 2
                dSumMean = dSumMean + dMean / dVariance
                dSumRecipVariance = dSumRecipVariance + 1# / dVariance
            Next i
        Else
            For i = 1 To nSamps
                dMean = rngVals.Cells(i, 1).Value
                dFracStdDev = rngVals.Cells(i, 2).Value
                dStdDev = dFracStdDev * dMean
                dVariance = dStdDev ^ 2
                dSumMean = dSumMean + dMean
                dSumVariance = dSumVariance + dVariance
            Next i
        End If
    Next rngVals
    
    dblSampsTot = CDbl(nSampsTot)
    If blnWeighted Then
        dMean = dSumMean / dSumRecipVariance
        dVariance = 1# / dSumRecipVariance
        dFracStdDev = Sqr(dVariance) / dMean
    Else
        dMean = dSumMean / dblSampsTot
        dVariance = dSumVariance / dblSampsTot
        dFracStdDev = Sqr(dVariance) / dMean
    End If
    
End Sub

'
' GetCombMeanArea
'
' Worksheet accessor for combineMeanArea sub
'
Function GetCombMeanArea(rngArea As Range, Optional blnWeighted As Boolean = True) As Double
    Dim dMean As Double
    Dim dFracStdDev As Double
    
    'MsgBox "Area count is " & rngArea.Areas.Count
    
    combineMeanArea rngArea, dMean, dFracStdDev, blnWeighted
    GetCombMeanArea = dMean
End Function
    
'
' GetCombFSDArea
'
' Worksheet accessor for combineMean sub
'
Function GetCombFSDArea(rngArea As Range, Optional blnWeighted As Boolean = True) As Double
    Dim dMean As Double
    Dim dFracStdDev As Double
    
    combineMeanArea rngArea, dMean, dFracStdDev, blnWeighted
    GetCombFSDArea = dFracStdDev
End Function

'
' CountUniq
'
Public Sub CountUniq()
    Dim rngVals As Range
    Dim rngDest As Range
    Dim i As Range
    Dim vntPrev As Variant
    Dim iDest As Integer
    Dim iCount As Integer
    
    If Selection.Areas.Count <> 2 Then
        MsgBox "Must select values and dest ranges"
        Exit Sub
    End If
    
    Set rngVals = Selection.Areas(1)
    Set rngDest = Selection.Areas(2).Cells(1, 1)
    
    iDest = 0
    iCount = 0
    vntPrev = rngVals.Cells(1).Value
    For Each i In rngVals
        If i.Value <> vntPrev Then
            rngDest.Offset(iDest, 0).Value = vntPrev
            rngDest.Offset(iDest, 1).Value = iCount
            iDest = iDest + 1
            iCount = 1
            If i.Text = "" Then
                GoTo Done
            End If
        Else
            iCount = iCount + 1
        End If
        vntPrev = i.Value
    Next i
    
Last:
    rngDest.Offset(iDest, 0).Value = vntPrev
    rngDest.Offset(iDest, 1).Value = iCount

Done:

End Sub

