Attribute VB_Name = "mdlCheck"
'
' Copyright (c) 2017 Henkel Technical Computing, LLC. All rights reserved.
'
Option Explicit

Public Function CheckNear(vntO As Variant, vntT As Variant, Optional vntTol As Variant = 0.001, Optional sTolFmt As String = "0.0%") As String
    Dim dblRelErr As Double
    
    dblRelErr = Relerr(vntO, vntT)
    If vntO = vntT Then
            CheckNear = "ok (=)"
    ElseIf Near(vntO, vntT, vntTol) Then
            CheckNear = "ok (<" & Format(vntTol, sTolFmt) & ")"
    Else
        CheckNear = "Error (" & Format(dblRelErr, sTolFmt) & ")"
    End If

End Function

Public Function Near(vntO As Variant, vntT As Variant, Optional vntTol As Variant = 0.001) As Boolean
    Dim dblCmp As Double
    Dim dblO As Double
    Dim dblT As Double
    Dim dblTol As Double
    
    dblO = CDbl(vntO)
    dblT = CDbl(vntT)
    dblTol = CDbl(vntTol)
    
    ' (Observed - True)/True
    If dblT <> 0 Then
        dblCmp = Abs((dblO - dblT) / dblT)
    Else
        dblCmp = Abs(dblO - dblT)
    End If

    If dblCmp < dblTol Then
        Near = True
    Else
        Near = False
    End If
    
End Function

Public Function Relerr(vntO As Variant, vntT As Variant) As Variant
    Dim dblO As Double
    Dim dblT As Double
    
    dblO = CDbl(vntO)
    dblT = CDbl(vntT)
    
    ' (Observed - True)/True
    If dblT <> 0 Then
        Relerr = (dblO - dblT) / dblT
    Else
        Relerr = dblO - dblT
    End If

End Function
