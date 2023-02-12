Attribute VB_Name = "mdlUtils"
Const TAGLEN As Integer = 30

Public Function DisplayString(sTag As String, _
                              sVal As String, _
                              Optional iTagLen As Integer = TAGLEN, _
                              Optional sUnit As String = "", _
                              Optional sComment As String = "") As String
    Dim sTagFormat As String
    
    sTagFormat = "!" & String(iTagLen, "@")
    DisplayString = sComment & Format(Left(sTag, iTagLen), sTagFormat) & ": " & sVal & " " & sUnit & vbCrLf
End Function

Public Function DisplayDouble(sTag As String, _
                              dblVal As Double, _
                              Optional iTagLen As Integer = TAGLEN, _
                              Optional sValFormat As String = "0.0000E+00", _
                              Optional sUnit As String = "", _
                              Optional sComment As String = "") As String
                              
    Dim sTagFormat As String
    
    sTagFormat = "!" & String(iTagLen, "@")
    DisplayDouble = sComment & Format(Left(sTag, iTagLen), sTagFormat) & ": " & Format(dblVal, sValFormat) & " " & sUnit & vbCrLf
End Function

Public Function DisplayLong(sTag As String, _
                              lVal As Long, _
                              Optional iTagLen As Integer = TAGLEN, _
                              Optional sValFormat As String = "0", _
                              Optional sUnit As String = "", _
                              Optional sComment As String = "") As String
                              
    Dim sTagFormat As String
    
    sTagFormat = "!" & String(iTagLen, "@")
    DisplayLong = sComment & Format(Left(sTag, iTagLen), sTagFormat) & ": " & Format(lVal, sValFormat) & " " & sUnit & vbCrLf
End Function

Public Function DisplayInteger(sTag As String, _
                              iVal As Integer, _
                              Optional iTagLen As Integer = TAGLEN, _
                              Optional sValFormat As String = "0", _
                              Optional sUnit As String = "", _
                              Optional sComment As String = "") As String
                              
    Dim sTagFormat As String
    
    sTagFormat = "!" & String(iTagLen, "@")
    DisplayInteger = sComment & Format(Left(sTag, iTagLen), sTagFormat) & ": " & Format(iVal, sValFormat) & " " & sUnit & vbCrLf
End Function

Public Function FtIn(dblVal As Double) As String
    Dim dblFt As Double
    Dim dblIn As Double
    
    ' No floating point mod function in vba
    dblFt = Int(dblVal / 12#)
    dblIn = dblVal - 12# * dblFt
    
    FtIn = Format(dblFt, "0") & " ft " & Format(dblIn, "0.0") & " in"

End Function

Function IfBorZ(val As Variant, alt As Variant) As Variant
    Application.Volatile False
    If val = "" Or val = 0 Then
        IfBorZ = alt
    Else
        IfBorZ = val
    End If
End Function

Function IsNaN(ByVal x As Variant) As Boolean
    IsNaN = Excel.WorksheetFunction.IsNA(x)
End Function

Function NaN() As Variant
    NaN = CVErr(xlErrNA)
End Function


Public Function StopCheck() As Boolean
    Dim strFileName As String
    Dim strFileExists As String
    Dim sDirRun As String
    Dim sVersion As String
 
    sDirRun = Application.Names("dirRun").RefersToRange.Text
    If Right(sDirRun, 1) = "\" Then sDirRun = Mid(sDirRun, 1, Len(sDirRun) - 1)
    sVersion = Application.Names("version").RefersToRange.Text
    strFileName = sDirRun & "\" & sVersion & "\stop.txt"
    strFileExists = Dir(strFileName)
 
    If strFileExists = "" Then
        StopCheck = False
    Else
        StopCheck = True
    End If
 
End Function

Public Function SwanRule(x As Variant, coef() As Variant) As Variant
    Dim i As Integer
    Dim dbl As Variant
    
    ' Evaluate a polynomial at x given coefficients coef
    ' Coefficients are ordered from highest power to constant.
    
    dbl = coef(LBound(coef))
    For i = LBound(coef) + 1 To UBound(coef)
        dbl = coef(i) + x * dbl
    Next i

    SwanRule = dbl

End Function


Public Function SwanRuleRange(rngX As Variant, rngCoef As Variant) As Variant
    Dim i As Integer
    Dim x As Variant
    Dim coef() As Variant
    
    x = rngX.Value2
    ReDim coef(1 To rngCoef.Cells.Count)
    
    For i = 1 To rngCoef.Cells.Count
        coef(i) = rngCoef.Cells(i).Value2
    Next i

    SwanRuleRange = SwanRule(x, coef)

End Function


Public Sub FixBP()
    Stop
End Sub
