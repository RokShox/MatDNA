VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMatCom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' Copyright (c) 2022 Henkel Technical Computing, LLC. All rights reserved.
'
Option Explicit
Option Compare Text

Dim mrngProtoCon As Range
Dim mdCon As Scripting.Dictionary
Dim mdFrac As Scripting.Dictionary
Dim mwksProto As Worksheet

Dim meMode As CompositionMode
Dim msName As String
Dim mblnSealed As Boolean
Dim mdblFrac As Double


Private Sub Class_Initialize()

    Set mwksProto = Worksheets("Materials")

    meMode = Atom
    Set mrngProtoCon = mwksProto.Names("protoCompositionConAtom").RefersToRange
    Set mdCon = New Scripting.Dictionary
    Set mdFrac = New Scripting.Dictionary
    mblnSealed = False
    
End Sub

Property Get mode() As CompositionMode
    mode = meMode
End Property

Property Let mode(e As CompositionMode)
    meMode = e
    Select Case meMode
        Case Atom
            Set mrngProtoCon = mwksProto.Names("protoCompositionConAtom").RefersToRange
        Case Mass
            Set mrngProtoCon = mwksProto.Names("protoCompositionConMass").RefersToRange
    End Select
End Property

Property Get ModeText() As String
    Select Case meMode
        Case Atom
            ModeText = "Atom"
        Case Mass
            ModeText = "Mass"
    End Select

End Property

Property Get Name() As String
    Name = msName
End Property

Property Let Name(s As String)
    msName = s
End Property

Property Get AValue() As Double
    
    If Not mblnSealed Then
        MsgBox "Constituent must be sealed", vbCritical, "clsMatCon"
        GoTo Fail:
    End If
    
    Select Case meMode
        Case Atom
            AValue = calculateAbyAtom()
        Case Mass
            AValue = calculateAbyMass()
    End Select
    Exit Property
    
Fail:
    AValue = -1#
    
End Property

Property Get IsSealed() As Boolean
    IsSealed = mblnSealed
End Property

Property Get Count() As Integer
    Count = mdCon.Count
End Property

Property Get Exists(sCon As String) As Boolean
    Exists = mdCon.Exists(sCon)
End Property

Property Get UniqueIsos() As ArrayList
    Dim alIso As ArrayList
    Dim alConIso As ArrayList
    Dim key As Variant
    Dim Iso As Variant
    Dim cCon As clsMatCon
        
    Set alIso = New ArrayList
        
    For Each key In mdCon.Keys
        Set cCon = mdCon.Item(key)
        Set alConIso = cCon.UniqueIsos
        
        For Each Iso In alConIso
            If Not alIso.Contains(Iso) Then
                alIso.Add Iso
            End If
        Next Iso
    Next key
        
    alIso.Sort
    Set UniqueIsos = alIso

End Property


Property Get Constituent(key As Variant) As clsMatCon

    If Not mdCon.Exists(key) Then
        MsgBox "Key " & key & " does not exist", vbCritical, "clsMatCom"
        GoTo Fail:
    End If
        
    Set Constituent = mdCon.Item(key)
    
    Exit Property
Fail:
    End

End Property

Property Get Fraction(key As Variant) As Double

    If Not mdFrac.Exists(key) Then
        MsgBox "Key " & key & " does not exist", vbCritical, "clsMatCom"
        GoTo Fail:
    End If
        
    Fraction = mdFrac.Item(key)
    
    Exit Property
Fail:
    End

End Property


Public Sub Add(cCon As clsMatCon, frac As Double)
    Dim sCon As String

    If Me.IsSealed Then
        MsgBox "Cannot add constituent. Component is sealed", vbCritical, "clsMatCom"
        GoTo Fail:
    End If
    
    sCon = cCon.Name
    If mdCon.Exists(sCon) Then
        mdFrac(sCon) = mdFrac(sCon) + frac
    Else
        mdCon.Add sCon, cCon
        mdFrac.Add sCon, frac
    End If
    Exit Sub
    
Fail:
    Exit Sub
    
End Sub

Public Sub Seal()
    Dim cCon As clsMatCon
    Dim key As Variant
    
    If Not mblnSealed Then
                
        For Each key In mdCon.Keys
            Set cCon = mdCon.Item(key)
            cCon.Seal
        Next
       
        normalize
        mblnSealed = True
    End If
End Sub

Public Sub InitFromRange(rng As Range)
    Dim cCon As clsMatCon
    Dim sCom As String
    Dim dblFrac As Double
    Dim dblATbl As Double
    Dim dblACalc As Double

    Me.mode = Atom

    Do Until rng.Cells(1, 1).Offset(0, oc_Comp_ComRows).Value2 = 0
        
        Set cCon = New clsMatCon
        cCon.InitFromRange rng
    
        Select Case meMode
            Case Atom
                dblFrac = rng.Cells(1, 1).Offset(0, oc_Comp_ConAfrac).Value2
            Case Mass
                dblFrac = rng.Cells(1, 1).Offset(0, oc_Comp_ConMfrac).Value2
        End Select
        Me.Add cCon, dblFrac

        Set rng = rng.Offset(1, 0)
    Loop

    sCom = rng.Cells(1, 1).Offset(0, oc_Comp_Com).Text
    Me.Name = sCom
    Me.Seal
    
    dblATbl = rng.Cells(1, 1).Offset(0, oc_Comp_AValue).Value2
    dblACalc = Me.AValue
    
    Debug.Print "A Value: Table " & Format(dblATbl, "0.0000E+00") & " Calculated " & Format(dblACalc, "0.0000E+00")

    If Abs(dblATbl - dblACalc) > DIST_TOL Then
        MsgBox "Calculated A differs from tabulated value", vbCritical, "clsMatCom"
        Exit Sub
    End If

End Sub


Public Function Display() As String
    
    Dim cCon As clsMatCon
    Dim dblFrac As Double
    Dim key As Variant
    Dim alCon As ArrayList
    
    Set alCon = ConList()
    
    Display = ""
    For Each key In alCon
        Set cCon = mdCon.Item(key)
        dblFrac = mdFrac.Item(key)
        
        Display = Display & cCon.Display
        
        Display = Display & vbTab & _
        cCon.Name & vbTab & vbTab & _
        Format(cCon.AValue, "0.000000") & vbTab & _
        Me.ModeText & vbTab & _
        Format(dblFrac, "0.000000") & vbCrLf
    Next

End Function

Public Sub DisplayExcel(rng As Range)
    Dim cCon As clsMatCon
    Dim dblFrac As Double
    Dim key As Variant
    Dim alCon As ArrayList
    
    Set alCon = ConList()
    
    For Each key In alCon
        Set cCon = mdCon.Item(key)
        dblFrac = mdFrac.Item(key)
        
        cCon.DisplayExcel rng
        
        mrngProtoCon.Copy rng
        rng.Cells(1, 1).Offset(0, oc_Comp_Con).Value2 = cCon.Name
        Select Case meMode
        Case Atom
            rng.Cells(1, 1).Offset(0, oc_Comp_ConAfrac).Value2 = dblFrac
        Case Mass
            rng.Cells(1, 1).Offset(0, oc_Comp_ConMfrac).Value2 = dblFrac
        End Select
        Set rng = rng.Offset(1, 0)
    Next key
    
End Sub


Public Function ConList() As ArrayList
    Dim key As Variant

    Set ConList = New ArrayList
    For Each key In mdCon.Keys
        ConList.Add key
    Next
    ConList.Sort
    
End Function


Private Sub normalize()
    Dim cCon As clsMatCon
    Dim dblSum As Double
    Dim key As Variant
    
    dblSum = 0#
    For Each key In mdCon.Keys
        dblSum = dblSum + mdFrac.Item(key)
    Next
    
    If Abs(dblSum - 1#) > DIST_TOL Then
        For Each key In mdCon.Keys
            mdFrac.Item(key) = mdFrac.Item(key) / dblSum
        Next
    End If

End Sub


Private Function calculateAbyAtom() As Double
    Dim cCon As clsMatCon
    Dim dblA As Double
    Dim key As Variant
    
    dblA = 0#
    For Each key In mdCon.Keys
        Set cCon = mdCon.Item(key)
        dblA = dblA + mdFrac.Item(key) * cCon.AValue
    Next
    calculateAbyAtom = dblA
    Exit Function
    
Fail:
    calculateAbyAtom = 0#
    
End Function

Private Function calculateAbyMass() As Double
    Dim cCon As clsMatCon
    Dim dblA As Double
    Dim key As Variant
    
    dblA = 0#
    For Each key In mdCon.Keys
        Set cCon = mdCon.Item(key)
        dblA = dblA + mdFrac.Item(key) / cCon.AValue
    Next
    calculateAbyMass = 1# / dblA
    Exit Function
    
Fail:
    calculateAbyMass = 0#
    
End Function




