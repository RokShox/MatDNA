VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMatComRange"
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
Dim meCopyMode As MatFracCopyMode
Dim msName As String
Dim msThisClass As String
Dim mblnSealed As Boolean

Private Sub Class_Initialize()

    msThisClass = "clsMatComRange"
    Set mwksProto = Worksheets("Materials")

    meMode = Atom
    meCopyMode = ByValue
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

Property Get CopyMode() As MatFracCopyMode
    CopyMode = meCopyMode
End Property

Property Let CopyMode(e As MatFracCopyMode)
    meCopyMode = e
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
    
'    If Not mblnSealed Then
'        MsgBox "Must be sealed", vbCritical, msThisClass
'        GoTo Fail:
'    End If
    
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
    Dim cCon As clsMatConRange
        
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


Property Get FractionConInCom(sCon As String, eMode As CompositionMode) As Double
    Dim cCon As clsMatConRange

    If Not Me.Exists(sCon) Then
        FractionConInCom = 0#
        Exit Property
    End If
    
    If Me.mode = eMode Then
        FractionConInCom = Me.Fraction(sCon)
    Else
        Set cCon = Me.Constituent(sCon)
        Select Case eMode
            Case Atom
                ' Mat is composed by mass, need atom fraction
                FractionConInCom = Me.Fraction(sCon) / cCon.AValue * Me.AValue
            Case Mass
                ' Mat is composed by atom, need mass fraction
                FractionConInCom = Me.Fraction(sCon) * cCon.AValue / Me.AValue
        End Select
        Set cCon = Nothing
    End If

End Property


Property Get FractionIsoInCom(sIso As String, eMode As CompositionMode) As Double
    Dim cCon As clsMatConRange
    Dim alConKeys As ArrayList
    Dim vntKey As Variant
    Dim dblFracCom As Double
    Dim dblFracCon As Double
    Dim sCon As String
    
    Set alConKeys = Me.ConList
    
    For Each vntKey In alConKeys
        Set cCon = Me.Constituent(vntKey)
        sCon = cCon.Name

        If Not cCon.Exists(sIso) Then
            dblFracCon = 0#
        Else
            dblFracCon = cCon.FractionIsoInCon(sIso, eMode)
        End If
        dblFracCom = dblFracCom + dblFracCon * Me.FractionConInCom(sCon, eMode)
        Set cCon = Nothing
        
    Next vntKey
    
    Set alConKeys = Nothing
    FractionIsoInCom = dblFracCom

End Property

Public Sub Add(cCon As clsMatConRange, rngFrac As Range)
    Dim sCon As String

    If Me.IsSealed Then
        MsgBox "Cannot add constituent. Component is sealed", vbCritical, msThisClass
        GoTo Fail:
    End If
    
    sCon = cCon.Name
    If mdCon.Exists(sCon) Then
        MsgBox "Constituent already exists", vbCritical, msThisClass
        GoTo Fail:
    Else
        mdCon.Add sCon, cCon
        mdFrac.Add sCon, rngFrac
    End If
    Exit Sub
    
Fail:
    Exit Sub
    
End Sub

Public Sub Seal()
    Dim cCon As clsMatConRange
    Dim key As Variant
    
    If Not mblnSealed Then
        For Each key In mdCon.Keys
            Set cCon = mdCon.Item(key)
            cCon.Seal
        Next
'        normalize
        mblnSealed = True
    End If
End Sub

Public Sub InitFromRange(rng As Range)
    Dim cCon As clsMatConRange
    Dim sCom As String
    Dim rngFrac As Range
    Dim rngAValue As Range
    Dim dblATbl As Double
    Dim dblACalc As Double

    Do Until rng.Cells(1, 1).Offset(0, oc_Comp_ComRows).Value2 = 0
        
        Set cCon = New clsMatConRange
        With cCon
            .mode = Me.mode
            .CopyMode = Me.CopyMode
            .InitFromRange rng
        End With
        
        Select Case meMode
            Case Atom
                Set rngFrac = rng.Cells(1, 1).Offset(0, oc_Comp_ConAfrac)
            Case Mass
                Set rngFrac = rng.Cells(1, 1).Offset(0, oc_Comp_ConMfrac)
        End Select
        Me.Add cCon, rngFrac

        Set rng = rng.Offset(1, 0)
    Loop

    sCom = rng.Cells(1, 1).Offset(0, oc_Comp_Com).Text
    Me.Name = sCom
    Me.Seal
    
    Set rngAValue = rng.Cells(1, 1).Offset(0, oc_Comp_AValue)
    dblATbl = rngAValue.Value2
    dblACalc = Me.AValue
    
    Debug.Print msThisClass & ": " & sCom & " A Value: Table " & Format(dblATbl, "0.0000E+00") & " Calculated " & Format(dblACalc, "0.0000E+00")

    ' Calc engine may call this multiple times with some ranges uncalculated
    ' Only do this test if data is calculated
    If Not (IsEmpty(rngAValue.Value)) And (Abs(dblATbl - dblACalc) > DIST_TOL) Then
        MsgBox msThisClass & ": Calculated A differs from tabulated value", vbCritical, msThisClass
        Exit Sub
    End If

End Sub

Public Function Display() As String
    
    Dim cCon As clsMatConRange
    Dim dblFrac As Double
    Dim key As Variant
    Dim alCon As ArrayList
    
    Me.Seal
    Set alCon = ConList()
    
    Display = ""
    For Each key In alCon
        Set cCon = mdCon.Item(key)
        dblFrac = CDbl(mdFrac.Item(key).Value2)
        
        Display = Display & cCon.Display
        
        Display = Display & vbTab & _
        cCon.Name & vbTab & vbTab & _
        Format(cCon.AValue, "0.000000") & vbTab & _
        Me.ModeText & vbTab & _
        Format(dblFrac, "0.000000") & vbCrLf
    Next

End Function

Public Sub DisplayExcel(rng As Range)
    Dim cCon As clsMatConRange
    Dim rngFrac As Range
    Dim dblFrac As Double
    Dim key As Variant
    Dim alCon As ArrayList
    
    Set alCon = ConList()
    
    For Each key In alCon
        Set cCon = mdCon.Item(key)
        Set rngFrac = mdFrac.Item(key)
        dblFrac = CDbl(rngFrac.Value2)
        
        cCon.DisplayExcel rng
        
        mrngProtoCon.Copy rng
        ' Fix A value calc if modes differ
        If Me.mode = Mass And cCon.mode = Atom Then
            mwksProto.Names("protoCompositionConAtom").RefersToRange.Cells(1, 1).Offset(0, oc_Comp_AValue).Copy rng.Offset(0, oc_Comp_AValue)
        ElseIf Me.mode = Atom And cCon.mode = Mass Then
            mwksProto.Names("protoCompositionConMass").RefersToRange.Cells(1, 1).Offset(0, oc_Comp_AValue).Copy rng.Offset(0, oc_Comp_AValue)
        End If
        
        rng.Cells(1, 1).Offset(0, oc_Comp_Con).Value2 = cCon.Name
        Select Case meMode
            Case Atom
                Select Case meCopyMode
                    Case ByValue
                        rng.Cells(1, 1).Offset(0, oc_Comp_ConAfrac).Value2 = dblFrac
                    Case ByReference
                        rng.Cells(1, 1).Offset(0, oc_Comp_ConAfrac).Formula = "=" & rngFrac.Address(False, False)
                    Case ByFormula
                        rngFrac.Copy
                        rng.Cells(1, 1).Offset(0, oc_Comp_ConAfrac).PasteSpecial xlPasteFormulasAndNumberFormats
                End Select
            Case Mass
                Select Case meCopyMode
                    Case ByValue
                        rng.Cells(1, 1).Offset(0, oc_Comp_ConMfrac).Value2 = dblFrac
                    Case ByReference
                        rng.Cells(1, 1).Offset(0, oc_Comp_ConMfrac).Formula = "=" & rngFrac.Address(False, False)
                    Case ByFormula
                        rngFrac.Copy
                        rng.Cells(1, 1).Offset(0, oc_Comp_ConMfrac).PasteSpecial xlPasteFormulasAndNumberFormats
                End Select
        End Select
        Set rng = rng.Offset(1, 0)
    Next key

End Sub

Public Function Fraction(key As Variant) As Double

    If Not mdFrac.Exists(key) Then
        MsgBox "Key " & key & " does not exist", vbCritical, msThisClass
        GoTo Fail:
    End If
        
    Fraction = CDbl(mdFrac.Item(key).Value2)
    
    Exit Function
Fail:
    End
End Function

Public Function FractionRange(key As Variant) As Range

    If Not mdCon.Exists(key) Then
        MsgBox "con " & key & " does not exist", vbCritical, msThisClass
        GoTo Fail:
    End If
    Set FractionRange = mdFrac.Item(key)
    Exit Function
        
Fail:
    Set FractionRange = Nothing
    Exit Function
End Function

Property Get Constituent(key As Variant) As clsMatConRange

    If Not mdCon.Exists(key) Then
        MsgBox "Key " & key & " does not exist", vbCritical, msThisClass
        GoTo Fail:
    End If
        
    Set Constituent = mdCon.Item(key)
    
    Exit Property
Fail:
    End

End Property

Public Function ConList() As ArrayList
    Dim key As Variant

    Set ConList = New ArrayList
    For Each key In mdCon.Keys
        ConList.Add key
    Next
    ConList.Sort
    
End Function

Private Sub normalize()
    Dim cCon As clsMatConRange
    Dim dblSum As Double
    Dim key As Variant
    
'    dblSum = 0#
'    For Each key In mdCon.Keys
'        dblSum = dblSum + mdFrac.Item(key)
'    Next
'
'    If Abs(dblSum - 1#) > DIST_TOL Then
'        For Each key In mdCon.Keys
'            mdFrac.Item(key) = mdFrac.Item(key) / dblSum
'        Next
'    End If

End Sub

Private Function calculateAbyAtom() As Double
    Dim cCon As clsMatConRange
    Dim dblA As Double
    Dim key As Variant
    
    dblA = 0#
    For Each key In mdCon.Keys
        Set cCon = mdCon.Item(key)
        dblA = dblA + CDbl(mdFrac.Item(key).Value2) * cCon.AValue
    Next
    calculateAbyAtom = dblA
    Exit Function
    
Fail:
    calculateAbyAtom = 0#
    
End Function

Private Function calculateAbyMass() As Double
    Dim cCon As clsMatConRange
    Dim dblA As Double
    Dim key As Variant
    
    dblA = 0#
    For Each key In mdCon.Keys
        Set cCon = mdCon.Item(key)
        dblA = dblA + CDbl(mdFrac.Item(key).Value2) / cCon.AValue
    Next
    calculateAbyMass = 1# / dblA
    Exit Function
    
Fail:
    calculateAbyMass = 0#
    
End Function






