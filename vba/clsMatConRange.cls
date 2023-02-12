VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMatConRange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' Copyright (c) 2022 Henkel Technical Computing, LLC. All rights reserved.
'
Option Explicit
Option Compare Text

Dim mrngProtoIso As Range
Dim mdIso As Scripting.Dictionary
Dim mdFrac As Scripting.Dictionary
Dim mwksProto As Worksheet

Dim meMode As CompositionMode
Dim meCopyMode As MatFracCopyMode
Dim msName As String
Dim msThisClass As String
Dim mblnSealed As Boolean

Private Sub Class_Initialize()

    msThisClass = "clsMatConRange"
    Set mwksProto = Worksheets("Materials")

    meMode = Atom
    meCopyMode = ByValue
    Set mrngProtoIso = mwksProto.Names("protoCompositionIsoAtom").RefersToRange
    Set mdIso = New Scripting.Dictionary
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
            Set mrngProtoIso = mwksProto.Names("protoCompositionIsoAtom").RefersToRange
        Case Mass
            Set mrngProtoIso = mwksProto.Names("protoCompositionIsoMass").RefersToRange
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
'        MsgBox "Constituent must be sealed", vbCritical, msThisClass
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
    Count = mdIso.Count
End Property

Property Get Exists(sIso As String) As Boolean
    Exists = mdIso.Exists(sIso)
End Property

Property Get UniqueIsos() As ArrayList
    Set UniqueIsos = isoList()
End Property


Property Get FractionIsoInCon(sIso As String, eMode As CompositionMode) As Double
    Dim cIso As clsMatIso

    If Not Me.Exists(sIso) Then
        FractionIsoInCon = 0#
        Exit Property
    End If
    
    If Me.mode = eMode Then
        FractionIsoInCon = Me.Fraction(sIso)
    Else
        Set cIso = Me.Isotope(sIso)
        Select Case eMode
            Case Atom
                ' Mat is composed by mass, need atom fraction
                FractionIsoInCon = Me.Fraction(sIso) / cIso.MolecularWeight * Me.AValue
            Case Mass
                ' Mat is composed by atom, need mass fraction
                FractionIsoInCon = Me.Fraction(sIso) * cIso.MolecularWeight / Me.AValue
        End Select
        Set cIso = Nothing
    End If

End Property

Public Sub Add(cIso As clsMatIso, rngFrac As Range)
    Dim sIso As String

    If Me.IsSealed Then
        MsgBox "Cannot add iso. Constituent is sealed", vbCritical, msThisClass
        GoTo Fail:
    End If
    
    sIso = cIso.Iso
    If mdIso.Exists(sIso) Then
        MsgBox "Isotope already exists", vbCritical, msThisClass
        GoTo Fail:
    Else
        mdIso.Add sIso, cIso
        mdFrac.Add sIso, rngFrac
    End If
    Exit Sub
    
Fail:
    Exit Sub
    
End Sub

Public Sub Seal()
    If Not mblnSealed Then
'        normalize
        mblnSealed = True
    End If
End Sub

Public Sub InitFromRange(rng As Range)
    Dim cIso As clsMatIso
    Dim sIso As String
    Dim sCon As String
    Dim rngFrac As Range
    Dim rngAValue As Range
    Dim dblATbl As Double
    Dim dblACalc As Double

    Do Until rng.Cells(1, 1).Offset(0, oc_Comp_ConRows).Value2 = 0
    
        sIso = rng.Cells(1, 1).Offset(0, oc_Comp_Iso).Text
        Set cIso = New clsMatIso
        
        With cIso
            .Iso = sIso
        End With
        Select Case meMode
            Case Atom
                Set rngFrac = rng.Cells(1, 1).Offset(0, oc_Comp_IsoAfrac)
            Case Mass
                Set rngFrac = rng.Cells(1, 1).Offset(0, oc_Comp_IsoMfrac)
        End Select
        Me.Add cIso, rngFrac

        Set rng = rng.Offset(1, 0)
    Loop

    sCon = rng.Cells(1, 1).Offset(0, oc_Comp_Con).Text
    Me.Name = sCon
    Me.Seal
    
    Set rngAValue = rng.Cells(1, 1).Offset(0, oc_Comp_AValue)
    dblATbl = rngAValue.Value2
    dblACalc = Me.AValue
    
    Debug.Print msThisClass & ": " & sCon & " A Value: Table " & Format(dblATbl, "0.0000E+00") & " Calculated " & Format(dblACalc, "0.0000E+00")

    ' Calc engine may call this multiple times with some ranges uncalculated
    ' Only do this test if data is calculated
    If Not (IsEmpty(rngAValue.Value)) And (Abs(dblATbl - dblACalc) > DIST_TOL) Then
        MsgBox msThisClass & ": Calculated A differs from tabulated value", vbCritical, msThisClass
        Exit Sub
    End If

End Sub


Public Function Display() As String
    
    Dim cIso As clsMatIso
    Dim dblFrac As Double
    Dim key As Variant
    Dim alIso As ArrayList
    
    Set alIso = isoList()
    
    Display = ""
    For Each key In alIso
        Set cIso = mdIso.Item(key)
        dblFrac = CDbl(mdFrac.Item(key).Value2)
        
        Display = Display & vbTab & vbTab & _
        cIso.Iso & vbTab & _
        Format(cIso.MolecularWeight, "0.000000") & vbTab & _
        Me.ModeText & vbTab & _
        Format(dblFrac, "0.000000") & vbCrLf
    Next

End Function

Public Sub DisplayExcel(rng As Range)
    Dim cIso As clsMatIso
    Dim dblFrac As Double
    Dim rngFrac As Range
    Dim key As Variant
    Dim alIso As ArrayList
    Dim rngStart As Range
    
    Set alIso = isoList()
    Set rngStart = rng
    
    For Each key In alIso
        Set cIso = mdIso.Item(key)
        Set rngFrac = mdFrac.Item(key)
        dblFrac = CDbl(rngFrac.Value2)
        
        mrngProtoIso.Copy rng
        rng.Cells(1, 1).Offset(0, oc_Comp_Iso).Value2 = cIso.Iso
        Select Case meMode
        Case Atom
            Select Case meCopyMode
                Case ByValue
                    rng.Cells(1, 1).Offset(0, oc_Comp_IsoAfrac).Value2 = dblFrac
                Case ByReference
                    rng.Cells(1, 1).Offset(0, oc_Comp_IsoAfrac).Formula = "=" & rngFrac.Address(False, False)
                Case ByFormula
                    rngFrac.Copy
                    rng.Cells(1, 1).Offset(0, oc_Comp_IsoAfrac).PasteSpecial xlPasteFormulasAndNumberFormats
                    'rng.Cells(1, 1).Offset(0, oc_Comp_IsoAfrac).Formula = rngFrac.Formula
            End Select
        Case Mass
            Select Case meCopyMode
                Case ByValue
                    rng.Cells(1, 1).Offset(0, oc_Comp_IsoMfrac).Value2 = dblFrac
                Case ByReference
                    rng.Cells(1, 1).Offset(0, oc_Comp_IsoMfrac).Formula = "=" & rngFrac.Address(False, False)
                Case ByFormula
                    rngFrac.Copy
                    rng.Cells(1, 1).Offset(0, oc_Comp_IsoMfrac).PasteSpecial xlPasteFormulasAndNumberFormats
'                    rng.Cells(1, 1).Offset(0, oc_Comp_IsoMfrac).Formula = rngFrac.Formula
            End Select
        End Select
        Set rng = rng.Offset(1, 0)
    Next key
    
    ' Fix borders
    mdlExcel.BorderInside rngStart.Offset(0, oc_Comp_Iso).Resize(alIso.Count, 6)
    With rng.Offset(0, oc_Comp_Con).Resize(1, 7)
         With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
    End With
    
End Sub

Public Function Fraction(key As Variant) As Double
    If Not mdIso.Exists(key) Then
        MsgBox "iso " & key & " does not exist", vbCritical, msThisClass
        GoTo Fail:
    End If
    Fraction = CDbl(mdFrac.Item(key).Value2)
    Exit Function
        
Fail:
    Fraction = 0#
    Exit Function
End Function

Public Function FractionRange(key As Variant) As Range
    If Not mdIso.Exists(key) Then
        MsgBox "iso " & key & " does not exist", vbCritical, msThisClass
        GoTo Fail:
    End If
    Set FractionRange = mdFrac.Item(key)
    Exit Function
        
Fail:
    Set FractionRange = Nothing
    Exit Function
End Function

Property Get Isotope(key As Variant) As clsMatIso

    If Not mdIso.Exists(key) Then
        MsgBox "Key " & key & " does not exist", vbCritical, msThisClass
        GoTo Fail:
    End If
        
    Set Isotope = mdIso.Item(key)
    
    Exit Property
Fail:
    End

End Property

Private Function isoList() As ArrayList
    Dim key As Variant

    Set isoList = New ArrayList
    For Each key In mdIso.Keys
        isoList.Add key
    Next
    isoList.Sort
    
End Function


Private Sub normalize()
    Dim cIso As clsMatIso
    Dim dblSum As Double
    Dim key As Variant
    
'    dblSum = 0#
'    For Each key In mdIso.Keys
'        dblSum = dblSum + mdFrac.Item(key).Value2
'    Next
'
'    If Abs(dblSum - 1#) > DIST_TOL Then
'        For Each key In mdIso.Keys
'            mdFrac.Item(key) = mdFrac.Item(key) / dblSum
'        Next
'    End If

End Sub


Private Function calculateAbyAtom() As Double
    Dim cIso As clsMatIso
    Dim dblA As Double
    Dim key As Variant
    
    dblA = 0#
    For Each key In mdIso.Keys
        Set cIso = mdIso.Item(key)
        dblA = dblA + CDbl(mdFrac.Item(key).Value2) * cIso.MolecularWeight
    Next
    calculateAbyAtom = dblA
    Exit Function
    
Fail:
    calculateAbyAtom = 0#
    
End Function

Private Function calculateAbyMass() As Double
    Dim cIso As clsMatIso
    Dim dblA As Double
    Dim key As Variant
    
    dblA = 0#
    For Each key In mdIso.Keys
        Set cIso = mdIso.Item(key)
        dblA = dblA + CDbl(mdFrac.Item(key).Value2) / cIso.MolecularWeight
    Next
    calculateAbyMass = 1# / dblA
    Exit Function
    
Fail:
    calculateAbyMass = 0#
    
End Function




