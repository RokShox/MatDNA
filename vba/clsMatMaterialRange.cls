VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMatMaterialRange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' Copyright (c) 2022 Henkel Technical Computing, LLC. All rights reserved.
'
Option Explicit
Option Compare Text

Dim mrngProtoCom As Range
Dim mrngProtoMat As Range
Dim mrngProtoHeader As Range
Dim mrngProtoFooter As Range
Dim mrngAValue As Range
Dim mdCom As Scripting.Dictionary
Dim mdFrac As Scripting.Dictionary
Dim mdFracDisplay As Scripting.Dictionary
Dim mwksProto As Worksheet

Dim meMode As CompositionMode
Dim meCopyMode As MatFracCopyMode
Dim msName As String
Dim msThisClass As String
Dim mblnSealed As Boolean
Dim mnLastRow As Integer
Dim mnNumRows As Integer

Private Sub Class_Initialize()

    msThisClass = "clsMatMaterialRange"
    Set mwksProto = Worksheets("Materials")

    meMode = Mass
    meCopyMode = ByValue
    Set mrngProtoCom = mwksProto.Names("protoCompositionComMass").RefersToRange
    Set mrngProtoMat = mwksProto.Names("protoCompositionMatMass").RefersToRange
    Set mrngProtoHeader = mwksProto.Names("protoCompositionHeader").RefersToRange
    Set mrngProtoFooter = mwksProto.Names("protoCompositionFooter").RefersToRange
    Set mdCom = New Scripting.Dictionary
    Set mdFrac = New Scripting.Dictionary
    Set mdFracDisplay = New Scripting.Dictionary
    mblnSealed = False
    mnLastRow = 0
    mnNumRows = 0
    
End Sub

Property Get mode() As CompositionMode
    mode = meMode
End Property

Property Let mode(e As CompositionMode)
    meMode = e
    Select Case meMode
        Case Atom
            Set mrngProtoCom = mwksProto.Names("protoCompositionComAtom").RefersToRange
            Set mrngProtoMat = mwksProto.Names("protoCompositionMatAtom").RefersToRange
        Case Mass
            Set mrngProtoCom = mwksProto.Names("protoCompositionComMass").RefersToRange
            Set mrngProtoMat = mwksProto.Names("protoCompositionMatMass").RefersToRange
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

Property Get AValueRange() As Range
    Set AValueRange = mrngAValue
End Property

Property Get IsSealed() As Boolean
    IsSealed = mblnSealed
End Property

Property Get Count() As Integer
    Count = mdCom.Count
End Property

Property Get Exists(sCom As String) As Boolean
    Exists = mdCom.Exists(sCom)
End Property

Property Get LastRow() As Integer
    LastRow = mnLastRow
End Property

Property Get UniqueIsos() As ArrayList
    Dim alIso As ArrayList
    Dim alComIso As ArrayList
    Dim key As Variant
    Dim Iso As Variant
    Dim cCom As clsMatComRange
        
    Set alIso = New ArrayList
        
    For Each key In mdCom.Keys
        Set cCom = mdCom.Item(key)
        Set alComIso = cCom.UniqueIsos
        
        For Each Iso In alComIso
            If Not alIso.Contains(Iso) Then
                alIso.Add Iso
            End If
        Next Iso
    Next key
        
    alIso.Sort
    Set UniqueIsos = alIso
End Property

Property Get NumRows() As Integer
    NumRows = mnNumRows
End Property


Property Get FractionComInMat(sCom As String, eMode As CompositionMode) As Double
    Dim cCom As clsMatComRange

    If Not Me.Exists(sCom) Then
        FractionComInMat = 0#
        Exit Property
    End If
    
    If Me.mode = eMode Then
        FractionComInMat = Me.Fraction(sCom)
    Else
        Set cCom = Me.Component(sCom)
        Select Case eMode
            Case Atom
                ' Mat is composed by mass, need atom fraction
                FractionComInMat = Me.Fraction(sCom) / cCom.AValue * Me.AValue
            Case Mass
                ' Mat is composed by atom, need mass fraction
                FractionComInMat = Me.Fraction(sCom) * cCom.AValue / Me.AValue
        End Select
        Set cCom = Nothing
    End If

End Property

Property Get FractionConInMat(sCon As String, eMode As CompositionMode) As Double
    Dim cCom As clsMatComRange
    Dim alComKeys As ArrayList
    Dim vntKey As Variant
    Dim dblFracMat As Double
    Dim dblFracCom As Double
    Dim sCom As String
    
    Set alComKeys = Me.ComList
    
    For Each vntKey In alComKeys
        Set cCom = Me.Component(vntKey)
        sCom = cCom.Name

        If Not cCom.Exists(sCon) Then
            dblFracCom = 0#
        Else
            dblFracCom = cCom.FractionConInCom(sCon, eMode)
        End If
        dblFracMat = dblFracMat + dblFracCom * Me.FractionComInMat(sCom, eMode)
        Set cCom = Nothing
        
    Next vntKey
    
    Set alComKeys = Nothing
    FractionConInMat = dblFracMat

End Property


Property Get FractionIsoInMat(sIso As String, eMode As CompositionMode) As Double
    Dim cCom As clsMatComRange
    Dim alComKeys As ArrayList
    Dim vntKey As Variant
    Dim dblFracMat As Double
    Dim dblFracCom As Double
    Dim sCom As String
    
    Set alComKeys = Me.ComList
    
    For Each vntKey In alComKeys
        Set cCom = Me.Component(vntKey)
        sCom = cCom.Name

        dblFracCom = cCom.FractionIsoInCom(sIso, eMode)
        dblFracMat = dblFracMat + dblFracCom * Me.FractionComInMat(sCom, eMode)
        Set cCom = Nothing
        
    Next vntKey
    
    Set alComKeys = Nothing
    FractionIsoInMat = dblFracMat

End Property

Public Sub Add(cCom As clsMatComRange, rngFrac As Range)
    Dim sCom As String
    Dim sTemp As String
    Dim iDup As Integer

    If Me.IsSealed Then
        MsgBox "Cannot add component. Material is sealed", vbCritical, msThisClass
        GoTo Fail:
    End If
    
    sCom = cCom.Name
    If mdCom.Exists(sCom) Then
        For iDup = 2 To 99
            sTemp = sCom & "-" & Format(iDup, "00")
            If Not mdCom.Exists(sTemp) Then
                cCom.Name = sTemp
                Exit For
            ElseIf iDup = 9 Then
                MsgBox "Cannot add duplicate component. Max of 9 duplicates", vbCritical, msThisClass
                GoTo Fail:
            End If
        Next iDup
    End If

    mdCom.Add cCom.Name, cCom
    mdFrac.Add cCom.Name, rngFrac
    Exit Sub
    
Fail:
    Exit Sub
    
End Sub

Public Sub Seal()
    Dim cCom As clsMatComRange
    Dim key As Variant
    
    If Not mblnSealed Then
        For Each key In mdCom.Keys
            Set cCom = mdCom.Item(key)
            cCom.Seal
        Next
       
'        normalize
        mblnSealed = True
    End If
End Sub

Public Sub InitFromRange(rng As Range)
    Dim cCom As clsMatComRange
    Dim sMat As String
    Dim rngFrac As Range
    Dim dblATbl As Double
    Dim dblACalc As Double
    
    sMat = rng.Cells(1, 1).Offset(-4, oc_Comp_Com).Text

    Do Until rng.Cells(1, 1).Offset(0, oc_Comp_MatRows).Value2 = 0
        
        Set cCom = New clsMatComRange
        With cCom
            .mode = Atom
            .CopyMode = Me.CopyMode
            .InitFromRange rng
        End With
        
        ' When a material is created using InitFromRange, the component fraction range just points
        ' to the original component fraction range.
        Select Case meMode
            Case Atom
                Set rngFrac = rng.Cells(1, 1).Offset(0, oc_Comp_ComAfrac)
            Case Mass
                Set rngFrac = rng.Cells(1, 1).Offset(0, oc_Comp_ComMfrac)
        End Select
        Me.Add cCom, rngFrac

        Set rng = rng.Offset(1, 0)
    Loop

    Me.Name = sMat
    Me.Seal
    
    Set mrngAValue = rng.Cells(1, 1).Offset(0, oc_Comp_AValue)
    dblATbl = mrngAValue.Value2
    dblACalc = Me.AValue
    
    Debug.Print msThisClass & ": " & sMat & " A Value: Table " & Format(dblATbl, "0.0000E+00") & " Calculated " & Format(dblACalc, "0.0000E+00")

    ' Calc engine may call this multiple times with some ranges uncalculated
    ' Only do this test if data is calculated
    If Not (IsEmpty(mrngAValue.Value)) And (Abs(dblATbl - dblACalc) > DIST_TOL) Then
        MsgBox msThisClass & ": Calculated A differs from tabulated value", vbCritical, msThisClass
        Exit Sub
    End If

End Sub

Public Function Display() As String
    
    Dim cCom As clsMatComRange
    Dim dblFrac As Double
    Dim key As Variant
    Dim alCom As ArrayList
    
    Me.Seal
    Set alCom = ComList()
    
    Display = ""
    For Each key In alCom
        Set cCom = mdCom.Item(key)
        dblFrac = CDbl(mdFrac.Item(key).Value2)
        
        Display = Display & cCom.Display
        
        Display = Display & _
        cCom.Name & vbTab & vbTab & vbTab & _
        Format(cCom.AValue, "0.000000") & vbTab & _
        Me.ModeText & vbTab & _
        Format(dblFrac, "0.000000") & vbCrLf
    Next
    
    Display = Display & _
        Me.Name & vbTab & vbTab & vbTab & _
        Format(Me.AValue, "0.000000") & vbTab & _
        Me.ModeText & vbTab & _
        Format(dblFrac, "0.000000") & vbCrLf

End Function

Public Sub DisplayExcel(rng As Range)
    Dim cCom As clsMatComRange
    Dim rngFrac As Range
    Dim dblFrac As Double
    Dim key As Variant
    Dim alCom As ArrayList
    Dim rngNumRows As Range
    Dim rngTblNameIso As Range
    Dim rngTblNameCon As Range
    Dim rngTblDataBgn As Range
    
    Const oc_NumRows As Integer = 1
    Const or_TblName As Integer = 2
    Const or_NumRows As Integer = 4
    
    Me.Seal
    
    mrngProtoHeader.Copy rng
    
    ' Locations of cells that need fixed
    Set rngNumRows = rng.Cells(1, 1).Offset(or_NumRows, oc_NumRows)
    Set rngTblNameIso = rng.Cells(1, 1).Offset(or_TblName, oc_Comp_IsoMfrac)
    Set rngTblNameCon = rng.Cells(1, 1).Offset(or_TblName, oc_Comp_ConMfrac)
    
    Set rng = rng.Offset(mrngProtoHeader.Rows.Count, 0)
    Set rngTblDataBgn = rng.Cells(1, 1)
        
    Set alCom = ComList()
    
    For Each key In alCom
        Set cCom = mdCom.Item(key)
        Set rngFrac = mdFrac.Item(key)
        dblFrac = CDbl(rngFrac.Value2)
        
        cCom.DisplayExcel rng
        
        mrngProtoCom.Copy rng
        ' Fix A value calc if modes differ
        If Me.mode = Mass And cCom.mode = Atom Then
            mwksProto.Names("protoCompositionComAtom").RefersToRange.Cells(1, 1).Offset(0, oc_Comp_AValue).Copy rng.Offset(0, oc_Comp_AValue)
        ElseIf Me.mode = Atom And cCom.mode = Mass Then
            mwksProto.Names("protoCompositionComMass").RefersToRange.Cells(1, 1).Offset(0, oc_Comp_AValue).Copy rng.Offset(0, oc_Comp_AValue)
        End If
        
        ' Keep track of component fraction range
        If Me.mode = Atom Then
            mdFracDisplay.Add cCom.Name, rng.Cells(1, 1).Offset(0, oc_Comp_ComAfrac)
        Else
            mdFracDisplay.Add cCom.Name, rng.Cells(1, 1).Offset(0, oc_Comp_ComMfrac)
        End If
        
        rng.Cells(1, 1).Offset(0, oc_Comp_Com).Value2 = cCom.Name
        
        Select Case meMode
            Case Atom
                Select Case meCopyMode
                    Case ByValue
                        rng.Cells(1, 1).Offset(0, oc_Comp_ComAfrac).Value2 = dblFrac
                    Case ByReference
                        rng.Cells(1, 1).Offset(0, oc_Comp_ComAfrac).Formula = "=" & rngFrac.Address(False, False)
                    Case ByFormula
                        rngFrac.Copy
                        rng.Cells(1, 1).Offset(0, oc_Comp_ComAfrac).PasteSpecial xlPasteFormulasAndNumberFormats
                End Select
            Case Mass
                Select Case meCopyMode
                    Case ByValue
                        rng.Cells(1, 1).Offset(0, oc_Comp_ComMfrac).Value2 = dblFrac
                    Case ByReference
                        rng.Cells(1, 1).Offset(0, oc_Comp_ComMfrac).Formula = "=" & rngFrac.Address(False, False)
                    Case ByFormula
                        rngFrac.Copy
                        rng.Cells(1, 1).Offset(0, oc_Comp_ComMfrac).PasteSpecial xlPasteFormulasAndNumberFormats
                End Select
        End Select
        Set rng = rng.Offset(1, 0)
    Next key
    
    ' Material line (last)
    mrngProtoMat.Copy rng
    
    ' Fix num rows
    ' Formula =ROW(AL3560)-ROW(AL3556)
    rngNumRows.Formula = _
        "=ROW(" & rng.Address(False, False) & ")" & _
        "-ROW(" & rngTblDataBgn.Offset(-1, 0).Address(False, False) & ")"
    mnNumRows = rng.Row - rngTblDataBgn.Offset(-1, 0).Row
    Set mrngAValue = rng.Offset(0, oc_Comp_AValue)
    
    ' Remove iso and con table names
'    rngTblNameIso.ClearContents
'    rngTblNameCon.ClearContents

    ' Footer
    Set rng = rng.Offset(1, 0)
    mrngProtoFooter.Copy rng
    
    mnLastRow = rng.Row + mrngProtoFooter.Rows.Count - 1
    
    
    Exit Sub
    
Fail:
    End
    
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

    If Not mdCom.Exists(key) Then
        MsgBox "com " & key & " does not exist", vbCritical, msThisClass
        GoTo Fail:
    End If
    Set FractionRange = mdFrac.Item(key)
    Exit Function
        
Fail:
    Set FractionRange = Nothing
    Exit Function
End Function

Public Function FractionRangeDisplay(key As Variant) As Range
    If mdFracDisplay.Exists(key) Then
        Set FractionRangeDisplay = mdFracDisplay.Item(key)
    Else
        MsgBox "No such component " & key, vbCritical, msThisClass
        Set FractionRangeDisplay = Nothing
    End If
End Function

Property Get Component(key As Variant) As clsMatComRange

    If Not mdCom.Exists(key) Then
        MsgBox "Key " & key & " does not exist", vbCritical, msThisClass
        GoTo Fail:
    End If
        
    Set Component = mdCom.Item(key)
    
    Exit Property
Fail:
    End

End Property

Public Function ComList() As ArrayList
    Dim key As Variant

    Set ComList = New ArrayList
    For Each key In mdCom.Keys
        ComList.Add key
    Next
    ComList.Sort
End Function

Private Sub normalize()
    Dim cCom As clsMatCom
    Dim dblSum As Double
    Dim key As Variant
    
'    If Not Me.HasFormulas Then
'        dblSum = 0#
'        For Each key In mdCom.Keys
'            dblSum = dblSum + mdFrac.Item(key)
'        Next
'
'        If Abs(dblSum - 1#) > DIST_TOL Then
'            For Each key In mdCom.Keys
'                mdFrac.Item(key) = mdFrac.Item(key) / dblSum
'            Next
'        End If
'    End If
End Sub

Private Function calculateAbyAtom() As Double
    Dim cCom As clsMatComRange
    Dim dblA As Double
    Dim key As Variant
    
    dblA = 0#
    For Each key In mdCom.Keys
        Set cCom = mdCom.Item(key)
        dblA = dblA + CDbl(mdFrac.Item(key).Value2) * cCom.AValue
    Next
    calculateAbyAtom = dblA
    Exit Function
    
Fail:
    calculateAbyAtom = 0#
    
End Function

Private Function calculateAbyMass() As Double
    Dim cCom As clsMatComRange
    Dim dblA As Double
    Dim key As Variant
    
    dblA = 0#
    For Each key In mdCom.Keys
        Set cCom = mdCom.Item(key)
        dblA = dblA + CDbl(mdFrac.Item(key).Value2) / cCom.AValue
    Next
    calculateAbyMass = 1# / dblA
    Exit Function
    
Fail:
    calculateAbyMass = 0#
    
End Function








