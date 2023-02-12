VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMatMaterial"
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
Dim mdFormula As Scripting.Dictionary
Dim mwksProto As Worksheet

Dim meMode As CompositionMode
Dim msName As String
Dim mblnSealed As Boolean
Dim mblnHasFormulas As Boolean
Dim mdblFrac As Double
Dim mnLastRow As Integer
Dim mnNumRows As Integer


Private Sub Class_Initialize()

    Set mwksProto = Worksheets("Materials")

    meMode = Mass
    Set mrngProtoCom = mwksProto.Names("protoCompositionComMass").RefersToRange
    Set mrngProtoMat = mwksProto.Names("protoCompositionMatMass").RefersToRange
    Set mrngProtoHeader = mwksProto.Names("protoCompositionHeader").RefersToRange
    Set mrngProtoFooter = mwksProto.Names("protoCompositionFooter").RefersToRange
    Set mdCom = New Scripting.Dictionary
    Set mdFrac = New Scripting.Dictionary
    Set mdFormula = New Scripting.Dictionary
    mblnSealed = False
    mblnHasFormulas = True
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
        MsgBox "Constituent must be sealed", vbCritical, "clsMatMaterial"
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

Property Get AValueRange() As Range
    Set AValueRange = mrngAValue
End Property

Property Get IsSealed() As Boolean
    IsSealed = mblnSealed
End Property

Property Get HasFormulas() As Boolean
    HasFormulas = mblnHasFormulas
End Property

Property Let HasFormulas(bln As Boolean)
    If mdCom.Count <> 0 Then
        MsgBox "Attempt to change formulas", vbCritical, "clsMatMaterial"
        End
    End If
    mblnHasFormulas = bln
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
    Dim cCom As clsMatCom
        
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

Property Get Component(key As Variant) As clsMatCom

    If Not mdCom.Exists(key) Then
        MsgBox "Key " & key & " does not exist", vbCritical, "clsMatMaterial"
        GoTo Fail:
    End If
        
    Set Component = mdCom.Item(key)
    
    Exit Property
Fail:
    End

End Property


Property Get Fraction(key As Variant) As Double

    If Not mdFrac.Exists(key) Then
        MsgBox "Key " & key & " does not exist", vbCritical, "clsMatMaterial"
        GoTo Fail:
    End If
        
    Fraction = mdFrac.Item(key)
    
    Exit Property
Fail:
    End

End Property

Property Get Formula(key As Variant) As String

    If Not mdFormula.Exists(key) Then
        MsgBox "Key " & key & " does not exist", vbCritical, "clsMatMaterial"
        GoTo Fail:
    End If
        
    Formula = mdFormula.Item(key)
    
    Exit Property
Fail:
    End

End Property

Public Sub Add(cCom As clsMatCom, frac As Double, sForm As String)
    Dim sCom As String
    Dim sFormSum As String

    If Me.IsSealed Then
        MsgBox "Cannot add component. Material is sealed", vbCritical, "clsMatMaterial"
        GoTo Fail:
    End If
    
    If Me.HasFormulas And sForm = "" Then
        MsgBox "Formula expected", vbCritical, "clsMatMaterial"
        GoTo Fail:
    ElseIf Not Me.HasFormulas And sForm <> "" Then
        MsgBox "Formula unexpected", vbCritical, "clsMatMaterial"
        GoTo Fail:
    End If
    
    sCom = cCom.Name
    If Me.HasFormulas Then
        If mdCom.Exists(sCom) Then
            sFormSum = "=(" & Mid(mdFormula(sCom), 2) & ") + (" & Mid(sForm, 2) & ")"
            mdFormula(sCom) = sFormSum
        Else
            mdCom.Add sCom, cCom
            mdFrac.Add sCom, 0#
            mdFormula.Add sCom, sForm
        End If
    Else
        If mdCom.Exists(sCom) Then
            mdFrac(sCom) = mdFrac(sCom) + frac
        Else
            mdCom.Add sCom, cCom
            mdFrac.Add sCom, frac
            mdFormula.Add sCom, ""
        End If
    End If
    Exit Sub
    
Fail:
    Exit Sub
    
End Sub

Public Sub Seal()
    Dim cCom As clsMatCom
    Dim key As Variant
    
    If Not mblnSealed Then
                
        For Each key In mdCom.Keys
            Set cCom = mdCom.Item(key)
            cCom.Seal
        Next
       
        normalize
        mblnSealed = True
    End If
End Sub

Public Sub InitFromRange(rng As Range)
    Dim cCom As clsMatCom
    Dim sMat As String
    Dim dblFrac As Double
    Dim dblATbl As Double
    Dim dblACalc As Double

    Me.mode = Mass
    Me.HasFormulas = False
    
    sMat = rng.Cells(1, 1).Offset(-4, oc_Comp_Com).Text

    Do Until rng.Cells(1, 1).Offset(0, oc_Comp_MatRows).Value2 = 0
        
        Set cCom = New clsMatCom
        cCom.InitFromRange rng
    
        Select Case meMode
            Case Atom
                dblFrac = rng.Cells(1, 1).Offset(0, oc_Comp_ComAfrac).Value2
            Case Mass
                dblFrac = rng.Cells(1, 1).Offset(0, oc_Comp_ComMfrac).Value2
        End Select
        Me.Add cCom, dblFrac, ""

        Set rng = rng.Offset(1, 0)
    Loop

    Me.Name = sMat
    Me.Seal
    
    dblATbl = rng.Cells(1, 1).Offset(0, oc_Comp_AValue).Value2
    dblACalc = Me.AValue
    
    Debug.Print "A Value: Table " & Format(dblATbl, "0.0000E+00") & " Calculated " & Format(dblACalc, "0.0000E+00")

    If Abs(dblATbl - dblACalc) > DIST_TOL Then
        MsgBox "Calculated A differs from tabulated value", vbCritical, "clsMatMaterial"
        Exit Sub
    End If

End Sub


Public Function Display() As String
    
    Dim cCom As clsMatCom
    Dim dblFrac As Double
    Dim key As Variant
    Dim alCom As ArrayList
    
    Me.Seal
    
    Set alCom = ComList()
    
    Display = ""
    For Each key In alCom
        Set cCom = mdCom.Item(key)
        dblFrac = mdFrac.Item(key)
        
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
    Dim cCom As clsMatCom
    Dim dblFrac As Double
    Dim key As Variant
    Dim alCom As ArrayList
    Dim rngNumRows As Range
    Dim rngTblNameIso As Range
    Dim rngTblNameCon As Range
    Dim rngTblDataBgn As Range
    Dim sForm As String
    
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
        dblFrac = mdFrac.Item(key)
        sForm = mdFormula.Item(key)
        
        cCom.DisplayExcel rng
        
        mrngProtoCom.Copy rng
        ' Fix A value calc if modes differ
        If Me.mode = Mass And cCom.mode = Atom Then
            mwksProto.Names("protoCompositionComAtom").RefersToRange.Cells(1, 1).Offset(0, oc_Comp_AValue).Copy rng.Offset(0, oc_Comp_AValue)
        ElseIf Me.mode = Atom And cCom.mode = Mass Then
            mwksProto.Names("protoCompositionComMass").RefersToRange.Cells(1, 1).Offset(0, oc_Comp_AValue).Copy rng.Offset(0, oc_Comp_AValue)
        End If
        
        rng.Cells(1, 1).Offset(0, oc_Comp_Com).Value2 = cCom.Name
        
        ' Fraction value may be a formula
        If Me.HasFormulas Then
            Select Case meMode
                Case Atom
                    rng.Cells(1, 1).Offset(0, oc_Comp_ComAfrac).Formula = sForm
                Case Mass
                    rng.Cells(1, 1).Offset(0, oc_Comp_ComMfrac).Formula = sForm
            End Select
        Else
            Select Case meMode
                Case Atom
                    rng.Cells(1, 1).Offset(0, oc_Comp_ComAfrac).Value2 = dblFrac
                Case Mass
                    rng.Cells(1, 1).Offset(0, oc_Comp_ComMfrac).Value2 = dblFrac
            End Select
        End If
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
    rngTblNameIso.ClearContents
    rngTblNameCon.ClearContents

    ' Footer
    Set rng = rng.Offset(1, 0)
    mrngProtoFooter.Copy rng
    
    mnLastRow = rng.Row + mrngProtoFooter.Rows.Count - 1
    
    
    Exit Sub
    
Fail:
    End
    
End Sub


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
    
    If Not Me.HasFormulas Then
        dblSum = 0#
        For Each key In mdCom.Keys
            dblSum = dblSum + mdFrac.Item(key)
        Next
        
        If Abs(dblSum - 1#) > DIST_TOL Then
            For Each key In mdCom.Keys
                mdFrac.Item(key) = mdFrac.Item(key) / dblSum
            Next
        End If
    End If
End Sub


Private Function calculateAbyAtom() As Double
    Dim cCom As clsMatCom
    Dim dblA As Double
    Dim key As Variant
    
    dblA = 0#
    If Not Me.HasFormulas Then
        For Each key In mdCom.Keys
            Set cCom = mdCom.Item(key)
            dblA = dblA + mdFrac.Item(key) * cCom.AValue
        Next
    End If
    calculateAbyAtom = dblA
    Exit Function
    
Fail:
    calculateAbyAtom = 0#
    
End Function

Private Function calculateAbyMass() As Double
    Dim cCom As clsMatCom
    Dim dblA As Double
    Dim key As Variant
    
    dblA = 0#
    If Not Me.HasFormulas Then
        For Each key In mdCom.Keys
            Set cCom = mdCom.Item(key)
            dblA = dblA + mdFrac.Item(key) / cCom.AValue
        Next
    End If
    calculateAbyMass = 1# / dblA
    Exit Function
    
Fail:
    calculateAbyMass = 0#
    
End Function






