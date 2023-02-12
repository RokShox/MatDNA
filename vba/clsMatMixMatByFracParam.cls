VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMatMixMatByFracParam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' Copyright (c) 2022 Henkel Technical Computing, LLC. All rights reserved.
'
Option Explicit
Option Compare Text

Dim mwksProto As Worksheet
Dim mrngMatBase As Range
Dim mrngTblBase As Range
Dim mrngProto As Range
Dim mrngDataBase As Range
Dim mdMat As Scripting.Dictionary
Dim mcMatMaster As clsMatMaster

Dim meMode As CompositionMode
Dim mdblDenModel As Double
Dim mnMat As Integer
Dim msDescAux As String
Dim msSource As String

' Location of table rel to mat base
Const oc_Proto As Integer = 1
Const or_Proto As Integer = 3

' Row offsets rel to table
Const or_Desc As Integer = 0
Const or_Data As Integer = 3

' Column offsets rel to table
Const oc_Mat As Integer = 0
Const oc_DenTrue As Integer = 1
Const oc_MFrac As Integer = 2
Const oc_AFrac As Integer = 3
Const oc_DenModel As Integer = 4
Const oc_OoDenTrue As Integer = 5
Const oc_Ai As Integer = 6
Const oc_OoAi As Integer = 7

Dim msThisClass As String
Const q As String = """"

Private Sub Class_Initialize()
    msThisClass = "clsMatMixMatByFracParam"
    Set mcMatMaster = New clsMatMaster
    Set mwksProto = Worksheets("Materials")
    Set mrngProto = mwksProto.Names("tblMixMatByFrac").RefersToRange
    mnMat = mrngProto.Rows.Count - or_Data - 1
    
    msDescAux = ""
    msSource = ""
End Sub

Property Get Base() As Range
    Set Base = mrngMatBase
End Property

Property Set Base(rngBase As Range)
    Set mrngMatBase = rngBase
    Set mrngTblBase = mrngMatBase.Offset(or_Proto, oc_Proto)
    Set mrngDataBase = mrngTblBase.Offset(or_Data, 0)
    Set mdMat = New Scripting.Dictionary
End Property

Property Get mode() As CompositionMode
    mode = meMode
End Property

Property Let mode(e As CompositionMode)
    meMode = e
End Property

Property Get Density() As Double
    Density = CDbl(Me.DensityRange.Value2)
End Property

Property Get DensityRange() As Range
    Set DensityRange = mrngDataBase.Offset(mnMat, oc_DenModel)
End Property

Property Let ModelDensity(dbl As Double)
    mdblDenModel = dbl
End Property

Property Get DescriptiveText() As String
    DescriptiveText = msDescAux
End Property

Property Let DescriptiveText(s As String)
    msDescAux = s
End Property

Property Get Source() As String
    Source = msSource
End Property

Property Let Source(s As String)
    msSource = s
End Property

Property Get Description() As String
    Description = mrngTblBase.Offset(or_Desc, 0).Text
End Property

Property Get MatList() As ArrayList
    Dim key As Variant
    
    Set MatList = New ArrayList
    For Each key In mdMat.Keys
        MatList.Add key
    Next
    MatList.Sort
    
End Property

Property Get Material(key As Variant) As clsMatMaterialRange
    If mdMat.Exists(key) Then
        Set Material = mdMat.Item(key)
    Else
        MsgBox "No such material " & key, vbCritical, msThisClass
        Set Material = Nothing
    End If
End Property

Property Get MassFrac(keyMat As Variant) As Double
    MassFrac = CDbl(Me.MassFracRange(keyMat).Value2)
End Property

Property Get MassFracRange(keyMat As Variant) As Range
    Dim rng As Range
    
    Set rng = mrngDataBase.Resize(mnMat, 1).Find(CStr(keyMat), LookIn:=xlValues, Lookat:=xlWhole)
    Set MassFracRange = rng.Offset(0, oc_MFrac)
End Property

Property Get AtomFrac(keyMat As Variant) As Double
    AtomFrac = CDbl(Me.AtomFracRange(keyMat).Value2)
End Property

Property Get AtomFracRange(keyMat As Variant) As Range
    Dim rng As Range
    
    Set rng = mrngDataBase.Resize(mnMat, 1).Find(CStr(keyMat), LookIn:=xlValues, Lookat:=xlWhole)
    Set AtomFracRange = rng.Offset(0, oc_AFrac)
End Property

Property Get FractionRange(keyMat As Variant) As Range
            If Me.mode = Atom Then
                Set FractionRange = Me.AtomFracRange(keyMat)
            ElseIf Me.mode = Mass Then
                Set FractionRange = Me.MassFracRange(keyMat)
            End If
End Property


Private Sub getMaterials()
    Dim rngStanza As Range
    Dim rngData As Range
    Dim oMat As Integer
    Dim sMat As String
    Dim cMat As clsMatMaterialRange
    
    For oMat = 0 To mnMat - 1
        sMat = mrngDataBase.Offset(oMat, oc_Mat).Text
        
        If Not mcMatMaster.StanzaExists(sMat) Then
            MsgBox "Material " & sMat & " does not exist", vbCritical, msThisClass
            GoTo Fail:
        End If
        
        If mdMat.Exists(sMat) Then
            MsgBox "Material " & sMat & " already exists", vbCritical, msThisClass
            GoTo Fail:
        End If
                
        Set rngStanza = mcMatMaster.Stanza(sMat)
        Set rngData = rngStanza.Cells(1, 1).Offset(or_Stanza_Comp + or_Comp_Data, oc_Stanza_Comp)
        
        ' Create a material
        Set cMat = New clsMatMaterialRange
        With cMat
            .mode = meMode
            .CopyMode = ByReference
            .InitFromRange rngData
        End With
        
        mdMat.Add sMat, cMat
        
    Next oMat
    
    
    Exit Sub
    
Fail:
    Exit Sub
    
    
End Sub

Public Sub DisplayExcel()
    Dim xlCalc As XlCalculation
    Dim blnUpdate As Boolean
    Dim alMat As ArrayList
    Dim key As Variant
    Dim cMat As clsMatMaterialRange
    Dim oMat As Integer
    Dim sMat As String
    Dim rngMat As Range
    Dim rngDenTrue As Range
    Dim rngMFrac As Range
    Dim rngAFrac As Range
    Dim rngDenModel As Range
    Dim rngOoDenTrue As Range
    Dim rngAi As Range
    Dim rngOoAi As Range
    Dim rngA As Range
    Dim rngOoA As Range

    xlCalc = Application.Calculation
    Application.Calculation = xlCalculationManual
    blnUpdate = Application.ScreenUpdating
    Application.ScreenUpdating = False
    Application.StatusBar = mrngMatBase.Text & ": " & "Parameters"
    
    
    ' Clear data
    mrngTblBase.Resize(mrngProto.Rows.Count, mrngProto.Columns.Count).Clear
    
    ' Copy table
    mrngProto.Copy mrngTblBase
    
    ' get materials
    getMaterials
    
    Set alMat = Me.MatList
    Set rngMat = mrngDataBase.Offset(0, oc_Mat).Resize(mnMat, 1)
    Set rngDenTrue = mrngDataBase.Offset(0, oc_DenTrue).Resize(mnMat, 1)
    Set rngMFrac = mrngDataBase.Offset(0, oc_MFrac).Resize(mnMat, 1)
    Set rngAFrac = mrngDataBase.Offset(0, oc_AFrac).Resize(mnMat, 1)
    Set rngDenModel = mrngDataBase.Offset(0, oc_DenModel).Resize(mnMat, 1)
    Set rngOoDenTrue = mrngDataBase.Offset(0, oc_OoDenTrue).Resize(mnMat, 1)
    Set rngAi = mrngDataBase.Offset(0, oc_Ai).Resize(mnMat, 1)
    Set rngOoAi = mrngDataBase.Offset(0, oc_OoAi).Resize(mnMat, 1)
    Set rngA = mrngDataBase.Offset(mnMat, oc_Ai)
    Set rngOoA = mrngDataBase.Offset(mnMat, oc_OoAi)
    
    
    
    ' Debug
'    For Each key In alMat
'        Set cMat = mdMat.Item(key)
'        Debug.Print cMat.Display
'    Next key

'   Use formula in proto
'    mrngTblBase.Offset(or_Desc, 0).Formula = "=" & q & Me.DescriptiveText & q & " & " & q & " - " & q & " & " & Mid(mrngTblBase.Offset(or_Desc, 0).Formula, 2)
    
    For oMat = 0 To mnMat - 1
        sMat = mrngDataBase.Offset(oMat, oc_Mat).Text
        Set cMat = mdMat.Item(sMat)
        
        ' True density
        mrngDataBase.Offset(oMat, oc_DenTrue).Formula = "=rho" & sMat
        
        ' 1/True density
        mrngDataBase.Offset(oMat, oc_OoDenTrue).Formula = "=1/" & mrngDataBase.Offset(oMat, oc_DenTrue).Address(False, False)
        
        ' A_i
        mrngDataBase.Offset(oMat, oc_Ai).Formula = "=" & cMat.AValueRange.Address(False, False)
    
        ' 1/A_i
        mrngDataBase.Offset(oMat, oc_OoAi).Formula = "=1/" & mrngDataBase.Offset(oMat, oc_Ai).Address(False, False)

    Next oMat
        
    ' Calculate A
    If meMode = Atom Then
        rngA.Formula = "=sumproduct(" & rngAFrac.Address(False, False) & "," & rngAi.Address(False, False) & ")"
    ElseIf meMode = Mass Then
        rngA.Formula = "=1/sumproduct(" & rngMFrac.Address(False, False) & "," & rngOoAi.Address(False, False) & ")"
    End If
    rngOoA.Formula = "=1/" & rngA.Address(False, False)
    
    ' Calculate unspecified fraction
    If meMode = Atom Then
        For oMat = 0 To mnMat - 1
            rngMFrac.Cells(oMat + 1, 1).ClearContents
            rngMFrac.Cells(oMat + 1, 1).Formula = "=" & rngAFrac.Cells(oMat + 1, 1).Address(False, False) & "*" _
            & rngAi.Cells(oMat + 1, 1).Address(False, False) & "*" _
            & rngOoA.Address(False, False)
        Next oMat
    ElseIf meMode = Mass Then
        For oMat = 0 To mnMat - 1
            rngAFrac.Cells(oMat + 1, 1).ClearContents
            rngAFrac.Cells(oMat + 1, 1).Formula = "=" & rngMFrac.Cells(oMat + 1, 1).Address(False, False) & "*" _
            & rngA.Address(False, False) & "*" _
            & rngOoAi.Cells(oMat + 1, 1).Address(False, False)
        Next oMat
    End If
    
    ' True density
    mrngDataBase.Offset(mnMat, oc_DenTrue).Formula = "=1/sumproduct(" & rngMFrac.Address(False, False) & "," & rngOoDenTrue.Address(False, False) & ")"
    
    ' Total mass frac
    mrngDataBase.Offset(mnMat, oc_MFrac).Formula = "=sum(" & rngMFrac.Address(False, False) & ")"
    
    ' Total atom frac
    mrngDataBase.Offset(mnMat, oc_AFrac).Formula = "=sum(" & rngAFrac.Address(False, False) & ")"
    
    ' Model density
    If mdblDenModel = 0# Then
        mrngDataBase.Offset(mnMat, oc_DenModel).Formula = "=" & mrngDataBase.Offset(mnMat, oc_DenTrue).Address(False, False)
    Else
        mrngDataBase.Offset(mnMat, oc_DenModel).Value2 = mdblDenModel
    End If
    For oMat = 0 To mnMat - 1
        rngDenModel.Cells(oMat + 1, 1).Formula = "=" & rngMFrac.Cells(oMat + 1, 1).Address(False, False) & "*" & mrngDataBase.Offset(mnMat, oc_DenModel).Address(False, False)
    Next oMat

    ' 1/True density
    mrngDataBase.Offset(mnMat, oc_OoDenTrue).Formula = "=1/" & mrngDataBase.Offset(mnMat, oc_DenTrue).Address(False, False)

    ' Source
    If msSource <> "" Then
        mrngDataBase.Offset(mnMat + 1, oc_Mat).Value2 = "Source: " & msSource
    End If

    With rngMat.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With

    If Me.mode = Atom Then
        With rngAFrac.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent3
            .TintAndShade = 0.8
            .PatternTintAndShade = 0
        End With
        With rngMFrac.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
    Else
        With rngMFrac.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent3
            .TintAndShade = 0.8
            .PatternTintAndShade = 0
        End With
        With rngAFrac.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
    End If



    Application.Calculation = xlCalc
    Application.ScreenUpdating = blnUpdate
    Application.StatusBar = False

End Sub




