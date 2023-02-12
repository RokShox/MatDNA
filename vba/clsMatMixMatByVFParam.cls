VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMatMixMatByVFParam"
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

' Location of table rel to mat base
Const oc_Proto As Integer = 1
Const or_Proto As Integer = 3

' Row offsets rel to table
Const or_Desc As Integer = 0
Const or_Data As Integer = 3

' Column offsets rel to table
Const oc_Mat As Integer = 0
Const oc_DenTrue As Integer = 1
Const oc_VFrac As Integer = 2
Const oc_DenModel As Integer = 3
Const oc_MFrac As Integer = 4

Dim msThisClass As String
Const q As String = """"

Private Sub Class_Initialize()
    msThisClass = "clsMatMixMatByVFParam"
    Set mcMatMaster = New clsMatMaster
    Set mwksProto = Worksheets("Materials")
    Set mrngProto = mwksProto.Names("protoMixByVFrac").RefersToRange
    mnMat = mrngProto.Rows.Count - or_Data - 1
    meMode = Mass
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

Property Get Density() As Double
    Density = CDbl(Me.DensityRange.Value2)
End Property

Property Get DensityRange() As Range
    Set DensityRange = mrngDataBase.Offset(mnMat, oc_DenModel)
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

Property Get FractionRange(keyMat As Variant) As Range
    Set FractionRange = Me.MassFracRange(keyMat)
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
    
    
    Application.Calculation = xlCalc
    Application.ScreenUpdating = blnUpdate
    Application.StatusBar = False

End Sub






