VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMatByElmStanza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' Copyright (c) 2022 Henkel Technical Computing, LLC. All rights reserved.
'
Option Explicit
Option Compare Text

Dim mcMatMaterial As clsMatMaterialRange
Dim mcMatMaster As clsMatMaster
Dim mcMatElmIso As clsMatElmIso
Dim mcMatTblParam As clsMatTblParam
Dim mcMatTblFracCon As clsMatTblFracCon
Dim mcMatTblIso As clsMatTblIso
Dim mcMatTblComp As clsMatTblComp
Dim mcMatTblSpec As clsMatTblSpec
Dim mcMatTblMCNP As clsMatTblMCNP
Dim mcMatTblSCALE As clsMatTblSCALE
Dim mwksMat As Worksheet

Dim mrngMatByElm As Range
Dim mrngMatByElmOverride As Range
Dim mrngMatKey As Range
Dim mrngMatBase As Range
Dim msMat As String
Dim msDescAux As String
Dim msSource As String
Dim msSab As String
Dim mdblDensity As Double
Dim msThisClass As String
Dim meMode As CompositionMode

Private Sub Class_Initialize()
    Dim iColKey As Integer

    msThisClass = "clsMatByElmStanza"
    Set mwksMat = Worksheets("Materials")

    Set mcMatMaster = New clsMatMaster
    Set mcMatElmIso = New clsMatElmIso

    iColKey = mwksMat.Names("icol_Key").RefersToRange.Value2
    Set mrngMatKey = mwksMat.Columns(iColKey)

    Set mrngMatByElm = mwksMat.Names("tblMatByElm").RefersToRange
    Set mrngMatByElmOverride = mwksMat.Names("tblMatByElmOverride").RefersToRange

End Sub

Private Function nextRow() As Integer
    Dim iRow As Integer
    Dim iStart As Integer
'    NextRow = Application.WorksheetFunction.Match("Next", mrngMatKey, 0)
    
    iStart = mcMatMaster.MaxRow
    
    nextRow = 0
    For iRow = iStart To iStart + 1000
        If mrngMatKey.Cells(iRow, 1).Text = "Next" Then
            nextRow = iRow
            Exit For
        End If
    Next iRow
    
End Function

Property Get MatName() As String
    MatName = msMat
End Property

Property Let MatName(s As String)
    msMat = s
End Property

Property Get DescriptiveText() As String
    DescriptiveText = msDescAux
End Property

Property Let DescriptiveText(s As String)
    msDescAux = s
End Property

Property Get Sab() As String
    Sab = msSab
End Property

Property Let Sab(s As String)
    msSab = s
End Property

Property Get Source() As String
    Source = msSource
End Property

Property Let Source(s As String)
    msSource = s
End Property

Property Get Density() As Double
    Density = mdblDensity
End Property

Property Let Density(dbl As Double)
    mdblDensity = dbl
End Property

Property Get mode() As CompositionMode
    mode = meMode
End Property

Property Let mode(e As CompositionMode)
    meMode = e
End Property

Public Sub AddStanza()
    Dim iNextRow As Integer
    Dim iCon As Integer
    Dim sCon As String
    Dim alCon As ArrayList
    Dim rngProtoMatHeader As Range
    Dim rngFrac As Range
    Dim rngUnity As Range
    Dim key As Variant
    Dim xlCalc As XlCalculation
    Dim blnUpdate As Boolean
    Dim cCom As clsMatComRange
    Dim cCon As clsMatConRange
    Dim cMat As clsMatMaterialRange
    
    xlCalc = Application.Calculation
    Application.Calculation = xlCalculationManual
    blnUpdate = Application.ScreenUpdating
    Application.ScreenUpdating = False
    Application.StatusBar = "Adding " & Me.MatName

    mwksMat.Activate

    If mcMatMaster.StanzaExists(msMat) Then
        MsgBox "Material " & msMat & " already exists", vbOKOnly, msThisClass
        Exit Sub
    End If

    iNextRow = nextRow()
    If iNextRow < mcMatMaster.MaxRow Then
        MsgBox "There appears to be data after Next", vbOKOnly, msThisClass
        Exit Sub
    End If
    
    Set mrngMatBase = ActiveSheet.Cells(iNextRow, mrngMatKey.Column)
    mrngMatBase.Activate
    
    ' Move Next out of the way
    mrngMatBase.Offset(0, -1).Resize(1, 2).Copy mrngMatBase.Offset(-1, -1)
    mrngMatBase.Offset(0, -1).Resize(1, 2).Clear
    
    ' Stanza header
    Set rngProtoMatHeader = ActiveSheet.Names("protoMatHeader").RefersToRange
    rngProtoMatHeader.Copy mrngMatBase
    mrngMatBase.Cells(1, 1).Value2 = Me.MatName
    
    ' Parameters
    Set mcMatTblParam = New clsMatTblParam
    With mcMatTblParam
        Set .Base = mrngMatBase
        .Density = Me.Density
        .DescriptiveText = Me.DescriptiveText
        .Source = Me.Source
        .DisplayExcel
    End With
    
    ' FracCon
    Set mcMatTblFracCon = New clsMatTblFracCon
    With mcMatTblFracCon
        Set .Base = mrngMatBase
        .mode = meMode
        .CopyMode = ByFormula
        
        For iCon = 1 To mrngMatByElm.Rows.Count
            sCon = mrngMatByElm.Cells(iCon, 1).Text
            If sCon <> "" Then
                Set rngFrac = mrngMatByElm.Cells(iCon, 2)
                .AddConstituent sCon, rngFrac
            End If
        Next iCon
        .DisplayExcel
    End With
    
    ' Iso table
    Set mcMatTblIso = New clsMatTblIso
    With mcMatTblIso
        Set .Base = mrngMatBase
        Set .ElmIso = mcMatElmIso
        .LoadElements mcMatTblFracCon.ConList
        .Override mrngMatByElmOverride
        .DisplayExcel
    End With
    
    ' Generate material referencing table ranges
    Set cCom = New clsMatComRange
    With cCom
        .Name = Me.MatName
        .mode = mcMatTblFracCon.mode
        .CopyMode = ByReference
        Set alCon = mcMatTblFracCon.ConList
        For Each key In alCon
            Set cCon = mcMatTblIso.TableConstituent(key)
            Set rngFrac = mcMatTblFracCon.FractionRange(key)
            .Add cCon, rngFrac
        Next key
    End With
    
    Set rngUnity = mwksMat.Names("matFracUnity").RefersToRange
    Set cMat = New clsMatMaterialRange
    With cMat
        .Name = Me.MatName
        .mode = Mass
        .CopyMode = ByValue
        .Add cCom, rngUnity
    End With
    
    ' Composition table
    Set mcMatTblComp = New clsMatTblComp
    With mcMatTblComp
        Set .Base = mrngMatBase
        Set .Material = cMat
        .DisplayExcel
    End With
    
    ' Specification table
    Set mcMatTblSpec = New clsMatTblSpec
    With mcMatTblSpec
        Set .Base = mrngMatBase
        Set .DensityRange = mcMatTblParam.DensityRange
        Set .MatTblComp = mcMatTblComp
        .DisplayExcel
    End With

    ' MCNP input
    Set mcMatTblMCNP = New clsMatTblMCNP
    With mcMatTblMCNP
        Set .Base = mrngMatBase
        Set .MatTblComp = mcMatTblComp
        Set .MatTblSpec = mcMatTblSpec
        .Sab = Me.Sab
        .DisplayExcel
    End With
'
'    ' SCALE input
'    Set mcMatTblSCALE = New clsMatStoichTblSCALE
'    With mcMatTblSCALE
'        Set .Base = mrngMatBase
'        Set .MatTblComp = mcMatTblComp
'        Set .MatTblSpec = mcMatTblSpec
'        .DisplayExcel
'    End With
'
    
    mcMatMaster.Add Me.MatName
    
    ' Move next
    mrngMatBase.Offset(-1, -1).Resize(1, 2).Copy mrngMatKey.Cells(mcMatTblComp.LastRow + 3, 1).Offset(0, -1)
    mrngMatBase.Offset(-1, -1).Resize(1, 2).Clear
    
    ' Update
    Application.Calculate
    
    Application.Calculation = xlCalc
    Application.ScreenUpdating = blnUpdate
    Application.StatusBar = False
    
End Sub


