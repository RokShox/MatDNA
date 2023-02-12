VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMatMixMatByFracStanza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' Copyright (c) 2022 Henkel Technical Computing, LLC. All rights reserved.
'
Option Explicit
Option Compare Text

Dim mcMatMaster As clsMatMaster
Dim mcMatTblParam As clsMatMixMatByFracParam
Dim mcMatTblComp As clsMatTblComp
Dim mcMatTblSpec As clsMatTblSpec
Dim mcMatTblMCNP As clsMatTblMCNP
Dim mcMatTblSCALE As clsMatTblSCALE
Dim mwksMat As Worksheet

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

    msThisClass = "clsMatMixMatByFracStanza"
    Set mwksMat = Worksheets("Materials")
    
    msDescAux = ""
    msSab = ""
    msSource = ""

    Set mcMatMaster = New clsMatMaster

    iColKey = mwksMat.Names("icol_Key").RefersToRange.Value2
    Set mrngMatKey = mwksMat.Columns(iColKey)

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
    Dim keyMat As Variant
    Dim keyCom As Variant
    Dim xlCalc As XlCalculation
    Dim blnUpdate As Boolean
    Dim cCom As clsMatComRange
    Dim cMat As clsMatMaterialRange
    Dim cMatMix As clsMatMaterialRange
    Dim alMat As ArrayList
    Dim dComFracRange As Scripting.Dictionary
    Dim dMatFracRange As Scripting.Dictionary
    
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
    ' Creates a dictionary of materials representing each material in the parameter table
    ' initialized using "InitFromRange" method
    Set mcMatTblParam = New clsMatMixMatByFracParam
    With mcMatTblParam
        Set .Base = mrngMatBase
        .DescriptiveText = Me.DescriptiveText
        .mode = Me.mode
        .ModelDensity = Me.Density
        .Source = Me.Source
        .DisplayExcel
    End With
    
'    Debug.Print "MFrac range of H2O is " & mcMatTblParam.MassFracRange("H2O").Address(False, False)
'    Debug.Print "AFrac range of H2O is " & mcMatTblParam.AtomFracRange("H2O").Address(False, False)
'    Debug.Print "Density range is " & mcMatTblParam.DensityRange.Address(False, False)
'    Debug.Print mcMatTblParam.Material("H2O").Display
    
    
    ' Generate new material referencing components from the mixture list
    ' Note if a component is added with the same name as an existing component, it's name is amended with "-02", "-03", etc
    ' Hence, these duplicated components cannot be easily found in the original material once added to the mixture material.
    '
    ' So there are two types of materials. One type are materials defined from the materials being mixed via InitFromRange
    ' The other - the mixture material - is a single material that collects all the components in the mixed materials into
    ' a sinngle material with changes to the names of duplicated components.
    
    Set dComFracRange = New Scripting.Dictionary
    Set dMatFracRange = New Scripting.Dictionary
    Set cMatMix = New clsMatMaterialRange
    With cMatMix
        .Name = Me.MatName
        .mode = Me.mode
        .CopyMode = ByReference
    End With
    
    Set alMat = mcMatTblParam.MatList
    For Each keyMat In alMat
        Set cMat = mcMatTblParam.Material(keyMat)
        For Each keyCom In cMat.ComList
            Set cCom = cMat.Component(keyCom)
            ' Add the material to the mixture material using the param table fraction as a placeholder only
            ' The actual component fraction needs to be this value times the fraction of the component in the original material
            ' This is fixed after the Composition table is displayed so that these ranges can be captured in FractionRangeDisplay property
            ' of the mixture material
            cMatMix.Add cCom, mcMatTblParam.FractionRange(keyMat)
            
            ' Save original FractionRange (this is actually the cell where the component fraction
            ' is displayed in the original material)
            ' Note that cCom may have its name changed if it's a duplicate
            ' Also save the mixture material fraction for the material that this component belongs to.
            ' The actual component fraction in the mixture is the product of these two
            dComFracRange.Add cCom.Name, cMat.FractionRange(keyCom)
            dMatFracRange.Add cCom.Name, mcMatTblParam.FractionRange(keyMat)
        Next keyCom
    Next keyMat
    
    ' Composition table
    Set mcMatTblComp = New clsMatTblComp
    With mcMatTblComp
        Set .Base = mrngMatBase
        Set .Material = cMatMix
        .DisplayExcel
    End With
    
    ' Fix up composition fractions
    For Each keyCom In cMatMix.ComList
        ' cMatMix.FractionRangeDisplay(keyCom) is the cell in the mixture composition table where this component fraction is specified
        ' Its location is only known after the composition table has been displayed.
        ' The forumla in this cell is the component fraction in the original material times the material fraction in the mixture
        cMatMix.FractionRangeDisplay(keyCom).Formula = "=" & dComFracRange.Item(keyCom).Address(False, False) _
        & "*" & dMatFracRange.Item(keyCom).Address(False, False)
    Next keyCom

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



