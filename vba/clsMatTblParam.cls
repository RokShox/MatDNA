VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMatTblParam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' Copyright (c) 2022 Henkel Technical Computing, LLC. All rights reserved.
'
Option Explicit
Option Compare Text

Dim mrngMatBase As Range
Dim mrngTblBase As Range
Dim mrngProto As Range
Dim mwksProto As Worksheet

Dim mocProto As Integer
Dim morProto As Integer
Dim morDesc As Integer
Dim morDensity As Integer
Dim morSource As Integer
Dim msSource As String
Dim mdblDensity As Double
Dim msDescAux As String
Dim msThisClass As String
Const q As String = """"

Private Sub Class_Initialize()

    msThisClass = "clsMatTblParam"
    Set mwksProto = Worksheets("Materials")
    Set mrngProto = mwksProto.Names("protoParam").RefersToRange
    
    ' Relative to material stanza base cell
    mocProto = 1
    morProto = 3
    ' Relative to table
    morDesc = 0
    morDensity = 2
    morSource = 3
    
    msSource = ""
    msDescAux = ""
    mdblDensity = 0#
End Sub

Property Get Base() As Range
    Set Base = mrngMatBase
End Property

Property Set Base(rngBase As Range)
    Set mrngMatBase = rngBase
    Set mrngTblBase = mrngMatBase.Offset(morProto, mocProto)
End Property

Property Get Density() As Double
    Density = mdblDensity
End Property

Property Let Density(dblDen As Double)
    mdblDensity = dblDen
End Property

Property Get DensityRange() As Range
    Set DensityRange = mrngTblBase.Offset(morDensity, 1)
End Property

Property Get Description() As String
    Description = mrngTblBase.Offset(morDesc, 0).Text
End Property

Property Get DescriptiveText() As String
    DescriptiveText = msDescAux
End Property

Property Let DescriptiveText(sDesc As String)
    msDescAux = sDesc
End Property

Property Get Source() As String
    Source = msSource
End Property

Property Let Source(sSource As String)
    msSource = sSource
End Property

Public Sub DisplayExcel()
    Dim sPre As String
    Dim sPos As String
    Dim sForm As String
    Dim xlCalc As XlCalculation
    Dim blnUpdate As Boolean

    xlCalc = Application.Calculation
    Application.Calculation = xlCalculationManual
    blnUpdate = Application.ScreenUpdating
    Application.ScreenUpdating = False
    
    Application.StatusBar = mrngMatBase.Text & ": " & "Parameters"
    
    ' Clear data
    mrngTblBase.Resize(mrngProto.Rows.Count, mrngProto.Columns.Count).Clear
    
    mrngProto.Copy mrngTblBase
    
    ' Set density
    mrngTblBase.Offset(morDensity, 1).Value2 = mdblDensity

    ' Set source
    mrngTblBase.Offset(morSource, 0).Value2 = "Source: " & msSource
    
    ' Set description
    sPre = mrngMatBase.Text & " - " & msDescAux & " - rho("
    sPos = ")"
    sForm = "=" & q & sPre & q & "& TEXT(" _
        & mrngTblBase.Offset(morDensity, 1).Address(False, False) _
        & "," & q & "0.0000" & q & ") &" & q & sPos & q
    
    mrngTblBase.Offset(morDesc, 0).Formula = sForm

    Application.Calculation = xlCalc
    Application.ScreenUpdating = blnUpdate
    Application.StatusBar = False


End Sub


