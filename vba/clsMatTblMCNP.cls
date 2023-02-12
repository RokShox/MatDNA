VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMatTblMCNP"
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
Dim mcMatTblComp As clsMatTblComp
Dim mcMatTblSpec As clsMatTblSpec
Dim mrngMatBase As Range
Dim mrngTblBase As Range
Dim mrngProto As Range
Dim mocProto As Integer
Dim morProto As Integer

Dim morHead As Integer
Dim morIsoRow As Integer
Dim mocNU As Integer
Dim mocZAID As Integer
Dim mocAFrac As Integer
Dim mocComment As Integer
Dim morNlibRow As Integer
Dim msSab As String
Dim miLastRow As Integer
Const q As String = """"

' Need to set Base, MatTblSpec
Private Sub Class_Initialize()
        
    Set mwksProto = Worksheets("Materials")
        
    Set mrngProto = mwksProto.Names("protoMCNP").RefersToRange
        
    msSab = ""
    miLastRow = 0
        
    ' Relative to material stanza base cell
    mocProto = 55
    morProto = 2
    ' Relative to table
    morHead = 4
    morIsoRow = 7
    morNlibRow = 8
    mocNU = 0
    mocZAID = 1
    mocAFrac = 2
    mocComment = 3
    
End Sub


Property Get Base() As Range
    Set Base = mrngMatBase
End Property

Property Set Base(rngBase As Range)
    Set mrngMatBase = rngBase
    Set mrngTblBase = mrngMatBase.Offset(morProto, mocProto)
End Property

Property Set MatTblComp(c As clsMatTblComp)
    Set mcMatTblComp = c
End Property

Property Set MatTblSpec(c As clsMatTblSpec)
    Set mcMatTblSpec = c
End Property

Property Let Sab(s As String)
    msSab = s
End Property

Property Get Sab() As String
    Sab = msSab
End Property

Property Get LastRow() As Integer
    LastRow = miLastRow
End Property

Public Sub DisplayExcel()
    Dim rngProtoIso As Range
    Dim rngProtoNlib As Range
    Dim rngCurrent As Range
    Dim nIso As Integer
    Dim oIso As Integer
    Dim nCol As Integer
    Dim sName As String
    Dim sRefersTo As String
    Dim nm As Name
    Dim nAux As Integer
    Dim xlCalc As XlCalculation
    Dim blnUpdate As Boolean
    
    nIso = mcMatTblSpec.NumIso
    nCol = mrngProto.Columns.Count
    
    xlCalc = Application.Calculation
    Application.Calculation = xlCalculationManual
    blnUpdate = Application.ScreenUpdating
    Application.ScreenUpdating = False
    Application.StatusBar = mrngMatBase.Text & ": " & "MCNP"
   
    ' Clear old data
    mrngTblBase.Resize(mcMatTblComp.LastRow - mrngTblBase.Row + 1, nCol).Clear
    
    mrngProto.Copy mrngTblBase

    Set rngProtoIso = mrngTblBase.Offset(morIsoRow, 0).Resize(1, nCol)
    Set rngProtoNlib = mrngTblBase.Offset(morNlibRow, 0).Resize(1, nCol)
    
    For oIso = 0 To nIso - 1
        If oIso > 0 Then
            rngProtoNlib.Insert
            Set rngCurrent = rngProtoNlib.Cells(1, 1).Offset(-1, 0)
            rngProtoIso.Copy rngCurrent
        Else
            Set rngCurrent = rngProtoIso.Cells(1, 1)
        End If
    Next oIso

    ' Fix S(a,b)
    If msSab <> "" Then
        ' ="mt"&@INDIRECT("mat"&S3547)
        rngProtoNlib.Cells(1, 1).Offset(1, 0).Formula = "=" & q & "mt" & q & _
            "&INDIRECT(" & q & "mat" & q & "&" & mrngMatBase.Address(False, False) & ")"
'        rngProtoNlib.Cells(1, 1).Offset(1, 1).Value2 = msSab
        rngProtoNlib.Cells(1, 1).Offset(1, 1).Formula = msSab
    Else
        rngProtoNlib.Cells(1, 1).Offset(1, 0).Value2 = "n/u"
    End If
    
    
    ' Fix check
    ' Formula =SUM(BX3556:BX3556)
    rngProtoNlib.Cells(1, 1).Offset(3, 2).Formula = "=SUM(" & _
        rngProtoNlib.Cells(1, 1).Offset(-nIso, mocAFrac).Resize(nIso, 1).Address(False, False) & ")"
    
        
    ' Set name
    mwksProto.Calculate
    sName = rngProtoNlib.Cells(1, 1).Offset(2, mocComment).Text
    sRefersTo = "=" & mwksProto.Name & "!" & _
        Range(mrngTblBase.Offset(morHead + 1, 0), rngProtoNlib.Cells(1, 1).Offset(1, mocComment)).Address
    
    On Error Resume Next
    Set nm = ThisWorkbook.Names(sName)
    If Not nm Is Nothing Then
        ThisWorkbook.Names(sName).Delete
    End If
    ThisWorkbook.Names.Add sName, sRefersTo

    ' Record last row
    miLastRow = rngProtoNlib.Cells(1, 1).Offset(3, 0).Row
        
    ' Format below
    nAux = mcMatTblComp.LastRow - miLastRow
    If nAux > 0 Then
        FormatAux rngProtoNlib.Cells(1, 1).Offset(4, 0).Resize(nAux, nCol)
    End If
        
    Application.Calculation = xlCalc
    Application.ScreenUpdating = blnUpdate
    Application.StatusBar = False
        
        
End Sub





