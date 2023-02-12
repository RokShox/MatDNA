VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMatTblComp"
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
Dim mrngTblMatNumIso As Range
Dim mrngTblMatFracIso As Range
Dim mrngTblMatFracCon As Range
Dim mcMat As clsMatMaterialRange

Dim mocProto As Integer
Dim morProto As Integer
Dim mncLeft As Integer
Dim mncRight As Integer
Dim mnNumRows As Integer

Dim mocNumIso As Integer
Dim mocNumRows As Integer
Dim mocCom As Integer
Dim mocCon As Integer
Dim mocIso As Integer
Dim mocMass As Integer
Dim mocIsoWfrac As Integer
Dim mocIsoAfrac As Integer
Dim mocConWfrac As Integer
Dim mocConAfrac As Integer
Dim mocComWfrac As Integer
Dim mocComAfrac As Integer
Dim morTblName As Integer
Dim morTblComp As Integer
Dim mnrTblCompHeader As Integer

Dim msThisClass As String


Private Sub Class_Initialize()
    msThisClass = "clsMatTblComp"
        
    mocProto = 16
    morProto = 2
    mncLeft = 3
    mncRight = 12
    
    ' Relative to proto base
    mocNumIso = 0
    mocNumRows = 1
    mocCom = 3
    mocCon = 4
    mocIso = 5
    mocMass = 6
    mocIsoWfrac = 7
    mocIsoAfrac = 8
    mocConWfrac = 9
    mocConAfrac = 10
    mocComWfrac = 11
    mocComAfrac = 12
    morTblName = 2
    morTblComp = 4
    mnrTblCompHeader = 4
End Sub

Property Set Base(rngBase As Range)

    Set mrngMatBase = rngBase
    Set mrngTblBase = mrngMatBase.Offset(morProto, mocProto)
    
End Property

Property Set Material(cMat As clsMatMaterialRange)
    Set mcMat = cMat
End Property

Property Get Name() As String
    Name = mcMat.Name
End Property

Property Get LastRow() As Integer
    LastRow = mcMat.LastRow
End Property

Property Get UniqueIsos() As ArrayList
    Dim alIso As ArrayList
    Dim rngIso As Range
    Dim iIso As Integer
    Dim sIso As String
    
    Set alIso = New ArrayList
    Set rngIso = mrngTblBase.Offset(morTblComp + mnrTblCompHeader, mocIso).Resize(mnNumRows, 1)
    Debug.Print rngIso.Address
    
    For iIso = 1 To rngIso.Rows.Count
        sIso = rngIso.Cells(iIso, 1)
        If sIso <> "" And Not alIso.Contains(sIso) Then
            alIso.Add sIso
        End If
    Next iIso
    alIso.Sort

    Set UniqueIsos = alIso
End Property

Property Get AValueRange() As Range
    Set AValueRange = mrngTblBase.Offset(morTblComp + mnrTblCompHeader + mnNumRows - 1, mocMass)
End Property

Public Sub DisplayExcel()
    Dim xlCalc As XlCalculation
    Dim blnUpdate As Boolean
    Dim rngStart As Range
        
    xlCalc = Application.Calculation
    Application.Calculation = xlCalculationManual
    blnUpdate = Application.ScreenUpdating
    Application.ScreenUpdating = False
    Application.StatusBar = mrngMatBase.Text & ": " & "Composition"
   
    Set rngStart = mrngTblBase
    mcMat.DisplayExcel rngStart
    
    mnNumRows = mcMat.NumRows
    
    Application.Calculation = xlCalc
    Application.ScreenUpdating = blnUpdate
    Application.StatusBar = False


End Sub


