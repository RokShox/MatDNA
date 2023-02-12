VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMatTblSpec"
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
Dim mrngProtoSpec As Range
Dim mrngProtoDensities As Range
Dim mrngMatBase As Range
Dim mrngTblBase As Range
Dim mcMatTblComp As clsMatTblComp
Dim mrngDensityRange As Range

Dim mocProto As Integer
Dim morProto As Integer

Dim mocIsoRange As Integer
Dim morIsoRange As Integer

Dim mocSub As Integer
Dim mocActualAfrac  As Integer
Dim mocIso As Integer
Dim mocZAID As Integer
Dim mocAFrac As Integer
Dim mocMolWgt As Integer
Dim mocWfrac As Integer
Dim mocCopyA As Integer
Dim mocDensities As Integer
Dim mocDenMass As Integer
Dim mocDenAValue As Integer
Dim morDenMassDenActual As Integer
Dim morDenAtomDenActual As Integer
Dim morDenAValueActual As Integer
Dim morDenAValueModel As Integer
Dim morDenAtomDenModel As Integer

Dim mnIso As Integer
Dim miLastRow As Integer
Dim morTbl As Integer
Dim mnrTblHeader As Integer
Const q As String = """"

' Need to set Base, MatTblComp and MatTblParam
Private Sub Class_Initialize()

    Set mwksProto = Worksheets("Materials")

    Set mrngProtoSpec = mwksProto.Names("protoSpecification").RefersToRange
    Set mrngProtoDensities = mwksProto.Names("protoDensities").RefersToRange

    mocProto = 41
    morProto = 2
    morTbl = 4
    mnrTblHeader = 3
    
    mocIsoRange = 10
    morIsoRange = 13
    
    mocSub = 0
    mocActualAfrac = 1
    mocIso = 2
    mocZAID = 3
    mocAFrac = 4
    mocMolWgt = 5
    mocWfrac = 6
    mocCopyA = 7
    mocDensities = 9
    mocDenMass = 10
    mocDenAValue = 11
    morDenMassDenActual = 6
    morDenAtomDenActual = 7
    morDenAValueActual = 5
    morDenAValueModel = 8
    morDenAtomDenModel = 10
End Sub

Property Set Base(rngBase As Range)
    Set mrngMatBase = rngBase
    Set mrngTblBase = mrngMatBase.Offset(morProto, mocProto)
End Property

Property Set MatTblComp(c As clsMatTblComp)
    Set mcMatTblComp = c
End Property

Property Set DensityRange(rng As Range)
    Set mrngDensityRange = rng
End Property

Property Get NumIso() As Integer
    NumIso = mnIso
End Property

Property Get LastRow() As Integer
    LastRow = miLastRow
End Property

Public Sub DisplayExcel()
    Dim alIso As ArrayList
    Dim oIso As Integer
    Dim sIso As String
    Dim rngRowProto As Range
    Dim rngRowTotal As Range
    Dim rngCurrent As Range
    Dim sIsoRange As String
    Dim sAfracRange As String
    Dim sNiso As String
    Dim sSubStart As String
    Dim sActualStart As String
    Dim xlCalc As XlCalculation
    Dim blnUpdate As Boolean
    Dim rngDenLast As Range
    Dim iLastRowSpec As Integer
    Dim iLastRowDen As Integer
    Dim nAux As Integer
    Dim nColSpec As Integer
    Dim nColDen As Integer

    sIsoRange = mrngTblBase.Offset(morIsoRange, mocIsoRange).Address
    sAfracRange = mrngTblBase.Offset(morIsoRange + 1, mocIsoRange).Address
    sNiso = mrngTblBase.Offset(morTbl + 1, 1).Address(True, False)
    sSubStart = mrngTblBase.Offset(morTbl + mnrTblHeader, 0).Address(True, False)
    sActualStart = mrngTblBase.Offset(morTbl + mnrTblHeader, 1).Address(True, False)

    ' Clear old data
    nColSpec = mrngProtoSpec.Columns.Count
    mrngTblBase.Resize(mcMatTblComp.LastRow - mrngTblBase.Row + 1, nColSpec).Clear

    mrngProtoSpec.Copy mrngTblBase
    Set rngRowProto = mrngTblBase.Offset(morTbl + mnrTblHeader, 0).Resize(1, nColSpec)
    Set rngRowTotal = mrngTblBase.Offset(morTbl + mnrTblHeader + 1, 0).Resize(1, nColSpec)
    
    Set alIso = mcMatTblComp.UniqueIsos
    mnIso = alIso.Count

    xlCalc = Application.Calculation
    Application.Calculation = xlCalculationManual
    blnUpdate = Application.ScreenUpdating
    Application.ScreenUpdating = False
    Application.StatusBar = mrngMatBase.Text & ": " & "Specification"

    For oIso = 0 To mnIso - 1
        If oIso > 0 Then
            rngRowTotal.Insert
            Set rngCurrent = rngRowTotal.Cells(1, 1).Offset(-1, 0)
            rngRowProto.Copy rngCurrent
        Else
            Set rngCurrent = rngRowProto.Cells(1, 1)
        End If

        sIso = alIso(oIso)
        rngCurrent.Offset(0, mocIso).Value2 = sIso

        ' Check for subs
        rngCurrent.Offset(0, mocSub).Value2 = ""
        If sIso = "h-2" Then
            ' =IF(optSubIsos,"h-1","")
            With rngCurrent.Offset(0, mocSub)
                .Formula = "=IF(optSubIsos," & q & "h-1" & q & "," & q & q & ")"
                With .Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .ThemeColor = xlThemeColorAccent6
                    .TintAndShade = 0.599993896298105
                    .PatternTintAndShade = 0
                End With
            End With
        ElseIf sIso = "o-17" Then
            ' =IF(optSubIsos,"o-16","")
            With rngCurrent.Offset(0, mocSub)
                .Formula = "=IF(optSubIsos," & q & "o-16" & q & "," & q & q & ")"
                With .Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .ThemeColor = xlThemeColorAccent6
                    .TintAndShade = 0.599993896298105
                    .PatternTintAndShade = 0
                End With
            End With
        ElseIf sIso = "o-18" Then
            rngCurrent.Offset(0, mocSub).Value2 = "o-16"
        End If

        ' Fix Actual AF
        ' =SUMIF(INDIRECT($BR$3186),BJ3180,INDIRECT($BR$3187))
        rngCurrent.Offset(0, mocActualAfrac).Formula = _
            "=SUMIF(INDIRECT(" & sIsoRange & ")," & _
            rngCurrent.Offset(0, mocIso).Address(False, False) & "," & _
            "INDIRECT(" & sAfracRange & "))"

        ' Fix Afrac
        ' =IF(BH3180<>"",0,BI3180+SUMIF(OFFSET(BH$3180,0,0,BI$3178,1),BJ3180,OFFSET(BI$3180,0,0,BI$3178,1)))
        rngCurrent.Offset(0, mocAFrac).Formula = _
            "=IF(" & rngCurrent.Offset(0, mocSub).Address(False, False) & "<>" & q & q & ",0," & _
            rngCurrent.Offset(0, mocActualAfrac).Address(False, False) & "+" & _
            "SUMIF(OFFSET(" & sSubStart & ",0,0," & sNiso & ",1)," & _
            rngCurrent.Offset(0, mocIso).Address(False, False) & "," & _
            "OFFSET(" & sActualStart & ",0,0," & sNiso & ",1)))"

    Next oIso

    ' Fix CopyA
    Set rngCurrent = rngRowProto.Cells(1, 1)
    For oIso = 0 To mnIso - 1
        ' Fix CopyA
        ' =IF(BJ3180="Total",BM3180,BO3181)
        rngCurrent.Offset(0, mocCopyA).Formula = _
        "=IF(" & rngCurrent.Offset(0, mocIso).Address(False, False) & "=" & q & "Total" & q & "," & _
        rngCurrent.Offset(0, mocMolWgt).Address(False, False) & "," & _
        rngCurrent.Offset(1, mocCopyA).Address(False, False) & ")"

        Set rngCurrent = rngCurrent.Offset(1, 0)
    Next oIso

    ' Fix total row
    ' Total Actual AF =SUM(BI3180:BI3190)
    rngRowTotal.Cells(1, 1).Offset(0, mocActualAfrac).Formula = "=SUM(" & _
    rngRowTotal.Cells(1, 1).Offset(-mnIso, mocActualAfrac).Resize(mnIso, 1).Address(False, False) & ")"

    ' Total AF
    rngRowTotal.Cells(1, 1).Offset(0, mocAFrac).Formula = "=SUM(" & _
    rngRowTotal.Cells(1, 1).Offset(-mnIso, mocAFrac).Resize(mnIso, 1).Address(False, False) & ")"

    ' Total Mol Wgt
    ' Formula =SUMPRODUCT(BL3180:BL3190,BM3180:BM3190)
    rngRowTotal.Cells(1, 1).Offset(0, mocMolWgt).Formula = "=SUMPRODUCT(" & _
    rngRowTotal.Cells(1, 1).Offset(-mnIso, mocAFrac).Resize(mnIso, 1).Address(False, False) & "," & _
    rngRowTotal.Cells(1, 1).Offset(-mnIso, mocMolWgt).Resize(mnIso, 1).Address(False, False) & ")"

    ' Total WF
    rngRowTotal.Cells(1, 1).Offset(0, mocWfrac).Formula = "=SUM(" & _
    rngRowTotal.Cells(1, 1).Offset(-mnIso, mocWfrac).Resize(mnIso, 1).Address(False, False) & ")"

    ' Last row
    iLastRowSpec = rngRowTotal.Cells(1, 1).Row
    
    '
    ' Densities Table
    '
    
    ' Clear old data
    nColDen = mrngProtoDensities.Columns.Count
    mrngTblBase.Offset(0, mocDensities).Resize(mcMatTblComp.LastRow - mrngTblBase.Row + 1, nColDen).Clear
    
    mrngProtoDensities.Copy mrngTblBase.Offset(0, mocDensities)

    ' Actual A value
    mrngTblBase.Offset(morDenAValueActual, mocDenAValue).Formula = "=" & mcMatTblComp.AValueRange.Address(False, False)

    ' Actual mass density
    mrngTblBase.Offset(morDenMassDenActual, mocDenMass).Formula = "=" & mrngDensityRange.Address(False, False)

    ' Model A value
    mrngTblBase.Offset(morDenAValueModel, mocDenAValue).Formula = "=" & rngRowTotal.Cells(1, 1).Offset(0, mocMolWgt).Address(False, False)

    ' Model number density
    mrngTblBase.Offset(morDenAtomDenModel, mocDenMass).Formula = "=" & _
        rngRowTotal.Cells(1, 1).Offset(0, mocAFrac).Address(False, False) & "*" & _
        mrngTblBase.Offset(morDenAtomDenActual, mocDenMass).Address(False, False)

    ' Last row
    iLastRowDen = mrngTblBase.Offset(morIsoRange + 1, mocIsoRange).Row
    miLastRow = iLastRowSpec
    If iLastRowDen > iLastRowSpec Then miLastRow = iLastRowDen

    ' Format
    nAux = mcMatTblComp.LastRow - iLastRowSpec
    If nAux > 0 Then
        FormatAux rngRowTotal.Cells(1, 1).Offset(1, 0).Resize(nAux, nColSpec)
    End If
    nAux = mcMatTblComp.LastRow - iLastRowDen
    If nAux > 0 Then
        Set rngDenLast = mrngTblBase.Offset(mrngProtoDensities.Rows.Count - 1, mocDensities)
        FormatAux rngDenLast.Offset(1, 0).Resize(mcMatTblComp.LastRow - iLastRowDen, nColDen)
    End If
    
    BorderInside rngRowProto.Cells(1, 1).Resize(mnIso, 7)

    Application.Calculation = xlCalc
    Application.ScreenUpdating = blnUpdate
    Application.StatusBar = False

End Sub












