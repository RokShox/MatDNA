VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMatStoichTblFracCon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' Copyright (c) 2022 Henkel Technical Computing, LLC. All rights reserved.
'
Option Explicit
Option Compare Text

Dim mdStoich As Scripting.Dictionary
Dim mrngMatBase As Range
Dim mrngTblBase As Range
Dim meMode As CompositionMode

Dim mocProto As Integer
Dim morProto As Integer
Dim mocCon As Integer
Dim mocStoich As Integer
Dim mocFrac As Integer
Dim msTblName As String
Dim msThisClass As String
Const q As String = """"

Private Sub Class_Initialize()
    msThisClass = "clsMatStoichTblFracCon"
    Set mdStoich = New Scripting.Dictionary
    mdStoich.CompareMode = TextCompare
    meMode = Atom
    
    ' Mode is fixed Atom
    ' No CopyMode because stoich is explicit values not ranges
    
    ' Relative to material stanza base cell
    mocProto = 5
    morProto = 5
    ' Relative to table
    mocCon = 0
    mocStoich = 1
    mocFrac = 2
End Sub

Property Get Base() As Range
    Set Base = mrngMatBase
End Property

Property Set Base(rngBase As Range)
    Set mrngMatBase = rngBase
    Set mrngTblBase = mrngMatBase.Offset(morProto, mocProto)
    msTblName = "tbl" & mrngMatBase.Text & "FracCon"
End Property

Property Get Fraction(key As Variant) As Double
    Fraction = CDbl(Me.FractionRange(key).Value2)
End Property

Property Get FractionRange(key As Variant) As Range
    Dim rngTbl As Range
    Dim iCon As Integer
    
    ' Return range from table, not original range!
    Set rngTbl = ActiveSheet.Names(msTblName).RefersToRange
    Set FractionRange = Nothing
    For iCon = 1 To rngTbl.Rows.Count
        If rngTbl.Cells(iCon, mocCon + 1).Text = key Then
            Set FractionRange = rngTbl.Cells(iCon, mocFrac + 1)
            Exit For
        End If
    Next iCon
End Property

Property Get mode() As CompositionMode
    mode = meMode
End Property

' No set for Mode

' No CopyMode - uses explicit values

Property Get ConList() As ArrayList
    Dim alCons As ArrayList
    Dim key As Variant
    
    Set alCons = New ArrayList
    For Each key In mdStoich.Keys
        alCons.Add key
    Next key
    alCons.Sort
    Set ConList = alCons
End Property

Public Sub AddConstituent(sCon As String, dblStoich As Double)
        mdStoich.Add sCon, dblStoich
End Sub

Public Sub DisplayExcel()
    Dim alKeys As ArrayList
    Dim dblStoich As Double
    Dim oCon As Integer
    Dim nCon As Integer
    Dim xlCalc As XlCalculation
    Dim blnUpdate As Boolean
    
    xlCalc = Application.Calculation
    Application.Calculation = xlCalculationManual
    blnUpdate = Application.ScreenUpdating
    Application.ScreenUpdating = False
    Application.StatusBar = mrngMatBase.Text & ": " & "FracCon"

    Set alKeys = Me.ConList
    nCon = alKeys.Count
    
    With mrngTblBase
        ' Header
        .Offset(-2, 0).Value2 = "Stoichiometry"
        .Offset(-1, 0).Value2 = "Element"
        .Offset(-1, 1).Value2 = "Stoich"
        .Offset(-1, 2).Value2 = "Atom Frac"
        
        ' First two columns
        For oCon = 0 To nCon - 1
            .Offset(oCon, mocCon).Value2 = alKeys(oCon)
            dblStoich = CDbl(mdStoich.Item(alKeys(oCon)))
            .Offset(oCon, mocStoich).Value2 = dblStoich
        Next oCon
        
        .Offset(nCon, mocCon).Value2 = "total"
        .Offset(nCon, mocStoich).Formula = "=SUM(" & .Offset(0, mocStoich).Resize(nCon, 1).Address(False, False) & ")"
    
        ' Afrac column
        For oCon = 0 To nCon - 1
            .Offset(oCon, mocFrac).Formula = "=" & .Offset(oCon, mocStoich).Address(False, False) & _
            "/" & .Offset(nCon, mocStoich).Address(False, False)
        Next oCon
        .Offset(nCon, mocFrac).Formula = "=SUM(" & .Offset(0, mocFrac).Resize(nCon, 1).Address(False, False) & ")"
        
        ' Table name
        .Offset(nCon + 1, mocFrac).Value2 = "!" & msTblName
    
        On Error Resume Next
        ActiveSheet.Names(msTblName).Delete
        ActiveSheet.Names.Add msTblName, .Offset(0, 0).Resize(nCon + 1, 3)
    End With
    
    ' Formatting
    With mrngTblBase
        mdlExcel.FormatEntries .Offset(-2, 0).Resize(1, 3)
        mdlExcel.BorderExtent .Offset(-2, 0).Resize(1, 3)
        .Offset(-2, 0).HorizontalAlignment = xlLeft
        mdlExcel.BorderInside .Offset(-1, 0).Resize(1, 3)
        mdlExcel.FormatEntries .Offset(-1, 0).Resize(1, 3)
        mdlExcel.BorderInside .Offset(0, 0).Resize(nCon, 3)
        mdlExcel.FormatEntries .Offset(0, 0).Resize(nCon, 3)
        mdlExcel.BorderInside .Offset(nCon, 0).Resize(1, 3)
        mdlExcel.FormatEntries .Offset(nCon, 0).Resize(1, 3)
        .Offset(0, mocFrac).Resize(nCon, 1).NumberFormat = "0.0000"
        With .Offset(nCon + 1, mocFrac)
            .Font.Color = 32768
            .HorizontalAlignment = xlRight
        End With
    End With
    
    Application.Calculation = xlCalc
    Application.ScreenUpdating = blnUpdate
    Application.StatusBar = False

End Sub
