VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMatTblFracCon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' Copyright (c) 2022 Henkel Technical Computing, LLC. All rights reserved.
'
Option Explicit
Option Compare Text

Dim mdFrac As Scripting.Dictionary
Dim mrngMatBase As Range
Dim mrngTblBase As Range
Dim meMode As CompositionMode
Dim meCopyMode As MatFracCopyMode
Dim msFracHead As String

Dim mocProto As Integer
Dim morProto As Integer
Dim mocCon As Integer
Dim mocFrac As Integer
Dim msTblName As String
Dim msThisClass As String
Const q As String = """"

Private Sub Class_Initialize()
    msThisClass = "clsMatTblFracCon"
    Set mdFrac = New Scripting.Dictionary
    mdFrac.CompareMode = TextCompare
    meMode = Mass
    meCopyMode = ByValue
    msFracHead = "Mass Fraction"
    
    ' Relative to material stanza base cell
    mocProto = 5
    morProto = 5
    ' Relative to table
    mocCon = 0
    mocFrac = 1
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

Property Let mode(e As CompositionMode)
    meMode = e
    Select Case meMode
        Case Atom
            msFracHead = "Atom Fraction"
        Case Mass
            msFracHead = "Mass Fraction"
    End Select
End Property

Property Get CopyMode() As MatFracCopyMode
    CopyMode = meCopyMode
End Property

Property Let CopyMode(e As MatFracCopyMode)
    meCopyMode = e
End Property

Public Function ConList() As ArrayList
    Dim alCons As ArrayList
    Dim key As Variant
    
    Set alCons = New ArrayList
    For Each key In mdFrac.Keys
        alCons.Add key
    Next key
    alCons.Sort
    Set ConList = alCons
End Function

Public Sub AddConstituent(sCon As String, rngFrac As Range)
        mdFrac.Add sCon, rngFrac
End Sub

Public Sub DisplayExcel()
    Dim alKeys As ArrayList
    Dim rngFrac As Range
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
        .Offset(-2, 0).Value2 = "Constituent Fractions"
        .Offset(-1, 0).Value2 = "Element"
        .Offset(-1, 1).Value2 = msFracHead
        
        ' Data
        For oCon = 0 To nCon - 1
            .Offset(oCon, mocCon).Value2 = alKeys(oCon)
            Set rngFrac = mdFrac.Item(alKeys(oCon))
            Select Case meCopyMode
                Case ByValue
                    .Offset(oCon, mocFrac).Value2 = rngFrac.Value2
                Case ByReference
                    .Offset(oCon, mocFrac).Formula = "=" & rngFrac.Address(False, False)
                Case ByFormula
                    rngFrac.Copy
                    .Offset(oCon, mocFrac).PasteSpecial xlPasteFormulasAndNumberFormats
            End Select
        Next oCon
        
        ' Table name
        .Offset(nCon, mocFrac).Value2 = "!" & msTblName
    
        On Error Resume Next
        ActiveSheet.Names(msTblName).Delete
        ActiveSheet.Names.Add msTblName, .Offset(0, 0).Resize(nCon, 2)
    End With
    
    ' Formatting
    With mrngTblBase
        mdlExcel.FormatEntries .Offset(-2, 0).Resize(1, 2)
        mdlExcel.BorderExtent .Offset(-2, 0).Resize(1, 2)
        .Offset(-2, 0).HorizontalAlignment = xlLeft
        mdlExcel.BorderInside .Offset(-1, 0).Resize(1, 2)
        mdlExcel.FormatEntries .Offset(-1, 0).Resize(1, 2)
        mdlExcel.BorderInside .Offset(0, 0).Resize(nCon, 2)
        mdlExcel.FormatEntries .Offset(0, 0).Resize(nCon, 2)
        .Offset(0, mocFrac).Resize(nCon, 1).NumberFormat = "0.0000"
        With .Offset(nCon, mocFrac)
            .Font.Color = 32768
            .HorizontalAlignment = xlRight
        End With
    End With
    
    Application.Calculation = xlCalc
    Application.ScreenUpdating = blnUpdate
    Application.StatusBar = False

End Sub
