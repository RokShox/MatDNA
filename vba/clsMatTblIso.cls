VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMatTblIso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' Copyright (c) 2022 Henkel Technical Computing, LLC. All rights reserved.
'
Option Explicit
Option Compare Text

Dim mwksMat As Worksheet
Dim mdElm As Scripting.Dictionary
Dim mdOverride As Scripting.Dictionary
Dim mcElmIso As clsMatElmIso
Dim mrngMatBase As Range
Dim mrngTblBase As Range

' Table offsets
Dim mocProto As Integer
Dim morProto As Integer
Dim mocElm As Integer
Dim mocNumIso As Integer
Dim mocMode As Integer
Dim mocIso As Integer
Dim mocFrac As Integer

Dim msTblNumIso As String
Dim msTblFracIso As String
Dim msThisClass As String
Dim meMode As CompositionMode

Const q As String = """"

Private Sub Class_Initialize()
    
    msThisClass = "clsMatTblIso"
    
    Set mwksMat = Worksheets("Materials")
    Set mdElm = New Scripting.Dictionary
    Set mdOverride = New Scripting.Dictionary
    
    ' Relative to material stanza base cell
    mocProto = 10
    morProto = 5
    ' Relative to table
    mocElm = 0
    mocNumIso = 1
    mocMode = 2
    mocIso = 3
    mocFrac = 4
End Sub

Property Get Base() As Range
    Set Base = mrngMatBase
End Property

Property Set Base(rngBase As Range)
    Set mrngMatBase = rngBase
    Set mrngTblBase = mrngMatBase.Offset(morProto, mocProto)
    msTblNumIso = "tbl" & mrngMatBase.Text & "NumIso"
    msTblFracIso = "tbl" & mrngMatBase.Text & "FracIso"
End Property

Property Set ElmIso(c As clsMatElmIso)
    Set mcElmIso = c
End Property

Public Sub LoadElements(alElm As ArrayList)
    Dim cCon As clsMatConRange
    Dim oElm As Integer
    Dim sElm As String
    Dim result As VbMsgBoxResult
    
    If mcElmIso Is Nothing Then
        result = MsgBox("Forgot to set ElmIso. Create one?", vbOKCancel, msThisClass)
        If result = vbOK Then
            Set mcElmIso = New clsMatElmIso
        Else
            End
        End If
    End If
        
    mdElm.RemoveAll
    For oElm = 0 To alElm.Count - 1
        sElm = alElm(oElm)
        Set cCon = mcElmIso.ConstituentRange(sElm)
        mdElm.Add sElm, cCon
    Next oElm
    
End Sub

' Must LoadElements first
Public Sub Override(rng As Range)
    Dim iRow As Integer
    Dim nRow As Integer
    Dim sElm As String
    Dim sMode As String
    Dim sIso As String
    Dim rngFrac As Range
    Dim cCon As clsMatConRange
    Dim cIso As clsMatIso
    Dim eMode As CompositionMode
    Dim key As Variant
    Const ic_Elm As Integer = 1
    Const ic_Mode As Integer = 2
    Const ic_Iso As Integer = 3
    Const ic_Frac As Integer = 4
    
    Set cCon = Nothing
    nRow = rng.Rows.Count
    For iRow = 1 To nRow
        sElm = rng.Cells(iRow, ic_Elm).Text
        If sElm <> "" Then
            ' Seal current con
            If Not cCon Is Nothing Then
                cCon.Seal
                mdOverride.Add cCon.Name, cCon
            End If
            
            sMode = rng.Cells(iRow, ic_Mode).Text
            If sMode = "Atom" Then
                eMode = Atom
            ElseIf sMode = "Mass" Or sMode = "Weight" Then
                eMode = Mass
            Else
                eMode = Atom
            End If
            
            ' Create new con
            Set cCon = New clsMatConRange
            With cCon
                .Name = sElm
                .mode = eMode
                .CopyMode = ByFormula
            End With
        End If
        
        sIso = rng.Cells(iRow, ic_Iso).Text
        Set rngFrac = rng.Cells(iRow, ic_Frac)
        
        If sIso <> "" Then
            Set cIso = New clsMatIso
            With cIso
                .Iso = sIso
            End With
            cCon.Add cIso, rngFrac
        End If
        
    Next iRow

    ' Seal current con
    If Not cCon Is Nothing Then
        cCon.Seal
        mdOverride.Add cCon.Name, cCon
    End If
    
    ' Do overrides
    For Each key In mdElm.Keys
        If mdOverride.Exists(key) Then
            Set cCon = mdOverride.Item(key)
            mdElm.Remove key
            mdElm.Add cCon.Name, cCon
        End If
    Next key
End Sub

Public Function ConList() As ArrayList
    Dim key As Variant

    Set ConList = New ArrayList
    For Each key In mdElm.Keys
        ConList.Add key
    Next
    ConList.Sort
    
End Function

' Constituents in mdElm should not be exposed
'Property Get Constituent(key As Variant) As clsMatConRange
'
'    If Not mdElm.Exists(key) Then
'        MsgBox "Key " & key & " does not exist", vbCritical, msThisClass
'        GoTo Fail:
'    End If
'
'    Set Constituent = mdElm.Item(key)
'
'    Exit Property
'Fail:
'    End
'
'End Property


Public Sub DisplayExcel()
    Dim cCon As clsMatConRange
    Dim alKey As ArrayList
    Dim elmkey As Variant
    Dim isokey As Variant
    Dim sElm As String
    Dim sIso As String
    Dim oElm As Integer
    Dim nElm As Integer
    Dim nIso As Integer
    Dim oIso As Integer
    Dim oRow As Integer
    Dim eMode As CompositionMode
    Dim xlCalc As XlCalculation
    Dim blnUpdate As Boolean

    nElm = mdElm.Count

    xlCalc = Application.Calculation
    Application.Calculation = xlCalculationManual
    blnUpdate = Application.ScreenUpdating
    Application.ScreenUpdating = False
    Application.StatusBar = mrngMatBase.Text & ": " & "Isotopics"
        
    With mrngTblBase
        ' Header
        .Offset(-2, mocElm).Value2 = "Isotopic Distributions"
        .Offset(-1, mocElm).Value2 = "Element"
        .Offset(-1, mocNumIso).Value2 = "Num Iso"
        .Offset(-1, mocMode).Value2 = "Mode"
        .Offset(-1, mocIso).Value2 = "Iso"
        .Offset(-1, mocFrac).Value2 = "Fraction"
        
        mdlExcel.BorderExtent .Offset(-2, mocElm).Resize(1, 5)
        mdlExcel.FormatEntries .Offset(-2, mocElm).Resize(1, 5)
        .Offset(-2, 0).HorizontalAlignment = xlLeft
        
        mdlExcel.BorderInside .Offset(-1, mocElm).Resize(1, 5)
        mdlExcel.FormatEntries .Offset(-1, mocElm).Resize(1, 5)
                
        oRow = 0
        For Each elmkey In mdElm.Keys
            Set cCon = mdElm(elmkey)
            sElm = CStr(elmkey)
            nIso = cCon.Count
            eMode = cCon.mode
            
            .Offset(oRow, mocElm).Value2 = sElm
            .Offset(oRow, mocNumIso).Value2 = nIso
            Select Case eMode
                Case Atom
                    .Offset(oRow, mocMode).Value2 = "Atom"
                Case Mass
                    .Offset(oRow, mocMode).Value2 = "Mass"
            End Select
            
            Set alKey = cCon.UniqueIsos
            oIso = 0
            For Each isokey In alKey
                sIso = CStr(isokey)
                .Offset(oRow + oIso, mocIso).Value2 = sIso
                Select Case cCon.CopyMode
                    Case ByValue
                        .Offset(oRow + oIso, mocFrac).Value2 = cCon.Fraction(sIso)
                    Case ByReference
                        .Offset(oRow + oIso, mocFrac).Formula = "=" & cCon.FractionRange(sIso).Address(False, False)
                    Case ByFormula
                        cCon.FractionRange(sIso).Copy
                        .Offset(oRow + oIso, mocFrac).PasteSpecial xlPasteFormulasAndNumberFormats
'                        .Offset(oRow + oIso, mocFrac).Formula = cCon.FractionRange(sIso).Formula
                End Select
                oIso = oIso + 1
            Next isokey
            
            ' Format
            mdlExcel.BorderInside .Offset(oRow, mocElm).Resize(nIso, 5)
            mdlExcel.FormatEntries .Offset(oRow, mocElm).Resize(nIso, 5)
            
            ' Move to next elm
            oRow = oRow + nIso
        Next elmkey
        
        ' Table names
        .Offset(oRow, mocMode).Value2 = "!" & msTblNumIso
        .Offset(oRow, mocFrac).Value2 = "!" & msTblFracIso
        With .Offset(oRow, mocMode)
            .Font.Color = 32768
            .HorizontalAlignment = xlRight
        End With
        With .Offset(oRow, mocFrac)
            .Font.Color = 32768
            .HorizontalAlignment = xlRight
        End With
    
        On Error Resume Next
        ActiveSheet.Names(msTblNumIso).Delete
        ActiveSheet.Names.Add msTblNumIso, .Offset(0, mocElm).Resize(oRow, 3)
        On Error Resume Next
        ActiveSheet.Names(msTblFracIso).Delete
        ActiveSheet.Names.Add msTblFracIso, .Offset(0, mocIso).Resize(oRow, 2)
    End With
    
    Application.Calculation = xlCalc
    Application.ScreenUpdating = blnUpdate
    Application.StatusBar = False
    
End Sub


Property Get TableConstituent(key As Variant) As clsMatConRange
    Dim cCon As clsMatConRange
    Dim cIso As clsMatIso
    Dim iRow As Integer
    Dim iStart As Integer
    Dim nIso As Integer
    Dim sMode As String
    Dim sIso As String
    Dim sCon As String
    Dim rngFrac As Range
    Dim rngTblNumIso As Range
    Dim rngTblFracIso As Range
    
    Set rngTblNumIso = ActiveSheet.Names(msTblNumIso).RefersToRange
    Set rngTblFracIso = ActiveSheet.Names(msTblFracIso).RefersToRange
    
    ' Find constituent
    iStart = 0
    sCon = CStr(key)
    For iRow = 1 To rngTblNumIso.Rows.Count
        If rngTblNumIso.Cells(iRow, 1).Text = sCon Then
            iStart = iRow
            Exit For
        End If
    Next iRow
    
    If iStart = 0 Then
        MsgBox "Unable to find constituent " & sCon, vbCritical, msThisClass
        GoTo Fail:
    End If

    nIso = rngTblNumIso.Cells(iRow, 2).Value2
    sMode = rngTblNumIso.Cells(iRow, 3).Text

    Set cCon = New clsMatConRange
    With cCon
        .Name = sCon
        Select Case sMode
            Case "Atom"
                .mode = Atom
            Case "Mass"
                .mode = Mass
        End Select
        .CopyMode = ByReference
        
        For iRow = iStart To iStart + nIso - 1
            sIso = rngTblFracIso.Cells(iRow, 1)
            Set rngFrac = rngTblFracIso.Cells(iRow, 2)
            Set cIso = New clsMatIso
            cIso.Iso = sIso
            .Add cIso, rngFrac
        Next iRow
        .Seal
    End With
    
    Set TableConstituent = cCon
    
    Exit Property

Fail:
    Set TableConstituent = Nothing
        
End Property


