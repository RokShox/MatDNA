VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMatElmIso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' Copyright (c) 2022 Henkel Technical Computing, LLC. All rights reserved.
'
Option Explicit
Option Compare Text

Dim mdZbyElm As Scripting.Dictionary
Dim mdRowbyElm As Scripting.Dictionary
Dim mrngTblElmIso As Range
Dim micZ As Integer
Dim micElm As Integer
Dim micNumIso As Integer
Dim micIso As Integer
Dim micAfrac As Integer
Dim msThisClass As String

Private Sub Class_Initialize()
    Dim iElm As Integer
    Dim iZ As Integer
    Dim sElm As String

    msThisClass = "clsMatElmIso"

    Set mrngTblElmIso = ThisWorkbook.Names("tblNatIso").RefersToRange
    ' Relative to table
    micZ = 1
    micElm = 2
    micNumIso = 3
    micIso = 6
    micAfrac = 8
    
    ' Create dictionary giving Z by Element
    Set mdZbyElm = New Scripting.Dictionary
    mdZbyElm.CompareMode = TextCompare
    Set mdRowbyElm = New Scripting.Dictionary
    mdRowbyElm.CompareMode = TextCompare
    
    For iElm = 1 To mrngTblElmIso.Rows.Count
        sElm = mrngTblElmIso.Cells(iElm, micElm).Text
        
        If sElm <> "" Then
            iZ = mrngTblElmIso.Cells(iElm, micZ).Value2
            mdZbyElm.Add sElm, iZ
            mdRowbyElm.Add sElm, iElm
        End If
    Next iElm

End Sub

Property Get ZbyElm(sElm As String) As Integer
    If mdZbyElm.Exists(sElm) Then
        ZbyElm = mdZbyElm.Item(sElm)
    Else
        ZbyElm = Null
    End If
End Property

Property Get ElmbyZ(iZ As Integer) As String
    ElmbyZ = Application.WorksheetFunction.VLookup(iZ, mrngTblElmIso, micElm, False)
End Property

Property Get NumIso(sElm As String) As Integer
    If mdRowbyElm.Exists(sElm) Then
        NumIso = mrngTblElmIso.Cells(mdRowbyElm.Item(sElm), micNumIso).Value2
    Else
        NumIso = Null
    End If
End Property

Property Get Isos(sElm As String) As Scripting.Dictionary
    Dim c As Scripting.Dictionary
    Dim iRow As Integer
    Dim iIso As Integer
    Dim nIso As Integer
    Dim sIso As String
    Dim dblAfrac As Double
    
    If mdRowbyElm.Exists(sElm) Then
    
        Set c = New Scripting.Dictionary
        c.CompareMode = TextCompare
        iRow = mdRowbyElm(sElm)
        nIso = mrngTblElmIso.Cells(iRow, micNumIso).Value2
        For iIso = 0 To nIso - 1
            sIso = mrngTblElmIso.Cells(iRow + iIso, micIso).Text
            dblAfrac = CDbl(mrngTblElmIso.Cells(iRow + iIso, micAfrac).Value2)
            c.Add sIso, dblAfrac
        Next iIso

    Else
        Set c = Nothing
    End If
    Set Isos = c
End Property

Property Get Constituent(sElm As String) As clsMatCon
    Dim cCon As clsMatCon
    Dim cIso As clsMatIso
    Dim iRow As Integer
    Dim iIso As Integer
    Dim nIso As Integer
    Dim sIso As String
    Dim dblAfrac As Double
    
    If mdRowbyElm.Exists(sElm) Then
    
        Set cCon = New clsMatCon
        With cCon
            .Name = sElm
            .mode = Atom
        End With
        iRow = mdRowbyElm(sElm)
        nIso = mrngTblElmIso.Cells(iRow, micNumIso).Value2
        For iIso = 0 To nIso - 1
            sIso = mrngTblElmIso.Cells(iRow + iIso, micIso).Text
            dblAfrac = CDbl(mrngTblElmIso.Cells(iRow + iIso, micAfrac).Value2)
            Set cIso = New clsMatIso
            cIso.Iso = sIso
            cCon.Add cIso, dblAfrac
        Next iIso
        cCon.Seal
    Else
        Set cCon = Nothing
    End If
    Set Constituent = cCon

End Property


Property Get ConstituentRange(sElm As String) As clsMatConRange
    Dim cCon As clsMatConRange
    Dim cIso As clsMatIso
    Dim iRow As Integer
    Dim iIso As Integer
    Dim nIso As Integer
    Dim sIso As String
    Dim rngAFrac As Range
    
    If mdRowbyElm.Exists(sElm) Then
    
        Set cCon = New clsMatConRange
        With cCon
            .Name = sElm
            .mode = Atom
            .CopyMode = ByValue
        End With
        iRow = mdRowbyElm(sElm)
        nIso = mrngTblElmIso.Cells(iRow, micNumIso).Value2
        For iIso = 0 To nIso - 1
            sIso = mrngTblElmIso.Cells(iRow + iIso, micIso).Text
            Set rngAFrac = mrngTblElmIso.Cells(iRow + iIso, micAfrac)
            Set cIso = New clsMatIso
            cIso.Iso = sIso
            cCon.Add cIso, rngAFrac
        Next iIso
        cCon.Seal
    Else
        Set cCon = Nothing
    End If
    Set ConstituentRange = cCon

End Property


