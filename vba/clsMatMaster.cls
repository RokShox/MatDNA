VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMatMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' Copyright (c) 2022 Henkel Technical Computing, LLC. All rights reserved.
'
Option Explicit
Option Compare Text

Dim mrngTblMaster As Range
Dim mrngTblMasterExtended As Range
Dim mrngMatKey
Dim micMat As Integer
Dim micDesc As Integer
Dim micName As Integer
Dim micMatID As Integer
Dim micNameMDen As Integer
Dim micMassDen As Integer
Dim micNameADen As Integer
Dim micAtomDen As Integer
Dim mnExtraColsRight As Integer
Dim mnExtraColsLeft As Integer
Dim mnStanzaCols As Integer
Dim msThisClass As String

Private Sub Class_Initialize()
    Dim iMat As Integer
    Dim iRow As Integer
    Dim sMat As String
    Dim iColKey As Integer

    msThisClass = "clsMatMaster"
    
    ' Relative to table
    micMat = 1
    micDesc = 2
    micName = 3
    micMatID = 4
    micNameMDen = 5
    micMassDen = 6
    micNameADen = 7
    micAtomDen = 8
    mnExtraColsRight = 8
    mnExtraColsLeft = 1
    Set mrngTblMaster = ThisWorkbook.Names("tblMaterials").RefersToRange
    Set mrngTblMasterExtended = Range(mrngTblMaster.Cells(1, 1).Offset(0, -mnExtraColsLeft), _
        mrngTblMaster.Cells(mrngTblMaster.Rows.Count, mrngTblMaster.Columns.Count).Offset(0, mnExtraColsRight))

    
    iColKey = mrngTblMaster.Worksheet.Names("icol_Key").RefersToRange.Value2
    mnStanzaCols = Range("DI:DI").Column - iColKey + 1
    Set mrngMatKey = mrngTblMaster.Worksheet.Columns(iColKey)
    
End Sub


Property Get Description(sMat As String) As String
    Description = Application.WorksheetFunction.VLookup(sMat, mrngTblMaster, micDesc, False)
End Property


Property Get MatID(sMat As String) As Integer
    MatID = Application.WorksheetFunction.VLookup(sMat, mrngTblMaster, micMatID, False)
End Property


Property Get MassDen(sMat As String) As Double
    MassDen = CDbl(Application.WorksheetFunction.VLookup(sMat, mrngTblMaster, micMassDen, False))
End Property


Property Get AtomDen(sMat As String) As Double
    AtomDen = CDbl(Application.WorksheetFunction.VLookup(sMat, mrngTblMaster, micAtomDen, False))
End Property

Property Get MaxRow() As Integer
    MaxRow = Application.WorksheetFunction.Max( _
    mrngTblMaster.Cells(1, 1).Offset(1, -1).Resize(mrngTblMaster.Rows.Count - 1, 1))
End Property


Public Sub Add(sMat As String)
    Dim result As VbMsgBoxResult
    Dim iRowBlank As Integer
    Dim iRow As Integer
    Dim iRowKey As Integer
    Dim sName As String
    
    If Me.StanzaExists(sMat) Then
    
        ' Find first blank row
        iRowBlank = 0
        For iRow = 1 To mrngTblMaster.Rows.Count
            If mrngTblMaster.Cells(iRow, 1).Text = "" Then
                iRowBlank = iRow
                Exit For
            End If
        Next iRow
        
        If iRowBlank = 0 Then
            result = MsgBox("No blank rows in table", vbOKOnly, msThisClass)
            End
        End If
        
        ' Insert a row in extended table
        mrngTblMasterExtended.Rows(iRowBlank).Insert
        
        ' Copy previous row
        mrngTblMasterExtended.Rows(iRowBlank - 1).Copy mrngTblMasterExtended.Rows(iRowBlank)
        mrngTblMaster.Cells(iRowBlank, 1).Value2 = sMat
        mrngTblMasterExtended.Rows(iRowBlank).Calculate
        
        ' Set names
        sName = mrngTblMaster.Cells(iRowBlank, micName).Text
        ThisWorkbook.Names.Add sName, mrngTblMaster.Cells(iRowBlank, micMatID)
        sName = mrngTblMaster.Cells(iRowBlank, micNameMDen).Text
        ThisWorkbook.Names.Add sName, mrngTblMaster.Cells(iRowBlank, micMassDen)
        sName = mrngTblMaster.Cells(iRowBlank, micNameADen).Text
        ThisWorkbook.Names.Add sName, mrngTblMaster.Cells(iRowBlank, micAtomDen)
        
        ' Update hyperlink
        Me.UpdateLink sMat
    
    Else
        result = MsgBox("No material stanza found", vbOKOnly, msThisClass)
    End If
    
End Sub

Public Sub UpdateLinks()
    Dim iRow As Integer
    Dim iColLink As Integer
    Dim iRowLink As Integer
    Dim iLink As Integer
    
    iColLink = mrngMatKey.Column + 19
    
    ' Skip void entry
    For iRow = 2 To mrngTblMaster.Rows.Count
        With mrngTblMaster.Cells(iRow, 1).Offset(0, -1)
            If .Text <> "" Then
                iRowLink = .Value2 + 6
                For iLink = .Hyperlinks.Count To 1 Step -1
                    .Hyperlinks(iLink).Delete
                Next iLink
                .Hyperlinks.Add Anchor:=.Range("A1"), Address:="", _
                    SubAddress:=ActiveSheet.Name & "!" & ActiveSheet.Cells(iRowLink, iColLink).Address
                FormatAux .Range("A1")
            End If
        End With
    Next iRow
End Sub


Public Sub UpdateLink(sMat As String)
    Dim iRow As Integer
    Dim iColLink As Integer
    Dim iRowLink As Integer
    Dim iLink As Integer
    
    iColLink = mrngMatKey.Column + 19
    
    ' Skip void entry
    For iRow = 2 To mrngTblMaster.Rows.Count
        If mrngTblMaster.Cells(iRow, 1).Text = sMat Then
            With mrngTblMaster.Cells(iRow, 1).Offset(0, -1)
                If .Text <> "" Then
                    iRowLink = .Value2 + 6
                    For iLink = .Hyperlinks.Count To 1 Step -1
                        .Hyperlinks(iLink).Delete
                    Next iLink
                    .Hyperlinks.Add Anchor:=.Range("A1"), Address:="", _
                        SubAddress:=ActiveSheet.Name & "!" & ActiveSheet.Cells(iRowLink, iColLink).Address
                    FormatAux .Range("A1")
                End If
            End With
            Exit Sub
        End If
    Next iRow
End Sub


Public Function StanzaExists(sMat As String) As Boolean
     Dim iRow As Integer
     
     iRow = 0
     On Error Resume Next
     iRow = Application.WorksheetFunction.Match(sMat, mrngMatKey, 0)
     If iRow = 0 Then
        StanzaExists = False
     Else
        StanzaExists = True
     End If

End Function

Public Sub Delete(sMat As String)
    Dim iRow As Integer
    Dim rngStanza As Range

    On Error GoTo Stanza:
    iRow = Application.WorksheetFunction.Match(sMat, mrngTblMaster.Columns(1), 0)

    deleteName "mat" & sMat
    deleteName "rho" & sMat
    deleteName "aden" & sMat
    
    mrngTblMasterExtended.Rows(iRow).Delete xlShiftUp

Stanza:
    Set rngStanza = Me.Stanza(sMat)
    
    If Not rngStanza Is Nothing Then
        ' Delete names
        deleteName "!tbl" & sMat & "FracCon"
        deleteName "!tbl" & sMat & "NumIso"
        deleteName "!tbl" & sMat & "FracIso"
        deleteName "inpMCNPMat" & sMat
        deleteName "inpSCALEMat" & sMat
    
        ' Delete stanza
        rngStanza.Delete xlShiftUp
        
        Me.UpdateLinks
    End If

End Sub

Public Function Stanza(sMat As String) As Range
    Dim iRowStart As Integer
    Dim iRowCount As Integer
    Dim iRow As Integer
    
    If Not Me.StanzaExists(sMat) Then
'        MsgBox "Material " & sMat & " does not exist", vbOKOnly, msThisClass
        Set Stanza = Nothing
        Exit Function
    End If
    
    iRowStart = Application.WorksheetFunction.Match(sMat, mrngMatKey, 0)

    iRow = iRowStart + 1
    While mrngMatKey.Cells(iRow, 1).Text = ""
        iRow = iRow + 1
    Wend
    
    iRowCount = iRow - iRowStart

    Set Stanza = mrngMatKey.Cells(iRowStart, 1).Resize(iRowCount, mnStanzaCols)
End Function


Private Sub deleteName(sName As String)
    
    If Left(sName, 1) = "!" Then
        sName = Mid(sName, 2)
        On Error GoTo Done
        ActiveSheet.Names(sName).Delete
    Else
        On Error GoTo Done
        ThisWorkbook.Names(sName).Delete
    End If

Done:

End Sub
