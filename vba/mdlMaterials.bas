Attribute VB_Name = "mdlMaterials"
'
' Copyright (c) 2022 Henkel Technical Computing, LLC. All rights reserved.
'
Option Explicit
Option Compare Text


Public Sub DeleteList()
    Dim rngDelete As Range
    Dim iRow As Integer
    Dim sMat As String
    Dim cMatMaster As clsMatMaster
    
    Set cMatMaster = New clsMatMaster
        
    Set rngDelete = ActiveSheet.Names("tblDelete").RefersToRange
    
    For iRow = 1 To rngDelete.Rows.Count
        sMat = rngDelete.Cells(iRow, 1).Text
        If sMat <> "" Then
            Application.StatusBar = sMat
            cMatMaster.Delete sMat
        Else
            Exit For
        End If
    Next iRow
    
    Application.StatusBar = False
    
End Sub

Public Sub UpdateLinks()
    Dim cMatMaster As clsMatMaster
    
    Set cMatMaster = New clsMatMaster
    cMatMaster.UpdateLinks
    
End Sub


Public Sub MakeStanzas()
    Dim rngTblNeeded As Range
    Dim rngLnkMat As Range
    Dim wksWDS As Worksheet
    Dim iRow As Integer
    Dim sExists As String
    
    Set wksWDS = Worksheets("WDS")
    Set rngTblNeeded = wksWDS.Names(wksWDS.Names("matTblNeeded").RefersToRange.Text).RefersToRange
    Set rngLnkMat = wksWDS.Names("lnkMat").RefersToRange
    
    For iRow = 1 To rngTblNeeded.Rows.Count
        If rngTblNeeded.Cells(iRow, 6).Text = "no" Then
            rngLnkMat.Value2 = iRow
            wksWDS.Calculate
            wksWDS.Activate
            AddMatByStoich
        Else
            Debug.Print "Skipping " & rngTblNeeded.Cells(iRow, 4).Text
        End If
    Next iRow
    
End Sub

Public Sub AddMatByStoich()
    Dim wksMat As Worksheet
    Dim cMatStanza As clsMatStoichStanza
    Dim cMatUtil As clsMatUtil
    Dim dStoich As Scripting.Dictionary
    
    Dim sName As String
    Dim sDesc As String
    Dim sFormula As String
    Dim dblDen As Double
    Dim sSab As String
    Dim sSource As String
    Dim vntKey As Variant
    
    Set wksMat = Worksheets("Materials")
    Set cMatUtil = New clsMatUtil
    Set cMatStanza = New clsMatStoichStanza
    
    With wksMat
        sName = .Names("matStoichName").RefersToRange.Text
        sDesc = .Names("matStoichDesc").RefersToRange.Text
        sFormula = .Names("matStoichFormula").RefersToRange.Text
        dblDen = CDbl(.Names("matStoichDen").RefersToRange.Value2)
        sSab = .Names("matStoichSab").RefersToRange.Text
        sSource = .Names("matStoichSource").RefersToRange.Text
    End With

    cMatUtil.ParseFormula sFormula, dStoich
    
    With cMatStanza
        .MatName = sName
        .DescriptiveText = sDesc
        .Density = dblDen
        .Sab = sSab
        .Source = sSource
        
        For Each vntKey In dStoich.Keys
            .AddElement CStr(vntKey), dStoich(vntKey)
        Next vntKey
        .AddStanza
    End With
    
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

End Sub

Public Sub MixMatByVF()
    Dim wksMat As Worksheet
    Dim cStanza As clsMatMixMatByVFStanza
    Dim iMat As Integer
    Dim sMat As String
    
    Set wksMat = Worksheets("Materials")
    wksMat.Activate
                
    Set cStanza = New clsMatMixMatByVFStanza
    
    With cStanza
        .MatName = wksMat.Names("matMixByVFracName").RefersToRange.Text
        .Sab = wksMat.Names("matMixByVFracSab").RefersToRange.Formula
        .AddStanza
    End With
    
    Set cStanza = Nothing

Done:
    
End Sub

Public Sub MakeMatVF()
    Dim wksMat As Worksheet
    Dim rngTblMatList As Range
    Dim rngProtoMixByVFrac As Range
    Dim rngMatMixByVFracName As Range
    
    Dim iMat As Integer
    Dim sMat As String
    Dim sMatVF As String

    Set wksMat = Worksheets("Materials")
    wksMat.Activate
        
    Set rngTblMatList = wksMat.Names("tblMatList").RefersToRange
    Set rngProtoMixByVFrac = wksMat.Names("protoMixByVFrac").RefersToRange
    Set rngMatMixByVFracName = wksMat.Names("matMixByVFracName").RefersToRange

    For iMat = 1 To rngTblMatList.Rows.Count
        sMat = rngTblMatList.Cells(iMat, 1).Text
        If sMat <> "" Then
            sMatVF = sMat & "_VF"
            rngMatMixByVFracName.Value2 = sMatVF
            rngProtoMixByVFrac.Cells(4, 1).Value2 = sMat
            rngProtoMixByVFrac.Calculate
            
            MixMatByVF
            
        End If
    Next iMat

End Sub


Public Sub AddMatByElm()
    Dim cMatByElm As clsMatByElmStanza
    Dim wksMat As Worksheet
    
    Dim sName As String
    Dim sDesc As String
    Dim sSab As String
    Dim dblDensity As Double
    Dim sSrc As String
    Dim eMode As CompositionMode
    
    Set wksMat = Worksheets("Materials")
    Set cMatByElm = New clsMatByElmStanza

    sName = wksMat.Names("matMatByElmName").RefersToRange.Text
    sDesc = wksMat.Names("matMatByElmDesc").RefersToRange.Text
    sSab = wksMat.Names("matMatByElmSab").RefersToRange.Text
    dblDensity = CDbl(wksMat.Names("matMatByElmDen").RefersToRange.Value2)
    sSrc = wksMat.Names("matMatByElmSource").RefersToRange.Text

    Select Case wksMat.Names("matMatByElmMode").RefersToRange.Text
        Case "Atom"
            eMode = Atom
        Case "Mass"
            eMode = Mass
    End Select

    With cMatByElm
        .MatName = sName
        .DescriptiveText = sDesc
        .Sab = sSab
        .Density = dblDensity
        .Source = sSrc
        .mode = eMode
        .AddStanza
    End With
End Sub

Public Sub MixMatByFrac()
    Dim cMixMatByFrac As clsMatMixMatByFracStanza
    Dim wksMat As Worksheet
    
    Dim sName As String
    Dim sDesc As String
    Dim sSab As String
    Dim dblDensity As Double
    Dim eMode As CompositionMode
    
    Set wksMat = Worksheets("Materials")
    Set cMixMatByFrac = New clsMatMixMatByFracStanza

    sName = wksMat.Names("matMixMatByFracName").RefersToRange.Text
    sDesc = wksMat.Names("matMixMatByFracDesc").RefersToRange.Text
    sSab = wksMat.Names("matMixMatByFracSab").RefersToRange.Formula
    dblDensity = CDbl(wksMat.Names("matMixMatByFracDen").RefersToRange.Value2)

    Select Case wksMat.Names("matMixMatByFracMode").RefersToRange.Text
        Case "Atom"
            eMode = Atom
        Case "Mass"
            eMode = Mass
    End Select

    With cMixMatByFrac
        .MatName = sName
        .DescriptiveText = sDesc
        .Sab = sSab
        .Density = dblDensity
        .mode = eMode
        .AddStanza
    End With
End Sub

Public Function AtomDenIsoInMat(sIso As String, sMat As String) As Double
    Dim rngStanza As Range
    Dim rngData As Range
    Dim cMat As clsMatMaterialRange
    Dim dblAtomDen As Double

    InitMaterials
    
    Set rngStanza = gblMatMaster.Stanza(sMat)
    Set rngData = rngStanza.Cells(1, 1).Offset(or_Stanza_Comp + or_Comp_Data, oc_Stanza_Comp)

    Set cMat = New clsMatMaterialRange
    With cMat
        .mode = Atom
        .CopyMode = ByValue
        .InitFromRange rngData
    End With

    dblAtomDen = gblMatMaster.AtomDen(sMat)

    AtomDenIsoInMat = cMat.FractionIsoInMat(sIso, Atom) * dblAtomDen

    Set cMat = Nothing

End Function


Public Function AtomDenConInMat(sCon As String, sMat As String) As Double
    Dim rngStanza As Range
    Dim rngData As Range
    Dim cMat As clsMatMaterialRange
    Dim dblAtomDen As Double

    InitMaterials
    
    Set rngStanza = gblMatMaster.Stanza(sMat)
    Set rngData = rngStanza.Cells(1, 1).Offset(or_Stanza_Comp + or_Comp_Data, oc_Stanza_Comp)

    Set cMat = New clsMatMaterialRange
    With cMat
        .mode = Atom
        .CopyMode = ByValue
        .InitFromRange rngData
    End With

    dblAtomDen = gblMatMaster.AtomDen(sMat)

    AtomDenConInMat = cMat.FractionConInMat(sCon, Atom) * dblAtomDen

    Set cMat = Nothing

End Function

