Attribute VB_Name = "mdlGlobal"
'
' Copyright (c) 2017 Henkel Technical Computing, LLC. All rights reserved.
'
Option Explicit

Public gblCalcState As XlCalculationState
Public Const PI As Double = 3.14159265358979

Public Enum enumSenVariationType
    FixedValue = 0
    Additive = 1
    Multiplicative = 2
    PercentChange = 3
End Enum

Public gblSurface As Scripting.Dictionary


'
' Materials
'
Public gblMatUtil As clsMatUtil
Public gblMatMaster As clsMatMaster

' Stanza offsets
Public Const or_Stanza_Comp = 2
Public Const oc_Stanza_Comp = 16

' Composition table offsets
' Row offsets rel to table base
Public Const or_Comp_Data As Integer = 8
' Column offsets rel to table base
Public Const oc_Comp_MatRows As Integer = 0
Public Const oc_Comp_ComRows As Integer = 1
Public Const oc_Comp_ConRows As Integer = 2
Public Const oc_Comp_Com As Integer = 3
Public Const oc_Comp_Con As Integer = 4
Public Const oc_Comp_Iso As Integer = 5
Public Const oc_Comp_AValue As Integer = 6
Public Const oc_Comp_IsoMfrac As Integer = 7
Public Const oc_Comp_IsoAfrac As Integer = 8
Public Const oc_Comp_ConMfrac As Integer = 9
Public Const oc_Comp_ConAfrac As Integer = 10
Public Const oc_Comp_ComMfrac As Integer = 11
Public Const oc_Comp_ComAfrac As Integer = 12

Public Const DIST_TOL As Double = 0.0000000001

'
'
'
Public Enum CompositionMode
    Mass
    Atom
End Enum


Public Enum MatFracCopyMode
    ByValue
    ByReference
    ByFormula
End Enum

Public Sub InitMaterials(Optional blnForce As Boolean = False)
    If blnForce Or (gblMatUtil Is Nothing) Then
        Set gblMatUtil = New clsMatUtil
    End If
    If blnForce Or (gblMatMaster Is Nothing) Then
        Set gblMatMaster = New clsMatMaster
    End If
End Sub
