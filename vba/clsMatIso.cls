VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMatIso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' Copyright (c) 2022 Henkel Technical Computing, LLC. All rights reserved.
'
Option Explicit
Option Compare Text

Dim mdblMolWgt As Double
Dim msIso As String
Dim msThisClass As String

Private Sub Class_Initialize()
    InitMaterials
    msThisClass = "clsMatIso"

    mdblMolWgt = 0#
End Sub

Property Get Iso() As String
    Iso = msIso
End Property

Property Let Iso(s As String)
    msIso = LCase(s)
    mdblMolWgt = gblMatUtil.MolWgt(msIso)
End Property

Property Get MolecularWeight() As Double
    MolecularWeight = mdblMolWgt
End Property

