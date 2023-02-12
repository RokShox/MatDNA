VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsParameter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' Copyright (c) 2017 Henkel Technical Computing, LLC. All rights reserved.
'
Option Explicit

Dim mrngTarget As Range
Dim mrngVariation As Range
Dim mdblPrevious As Double
Dim mdblCurrent As Double
Dim menuVariationType As enumSenVariationType
Dim msDescription As String
Const BADDBL As Double = -1E-99

Property Get TargetRange() As Range
    Set TargetRange = mrngTarget
End Property

Property Set TargetRange(rng As Range)
    Set mrngTarget = rng
End Property

Property Get Previous() As Double
    Previous = mdblPrevious
End Property

Property Get Current() As Double
    Select Case menuVariationType
    Case enumSenVariationType.FixedValue
        Current = Me.VariationParameter
    Case enumSenVariationType.Additive
        Current = mdblPrevious + Me.VariationParameter
    Case enumSenVariationType.Multiplicative
        Current = mdblPrevious * Me.VariationParameter
    Case enumSenVariationType.PercentChange
        Current = mdblPrevious * (1# + Me.VariationParameter)
    Case Else
        MsgBox "Invalid variation type"
        End
    End Select
End Property

Property Get VariationType() As enumSenVariationType
    VariationType = menuVariationType
End Property

Property Let VariationType(enu As enumSenVariationType)
    menuVariationType = enu
End Property

Property Get VariationRange() As Range
    Set VariationRange = mrngVariation
End Property

Property Set VariationRange(rng As Range)
    Set mrngVariation = rng
End Property

Property Get VariationParameter() As Double
    VariationParameter = CDbl(mrngVariation.Value2)
End Property

Property Get VariationText() As String
    VariationText = mrngVariation.Text
End Property

Property Get Description() As String
    Dim sDesc As String
    sDesc = msDescription
    If menuVariationType = Additive Then
            sDesc = sDesc & " + " & Me.VariationText
    ElseIf menuVariationType = FixedValue Then
        sDesc = sDesc & " = " & Me.VariationText
    ElseIf menuVariationType = Multiplicative Then
        sDesc = sDesc & " " & ChrW(&HD7) & Me.VariationText
    ElseIf menuVariationType = PercentChange Then
            sDesc = sDesc & " + " & Me.VariationText
    End If
    Description = sDesc
End Property

Property Let Description(sDesc As String)
    msDescription = sDesc
End Property


Sub SetTargetRange(srng As String)
    Dim rng As Range
    
    Set rng = Range(srng)
    If rng Is Nothing Then
        MsgBox "Error: Unable to grok " & srng & " as a range"
        End
    Else
        If rng.HasFormula Then
            MsgBox "Desired range " & srng & " contains a formula"
            End
        End If
        Set mrngTarget = rng
        If mdblPrevious <> BADDBL Then
            MsgBox "Previous value already stored. Attempt to overwrite"
            End
        Else
            mdblPrevious = CDbl(mrngTarget.Value2)
        End If
    End If
    
End Sub

Sub Restore()
    mrngTarget.Value2 = mdblPrevious
End Sub

Sub SetRange()
    mrngTarget.Value2 = Me.Current
End Sub

Private Sub Class_Initialize()
    Set mrngTarget = Nothing
    mdblPrevious = BADDBL
End Sub

Private Sub Class_Terminate()
    If CDbl(mrngTarget.Value2) <> mdblPrevious Then
        Debug.Print "Restoring paramater value on terminate"
        Me.Restore
    End If
End Sub
