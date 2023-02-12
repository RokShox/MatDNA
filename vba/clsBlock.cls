VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsBlock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' Copyright (c) 2017 Henkel Technical Computing, LLC. All rights reserved.
'
Option Explicit

Dim mdParameters As Scripting.Dictionary
Dim mrngBlock As Range
Dim miParameterOffset As Integer

Private Sub Class_Initialize()
    Set mdParameters = New Scripting.Dictionary
    miParameterOffset = -2
End Sub

Property Get BlockRange() As Range
    Set BlockRange = mrngBlock
End Property

Property Set BlockRange(rng As Range)
    Set mrngBlock = rng
End Property

Property Get ParameterOffset() As Integer
    ParameterOffset = miParameterOffset
End Property

Property Let ParameterOffset(iVal As Integer)
    If iVal > 0 Then
        iVal = -iVal
    End If
    miParameterOffset = iVal
End Property

Property Get Parameter(s As String) As Variant
    If Not mdParameters.Exists(s) Then
        MsgBox "No parameter " & s & " in block", vbCritical
        End
    End If
    Parameter = mdParameters(s)
End Property

Property Let Parameter(s As String, vnt As Variant)
    If mdParameters.Exists(s) Then
        mdParameters.Remove s
    End If
    mdParameters.Add s, vnt
End Property

Property Get Universe() As Long
    Dim sKey As String
    
    sKey = "universe"
    If mdParameters.Exists(sKey) Then
        Universe = CLng(Me.Parameter(sKey))
    Else
        MsgBox "No parameter " & sKey & "defined in this block", vbCritical
        End
    End If
End Property

Property Get LastCell() As Long
    Dim sKey As String
    
    sKey = "lastCell"
    If mdParameters.Exists(sKey) Then
        LastCell = CLng(Me.Parameter(sKey))
    Else
        MsgBox "No parameter " & sKey & "defined in this block", vbCritical
        End
    End If
End Property

Property Get FirstCell() As Long
    Dim sKey As String
    
    sKey = "firstCell"
    If mdParameters.Exists(sKey) Then
        FirstCell = CLng(Me.Parameter(sKey))
    Else
        MsgBox "No parameter " & sKey & "defined in this block", vbCritical
        End
    End If
End Property

Public Sub SetInputParameters()
    Dim rngStart As Range
    Dim imax As Integer
    Const MaxRow As Integer = 100
    Dim sKey As String
    
    Set rngStart = mrngBlock.Cells(1, 1).Offset(0, miParameterOffset)
    imax = 0
    
    While rngStart.Text <> "Input Parameters" And imax < MaxRow
        Set rngStart = rngStart.Offset(1, 0)
        imax = imax + 1
    Wend
    
    If imax = MaxRow Then
        MsgBox "Input parameters not found", vbCritical
        End
    End If
    
    Set rngStart = rngStart.Offset(1, 0)
    imax = 0
    While rngStart.Text <> "" And imax < MaxRow
        sKey = rngStart.Text
        If Not mdParameters.Exists(sKey) Then
            MsgBox "No parameter " & sKey & " has been defined", vbCritical
            End
        Else
            rngStart.Offset(0, 1).Value2 = mdParameters(sKey)
        End If
        imax = imax + 1
        Set rngStart = rngStart.Offset(1, 0)
    Wend
    
    mrngBlock.Calculate
End Sub


Public Sub GetOutputParameters()
    Dim rngStart As Range
    Dim imax As Integer
    Const MaxRow As Integer = 100
    Dim sKey As String
    
    Set rngStart = mrngBlock.Cells(1, 1).Offset(0, miParameterOffset)
    imax = 0
    
    While rngStart.Text <> "Output Parameters" And imax < MaxRow
        Set rngStart = rngStart.Offset(1, 0)
        imax = imax + 1
    Wend
    
    If imax = MaxRow Then
        MsgBox "Output parameters not found", vbCritical
        End
    End If
    
    Set rngStart = rngStart.Offset(1, 0)
    imax = 0
    While rngStart.Text <> "" And imax < MaxRow
        sKey = rngStart.Text
        If mdParameters.Exists(sKey) Then
            mdParameters.Remove sKey
        End If
        mdParameters.Add sKey, rngStart.Offset(0, 1).Value2
        imax = imax + 1
        Set rngStart = rngStart.Offset(1, 0)
    Wend
End Sub

Public Function Display() As String
    Dim sText As String
    Dim s As Variant
    
    sText = "Parameters: " & vbCrLf
    For Each s In mdParameters.Keys
        sText = sText & DisplayString(CStr(s), mdParameters(s))
    Next
    Display = sText
End Function

Public Function WriteBlock(iFile As Integer)
    Dim colPatSubst As Collection
    
    Set colPatSubst = Nothing
    writeMCNPRangeToFile mrngBlock, iFile, colPatSubst
End Function

Public Function WriteBlockExcel(rngStart As Range)
    Dim colPatSubst As Collection
    
    Set colPatSubst = Nothing
    writeMCNPRangeToExcel Me.BlockRange, rngStart, colPatSubst
End Function

Public Function WriteBlockText(iFile As Integer)
    Dim iRow As Integer
    Dim iCol As Integer
    Dim sLine As String
    
    For iRow = 1 To Me.BlockRange.Rows.Count
        sLine = ""
        For iCol = 1 To Me.BlockRange.Columns.Count
            If iCol > 1 Then
                sLine = sLine & " "
            End If
            sLine = sLine & Me.BlockRange.Cells(iRow, iCol).Text
        Next iCol
        Print #iFile, sLine
    Next iRow
End Function

Public Sub CopyBlockAsValues(rngStart As Range)
    mrngBlock.Copy
    rngStart.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    rngStart.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Set rngStart = rngStart.Offset(mrngBlock.Rows.Count, 0)

End Sub
