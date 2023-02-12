VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLexToken"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' Copyright (c) 2017 Henkel Technical Computing, LLC. All rights reserved.
'
Option Explicit

Private menuDataType As LexicalDataType
Private msVBFormat As String
Private mlLength As Long
Private mblnPadLeft As Boolean

Property Get DataType() As LexicalDataType
    DataType = menuDataType
End Property

Property Let DataType(enu As LexicalDataType)
    menuDataType = enu
    ' Default padding for strings is right
    If menuDataType = LDT_String Then
        Me.PadLeft = False
    End If
End Property

Property Get Length() As Long
    Length = mlLength
End Property

Property Let Length(l As Long)
    mlLength = l
End Property

Property Get VBFormat() As String
    VBFormat = msVBFormat
End Property

Property Let VBFormat(s As String)
    msVBFormat = s
End Property

Property Get PadLeft() As Boolean
    PadLeft = mblnPadLeft
End Property

Property Let PadLeft(bln As Boolean)
    mblnPadLeft = bln
End Property

' Convert string to value

Public Function ValueAsLong(s As String) As Long
    checkFieldLength s
    ValueAsLong = CLng(Trim(s))
End Function

Public Function ValueAsDouble(s As String) As Double
    checkFieldLength s
    ValueAsDouble = CDbl(Trim(s))
End Function

Public Function ValueAsString(s As String) As String
    checkFieldLength s
    ValueAsString = Trim(s)
End Function

Private Sub checkFieldLength(s As String)
    If Len(s) <> mlLength Then
        MsgBox "Expected field of length " & mlLength & " got " & s, vbCritical
        End
    End If
End Sub

' Convert value to formatted string

Public Function FieldOfDouble(dbl As Double) As String
    Dim s As String
    s = Format(dbl, msVBFormat)
    s = padField(s)
    FieldOfDouble = s
End Function

Public Function FieldOfLong(l As Long) As String
    Dim s As String
    s = Format(l, msVBFormat)
    s = padField(s)
    FieldOfLong = s
End Function

Public Function FieldOfInteger(i As Integer) As String
    Dim s As String
    s = Format(i, msVBFormat)
    s = padField(s)
    FieldOfInteger = s
End Function

Public Function FieldOfString(sIn As String) As String
    Dim s As String
    s = Trim(Format(sIn, msVBFormat))
    s = padField(s)
    FieldOfString = s
End Function

Private Function padField(s As String) As String
    While Len(s) < mlLength
        If mblnPadLeft Then
            s = " " & s
        Else
            s = s & " "
        End If
    Wend
    padField = s
End Function

Private Sub Class_Initialize()
    mblnPadLeft = True
End Sub
