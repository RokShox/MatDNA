VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLexOp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' Copyright (c) 2017 Henkel Technical Computing, LLC. All rights reserved.
'
Option Explicit

Private Const MAXTOKEN As Long = 50
Private moaLexToken(MAXTOKEN) As clsLexToken
Private msaField(MAXTOKEN) As String
Private mnVal As Long
Private mnToken As Long
Private mnLineLength As Long
Private mblnIsFormatted As Boolean


Public Enum LexicalDataType
    LDT_Integer = 0
    LDT_Float = 1
    LDT_Scientific = 2
    LDT_String = 3
    LDT_Space = 4
End Enum

' Lexical operations
' FormatString is a sequence of comma delimited Fortran-style format codes
' The format codes may include a prefixing repeat parameter or may themselves be
' parenthesized lists of format codes.

' Use:
' First call SetFormat to define the expected format. The number of recognized format codes (as expanded
' by any repeats) is the total number of fields that can be present (mnToken) in a line of text. mnLineLength is
' the expected total line length of the described data. Any additional characters will be ignored in parsing.
'
' Then call Parse with a line of text in the corresponding format. The number of values obtained is
' mnVal. The entire line does not have to be populated with data, but the data fields must be continuously
' occupied from the beginning of the line.


Public Sub SetFormat(sFmt As String)
    Me.ClearValues
    parseFormatString sFmt, mnToken, mnLineLength
    mblnIsFormatted = True
End Sub

Public Sub ResetFormat()
    Me.ClearValues
    Erase moaLexToken
    mblnIsFormatted = False
End Sub

Public Sub ClearValues()
    Erase msaField
    mnVal = 0
End Sub

Public Sub Parse(sLine As String, nVal As Long)
    Dim iField As Long
    Dim sField As String
    Dim iStart As Long
    Dim sTmp As String
    Dim cToken As clsLexToken
    Dim nLineLength As Long
    
    If Not Me.IsFormatted Then
        MsgBox "Must call SetFormat first", vbCritical
        End
    End If
    
    Me.ClearValues
    iStart = 1
    iField = 0
    nLineLength = Len(sLine)
    While iStart < IIf(nLineLength < mnLineLength, nLineLength, mnLineLength)
        Set cToken = moaLexToken(iField)
        sField = Mid(sLine, iStart, cToken.Length)
'        Debug.Print "Field " & iField & ": " & sField
        msaField(iField) = sField
        iStart = iStart + cToken.Length
        
        iField = iField + 1
        If iField > mnToken Then
            MsgBox "Expected only " & mnToken & " fields", vbCritical
            End
        End If
    Wend
    nVal = iField
    mnVal = iField
End Sub

Public Function Compose() As String
    Dim lIndex As Long
    Dim sLine As String
    Dim cToken As clsLexToken
    
    sLine = ""
    For lIndex = 0 To mnVal - 1
        If moaLexToken(lIndex).DataType = LDT_Space Then
            sLine = sLine & moaLexToken(lIndex).FieldOfString(" ")
        Else
            sLine = sLine & msaField(lIndex)
        End If
    Next
    Compose = sLine & vbCrLf
End Function

Property Get IsFormatted() As Boolean
    IsFormatted = mblnIsFormatted
End Property

Property Get NumVal() As Long
    NumVal = mnVal
End Property

Property Get NumToken() As Long
    NumToken = mnToken
End Property

Property Get ValueAsLong(lIndex As Long) As Long
    checkReadIndex lIndex
    ValueAsLong = moaLexToken(lIndex).ValueAsLong(msaField(lIndex))
End Property

Property Let ValueAsLong(lIndex As Long, lValue As Long)
    checkWriteIndex lIndex
    msaField(lIndex) = moaLexToken(lIndex).FieldOfLong(lValue)
End Property

Property Get ValueAsDouble(lIndex As Long) As Double
    checkReadIndex lIndex
    ValueAsDouble = moaLexToken(lIndex).ValueAsDouble(msaField(lIndex))
End Property

Property Let ValueAsDouble(lIndex As Long, dblValue As Double)
    checkWriteIndex lIndex
    msaField(lIndex) = moaLexToken(lIndex).FieldOfDouble(dblValue)
End Property

Property Get ValueAsString(lIndex As Long) As String
    checkReadIndex lIndex
    ValueAsString = moaLexToken(lIndex).ValueAsString(msaField(lIndex))
End Property

Property Let ValueAsString(lIndex As Long, sValue As String)
    checkWriteIndex lIndex
    msaField(lIndex) = moaLexToken(lIndex).FieldOfString(sValue)
End Property

Property Get LexToken(lIndex As Long) As clsLexToken
    If Not Me.IsFormatted Then
        Set LexToken = Nothing
    Else
        Set LexToken = moaLexToken(lIndex)
    End If
End Property

Private Sub checkWriteIndex(lIndex As Long)
    Dim l As Long
    
    If lIndex > mnToken - 1 Then
        MsgBox "Invalid value index: " & lIndex & " > " & mnToken - 1, vbCritical
        End
    ElseIf lIndex < 0 Then
        MsgBox "Invalid value index: " & lIndex & " < 0", vbCritical
        End
    ElseIf lIndex > MAXTOKEN Then
        MsgBox "Invalid value index: " & lIndex & " > MAXTOKEN", vbCritical
        End
    ElseIf lIndex > mnVal - 1 Then
        For l = mnVal To lIndex
            msaField(l) = ""
        Next l
        mnVal = lIndex + 1
    End If
End Sub

Private Sub checkReadIndex(lIndex As Long)
    If lIndex < 0 Then
        MsgBox "Invalid value index for read: " & lIndex, vbCritical
        End
    ElseIf lIndex >= mnVal Then
        MsgBox "Invalid value index for read: " & lIndex, vbCritical
        End
    End If
End Sub

' Expands a string containing comma-delimited format codes of the form "\d?(?([aix]\d+|[de][\d.]+)))?"
' into a single list of non-repeated format codes
'
' Example: (4a10,3(i5,2e10.4))
' Becomes: (a10,a10,a10,a10,i5,e10.4,e10.4,i5,e10.4,e10.4,i5,e10.4,e10.4)
'
' The expanded format string is then parsed into arrays describing field data type, length, and required
' VBA format description
'
Private Sub parseFormatString(ByVal sFmt As String, ByRef nToken As Long, ByRef lTotalLength As Long)
    Dim re As RegExp
    Dim mS As MatchCollection
    Dim M As Match
    Dim enu As LexicalDataType
    Dim lLen As Long
    Dim sVB As String
    Dim iToken As Long
    Dim cToken As clsLexToken
    
    Set re = New RegExp
    re.Pattern = "[^,]+"
    re.Global = True
    
    sFmt = LCase(Trim(sFmt))
    If Left(sFmt, 1) = "(" Then sFmt = Mid(sFmt, 2)
    If Right(sFmt, 1) = ")" Then sFmt = Mid(sFmt, 1, Len(sFmt) - 1)
    
    sFmt = removeChar(sFmt, " ")
    sFmt = expandRepeats(sFmt)
    Set mS = re.Execute(sFmt)
'    Debug.Print "Found " & ms.Count & " tokens"
    nToken = mS.Count
    
    lTotalLength = 0
    For iToken = 0 To nToken - 1
        Set M = mS.Item(iToken)
        parseFormatCode M.Value, enu, lLen, sVB
        Set cToken = New clsLexToken
        With cToken
            .DataType = enu
            .Length = lLen
            .VBFormat = sVB
        End With
        Set moaLexToken(iToken) = cToken
        lTotalLength = lTotalLength + lLen
    Next iToken
    
End Sub

Private Sub parseFormatCode(sFmt As String, ByRef enuType As LexicalDataType, ByRef lLength As Long, ByRef sVBFormat As String)
    Dim lDot As Long
    Dim lPrec As Long

    sFmt = LCase(sFmt)
    Select Case Left(sFmt, 1)
    Case "a"
        enuType = LDT_String
        lLength = CLng(Mid(sFmt, 2))
        sVBFormat = String(lLength, "@")
    Case "i"
        enuType = LDT_Integer
        lLength = CLng(Mid(sFmt, 2))
        sVBFormat = String(lLength - 1, "#") & "0"
    Case "f"
        enuType = LDT_Float
        lDot = InStr(1, sFmt, ".")
        If lDot = 0 Then
            lLength = CLng(Mid(sFmt, 2))
            lPrec = 0
        Else
            lLength = CLng(Mid(sFmt, 2, lDot - 2))
            lPrec = CLng(Mid(sFmt, lDot + 1))
        End If
        sVBFormat = String(lLength - lPrec - 2, "#") & "0." & String(lPrec, "0") & ";" & _
                    String(lLength - lPrec - 3, "#") & "0." & String(lPrec, "0")
    Case "d", "e"
        enuType = LDT_Scientific
        lDot = InStr(1, sFmt, ".")
        If lDot = 0 Then
            lLength = CLng(Mid(sFmt, 2))
            lPrec = 0
        Else
            lLength = CLng(Mid(sFmt, 2, lDot - 2))
            lPrec = CLng(Mid(sFmt, lDot + 1))
        End If
        sVBFormat = "0." & String(lPrec, "0") & "E+" & "00"
    Case "x"
        enuType = LDT_Space
        lLength = CLng(Mid(sFmt, 2))
        sVBFormat = String(lLength, "@")
    Case Else
        MsgBox "Unrecognized type key: " & Left(sFmt, 1), vbCritical
        End
    End Select
End Sub

Private Function removeChar(ByVal sText As String, ByVal sChar As String) As String
    Dim re As RegExp
    
    sChar = Left(sChar, 1)
    
    Set re = New RegExp
    re.Pattern = sChar
    re.IgnoreCase = True
    re.Global = True
    removeChar = re.Replace(sText, "")
End Function

Private Function expandRepeats(sFmt As String) As String
    Dim re As RegExp
    Dim re2 As RegExp
    Dim mS As MatchCollection
    Dim M As Match
    Dim s As String
    Dim l As Long
    Dim lRep As Long
    Dim nRep As Long
    Dim nMatch As Long
    Dim sTmp As String
    
    ' Parenthesized repeats
    Set re = New RegExp
    re.Pattern = "(\d+)\(([^()]+)\)"
    re.Global = False
    
    Set mS = re.Execute(sFmt)
    nMatch = mS.Count
    
    While nMatch > 0
        For Each M In mS
'            Debug.Print m.Value & " at " & m.FirstIndex & " len " & m.Length & " with " & m.SubMatches.Count & " submatches"
'            For l = 0 To m.SubMatches.Count - 1
'                Debug.Print vbTab & m.SubMatches.Item(l)
'            Next l
            
            nRep = CLng(M.SubMatches.Item(0))
            sTmp = ""
            For lRep = 0 To nRep - 1
                sTmp = sTmp & M.SubMatches.Item(1) & IIf(lRep < nRep - 1, ",", "")
            Next lRep
            sFmt = re.Replace(sFmt, sTmp)
        Next
        Set mS = Nothing
        Set mS = re.Execute(sFmt)
        nMatch = mS.Count
    Wend
    
    Set re2 = New RegExp
    re2.Pattern = "(\d+)([adefix][\d.]+)"
    re2.Global = False
    re2.IgnoreCase = True
    
    Set mS = re2.Execute(sFmt)
    nMatch = mS.Count
    
    While nMatch > 0
        For Each M In mS
            nRep = CLng(M.SubMatches.Item(0))
            sTmp = ""
            For lRep = 0 To nRep - 1
                sTmp = sTmp & M.SubMatches.Item(1) & IIf(lRep < nRep - 1, ",", "")
            Next lRep
            sFmt = re2.Replace(sFmt, sTmp)
        Next
        Set mS = re2.Execute(sFmt)
        nMatch = mS.Count
    Wend
    expandRepeats = sFmt
End Function

Public Sub Test()
    Dim sFmt As String
    Dim sLine As String
    Dim nToken As Long
    Dim s As String
    Dim enu As LexicalDataType
    Dim sVB As String
    Dim lLen As Long
    Dim dbl As Double
    Dim nVal As Long
    Dim lTot As Long
    
    
'    sFmt = "(4a10,3(i5,2e10.4))"
'    parseFormatString sFmt, nToken, lTot
    
'    sFmt = removeChar(sFmt, " ")
'    Debug.Print sFmt
'    sFmt = expandRepeats(sFmt)
'    Debug.Print sFmt
    
    sFmt = "(2A8,A19,I5,I13,I15)"
    sLine = "mcnpx      2.7.0  03/19/12 03:15:41    9     40476839    64558280015    "
'
'
    Debug.Print sLine
    Me.SetFormat sFmt
    Me.Parse sLine, nVal
    Me.ValueAsLong(3) = 10
    Me.ValueAsDouble(4) = 1.234
'    Debug.Print Me.ValueAsString(0)
'    Debug.Print Me.ValueAsLong(4)
    Debug.Print Me.Compose
'    Debug.Print "Found " & nVal & " values in line"
'    Debug.Print "Code: " & Me.ValueAsString(0)
'    Debug.Print "Version: " & Me.ValueAsString(1)
'    Debug.Print "RunDate: " & Me.ValueAsString(2)
'    Debug.Print "Dump: " & Me.ValueAsLong(3)
'    Debug.Print "NPS: " & Me.ValueAsLong(4)
'    Debug.Print "RNR: " & Me.ValueAsString(5)
End Sub

Private Sub Class_Initialize()
    mblnIsFormatted = False
End Sub
