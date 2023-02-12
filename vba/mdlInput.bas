Attribute VB_Name = "mdlInput"
'
' Copyright (c) 2017 Henkel Technical Computing, LLC. All rights reserved.
'
Option Explicit
'
' doPatSubst
'
Public Sub doPatSubst(sLine As String, sOut As String, colPatSubst As Collection)
    Const sPercent As String = "%"
    
    Dim i As Integer
    Dim sToken As String
    Dim sValue As String
    Dim sChar As String
    Dim sNext As String
    Dim blnInPattern As Boolean
    
    If colPatSubst Is Nothing Then
        ' Nothing to do
        sOut = sLine
        Exit Sub
    End If
    
'
'   March through each character in line. Four cases:
'       1) Initial "%" in an opening "%%" token
'       2) Initial "%" in a closing "%%" token
'       3) An innocent "%"
'       4) Not a "%" at all
'
    On Error Resume Next

    blnInPattern = False
    sOut = ""
    sToken = ""
    i = 1
    While i <= Len(sLine)
        sChar = Mid(sLine, i, 1)
        ' If "%", peek at next char
        If sChar = sPercent Then
            sNext = Mid(sLine, i + 1, 1)
            ' If next char is a "%" also, then we have a token match and we are either
            ' opening or closing a pattern token capture
            If sNext = sPercent Then
        
                If Not blnInPattern Then
                    ' Open pattern token capture
                    blnInPattern = True
                    ' Skip over next char
                    i = i + 1
       
                ElseIf blnInPattern Then
                    ' Make substitution
                    sValue = colPatSubst.Item(sToken)
                    If Err.Number <> 0 Then
                        MsgBox "No value for pattern " & sToken & " in line:" & Chr(10) & sLine, vbCritical
                        End
                    End If
                    sOut = sOut & sValue
                    ' Close token capture
                    sToken = ""
                    blnInPattern = False
                    ' Skip over next char
                    i = i + 1
                End If
            Else
                ' Just an innocent "%"
                sOut = sOut & sChar
            End If
            
        ElseIf blnInPattern Then
            ' Capture token text
            sToken = sToken & sChar
        Else
            ' Capture other line text
            sOut = sOut & sChar
        End If
        i = i + 1
    Wend
    
    ' Check for mismatched patterns tokens
    If blnInPattern Then
        MsgBox "Error: mismatched %% tokens", vbCritical
        End
    End If
End Sub

'
' doPatSubstExcel
'
Public Sub doPatSubstExcel(rngStart As Range, nCol As Long, colPatSubst As Collection)
    Const sPercent As String = "%"
    
    Dim i As Integer
    Dim c As Long
    Dim sToken As String
    Dim sValue As String
    Dim sChar As String
    Dim sNext As String
    Dim sText As String
    Dim sOut As String
    Dim blnInPattern As Boolean
    Dim blnSubstitutionMade As Boolean
    
    If colPatSubst Is Nothing Then
        ' Nothing to do
        Exit Sub
    End If
    
'
'   March through each character in each cell. Four cases:
'       1) Initial "%" in an opening "%%" token
'       2) Initial "%" in a closing "%%" token
'       3) An innocent "%"
'       4) Not a "%" at all
'
    On Error Resume Next

    blnInPattern = False
    sOut = ""
    sToken = ""
    
    For c = 0 To nCol - 1
        blnSubstitutionMade = False
        sText = rngStart.Offset(0, c).Text
        i = 1
        While i <= Len(sText)
            sChar = Mid(sText, i, 1)
            ' If "%", peek at next char
            If sChar = sPercent Then
                sNext = Mid(sText, i + 1, 1)
                ' If next char is a "%" also, then we have a token match and we are either
                ' opening or closing a pattern token capture
                If sNext = sPercent Then
            
                    If Not blnInPattern Then
                        ' Open pattern token capture
                        blnInPattern = True
                        ' Skip over next char
                        i = i + 1
           
                    ElseIf blnInPattern Then
                        ' Make substitution
                        sValue = colPatSubst.Item(sToken)
                        If Err.Number <> 0 Then
                            MsgBox "No value for pattern " & sToken & " in line:" & Chr(10) & sText, vbCritical
                            End
                        End If
                        sOut = sOut & sValue
                        ' Close token capture
                        sToken = ""
                        blnInPattern = False
                        blnSubstitutionMade = True
                        ' Skip over next char
                        i = i + 1
                    End If
                Else
                    ' Just an innocent "%"
                    sOut = sOut & sChar
                End If
                
            ElseIf blnInPattern Then
                ' Capture token text
                sToken = sToken & sChar
            Else
                ' Capture other line text
                sOut = sOut & sChar
            End If
            i = i + 1
            
        Wend
        
        ' Check for mismatched patterns tokens
        If blnInPattern Then
            MsgBox "Error: mismatched %% tokens", vbCritical
            End
        End If
        
        ' If a substitution was made, change the cell value
        If blnSubstitutionMade Then
            ' A change was made
            rngStart.Offset(0, c).Value2 = sOut
        End If
    Next c
End Sub



'
' parseInsert
'
Public Sub parseInsert(sLine As String, rngInsert As Range, colPatSubst As Collection, wksCurrent As Worksheet)
    Const sLParen As String = "("
    Const sRParen As String = ")"
    Const sComma As String = ","
    Const sEqual As String = "="
    Const sPercent As String = "%"
    Const sPound As String = "#"
    Dim sSpace As String
    Dim sTab As String
    
    Dim iBgn As Integer
    Dim i As Integer
    Dim iEnd As Integer
    Dim sTemp As String
    Dim sChar As String
    Dim sToken As String
    Dim nPat As Integer
    Dim blnInPattern As Boolean
    Dim blnInValue  As Boolean
    Dim sPattern As String
    Dim sValue As String
    Dim sInsert As String
    
    sSpace = Chr(32)
    sTab = Chr(9)
    
    If Left(sLine, 1) <> sPound Then
        MsgBox "Invalid input, expected first char = #", vbCritical
        End
    End If
    
    '
    ' First resolve the range name
    '
    
    ' Chop off leading "#"
    sTemp = Trim(Mid(sLine, 2))
    iBgn = InStr(1, sTemp, sLParen, vbTextCompare)
    
    If iBgn = 0 Then
        ' No pat subst required
        sInsert = sTemp
    Else
        sInsert = Trim(Left(sTemp, iBgn - 1))
        
    End If
    
    If Left(sInsert, 1) = "!" Then
        ' Local sheet name
        sInsert = Mid(sInsert, 2)
        Set rngInsert = wksCurrent.Names(sInsert).RefersToRange
    Else
        ' Global name
        Set rngInsert = Names(sInsert).RefersToRange
    End If
    
    If rngInsert Is Nothing Then
        ' No match found for name
        MsgBox "Error: Unrecognized range " & sInsert
        End
    End If
   
    '
    ' Second, look for pattern substitutions
    '
    
    If iBgn = 0 Then
        ' No pattern substitution
        Set colPatSubst = Nothing
        Exit Sub
    End If
    
    iEnd = Len(sTemp)
    If Right(sTemp, 1) <> sRParen Then
        MsgBox "Must end in right paren", vbCritical
        End
    End If
        
    i = iBgn + 1
    sToken = ""
    nPat = 0
    blnInPattern = True
    blnInValue = False
    Set colPatSubst = New Collection
    
    While i <= iEnd
        sChar = Mid(sTemp, i, 1)
        If sChar = sEqual Then
            If blnInValue Then
                MsgBox "Invalid pattern format", vbCritical
                End
            End If
            ' Pattern identified
            sPattern = sToken
            sToken = ""
            blnInPattern = False
            blnInValue = True
            
        ElseIf sChar = sComma Or sChar = sRParen Then
            If blnInPattern Then
                MsgBox "Invalid pattern format", vbCritical
                End
            End If
            ' Value identified
            sValue = Trim(sToken)
            sToken = ""
            blnInPattern = True
            blnInValue = False
            
            ' Add pat subst to collection
            colPatSubst.Add sValue, sPattern
        
        ElseIf Not blnInValue And (sChar = sSpace Or sChar = sTab) Then
            ' Skip blank chars unless in a value
     
        ElseIf sChar = sPercent Then
            MsgBox "Invalid character in pattern or value", vbCritical
            End
     
        Else
            sToken = sToken & sChar
            
        End If
        i = i + 1
    Wend
    
End Sub

