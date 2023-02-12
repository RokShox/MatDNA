VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMatUtil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' Copyright (c) 2022 Henkel Technical Computing, LLC. All rights reserved.
'
Option Explicit
Option Compare Text

Const UpperA As Integer = 65
Const UpperZ As Integer = 90
Const LowerA As Integer = 97
Const LowerZ As Integer = 122
Const MiddleDot As Integer = 183
Const CurlyLeft As Integer = 123
Const CurlyRight As Integer = 125
Const SquareLeft As Integer = 91
Const SquareRight As Integer = 93
Const ParenLeft As Integer = 40
Const ParenRight As Integer = 41
Const AngleLeft As Integer = 60
Const AngleRight As Integer = 62
Const Numeral0 As Integer = 48
Const Numeral9 As Integer = 57
Const Period As Integer = 46

Dim mdElmBySym As Scripting.Dictionary
Dim mdSymByZ As Scripting.Dictionary
Dim mrngTblAtomicMass As Range
Dim msThisClass As String

Private Sub Class_Initialize()

    msThisClass = "clsMatUtil"
    Set mrngTblAtomicMass = ThisWorkbook.Names("tblMCNP6NuclideMass").RefersToRange

    initDict


End Sub

Public Function MolWgt(sIso As String) As Double
    On Error GoTo Fail:
    MolWgt = Application.WorksheetFunction.VLookup(sIso, mrngTblAtomicMass, 4, False)
    Exit Function
    
Fail:
    MsgBox sIso & " not found in mass table", vbCritical, msThisClass
    MolWgt = 0#
End Function



Public Sub ParseFormula(sFormula As String, dStoich As Scripting.Dictionary, Optional iLevel As Integer = 0)
    Dim ic As Integer
    Dim c As String
    Const LookForSomething As Integer = 0
    Const LookForSymbol As Integer = 1
    Const LookForGroup As Integer = 2
    Const LookForNumber As Integer = 3
    Const PrevStateNull = 0
    Const PrevStateSymbol As Integer = 1
    Const PrevStateGroup As Integer = 2
    Dim sLastOpen As String
    Dim iLook As Integer
    Dim iPrevState As Integer
    Dim sNum As String
    Dim sSym As String
    Dim sGroup As String
    Dim dblNum As Double
    Dim icLast As Integer
    Dim vntKey As Variant
    Dim currentLevel As Integer
    Dim dStoichGroup As Scripting.Dictionary
    
    Debug.Print vbCrLf & "Testing " & sFormula & " at Level " & iLevel
    
    
    ' If recursive call, create a new dict
    If dStoich Is Nothing Then
        Set dStoich = New Scripting.Dictionary
        dStoich.CompareMode = TextCompare
    End If
    
    ' Four states:
    ' 0) Looking for anything
    ' 1) Looking for symbol
    ' 2) Collecting a group
    ' 3) Looking for number
    '
    ' If looking for numeral, may be applied to previous symbol or group
    ' Need to handle end of string
    currentLevel = iLevel
    iLook = LookForSomething
    icLast = Len(sFormula)
    ic = 1
    Do While ic <= icLast
        c = Mid(sFormula, ic, 1)
                
        ' No matter what
        If Me.IsMiddleDot(c) Then
            ' Worry about this later
            ic = ic + 1
            If ic > icLast Then
                Exit Do
            End If
        ElseIf iLook = LookForSomething Then
            sNum = ""
            sSym = ""
            sGroup = ""
            ' Found symbol
            If Me.IsUpper(c) Then
                iLook = LookForSymbol
                iPrevState = PrevStateSymbol
                sSym = c
                ic = ic + 1
                ' If last char, process
                If ic > icLast Then
                    If Me.IsSymbol(sSym) Then
                        sSym = LCase(sSym)
                        If dStoich.Exists(sSym) Then
                            dStoich(sSym) = dStoich(sSym) + 1#
                        Else
                            dStoich.Add sSym, 1#
                        End If
                    Else
                        MsgBox "Invalid symbol " & sSym, vbOKOnly, msThisClass
                    End If
                End If
                
            ' Found a grouping char
            ElseIf Me.IsGrouping(c) Then
                ' Opening group
                If Me.IsOpening(c) Then
                    iLook = LookForGroup
                    iPrevState = PrevStateGroup
                    currentLevel = currentLevel + 1
                    sLastOpen = c
                    ic = ic + 1
                    If ic > icLast Then
                        MsgBox "Invalid format", vbOKOnly, msThisClass
                        Exit Sub
                    End If
                ' Closing group (unexpected)
                ElseIf Me.IsClosing(c) Then
                    MsgBox "Invalid format", vbOKOnly, msThisClass
                    Exit Sub
                End If
            ' Expected either symbol or group to start
            Else
                MsgBox "Invalid format", vbOKOnly, msThisClass
                Exit Sub
            End If
        
        
        ElseIf iLook = LookForSymbol Then
            ' Already have first char, see if two letter symbol
            ' If so, advance
            If Me.IsLower(c) Then
                sSym = sSym & c
                ic = ic + 1
                ' Check if at end of string
                If ic > icLast Then
                    If Me.IsSymbol(sSym) Then
                        sSym = LCase(sSym)
                        If dStoich.Exists(sSym) Then
                            dStoich(sSym) = dStoich(sSym) + 1#
                        Else
                            dStoich.Add sSym, 1#
                        End If
                    Else
                        MsgBox "Invalid symbol " & sSym, vbOKOnly, msThisClass
                    End If
                End If
            End If
            
            If Not Me.IsSymbol(sSym) Then
                MsgBox "Invalid symbol " & sSym, vbOKOnly, msThisClass
            End If
            
            iLook = LookForNumber
                
        
        ElseIf iLook = LookForGroup Then
            ' Opening grouping char has already been found for current group
            ' Check if a grouping character closes the current group
            If Me.IsGrouping(c) Then
            
                If Me.IsOpening(c) Then
                    ' A subgroup...will be handled by recursion, but track current level
                    ' to know when current group ends
                    currentLevel = currentLevel + 1
                    sGroup = sGroup & c
                
                ElseIf Me.IsClosing(c) Then
                    ' A closing char. May be closing current group or a subgroup.
                    ' Closes current group if currentLevel is iLevel+1 and char is correct
                    currentLevel = currentLevel - 1
                    If sLastOpen = Me.Opening(c) And currentLevel = iLevel Then
                        ' Found group. Make recursive call
                        ' Defer processing the returned dictionary until we see if
                        ' there is a multiplier following the group
                        Set dStoichGroup = Nothing
                        Me.ParseFormula sGroup, dStoichGroup, iLevel + 1
                        iLook = LookForNumber
                        sLastOpen = ""
                        sGroup = ""
                    Else
                        ' closes a subgroup
                        sGroup = sGroup & c
                    End If
                Else
                    MsgBox "Shouldn't happen", vbOKOnly, msThisClass
                End If
            Else
                sGroup = sGroup & c
            End If
            
            ic = ic + 1
            ' Check for end of string with open group
            If ic > icLast Then
                If currentLevel <> iLevel Then
                    MsgBox "End of string with open group", vbOKOnly, msThisClass
                    Exit Sub
                Else
                    ' Need to process returned dictionary
                    ' Look through the dictionary returned from recursive call
                    ' and collect symbols and stoich
                    ' vntKey will already be lower case
                    For Each vntKey In dStoichGroup.Keys
                        Debug.Print "Level " & iLevel + 1 & vbTab & vntKey & vbTab & dStoichGroup(vntKey)
                        If dStoich.Exists(vntKey) Then
                            dStoich(vntKey) = dStoich(vntKey) + dblNum * dStoichGroup(vntKey)
                        Else
                            dStoich.Add vntKey, dblNum * dStoichGroup(vntKey)
                        End If
                    Next
                    Set dStoichGroup = Nothing
                End If
            End If
            
        ElseIf iLook = LookForNumber Then
            ' Looking for number and found one. Add to current number string and advance
            If Me.IsNumeral(c) Or Me.IsPeriod(c) Then
                sNum = sNum & c
                ic = ic + 1
                ' Check if at end of formula
                If ic > icLast Then
                
                    dblNum = CDbl(val(sNum))
                    
                    ' End of string while looking for symbol mult
                    If iPrevState = PrevStateSymbol Then
                        sSym = LCase(sSym)
                        If dStoich.Exists(sSym) Then
                            dStoich(sSym) = dStoich(sSym) + dblNum
                        Else
                            dStoich.Add sSym, dblNum
                        End If
                        
                    ' End of string while looking for group multiplier
                    ElseIf iPrevState = PrevStateGroup Then
                        ' Look through the dictionary returned from recursive call
                        ' and collect symbols and stoich
                        ' vntKey will already be lower case
                        For Each vntKey In dStoichGroup.Keys
                            Debug.Print "Level " & iLevel + 1 & vbTab & vntKey & vbTab & dStoichGroup(vntKey)
                            If dStoich.Exists(vntKey) Then
                                dStoich(vntKey) = dStoich(vntKey) + dblNum * dStoichGroup(vntKey)
                            Else
                                dStoich.Add vntKey, dblNum * dStoichGroup(vntKey)
                            End If
                        Next
                        Set dStoichGroup = Nothing
                    End If
                End If
            ' Looking for number and didn't find one
            Else
                ' If no numbers seen, implies mult of 1
                If sNum = "" Then
                    dblNum = 1#
                Else
                    dblNum = CDbl(val(sNum))
                End If
                
                If iPrevState = PrevStateSymbol Then
                    
                    sSym = LCase(sSym)
                    If dStoich.Exists(sSym) Then
                        dStoich(sSym) = dStoich(sSym) + dblNum
                    Else
                        dStoich.Add sSym, dblNum
                    End If
                    
                ElseIf iPrevState = PrevStateGroup Then
                    ' Look through the dictionary returned from recursive call
                    ' and collect symbols and stoich
                    ' vntKey will already be lower case
                    For Each vntKey In dStoichGroup.Keys
                        Debug.Print "Level " & iLevel + 1 & vbTab & vntKey & vbTab & dStoichGroup(vntKey)
                        If dStoich.Exists(vntKey) Then
                            dStoich(vntKey) = dStoich(vntKey) + dblNum * dStoichGroup(vntKey)
                        Else
                            dStoich.Add vntKey, dblNum * dStoichGroup(vntKey)
                        End If
                    Next
                    Set dStoichGroup = Nothing
                End If
                
                ' Don't advance char
                sNum = ""
                iPrevState = PrevStateNull
                iLook = LookForSomething
            End If
        
        Else
            MsgBox "Can't happen", vbOKOnly, msThisClass
            Exit Sub
        End If
        
    Loop
    
    Debug.Print "Exit level " & iLevel
    Debug.Print "Symbol" & vbTab & "Stoich"
    For Each vntKey In dStoich.Keys
        Debug.Print vntKey & vbTab & dStoich(vntKey)
    Next
    
End Sub

Public Static Function SymbolByZ(z As Integer) As String
    If mdSymByZ.Exists(z) Then
        SymbolByZ = mdSymByZ.Item(z)
    Else
        SymbolByZ = ""
    End If
End Function

Public Static Function ZBySymbol(s As String) As Integer
    Dim i As Integer
    For i = 0 To mdSymByZ.Count - 1
        ' Note carefully must index by Keys
        If LCase(s) = LCase(mdSymByZ(mdSymByZ.Keys(i))) Then
            ZBySymbol = mdSymByZ.Keys(i)
            Exit Function
        End If
    Next i
    ZBySymbol = 0
End Function

Public Static Function ElemenBySymbol(s As String) As String
    If mdElmBySym.Exists(s) Then
        ElemenBySymbol = mdElmBySym.Item(s)
    Else
        ElemenBySymbol = ""
    End If
End Function


Public Static Function SymbolByElement(s As String) As String
    Dim i As Integer
    For i = 0 To mdElmBySym.Count - 1
        ' Note carefully must index by Keys
        If LCase(s) = LCase(mdElmBySym(mdElmBySym.Keys(i))) Then
            SymbolByElement = mdElmBySym.Keys(i)
            Exit Function
        End If
    Next i
    SymbolByElement = 0
End Function

Public Static Function IsSymbol(s As String)
    IsSymbol = mdElmBySym.Exists(s)
End Function

Public Static Function IsUpper(c As String) As Boolean
    Dim iCode As Integer
    
    iCode = Asc(c)
    If iCode <= UpperZ And iCode >= UpperA Then
        IsUpper = True
    Else
        IsUpper = False
    End If
End Function

Public Static Function IsLower(c As String) As Boolean
    Dim iCode As Integer
    
    iCode = Asc(c)
    If iCode <= LowerZ And iCode >= LowerA Then
        IsLower = True
    Else
        IsLower = False
    End If
End Function

Public Static Function IsNumeral(c As String) As Boolean
    Dim iCode As Integer
    
    iCode = Asc(c)
    If iCode <= Numeral9 And iCode >= Numeral0 Then
        IsNumeral = True
    Else
        IsNumeral = False
    End If
End Function

Public Static Function IsMiddleDot(c As String) As Boolean
    Dim iCode As Integer
    
    iCode = Asc(c)
    If iCode = MiddleDot Then
        IsMiddleDot = True
    Else
        IsMiddleDot = False
    End If
End Function

Public Static Function IsPeriod(c As String) As Boolean
    Dim iCode As Integer
    
    iCode = Asc(c)
    If iCode = Period Then
        IsPeriod = True
    Else
        IsPeriod = False
    End If
End Function

Public Static Function IsGrouping(c As String) As Boolean
    Dim iCode As Integer
    
    iCode = Asc(c)
    Select Case iCode
    Case CurlyLeft
        IsGrouping = True
    Case CurlyRight
        IsGrouping = True
    Case SquareLeft
        IsGrouping = True
    Case SquareRight
        IsGrouping = True
    Case ParenLeft
        IsGrouping = True
    Case ParenRight
        IsGrouping = True
    Case AngleLeft
        IsGrouping = True
    Case AngleRight
        IsGrouping = True
    Case Else
        IsGrouping = False
    End Select

End Function

Public Static Function IsOpening(c As String) As Boolean
    Dim iCode As Integer
    
    iCode = Asc(c)
    Select Case iCode
    Case CurlyLeft
        IsOpening = True
    Case SquareLeft
        IsOpening = True
    Case ParenLeft
        IsOpening = True
    Case AngleLeft
        IsOpening = True
    Case Else
        IsOpening = False
    End Select

End Function

Public Static Function IsClosing(c As String) As Boolean
    Dim iCode As Integer
    
    iCode = Asc(c)
    Select Case iCode
    Case CurlyRight
        IsClosing = True
    Case SquareRight
        IsClosing = True
    Case ParenRight
        IsClosing = True
    Case AngleRight
        IsClosing = True
    Case Else
        IsClosing = False
    End Select
End Function


Public Static Function Closing(c As String) As String
    Dim iCode As Integer
    
    iCode = Asc(c)
    Select Case iCode
    Case CurlyLeft
        Closing = Chr(CurlyRight)
    Case SquareLeft
        Closing = Chr(SquareRight)
    Case ParenLeft
        Closing = Chr(ParenRight)
    Case AngleLeft
        Closing = Chr(AngleRight)
    Case Else
        Closing = ""
    End Select
End Function


Public Static Function Opening(c As String) As String
    Dim iCode As Integer
    
    iCode = Asc(c)
    Select Case iCode
    Case CurlyRight
        Opening = Chr(CurlyLeft)
    Case SquareRight
        Opening = Chr(SquareLeft)
    Case ParenRight
        Opening = Chr(ParenLeft)
    Case AngleRight
        Opening = Chr(AngleLeft)
    Case Else
        Opening = ""
    End Select
End Function

Private Sub initDict()

    Set mdElmBySym = New Scripting.Dictionary
    mdElmBySym.CompareMode = TextCompare
    With mdElmBySym
        .Add "h", "hydrogen"
        .Add "he", "helium"
        .Add "li", "lithium"
        .Add "be", "beryllium"
        .Add "b", "boron"
        .Add "c", "carbon"
        .Add "n", "nitrogen"
        .Add "o", "oxygen"
        .Add "f", "fluorine"
        .Add "ne", "neon"
        .Add "na", "sodium"
        .Add "mg", "magnesium"
        .Add "al", "aluminum"
        .Add "si", "silicon"
        .Add "p", "phosphorus"
        .Add "s", "sulfur"
        .Add "cl", "chlorine"
        .Add "ar", "argon"
        .Add "k", "potassium"
        .Add "ca", "calcium"
        .Add "sc", "scandium"
        .Add "ti", "titanium"
        .Add "v", "vanadium"
        .Add "cr", "chromium"
        .Add "mn", "manganese"
        .Add "fe", "iron"
        .Add "co", "cobalt"
        .Add "ni", "nickel"
        .Add "cu", "copper"
        .Add "zn", "zinc"
        .Add "ga", "gallium"
        .Add "ge", "germanium"
        .Add "as", "arsenic"
        .Add "se", "selenium"
        .Add "br", "bromine"
        .Add "kr", "krypton"
        .Add "rb", "rubidium"
        .Add "sr", "strontium"
        .Add "y", "yttrium"
        .Add "zr", "zirconium"
        .Add "nb", "niobium"
        .Add "mo", "molybdenum"
        .Add "tc", "technetium"
        .Add "ru", "ruthenium"
        .Add "rh", "rhodium"
        .Add "pd", "palladium"
        .Add "ag", "silver"
        .Add "cd", "cadmium"
        .Add "in", "indium"
        .Add "sn", "tin"
        .Add "sb", "antimony"
        .Add "te", "tellurium"
        .Add "i", "iodine"
        .Add "xe", "xenon"
        .Add "cs", "cesium"
        .Add "ba", "barium"
        .Add "la", "lanthanum"
        .Add "ce", "cerium"
        .Add "pr", "praseodymium"
        .Add "nd", "neodymium"
        .Add "pm", "promethium"
        .Add "sm", "samarium"
        .Add "eu", "europium"
        .Add "gd", "gadolinium"
        .Add "tb", "terbium"
        .Add "dy", "dysprosium"
        .Add "ho", "holmium"
        .Add "er", "erbium"
        .Add "tm", "thulium"
        .Add "yb", "ytterbium"
        .Add "lu", "lutetium"
        .Add "hf", "hafnium"
        .Add "ta", "tantalum"
        .Add "w", "tungsten"
        .Add "re", "rhenium"
        .Add "os", "osmium"
        .Add "ir", "iridium"
        .Add "pt", "platinum"
        .Add "au", "gold"
        .Add "hg", "mercury"
        .Add "tl", "thallium"
        .Add "pb", "lead"
        .Add "bi", "bismuth"
        .Add "po", "polonium"
        .Add "at", "astatine"
        .Add "rn", "radon"
        .Add "fr", "francium"
        .Add "ra", "radium"
        .Add "ac", "actinium"
        .Add "th", "thorium"
        .Add "pa", "protactinium"
        .Add "u", "uranium"
        .Add "np", "neptunium"
        .Add "pu", "plutonium"
        .Add "am", "americium"
        .Add "cm", "curium"
        .Add "bk", "berkelium"
        .Add "cf", "californium"
        .Add "es", "einsteinium"
        .Add "fm", "fermium"
        .Add "md", "mendelevium"
        .Add "no", "nobelium"
        .Add "lr", "lawrencium"
        .Add "rf", "rutherfordium"
        .Add "db", "dubnium"
        .Add "sg", "seaborgium"
        .Add "bh", "bohrium"
        .Add "hs", "hassium"
        .Add "mt", "meitnerium"
        .Add "ds", "darmstadtium"
        .Add "rg", "roentgenium"
        .Add "cn", "copernicium"
        .Add "uut", "ununtrium"
        .Add "fl", "flerovium"
        .Add "uup", "ununpentium"
        .Add "lv", "livermorium"
        .Add "uus", "ununseptium"
        .Add "uuo", "ununoctium"
    End With

    Set mdSymByZ = New Scripting.Dictionary
    mdSymByZ.CompareMode = TextCompare
    With mdSymByZ
        .Add 1, "h"
        .Add 2, "he"
        .Add 3, "li"
        .Add 4, "be"
        .Add 5, "b"
        .Add 6, "c"
        .Add 7, "n"
        .Add 8, "o"
        .Add 9, "f"
        .Add 10, "ne"
        .Add 11, "na"
        .Add 12, "mg"
        .Add 13, "al"
        .Add 14, "si"
        .Add 15, "p"
        .Add 16, "s"
        .Add 17, "cl"
        .Add 18, "ar"
        .Add 19, "k"
        .Add 20, "ca"
        .Add 21, "sc"
        .Add 22, "ti"
        .Add 23, "v"
        .Add 24, "cr"
        .Add 25, "mn"
        .Add 26, "fe"
        .Add 27, "co"
        .Add 28, "ni"
        .Add 29, "cu"
        .Add 30, "zn"
        .Add 31, "ga"
        .Add 32, "ge"
        .Add 33, "as"
        .Add 34, "se"
        .Add 35, "br"
        .Add 36, "kr"
        .Add 37, "rb"
        .Add 38, "sr"
        .Add 39, "y"
        .Add 40, "zr"
        .Add 41, "nb"
        .Add 42, "mo"
        .Add 43, "tc"
        .Add 44, "ru"
        .Add 45, "rh"
        .Add 46, "pd"
        .Add 47, "ag"
        .Add 48, "cd"
        .Add 49, "in"
        .Add 50, "sn"
        .Add 51, "sb"
        .Add 52, "te"
        .Add 53, "i"
        .Add 54, "xe"
        .Add 55, "cs"
        .Add 56, "ba"
        .Add 57, "la"
        .Add 58, "ce"
        .Add 59, "pr"
        .Add 60, "nd"
        .Add 61, "pm"
        .Add 62, "sm"
        .Add 63, "eu"
        .Add 64, "gd"
        .Add 65, "tb"
        .Add 66, "dy"
        .Add 67, "ho"
        .Add 68, "er"
        .Add 69, "tm"
        .Add 70, "yb"
        .Add 71, "lu"
        .Add 72, "hf"
        .Add 73, "ta"
        .Add 74, "w"
        .Add 75, "re"
        .Add 76, "os"
        .Add 77, "ir"
        .Add 78, "pt"
        .Add 79, "au"
        .Add 80, "hg"
        .Add 81, "tl"
        .Add 82, "pb"
        .Add 83, "bi"
        .Add 84, "po"
        .Add 85, "at"
        .Add 86, "rn"
        .Add 87, "fr"
        .Add 88, "ra"
        .Add 89, "ac"
        .Add 90, "th"
        .Add 91, "pa"
        .Add 92, "u"
        .Add 93, "np"
        .Add 94, "pu"
        .Add 95, "am"
        .Add 96, "cm"
        .Add 97, "bk"
        .Add 98, "cf"
        .Add 99, "es"
        .Add 100, "fm"
        .Add 101, "md"
        .Add 102, "no"
        .Add 103, "lr"
        .Add 104, "rf"
        .Add 105, "db"
        .Add 106, "sg"
        .Add 107, "bh"
        .Add 108, "hs"
        .Add 109, "mt"
        .Add 110, "ds"
        .Add 111, "rg"
        .Add 112, "cn"
        .Add 113, "uut"
        .Add 114, "fl"
        .Add 115, "uup"
        .Add 116, "lv"
        .Add 117, "uus"
        .Add 118, "uuo"
    End With

End Sub
