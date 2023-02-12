Attribute VB_Name = "mdlORIGIN"
'
' Copyright (c) 2017 Henkel Technical Computing, LLC. All rights reserved.
'
Option Explicit

Public Sub MakeCase()
    Dim sCaseName As String
    Dim sFileName As String
    Dim sDirSrc As String
    Dim sDirCase As String
    Dim sRev As String
    Dim rngInput As Range
    Dim iFile As Integer
    
    sDirSrc = Names("dirSrc").RefersToRange.Value
    sCaseName = ActiveSheet.Names("selCase").RefersToRange.Value
    sRev = Names("version").RefersToRange.Value
    sDirCase = sDirSrc & "\" & sRev
    SafeMkDir sDirSrc, True
    SafeMkDir sDirCase, True
    sFileName = sDirCase & "\" & sCaseName & ".inp"
    Set rngInput = ActiveSheet.Names("inpCase").RefersToRange
    
    iFile = FreeFile
    Open sFileName For Output As #iFile
    WriteSCALERangeToFile rngInput, iFile, Nothing, Chr(32)
    Close #iFile
End Sub

