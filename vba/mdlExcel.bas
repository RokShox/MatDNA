Attribute VB_Name = "mdlExcel"
'
' Copyright (c) 2017 Henkel Technical Computing, LLC. All rights reserved.
'
Option Explicit

Public Sub AddUniqueWorksheet(sName As String, Optional wksProto As Worksheet = Nothing, Optional wksAfter As Worksheet = Nothing)
    Dim wks As Worksheet
    Dim blnDisplayAlerts As Boolean
    Dim wksLast As Worksheet

    blnDisplayAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    
    For Each wks In Worksheets
        If wks.Name = sName Then
            wks.Delete
            GoTo Deleted
        End If
    Next wks
    
Deleted:
    Set wksLast = Worksheets(Worksheets.Count)
    If wksAfter Is Nothing Then
        Set wksAfter = wksLast
    End If

'    Debug.Print "wksProto name is " & wksProto.Name
'    Debug.Print "wksAfter name is " & wksAfter.Name

    If wksProto Is Nothing Then
        Worksheets.Add After:=wksAfter
    Else
        wksProto.Copy After:=wksAfter
    End If
    ActiveSheet.Name = sName
    ActiveWindow.Zoom = 75

    Application.DisplayAlerts = blnDisplayAlerts

End Sub

Public Function GetWorkbook(sName As String) As Workbook
    Dim wkb As Workbook
    
    For Each wkb In Workbooks
        If wkb.Name = sName Then
            Set GetWorkbook = wkb
            Exit Function
        End If
    Next wkb
    Set GetWorkbook = Nothing
End Function

Public Function GetWorksheet(sName As String, Optional wkb As Workbook = Nothing) As Worksheet
    Dim wks As Worksheet
    
    If wkb Is Nothing Then
        Set wkb = ThisWorkbook
    End If
    
    For Each wks In wkb.Worksheets
        If wks.Name = sName Then
            Set GetWorksheet = wks
            Exit Function
        End If
    Next wks
    Set GetWorksheet = Nothing
End Function

Public Function GetRangebyName(sName As String) As Range
    Dim nm As Name
    For Each nm In ThisWorkbook.Names
        If UCase(nm.Name) = UCase(sName) Then
            Set GetRangebyName = Names(sName).RefersToRange
            Exit Function
        End If
    Next nm
    Set GetRangebyName = Nothing
End Function

Public Sub CalcOff()
    Application.Calculation = xlCalculationManual
End Sub

Public Sub CalcOn()
    Application.Calculation = xlCalculationAutomatic
End Sub

Public Sub CalcSave()
    gblCalcState = Application.Calculation
End Sub

Public Sub CalcRestore()
    Application.Calculation = gblCalcState
End Sub

Public Function RefersTo(rng As Range) As Variant
    Dim sText As String
    
    Application.Volatile Volatile:=True

    sText = rng.Address
    sText = rng.Worksheet.Name & "!" & sText
    RefersTo = sText
End Function

Public Function IsBorZ(rng As Range) As Boolean
'   Is blank or zero?
    Application.Volatile
    
    If rng.Value2 = 0 Or rng.Cells(1, 1).Text = "" Then
        IsBorZ = True
    Else
        IsBorZ = False
    End If
End Function

Public Sub FormatEntries(rng As Range)
    With rng
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
        With .Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
    End With
End Sub

Public Sub FormatAux(rng As Range)
    With rng
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
        With .Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 15132390
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
    End With
End Sub

Public Sub BorderExtent(rng As Range)
    With rng
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
        .Borders(xlEdgeLeft).LineStyle = xlNone
        .Borders(xlEdgeTop).LineStyle = xlNone
        .Borders(xlEdgeBottom).LineStyle = xlNone
        .Borders(xlEdgeRight).LineStyle = xlNone
        .Borders(xlInsideVertical).LineStyle = xlNone
        .Borders(xlInsideHorizontal).LineStyle = xlNone
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
        With .Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
    End With
End Sub

Public Sub BorderInside(rng As Range)
    With rng
        .Borders(xlInsideVertical).LineStyle = xlNone
        .Borders(xlInsideHorizontal).LineStyle = xlNone
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
        With .Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
    End With
End Sub

Public Function GetSortedKeys(d As Scripting.Dictionary) As ArrayList
    Dim al As ArrayList
    Dim sKey As Variant
    
    Set al = New ArrayList
    For Each sKey In d.Keys
        al.Add sKey
    Next sKey
    al.Sort
    Set GetSortedKeys = al
End Function


Public Sub ExportModules()
    Dim bExport As Boolean
    Dim wkbSource As Excel.Workbook
    Dim szSourceWorkbook As String
    Dim szExportPath As String
    Dim szFileName As String
    Dim cmpComponent As VBIDE.VBComponent
    Dim sPath As String

    ''' The code modules will be exported in a folder named.
    ''' VBAProjectFiles in the Documents folder.
    ''' The code below create this folder if it not exist
    ''' or delete all files in the folder if it exist.
    sPath = FolderWithVBAProjectFiles
    If sPath = "Error" Then
        MsgBox "Export Folder not exist"
        Exit Sub
    End If
    
'    On Error Resume Next
'        Kill sPath & "\*.*"
'    On Error GoTo 0

    ''' NOTE: This workbook must be open in Excel.
    szSourceWorkbook = ActiveWorkbook.Name
    Set wkbSource = Application.Workbooks(szSourceWorkbook)
    
    If wkbSource.VBProject.Protection = 1 Then
    MsgBox "The VBA in this workbook is protected," & _
        "not possible to export the code"
    Exit Sub
    End If
    
    szExportPath = sPath & "\"
    
    For Each cmpComponent In wkbSource.VBProject.VBComponents
        
        bExport = True
        szFileName = cmpComponent.Name

        ''' Concatenate the correct filename for export.
        Select Case cmpComponent.Type
            Case vbext_ct_ClassModule
                szFileName = szFileName & ".cls"
            Case vbext_ct_MSForm
                szFileName = szFileName & ".frm"
            Case vbext_ct_StdModule
                szFileName = szFileName & ".bas"
            Case vbext_ct_Document
                ''' This is a worksheet or workbook object.
                ''' Don't try to export.
                bExport = False
        End Select
        
        If bExport Then
            ''' Export the component to a text file.
            cmpComponent.Export szExportPath & szFileName
            
        ''' remove it from the project if you want
        '''wkbSource.VBProject.VBComponents.Remove cmpComponent
        
        End If
   
    Next cmpComponent

    MsgBox "Export is ready"
End Sub


Public Sub ImportComponent()
    Dim vbeComponents As VBIDE.VBComponents
    Dim vbeComponent As VBIDE.VBComponent
    Dim vbeComponentReplace As VBIDE.VBComponent
    Dim vbeProject As VBIDE.VBProject
    Dim objFSO As Scripting.FileSystemObject
    Dim objFile As Scripting.File
    Dim dlgFile As FileDialog
    Dim sPath As String
    Dim sBase As String
    Dim vbeCodePane As VBIDE.CodePane
    Dim vbeWindow As VBIDE.Window
    
    
    Set vbeProject = ActiveWorkbook.VBProject
    Set vbeComponents = vbeProject.VBComponents
    Set objFSO = New Scripting.FileSystemObject
    Set dlgFile = Application.FileDialog(msoFileDialogFilePicker)
    
    Debug.Print "Project name is " & vbeProject.Name
    Debug.Print "Project filename name is " & vbeProject.Filename
    
    
    With dlgFile
        .Filters.Clear
        .AllowMultiSelect = False
        .Show
        sPath = .SelectedItems(1)
    End With
    
    Set objFile = objFSO.GetFile(sPath)
    sBase = objFSO.GetBaseName(sPath)
    Debug.Print "Basename is " & sBase
    
    Set vbeComponentReplace = Nothing
    For Each vbeComponent In vbeComponents
        If vbeComponent.Name = sBase Then
            Set vbeComponentReplace = vbeComponent
            Debug.Print sBase & " exists"
        End If
    Next
    
    If Not vbeComponentReplace Is Nothing Then
        ' Close any code panes in this project
        For Each vbeCodePane In vbeProject.VBE.CodePanes
            Set vbeComponent = vbeCodePane.CodeModule.Parent
            If vbeComponent.Name = sBase Then
                ' Need to check if component is in this project
                If vbeComponent.Collection.Parent.Filename = vbeProject.Filename Then
                    Debug.Print sBase & " has a code pane"
                    vbeCodePane.Window.Close
                End If
            End If
        Next
        
        ' Remove old component
        vbeComponents.Remove vbeComponentReplace
        Set vbeComponentReplace = Nothing
    End If
    
    ' Import new component
    Set vbeComponent = vbeComponents.Import(sPath)

End Sub


Public Sub ImportComponents()
    Dim vbeComponents As VBIDE.VBComponents
    Dim vbeComponent As VBIDE.VBComponent
    Dim vbeComponentReplace As VBIDE.VBComponent
    Dim vbeProject As VBIDE.VBProject
    Dim objFSO As Scripting.FileSystemObject
    Dim objFolder As Scripting.Folder
    Dim objFile As Scripting.File
    Dim dlgFolder As FileDialog
    Dim sPath As String
    Dim sBase As String
    Dim vbeCodePane As VBIDE.CodePane
    Dim vbeWindow As VBIDE.Window
    
    
    Set vbeProject = ActiveWorkbook.VBProject
    Set vbeComponents = vbeProject.VBComponents
    Set objFSO = New Scripting.FileSystemObject
    Set dlgFolder = Application.FileDialog(msoFileDialogFolderPicker)
    
    Debug.Print "Project name is " & vbeProject.Name
    Debug.Print "Project filename name is " & vbeProject.Filename
    
    
    With dlgFolder
        .Filters.Clear
        .AllowMultiSelect = False
        .Show
        sPath = .SelectedItems(1)
    End With
    
    Set objFolder = objFSO.GetFolder(sPath)
    Debug.Print "Basename is " & sBase
    
    For Each objFile In objFolder.Files
    
        sPath = objFile.Path
        sBase = objFSO.GetBaseName(sPath)
        
        Set vbeComponentReplace = Nothing
        For Each vbeComponent In vbeComponents
            If vbeComponent.Name = sBase Then
                Set vbeComponentReplace = vbeComponent
                Debug.Print sBase & " exists"
            End If
        Next
        
        If Not vbeComponentReplace Is Nothing Then
            ' Close any code panes in this project
            For Each vbeCodePane In vbeProject.VBE.CodePanes
                Set vbeComponent = vbeCodePane.CodeModule.Parent
                If vbeComponent.Name = sBase Then
                    ' Need to check if component is in this project
                    If vbeComponent.Collection.Parent.Filename = vbeProject.Filename Then
                        Debug.Print sBase & " has a code pane"
                        vbeCodePane.Window.Close
                    End If
                End If
            Next
            
            ' Remove old component
            vbeComponents.Remove vbeComponentReplace
            Debug.Print sBase & " removed"
            Set vbeComponentReplace = Nothing
        End If
    
        ' Import new component
        Set vbeComponent = vbeComponents.Import(sPath)
        Debug.Print sBase & " replaced"

    Next objFile

End Sub


Public Sub TestVBIDE()
    Dim vbeComponents As VBIDE.VBComponents
    Dim vbeComponent As VBIDE.VBComponent
    Dim vbeComponentReplace As VBIDE.VBComponent
    Dim vbeProject As VBIDE.VBProject
    Dim objFSO As Scripting.FileSystemObject
    Dim objFile As Scripting.File
    Dim dlgFile As FileDialog
    Dim sPath As String
    Dim sBase As String
    Dim vbeCodePane As VBIDE.CodePane
    Dim vbeWindow As VBIDE.Window
    
    
    Set vbeProject = ActiveWorkbook.VBProject
    Set vbeComponents = vbeProject.VBComponents
    Set objFSO = New Scripting.FileSystemObject
    Set dlgFile = Application.FileDialog(msoFileDialogFilePicker)
    
    Debug.Print "Project name is " & vbeProject.Name
    Debug.Print "Project filename name is " & vbeProject.Filename
    
    
    With dlgFile
        .Filters.Clear
        .AllowMultiSelect = False
        .Show
        sPath = .SelectedItems(1)
    End With
    
    Set objFile = objFSO.GetFile(sPath)
    sBase = objFSO.GetBaseName(sPath)
    Debug.Print "Basename is " & sBase
    
    Set vbeComponentReplace = Nothing
    For Each vbeComponent In vbeComponents
'        If vbeComponent.Type = vbext_ct_ClassModule Then
'            Debug.Print "Class: " & vbeComponent.Name
'        ElseIf vbeComponent.Type = vbext_ct_StdModule Then
'            Debug.Print "Module: " & vbeComponent.Name
'        ElseIf vbeComponent.Type = vbext_ct_Document Then
'            Debug.Print "Document: " & vbeComponent.Name
'        End If
        If vbeComponent.Name = sBase Then
            Set vbeComponentReplace = vbeComponent
            Debug.Print sBase & " exists"
        End If
    Next
    
    If Not vbeComponentReplace Is Nothing Then
        ' Close any code panes in this project
        For Each vbeCodePane In vbeProject.VBE.CodePanes
            Set vbeComponent = vbeCodePane.CodeModule.Parent
            If vbeComponent.Name = sBase Then
                ' Need to check if component is in this project
                If vbeComponent.Collection.Parent.Filename = vbeProject.Filename Then
                    Debug.Print sBase & " has a code pane"
                    vbeCodePane.Window.Close
                End If
            End If
        Next
        
        ' Remove old component
        vbeComponents.Remove vbeComponentReplace
        Set vbeComponentReplace = Nothing
    End If
    
    ' Import new component
    Set vbeComponent = vbeComponents.Import(sPath)

End Sub

'
'Public Sub ImportModules()
'    Dim wkbTarget As Excel.Workbook
'    Dim objFSO As Scripting.FileSystemObject
'    Dim objFile As Scripting.File
'    Dim szTargetWorkbook As String
'    Dim szImportPath As String
'    Dim szFileName As String
'    Dim cmpComponents As VBIDE.VBComponents
'
'    If ActiveWorkbook.Name = ThisWorkbook.Name Then
'        MsgBox "Select another destination workbook" & _
'        "Not possible to import in this workbook "
'        Exit Sub
'    End If
'
'    'Get the path to the folder with modules
'    If FolderWithVBAProjectFiles = "Error" Then
'        MsgBox "Import Folder not exist"
'        Exit Sub
'    End If
'
'    ''' NOTE: This workbook must be open in Excel.
'    szTargetWorkbook = ActiveWorkbook.Name
'    Set wkbTarget = Application.Workbooks(szTargetWorkbook)
'
'    If wkbTarget.VBProject.Protection = 1 Then
'    MsgBox "The VBA in this workbook is protected," & _
'        "not possible to Import the code"
'    Exit Sub
'    End If
'
'    ''' NOTE: Path where the code modules are located.
'    szImportPath = FolderWithVBAProjectFiles
'
'    Set objFSO = New Scripting.FileSystemObject
'    If objFSO.GetFolder(szImportPath).Files.Count = 0 Then
'       MsgBox "There are no files to import"
'       Exit Sub
'    End If
'
'    'Delete all modules/Userforms from the ActiveWorkbook
'    Call DeleteVBAModulesAndUserForms
'
'    Set cmpComponents = wkbTarget.VBProject.VBComponents
'
'    ''' Import all the code modules in the specified path
'    ''' to the ActiveWorkbook.
'    For Each objFile In objFSO.GetFolder(szImportPath).Files
'
'        If (objFSO.GetExtensionName(objFile.Name) = "cls") Or _
'            (objFSO.GetExtensionName(objFile.Name) = "frm") Or _
'            (objFSO.GetExtensionName(objFile.Name) = "bas") Then
'            cmpComponents.Import objFile.Path
'        End If
'
'    Next objFile
'
'    MsgBox "Import is ready"
'End Sub

Function FolderWithVBAProjectFiles() As String
    Dim FSO As Object
    Dim SpecialPath As String
    Dim iDot As Integer
    Dim sName As String
    Dim sPath As String

    Set FSO = CreateObject("scripting.filesystemobject")
    iDot = InStr(ActiveWorkbook.Name, ".")
    sName = "vba_" & Mid(ActiveWorkbook.Name, 1, iDot - 1)

    SpecialPath = ActiveWorkbook.Path

    If Right(SpecialPath, 1) <> "\" Then
        SpecialPath = SpecialPath & "\"
    End If
    sPath = SpecialPath & sName
    
    If FSO.FolderExists(sPath) = False Then
        On Error Resume Next
        MkDir sPath
        On Error GoTo 0
    End If
    
    If FSO.FolderExists(sPath) = True Then
        FolderWithVBAProjectFiles = sPath
    Else
        FolderWithVBAProjectFiles = "Error"
    End If
    
End Function

'
'Function DeleteVBAModulesAndUserForms()
'        Dim VBProj As VBIDE.VBProject
'        Dim VBComp As VBIDE.VBComponent
'
'        Set VBProj = ActiveWorkbook.VBProject
'
'        For Each VBComp In VBProj.VBComponents
'            If VBComp.Type = vbext_ct_Document Then
'                'Thisworkbook or worksheet module
'                'We do nothing
'            Else
'                VBProj.VBComponents.Remove VBComp
'            End If
'        Next VBComp
'End Function

