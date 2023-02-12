VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSurface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private md As Scripting.Dictionary
Private rngSurface As Range

Private Sub Class_Initialize()
    Set md = New Scripting.Dictionary
End Sub

Public Sub Initialize()
    Dim icolIndex As Integer
    Dim icolDollar As Integer
    Dim iColKey As Integer
    Dim iRow As Integer
    Dim sKey As String
    Dim sIndex As String
    Dim iLook As Integer
    
    md.RemoveAll
    icolIndex = 1
    iColKey = rngSurface.Columns.Count
    icolDollar = iColKey - 1
    
    For iRow = 1 To rngSurface.Rows.Count
        If rngSurface.Cells(iRow, icolDollar) = "$" Then
            sKey = Trim(rngSurface.Cells(iRow, iColKey).Text)
            If sKey <> "" Then
                
                sIndex = ""
                iLook = 0
                Do Until sIndex <> "" Or iLook = 10
                    sIndex = Trim(rngSurface.Cells(iRow - iLook, icolIndex).Text)
                    iLook = iLook + 1
                Loop
            
                If iLook >= 10 Then
                    MsgBox "Can't find index"
                End If
            
                If md.Exists(sKey) Then
'                    MsgBox "Duplicate key: " & sKey, vbOKOnly
                Else
                    md.Add sKey, sIndex
                End If
            End If
        End If
    Next iRow
End Sub

Public Property Get SurfaceRange() As Range
    Set SurfaceRange = rngSurface
End Property

Public Property Set SurfaceRange(rng As Range)
    Set rngSurface = rng
    Me.Initialize
End Property

Public Function Lookup(sSurf As String) As Long
    Dim sSign As String
    
    sSurf = Trim(sSurf)
    sSign = Left(sSurf, 1)
    If sSign = "-" Or sSign = "–" Then
        sSurf = Mid(sSurf, 2)
        sSign = "-"
    ElseIf sSign = "+" Then
        sSurf = Mid(sSurf, 2)
        sSign = ""
    Else
        sSign = ""
    End If

    If md.Exists(sSurf) Then
        Lookup = sSign & md.Item(sSurf)
    Else
        Lookup = "Error"
    End If
End Function
