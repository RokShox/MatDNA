Attribute VB_Name = "mdlMatTest"
Public Sub TestFractionComInMat()

    Dim cMat As clsMatMaterialRange
    Dim cMatMaster As clsMatMaster
    Dim cCom As clsMatComRange
    Dim cCon As clsMatConRange
    
    Dim rngStanza As Range
    Dim rngData As Range
    
    Dim sMat As String
    Dim sCom As String
    Dim sCon As String
    Dim sIso As String

    Set cMatMaster = New clsMatMaster
    sMat = "BerlSaddles"
    sCom = "AgNO3"
    sCon = "n"
    sIso = "n-14"
    
    Set rngStanza = cMatMaster.Stanza(sMat)
    Set rngData = rngStanza.Cells(1, 1).Offset(or_Stanza_Comp + or_Comp_Data, oc_Stanza_Comp)
    
    
    Set cMat = New clsMatMaterialRange
    With cMat
        .mode = Atom
        .CopyMode = ByValue
        .InitFromRange rngData
    End With
    
    Debug.Print sMat & ": Mass fraction " & sCom & " = " & cMat.FractionComInMat(sCom, Mass)
    Debug.Print sMat & ": Atom fraction " & sCom & " = " & cMat.FractionComInMat(sCom, Atom)
    

    Set cCom = cMat.Component(sCom)
    Debug.Print sMat & " - " & sCom & ": Mass fraction " & sCon & " = " & cCom.FractionConInCom(sCon, Mass)
    Debug.Print sMat & " - " & sCom & ": Atom fraction " & sCon & " = " & cCom.FractionConInCom(sCon, Atom)

    Set cCon = cCom.Constituent(sCon)
    Debug.Print sMat & " - " & sCom & " - " & sCon & ": Mass fraction " & sIso & " = " & cCon.FractionIsoInCon(sIso, Mass)
    Debug.Print sMat & " - " & sCom & " - " & sCon & ": Atom fraction " & sIso & " = " & cCon.FractionIsoInCon(sIso, Atom)

    ' Frac iso in com
    Debug.Print sMat & " - " & sCom & ": Mass fraction " & sIso & " = " & cCom.FractionIsoInCom(sIso, Mass)
    Debug.Print sMat & " - " & sCom & ": Atom fraction " & sIso & " = " & cCom.FractionIsoInCom(sIso, Atom)
        
    ' Frac con in mat
    Debug.Print sMat & ": Mass fraction " & sCon & " = " & cMat.FractionConInMat(sCon, Mass)
    Debug.Print sMat & ": Atom fraction " & sCon & " = " & cMat.FractionConInMat(sCon, Atom)

    ' Frac iso in mat
    Debug.Print sMat & ": Mass fraction " & sIso & " = " & cMat.FractionIsoInMat(sIso, Mass)
    Debug.Print sMat & ": Atom fraction " & sIso & " = " & cMat.FractionIsoInMat(sIso, Atom)


    Set cCon = Nothing
    Set cCom = Nothing
    Set cMat = Nothing

End Sub



