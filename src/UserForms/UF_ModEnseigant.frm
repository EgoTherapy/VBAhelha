VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_ModEnseigant 
   Caption         =   "Modifier le nom d'un enseignant"
   ClientHeight    =   3885
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   6840
   OleObjectBlob   =   "UF_ModEnseigant.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UF_ModEnseigant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CMB_cancelMod_Click()
    UF_ModEnseigant.Hide
End Sub

Private Sub CMB_confirmMod_Click()
    If NomModEnseignant.value = "" Or PrenomModEnseignant.value = "" Then
    MsgBox "Veuillez compléter tout les champs", vbCritical
    Else
   
   'Confirmation de la suppression
    If MsgBox("Etes-vous sûr de vouloir modifier " & UF_enseignants.LBenseignants.value & " ?", vbYesNo + vbQuestion, "Suppression") = vbYes Then
        'Récupération de l'index
        Dim listIndex As Integer
        listIndex = UF_enseignants.LBenseignants.listIndex
        
        'Récupération de l'enseignant sélectionné
        Dim teacherValue As String
        teacherValue = UF_enseignants.LBenseignants.value
      
        
        
        'modification sur la feuille
        SheetEnseignants.Cells(listIndex + 2, 1) = NomModEnseignant.value & " " & PrenomModEnseignant.value
    
    'tri par ordre alphabetique
    Dim numRows As Integer
    numRows = GetSheetNumRows(SheetEnseignants)
    With SheetEnseignants
        'Tri sur les enseignants
        .Range("A2:A" & numRows).Sort Key1:=.Range("A2"), Order1:=xlAscending, _
            Header:=xlGuess, OrderCustom:=1, MatchCase _
            :=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal, _
            DataOption2:=xlSortNormal
    End With
        MsgBox "L'enseignant a bien été modifié", vbInformation
        
        NomModEnseignant = Empty
        PrenomModEnseignant = Empty
        
        UF_ModEnseigant.Hide
        
    End If
    End If
End Sub



Private Sub UserForm_Initialize()
    Cadre_modEnseignant.Caption = "modifier " & UF_enseignants.LBenseignants.value
End Sub
