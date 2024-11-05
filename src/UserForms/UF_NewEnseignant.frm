VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_NewEnseignant 
   Caption         =   "Nouvel enseignant"
   ClientHeight    =   4008
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   6630
   OleObjectBlob   =   "UF_NewEnseignant.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UF_NewEnseignant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CMB_AddEnseignant_Click()

 If NomNvEnseignant.value = "" Or PrenomNvEnseignant.value = "" Then
    MsgBox "Veuillez compléter tout les champs", vbCritical
 Else

            Dim ens As New Enseignant
            
            'Création de l'enseignant
           ens.NomPrenom = NomNvEnseignant.value & " " & PrenomNvEnseignant.value
            UF_enseignants.AddTeacher ens
            
            
          
End If
    
    'Mise à jour des données
    NomNvEnseignant = Empty
    PrenomNvEnseignant = Empty
    
End Sub


Private Sub CMB_Annuler_Click()
    UF_NewEnseignant.Hide
End Sub
