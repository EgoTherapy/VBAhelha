VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_NewFournisseur 
   Caption         =   "Ajouter un fournisseur"
   ClientHeight    =   5505
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   5865
   OleObjectBlob   =   "UF_NewFournisseur.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UF_NewFournisseur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CMB_addFourn_Click()
    Dim fourn As New Fournisseur
    If TB_mail.value <> "" Then
        If IsValidEmail(TB_mail.value) = False Then
            MsgBox "L'adresse mail est invalide", vbCritical
            TB_mail = Empty
        Else
            With fourn
                .societe = TB_nom.value
                .telephone = TB_telephone.value
                .mail = TB_mail.value
                .domaine = TB_domaine.value
            End With
            UF_fournisseur.AddFourn fourn
        End If
    Else
        With fourn
                .societe = TB_nom.value
                .telephone = TB_telephone.value
                .mail = TB_mail.value
                .domaine = TB_domaine.value
        End With
        UF_fournisseur.AddFourn fourn
    End If
    UF_fournisseur.UpdateLBFournisseurs
        
End Sub

Private Sub CMB_cancel_Click()
    UF_NewFournisseur.Hide
End Sub


Private Sub UserForm_Click()

End Sub
