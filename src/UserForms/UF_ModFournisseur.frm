VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_ModFournisseur 
   Caption         =   "Modifier le fournisseur"
   ClientHeight    =   5670
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   8010
   OleObjectBlob   =   "UF_ModFournisseur.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UF_ModFournisseur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CMB_cancel_Click()
    UF_ModFournisseur.Hide
End Sub


Private Sub CMB_Modify_Click()
    Dim modFourn As New Fournisseur
    
    If TB_domaine.value = "" Or TB_societe.value = "" Then
    MsgBox "Veuillez compléter tout les champs obligatoires", vbCritical
    Else
   
   
        If TB_mail.value <> "" Then
        If IsValidEmail(TB_mail.value) = False Then
            MsgBox "L'adresse mail est invalide", vbCritical
            TB_mail = Empty
        Else
            'Confirmation de la modif
    If MsgBox("Etes-vous sûr de vouloir modifier " & UF_enseignants.LBenseignants.value & " ?", vbYesNo + vbQuestion, "Suppression") = vbYes Then
        fournDictionary.Remove (UF_fournisseur.LB_fournisseurs.List(UF_fournisseur.LB_fournisseurs.listIndex, 0))
        
        
        
        
        With modFourn
            .societe = TB_societe.value
            .telephone = TB_tel.value
            .mail = TB_mail.value
            .domaine = TB_domaine.value
        End With
            
            
         'modification du fournisseur
         ModFournisseur modFourn
         
        MsgBox "Le fournisseur a bien été modifié", vbInformation
        
        TB_societe = Empty
        TB_tel = Empty
        TB_mail = Empty
        TB_domaine = Empty
        
        UF_ModFournisseur.Hide
        
    End If
        End If
    Else
 'Confirmation de la modif
    If MsgBox("Etes-vous sûr de vouloir modifier " & UF_enseignants.LBenseignants.value & " ?", vbYesNo + vbQuestion, "Suppression") = vbYes Then
        fournDictionary.Remove (UF_fournisseur.LB_fournisseurs.List(UF_fournisseur.LB_fournisseurs.listIndex, 0))
        
        
        
        
        With modFourn
            .societe = TB_societe.value
            .telephone = TB_tel.value
            .mail = TB_mail.value
            .domaine = TB_domaine.value
        End With
            
            
         'modification du fournisseur
         ModFournisseur modFourn
         
        MsgBox "Le fournisseur a bien été modifié", vbInformation
        
        TB_societe = Empty
        TB_tel = Empty
        TB_mail = Empty
        TB_domaine = Empty
        
        UF_ModFournisseur.Hide
        
    End If
    End If
    End If
End Sub

'réinitialisation du formulaire
Private Sub CMB_ReInit_Click()
    UserForm_Initialize
End Sub

'on initialise le formulaire afin qu'il soit pré-rempli par les informations du fournisseur à modifier
Private Sub UserForm_Initialize()
    TB_societe.value = UF_fournisseur.LB_fournisseurs.List(UF_fournisseur.LB_fournisseurs.listIndex, 0)
    TB_tel.value = UF_fournisseur.LB_fournisseurs.List(UF_fournisseur.LB_fournisseurs.listIndex, 1)
    TB_mail.value = UF_fournisseur.LB_fournisseurs.List(UF_fournisseur.LB_fournisseurs.listIndex, 2)
    TB_domaine.value = UF_fournisseur.LB_fournisseurs.List(UF_fournisseur.LB_fournisseurs.listIndex, 3)
End Sub

'méthode pour modifier un fournisseur
Private Sub ModFournisseur(fourn As Fournisseur)
        If (fournDictionary.Exists(fourn.societe)) Then
            MsgBox "Ce fournisseur existe déjà dans la base données", vbCritical
        Else
        Dim numRows As Integer
        numRows = GetSheetNumRows(SheetFournisseurs)
        
        
        'Récupération de l'index
        Dim listIndex As Integer
        listIndex = UF_fournisseur.LB_fournisseurs.listIndex
        
        'Récupération du fournisseur sélectionné
        Dim teacherValue As String
        teacherValue = UF_fournisseur.LB_fournisseurs.value
        
        
        'modifications sur la feuille
        With SheetFournisseurs
            .Cells(listIndex + 2, 1) = fourn.societe
            .Cells(listIndex + 2, 2) = fourn.telephone
            .Cells(listIndex + 2, 3) = fourn.mail
            .Cells(listIndex + 2, 4) = fourn.domaine
        
            
            'Tri sur les fournisseurs
            .Range("A2:D" & numRows).Sort Key1:=.Range("A2"), Order1:=xlAscending, _
            Header:=xlGuess, OrderCustom:=1, MatchCase _
            :=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal, _
            DataOption2:=xlSortNormal
        End With
      
        'ajout au dictionaire
        fournDictionary.Add Key:=fourn.societe, Item:=fourn
        End If
        
End Sub
