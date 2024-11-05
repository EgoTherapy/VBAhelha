VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_fournisseur 
   Caption         =   "Liste des fournisseurs"
   ClientHeight    =   7635
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   16890
   OleObjectBlob   =   "UF_fournisseur.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UF_fournisseur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CMB_AddFournisseur_Click()
    UF_NewFournisseur.Show
End Sub

Private Sub CMB_modifier_Click()
    UF_ModFournisseur.Show
End Sub

Private Sub CMB_PDF_Click()
    ExportAsPDF "Fournisseurs"
End Sub

Private Sub CMB_Print_Click()
    PrintSheet "Fournisseurs"
End Sub

Private Sub CMB_supprimer_Click()
    'Confirmation de la suppression
    If MsgBox("Etes-vous sûr de vouloir supprimer " & LB_fournisseurs.List(UF_fournisseur.LB_fournisseurs.listIndex, 0) & " de la base de données?", vbYesNo + vbQuestion, "Suppression") = vbYes Then
        'Récupération de l'index
        Dim listIndex As Integer
        listIndex = LB_fournisseurs.listIndex
        
        'Récupération de l'enseignant sélectionné
        Dim fournValue As String
        fournValue = LB_fournisseurs.List(UF_fournisseur.LB_fournisseurs.listIndex, 0)
      
        
        LB_fournisseurs.value = Empty
        
        'Suppression de l'enseignant
        SheetFournisseurs.Rows(listIndex + 2).EntireRow.Delete
        fournDictionary.Remove (fournValue)
        
        MsgBox "Le fournisseur a bien été supprimé", vbInformation
        
        'Mise à jour du userform
        UpdateLBFournisseurs
        
    End If
End Sub

Private Sub Cmd_quitter_Click()
    UF_fournisseur.Hide
End Sub

Private Sub LB_fournisseurs_Click()
    CMB_modifier.Enabled = True
    CMB_supprimer.Enabled = True
End Sub

Private Sub UserForm_Initialize()
        'remplir la list box des fournisseurs
        UpdateLBFournisseurs
    
        'Nombre de colonnes dans la ListBox
        LB_fournisseurs.ColumnCount = 4
        'Largeur des colonnes de la ListBox
        LB_fournisseurs.ColumnWidths = "170;70;150;170"
        
        CMB_modifier.Enabled = False
        CMB_supprimer.Enabled = False
End Sub
Sub UpdateLBFournisseurs()
    'méthode pour rafraichir la listbox contenant les fournisseurs
    LB_fournisseurs = Empty
    Dim nr As Integer
    
    nr = GetSheetNumRows(SheetFournisseurs)
    LB_fournisseurs.ColumnHeads = True
    LB_fournisseurs.RowSource = SheetFournisseurs.name & "!A2:D" & nr
    

End Sub

Sub AddFourn(nvFourn As Fournisseur)
    'Nombre de fournisseur
    Dim numRows As Integer
    numRows = GetSheetNumRows(SheetFournisseurs) + 1

        'Vérification que le fournisseur n'existe pas déjà
            If fournDictionary.Exists(nvFourn.societe) Then
                MsgBox "Ce fournisseur est déjà dans la base de données!", vbCritical
            Else
                'ajout du fournisseur dans le dictionaire
                fournDictionary.Add Key:=nvFourn.societe, Item:=nvFourn
                
                 'ajout de l'enseignant sur la feuille qui sert de base de données
                 With SheetFournisseurs
                     .Cells(numRows, 1) = nvFourn.societe
                     .Cells(numRows, 2) = nvFourn.telephone
                     .Cells(numRows, 3) = nvFourn.mail
                     .Cells(numRows, 4) = nvFourn.domaine
                        
                     'Tri sur les fournisseurs, par ordre alphabethique
                     .Range("A2:D" & numRows).Sort Key1:=.Range("A2"), Order1:=xlAscending, _
                         Header:=xlGuess, OrderCustom:=1, MatchCase _
                         :=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal, _
                         DataOption2:=xlSortNormal
                 End With
            End If
            
    'Mise à jour des données
    UpdateLBFournisseurs
    
End Sub
