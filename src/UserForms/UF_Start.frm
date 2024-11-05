VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_Start 
   Caption         =   "Que voulez-vous gérer?"
   ClientHeight    =   5355
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   10110
   OleObjectBlob   =   "UF_Start.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UF_Start"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CMB_Budget_Click()
    UF_Budget.Show
End Sub

Private Sub CMB_Fournisseurs_Click()
    UF_fournisseur.Show
End Sub

Private Sub CMD_enseignants_Click()
    UF_enseignants.Show
End Sub

Private Sub CMD_factures_Click()
UF_Factures.Show
End Sub

Private Sub UserForm_Initialize()
    Dim nbAnnees As Integer
    Dim i As Integer
    Dim annee As String
    nbAnnees = GetSheetNumRows(SheetAnnees)

    'initialisation des dictionnaires
    ComputeEnsDictionary
    computeFournDictionary
    For i = 2 To nbAnnees
       annee = SheetAnnees.Cells(i, 1).value
       If WorksheetExists(annee) Then
        computeFactDictionary (annee)
       End If
    Next
    
    
End Sub
