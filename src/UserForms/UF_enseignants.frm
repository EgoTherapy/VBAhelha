VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_enseignants 
   Caption         =   "Enseignants"
   ClientHeight    =   8880.001
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   12825
   OleObjectBlob   =   "UF_enseignants.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UF_enseignants"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Méthode pour mettre à jour les enseignants
Private Sub UpdateLBEnseignants()
    LBenseignants = Empty
    Dim nr As Integer
    
    nr = GetSheetNumRows(SheetEnseignants)
    
    LBenseignants.RowSource = SheetEnseignants.name & "!A2:A" & nr
  
End Sub



Private Sub Cadre_fraisEnseignants_Click()

End Sub

Private Sub CMB_AddEnseignant_Click()
UF_NewEnseignant.Show
End Sub

Private Sub CMB_backStart_Click()
    UF_enseignants.Hide
End Sub

Private Sub CMB_DelEnseignant_Click()
    
    'Confirmation de la suppression
    If MsgBox("Etes-vous sûr de vouloir supprimer " & LBenseignants.value & " de la base de données?", vbYesNo + vbQuestion, "Suppression") = vbYes Then
        'Récupération de l'index
        Dim listIndex As Integer
        listIndex = LBenseignants.listIndex
        
        'Récupération de l'enseignant sélectionné
        Dim teacherValue As String
        teacherValue = LBenseignants.value
      
        
        LBenseignants.value = Empty
        
        'Suppression de l'enseignant
        SheetEnseignants.Rows(listIndex + 2).EntireRow.Delete
        ensDictionary.Remove (teacherValue)
        
        MsgBox "L'enseignant a bien été supprimé", vbInformation
        
        'Mise à jour du userform
        UpdateLBEnseignants
        
    End If
    
End Sub


Private Sub CMB_ModEnseignant_Click()
    UF_ModEnseigant.Show
End Sub

Private Sub CMB_PDF_Click()
    ExportAsPDF Replace(LBenseignants.value, " ", "")
End Sub

Private Sub CMB_Print_Click()
    PrintSheet Replace(LBenseignants.value, " ", "")
End Sub

Private Sub LBenseignants_Click()

    CMB_DelEnseignant.Enabled = True
    CMB_ModEnseignant.Enabled = True
    LB_listeFraisEns.RowSource = ""
    If WorksheetExists(Replace(LBenseignants.value, " ", "")) Then
        UpdateFraisList
        CMB_PDF.Enabled = True
    Else
        CMB_PDF.Enabled = False
    
    End If
    
    
    Cadre_fraisEnseignants.Caption = "Frais de " & LBenseignants.value
End Sub

Private Sub UserForm_Initialize()
    UpdateLBEnseignants
    
    CMB_DelEnseignant.Enabled = False
    CMB_ModEnseignant.Enabled = False
    CMB_PDF.Enabled = False

        'Nombre de colonnes dans la ListBox
        LB_listeFraisEns.ColumnCount = 10
        'Largeur des colonnes de la ListBox
        LB_listeFraisEns.ColumnWidths = "50;60;70;170;200;200;200;110"
    
End Sub

Sub AddTeacher(nvEnseignant As Enseignant)
    'Nombre d'enseignants
    Dim numRows As Integer
    numRows = GetSheetNumRows(SheetEnseignants) + 1

        'Vérification que l'enseignant n'existe pas déjà
            If ensDictionary.Exists(nvEnseignant.NomPrenom) Then
                MsgBox "Il existe déjà un professeur avec ce nom et prénom!", vbCritical
            Else
                'ajout de l'enseignant dans le dictionaire
                ensDictionary.Add Key:=nvEnseignant.NomPrenom, Item:=nvEnseignant
                
                 'ajout de l'enseignant sur la feuille qui sert de base de données
                 With SheetEnseignants
                     .Cells(numRows, 1) = nvEnseignant.NomPrenom
                        
                     'Tri sur les enseignants
                     .Range("A2:A" & numRows).Sort Key1:=.Range("A2"), Order1:=xlAscending, _
                         Header:=xlGuess, OrderCustom:=1, MatchCase _
                         :=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal, _
                         DataOption2:=xlSortNormal
                 End With
            End If
    'Mise à jour des données
    UpdateLBEnseignants

End Sub

Sub UpdateFraisList()
    
    
    Dim numRows As Integer
    numRows = GetSheetNumRows(Worksheets(Replace(LBenseignants.value, " ", "")))
    LB_listeFraisEns.ColumnHeads = True
    LB_listeFraisEns.RowSource = Worksheets(Replace(LBenseignants.value, " ", "")).name & "!A2:J" & numRows
   
   

End Sub


