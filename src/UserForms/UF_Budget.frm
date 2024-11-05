VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_Budget 
   Caption         =   "Gestion du budget"
   ClientHeight    =   11850
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   12000
   OleObjectBlob   =   "UF_Budget.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UF_Budget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CB_annee_Change()
    If Not WorksheetExists(CB_annee.value) Then
        CreateSheetsFact CB_annee.value
    End If

    Worksheets("Budget" & CB_annee.value).Activate

    TB_entretiens.value = Format(Range("B2").value, "Standard")
    TB_telecom.value = Format(Range("B3").value, "Standard")
    TB_autresFourn.value = Format(Range("B4").value, "Standard")
    TB_retrib.value = Format(Range("B5").value, "Standard")
    TB_infos.value = Format(Range("B6").value, "Standard")
    TB_assurances.value = Format(Range("B7").value, "Standard")
    TB_autres.value = Format(Range("B8").value, "Standard")

    TB_entretiensDep.value = Format(Range("F2").value, "Standard")
    TB_telecomDep.value = Format(Range("F6").value, "Standard")
    TB_autresFournDep.value = Format(Range("F10").value, "Standard")
    TB_retribDep.value = Format(Range("F26").value, "Standard")
    TB_infosDep.value = Format(Range("F31").value, "Standard")
    TB_assurancesDep.value = Format(Range("F41").value, "Standard")
    TB_autresDep.value = Format(Range("F49").value, "Standard")
    
    CMB_modifier.Enabled = True
    
    
    Cadre_budget.Caption = "Budget prévisionnel en " & CB_annee.value
    Cadre_depenses.Caption = "Dépenses en " & CB_annee.value
    
    ShowGraph 1
    
End Sub

Private Sub CMB_back_Click()
    UF_Budget.Hide
End Sub

Private Sub CMB_compare_Click()
    UF_ComparaisonBudget.Show
End Sub

'modification du budget sur la feuille
Private Sub CMB_confirmModification_Click()
    
    
    Dim b As New Budget
    With b
        .entretiens = TB_entretiens.value
        .telecom = TB_telecom.value
        .autresFourn = TB_autresFourn.value
        .retrib = TB_retrib.value
        .infos = TB_infos.value
        .assurances = TB_assurances.value
        .autres = TB_autres.value
    End With
    
    With Worksheets("Budget" & CB_annee.value)
        .Range("B2").value = b.entretiens
        .Range("B3").value = b.telecom
        .Range("B4").value = b.autresFourn
        .Range("B5").value = b.retrib
        .Range("B6").value = b.infos
        .Range("B7").value = b.assurances
        .Range("B8").value = b.autres
    End With

    'rafraichir le graph
    ShowGraph 1
End Sub

Private Sub CMB_modifier_Click()
    If CMB_confirmModification.Enabled = False Then
        CMB_confirmModification.Enabled = True
        CMB_modifier.Caption = "Annuler la modification"
        
        TB_entretiens.Enabled = True
        TB_telecom.Enabled = True
        TB_autresFourn.Enabled = True
        TB_retrib.Enabled = True
        TB_infos.Enabled = True
        TB_assurances.Enabled = True
        TB_autres.Enabled = True
    Else
        CMB_confirmModification.Enabled = False
        CMB_modifier.Caption = "modifier"
        
        TB_entretiens.Enabled = False
        TB_telecom.Enabled = False
        TB_autresFourn.Enabled = False
        TB_retrib.Enabled = False
        TB_infos.Enabled = False
        TB_assurances.Enabled = False
        TB_autres.Enabled = False
    End If
End Sub

Private Sub CMB_PDF_Click()
    Worksheets("Budget" & CB_annee.value).Range("A1").Select
    
    ExportAsPDF ("Budget" & CB_annee.value)
End Sub

Private Sub CMB_Print_Click()
    PrintSheet ("Budget" & CB_annee.value)
End Sub

Private Sub TB_entretiens_AfterUpdate()
    If Not IsNumeric(TB_entretiens.value) Then
        MsgBox "Veuillez entrer une valeur numérique", vbCritical
        TB_entretiens.value = Format(Range("B2").value, "Standard")
    Else
        If TB_entretiens.value < 0 Then
            MsgBox "Veuillez entrer une valeur positive (supérieure ou égale à 0).", vbExclamation
            TB_entretiens.value = Format(Range("B2").value, "Standard")
        Else
            TB_entretiens.value = Format(TB_entretiens.value, "Standard")
        End If
    End If
End Sub

Private Sub TB_telecom_AfterUpdate()
    If Not IsNumeric(TB_telecom.value) Then
        MsgBox "Veuillez entrer une valeur numérique", vbCritical
        TB_telecom.value = Format(Range("B3").value, "Standard")
    Else
        If TB_telecom.value < 0 Then
            MsgBox "Veuillez entrer une valeur positive (supérieure ou égale à 0).", vbExclamation
            TB_telecom.value = Format(Range("B3").value, "Standard")
        Else
            TB_telecom.value = Format(TB_telecom.value, "Standard")
        End If
    End If
End Sub

Private Sub TB_autresFourn_AfterUpdate()
    If Not IsNumeric(TB_autresFourn.value) Then
        MsgBox "Veuillez entrer une valeur numérique", vbCritical
        TB_autresFourn.value = Format(Range("B4").value, "Standard")
    Else
        If TB_autresFourn.value < 0 Then
            MsgBox "Veuillez entrer une valeur positive (supérieure ou égale à 0).", vbExclamation
            TB_autresFourn.value = Format(Range("B4").value, "Standard")
        Else
            TB_autresFourn.value = Format(TB_autresFourn.value, "Standard")
        End If
    End If
End Sub

Private Sub TB_retrib_AfterUpdate()
    If Not IsNumeric(TB_retrib.value) Then
        MsgBox "Veuillez entrer une valeur numérique", vbCritical
        TB_retrib.value = Format(Range("B5").value, "Standard")
    Else
        If TB_retrib.value < 0 Then
            MsgBox "Veuillez entrer une valeur positive (supérieure ou égale à 0).", vbExclamation
            TB_retrib.value = Format(Range("B5").value, "Standard")
        Else
            TB_retrib.value = Format(TB_retrib.value, "Standard")
        End If
    End If
End Sub

Private Sub TB_infos_AfterUpdate()
    If Not IsNumeric(TB_infos.value) Then
        MsgBox "Veuillez entrer une valeur numérique", vbCritical
        TB_infos.value = Format(Range("B6").value, "Standard")
    Else
        If TB_infos.value < 0 Then
            MsgBox "Veuillez entrer une valeur positive (supérieure ou égale à 0).", vbExclamation
            TB_infos.value = Format(Range("B6").value, "Standard")
        Else
            TB_infos.value = Format(TB_infos.value, "Standard")
        End If
    End If
End Sub

Private Sub TB_assurances_AfterUpdate()
    If Not IsNumeric(TB_assurances.value) Then
        MsgBox "Veuillez entrer une valeur numérique", vbCritical
        TB_assurances.value = Format(Range("B2").value, "Standard")
    Else
        If TB_assurances.value < 0 Then
            MsgBox "Veuillez entrer une valeur positive (supérieure ou égale à 0).", vbExclamation
            TB_assurances.value = Format(Range("B7").value, "Standard")
        Else
            TB_assurances.value = Format(TB_assurances.value, "Standard")
        End If
    End If
End Sub

Private Sub TB_autres_AfterUpdate()
    If Not IsNumeric(TB_autres.value) Then
        MsgBox "Veuillez entrer une valeur numérique", vbCritical
        TB_autres.value = Format(Range("B2").value, "Standard")
    Else
        If TB_autres.value < 0 Then
            MsgBox "Veuillez entrer une valeur positive (supérieure ou égale à 0).", vbExclamation
            TB_autres.value = Format(Range("B8").value, "Standard")
        Else
            TB_autres.value = Format(TB_autres.value, "Standard")
        End If
    End If
End Sub

Private Sub UserForm_Initialize()
    UpdateCB SheetAnnees, CB_annee
    
    'Budget
    TB_entretiens.Enabled = False
    TB_telecom.Enabled = False
    TB_autresFourn.Enabled = False
    TB_retrib.Enabled = False
    TB_infos.Enabled = False
    TB_assurances.Enabled = False
    TB_autres.Enabled = False
    
    'dépenses
    TB_entretiensDep.Enabled = False
    TB_telecomDep.Enabled = False
    TB_autresFournDep.Enabled = False
    TB_retribDep.Enabled = False
    TB_infosDep.Enabled = False
    TB_assurancesDep.Enabled = False
    TB_autresDep.Enabled = False
    
    'bouttons
    CMB_confirmModification.Enabled = False
    CMB_modifier = False
End Sub

'verifier si la valeur d'une text box est numéérique
Private Sub TBIsNumeric(tb As TextBox)
    If Not IsNumeric(tb.value) Then
        MsgBox "Veuillez entrer une valeur numérique", vbCritical
        tb.value = Format(Range("B2").value, "Standard")
    Else
        tb.value = Format(tb.value, "Standard")
    End If
End Sub

'méthode pour afficher un graph dans le formulaire
Private Sub ShowGraph(nb As Integer)

    'il faut enregistrer l'image du graph dans un fichier à part puis afficher l'image dans le formulaire
    Set CurrentChart = Sheets("Budget" & CB_annee.value).ChartObjects(1).Chart
    
    Sheets("Budget" & CB_annee.value).ChartObjects("Graphique " & nb).Activate
    ActiveChart.ChartArea.Select
    
    fname = ThisWorkbook.Path & "\temp.gif"
    CurrentChart.Export Filename:=fname, FilterName:="GIF"

    IMG_graph1.Picture = LoadPicture(fname)
    
    'Suppression du fichier
    Kill fname
End Sub
