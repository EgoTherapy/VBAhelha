VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_NewFacture 
   Caption         =   "Nouvelle facture"
   ClientHeight    =   5970
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   7650
   OleObjectBlob   =   "UF_NewFacture.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UF_NewFacture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CB_CatFrais_Change()
    CB_TypeFrais.Enabled = True
    
    Select Case CB_CatFrais.value
    Case "ENTRETIENS ET REPARATIONS"
        CB_TypeFrais.RowSource = SheetTypeFrais.name & "!A2:A4"
    Case "TELECOMMUNICATIONS ET FRAIS DE PORT"
        CB_TypeFrais.RowSource = SheetTypeFrais.name & "!B2:B4"
    Case "AUTRES FOURNITURES"
        CB_TypeFrais.RowSource = SheetTypeFrais.name & "!C2:C16"
    
    Case "RETRIBUTIONS DE TIERS"
        CB_TypeFrais.RowSource = SheetTypeFrais.name & "!D2:D5"
    
    Case "INFORMATIONS, PUBLICITES, RECEPTIONS, ACTIVITES PEDAGOGIQUES"

        CB_TypeFrais.RowSource = SheetTypeFrais.name & "!E2:E10"
    Case "ASSURANCES ET DEPLACEMENTS"
        CB_TypeFrais.RowSource = SheetTypeFrais.name & "!f2:f8"
    Case "AUTRES"
        CB_TypeFrais.Enabled = False
End Select
    CB_TypeFrais = Empty
End Sub



Private Sub CB_concerne_Change()
    If (CB_concerne.value = "Campus" Or CB_concerne.value = "Département économique") Then
        CB_enseignants.Enabled = False
    Else
        CB_enseignants.Enabled = True
    End If
    
End Sub

Private Sub CMB_add_Click()
If TB_date = "" And TB_montant = "" And TB_objet = "" And CB_fournisseur = "" And CB_CatFrais = "" And CB_concerne = "" Then
    MsgBox "Veuillez completer les champs obligatoires!", vbCritical
    'réinitialisation du userForm
    UserForm_Initialize
Else
    Dim nr As Integer
    Dim year As String
    year = Format(TB_date, "yyyy")
    If Not WorksheetExists(year) Then
        CreateSheetsFact year
        AddYear year
    End If
    If CB_enseignants.value <> "" Then
        If Not WorksheetExists(Replace(CB_enseignants.value, " ", "")) Then
            CreateSheetsFact Replace(CB_enseignants.value, " ", "")
        End If
    End If
    nr = GetSheetNumRows(Worksheets(year))
    
    Dim fact As New Facture
    With fact
        .num = year & "-" & Format(nr, "000")
        .dateFact = TB_date.value
        .montant = TB_montant.value
        .Fournisseur = CB_fournisseur.value
        .categorieFrais = CB_CatFrais.value
        .typeFrais = CB_TypeFrais.value
        .objet = TB_objet
        .concerne = CB_concerne
        .ens = CB_enseignants.value
        .fichier = fact.num & ".pdf"
    End With
    
    
    
    AddFacture fact, year
    AddTypeFrais year
    
    MsgBox "La facture a bien été ajoutée à la base de données", vbInformation
    
    'réinitialisation du userForm
    UserForm_Initialize
End If
End Sub

Private Sub CMB_cancel_Click()
    UF_NewFacture.Hide
End Sub



Private Sub TB_date_AfterUpdate()

    If Not IsDate(TB_date) Then
        MsgBox "La date est incorrecte", vbCritical
        TB_date = ""
        Exit Sub
    End If
    
End Sub


Private Sub TB_montant_AfterUpdate()
    'mise de la textbox "montant" sous format monétaire
    TB_montant.value = Format(TB_montant.value, "Standard")
    If ((IsNumeric(TB_montant.value) = False) And (TB_montant.value <> "")) Then
    
        MsgBox "Le montant entré est invalide.", vbCritical
        TB_montant = Empty
        
    End If
    If TB_montant.value < 0 Then
        MsgBox "Le montant ne peut pas être négatif.", vbExclamation
        TB_montant = Empty
    End If
End Sub



Private Sub UserForm_Initialize()
   UpdateCB SheetEnseignants, CB_enseignants
   UpdateCB SheetFournisseurs, CB_fournisseur
   Sheets("TypeFrais").Activate
   CB_CatFrais.List = Application.Transpose(Range("A1:G1"))
   With CB_concerne
        .AddItem "Campus"
        .AddItem "Département économique"
        .AddItem "Section Comptabilité"
        .AddItem "Section Informatique de gestion"
        .AddItem "Section Assistance de direction"
   End With
   CB_TypeFrais.Enabled = False
   
   EmptyFields
End Sub

Private Sub EmptyFields()
    TB_date = Empty
    TB_montant = Empty
    CB_fournisseur = Empty
    CB_CatFrais = Empty
    CB_TypeFrais = Empty
    TB_objet = Empty
    CB_concerne = Empty
    CB_enseignants = Empty
End Sub

Private Sub AddTypeFrais(year As String)
    Worksheets("Budget" & year).Activate
    Select Case CB_TypeFrais.value
    'ENTRETIENS ET REPARATIONS
    Case "Entreprises extérieures de nettoyage"
        Range("F2").value = Range("F2").value + TB_montant.value
        Range("F3").value = Range("F3").value + TB_montant.value
        If Range("F2").value > Range("B2").value Then
            MsgBox "vous dépassez le budget alloué à 'ENTRETIENS ET REPARATIONS'", vbExclamation
        End If
    Case "Entretien mobilier et matériel"
        Range("F2").value = Range("F2").value + TB_montant.value
        Range("F4").value = Range("F4").value + TB_montant.value
        If Range("F2").value > Range("B2").value Then
            MsgBox "vous dépassez le budget alloué à 'ENTRETIENS ET REPARATIONS'", vbExclamation
        End If
    Case "Petits travaux"
        Range("F2").value = Range("F2").value + TB_montant.value
        Range("F5").value = Range("F5").value + TB_montant.value
        If Range("F2").value > Range("B2").value Then
            MsgBox "vous dépassez le budget alloué à 'ENTRETIENS ET REPARATIONS'", vbExclamation
        End If
        
    'TELECOMMUNICATIONS ET FRAIS DE PORT
    Case "Téléphone, Intenet, fax, télédistribution"
        Range("F6").value = Range("F6").value + TB_montant.value
        Range("F7").value = Range("F7").value + TB_montant.value
        If Range("F6").value > Range("B3").value Then
            MsgBox "vous dépassez le budget alloué à 'TELECOMMUNICATIONS ET FRAIS DE PORT'", vbExclamation
        End If
    Case "Gsm"
        Range("F6").value = Range("F6").value + TB_montant.value
        Range("F8").value = Range("F8").value + TB_montant.value
        If Range("F6").value > Range("B3").value Then
            MsgBox "vous dépassez le budget alloué à 'TELECOMMUNICATIONS ET FRAIS DE PORT'", vbExclamation
        End If
    Case "Frais postaux (timbres, recommandés, colis)"
        Range("F6").value = Range("F6").value + TB_montant.value
        Range("F9").value = Range("F9").value + TB_montant.value
        If Range("F6").value > Range("B3").value Then
            MsgBox "vous dépassez le budget alloué à 'TELECOMMUNICATIONS ET FRAIS DE PORT'", vbExclamation
        End If
        
    'AUTRES FOURNITURES
    Case "Photocopies faites à l'extérieur"
        Range("F10").value = Range("F10").value + TB_montant.value
        Range("F11").value = Range("F11").value + TB_montant.value
        If Range("F10").value > Range("B4").value Then
            MsgBox "vous dépassez le budget alloué à 'AUTRES FOURNITURES'", vbExclamation
        End If
    Case "Matériel de bureau"
        Range("F10").value = Range("F10").value + TB_montant.value
        Range("F12").value = Range("F12").value + TB_montant.value
        If Range("F10").value > Range("B4").value Then
            MsgBox "vous dépassez le budget alloué à 'AUTRES FOURNITURES'", vbExclamation
        End If
    Case "Matériel pédagogique"
        Range("F10").value = Range("F10").value + TB_montant.value
        Range("F13").value = Range("F13").value + TB_montant.value
        If Range("F10").value > Range("B4").value Then
            MsgBox "vous dépassez le budget alloué à 'AUTRES FOURNITURES'", vbExclamation
        End If
    Case "Matériel informatique de bureau"
        Range("F10").value = Range("F10").value + TB_montant.value
        Range("F14").value = Range("F14").value + TB_montant.value
        If Range("F10").value > Range("B4").value Then
            MsgBox "vous dépassez le budget alloué à 'AUTRES FOURNITURES'", vbExclamation
        End If
    Case "Matériel informatique pédagogique"
        Range("F10").value = Range("F10").value + TB_montant.value
        Range("F15").value = Range("F15").value + TB_montant.value
        If Range("F10").value > Range("B4").value Then
            MsgBox "vous dépassez le budget alloué à 'AUTRES FOURNITURES'", vbExclamation
        End If
    Case "Matériel de réunion et cours à distance"
        Range("F10").value = Range("F10").value + TB_montant.value
        Range("F16").value = Range("F16").value + TB_montant.value
        If Range("F10").value > Range("B4").value Then
            MsgBox "vous dépassez le budget alloué à 'AUTRES FOURNITURES'", vbExclamation
        End If
    Case "Mobilier de bureau"
        Range("F10").value = Range("F10").value + TB_montant.value
        Range("F17").value = Range("F17").value + TB_montant.value
        If Range("F10").value > Range("B4").value Then
            MsgBox "vous dépassez le budget alloué à 'AUTRES FOURNITURES'", vbExclamation
        End If
    Case "Mobilier pédagogique"
        Range("F10").value = Range("F10").value + TB_montant.value
        Range("F18").value = Range("F18").value + TB_montant.value
        If Range("F10").value > Range("B4").value Then
            MsgBox "vous dépassez le budget alloué à 'AUTRES FOURNITURES'", vbExclamation
        End If
    Case "Location mobilier"
        Range("F10").value = Range("F10").value + TB_montant.value
        Range("F19").value = Range("F19").value + TB_montant.value
        If Range("F10").value > Range("B4").value Then
            MsgBox "vous dépassez le budget alloué à 'AUTRES FOURNITURES'", vbExclamation
        End If
    Case "Location immobilier"
        Range("F10").value = Range("F10").value + TB_montant.value
        Range("F20").value = Range("F20").value + TB_montant.value
        If Range("F10").value > Range("B4").value Then
            MsgBox "vous dépassez le budget alloué à 'AUTRES FOURNITURES'", vbExclamation
        End If
    Case "Abonnements informatiques et licences de bureau/informatiques"
        Range("F10").value = Range("F10").value + TB_montant.value
        Range("F21").value = Range("F21").value + TB_montant.value
        If Range("F10").value > Range("B4").value Then
            MsgBox "vous dépassez le budget alloué à 'AUTRES FOURNITURES'", vbExclamation
        End If
    Case "Abonnements informatiques et licences pédagogiques"
        Range("F10").value = Range("F10").value + TB_montant.value
        Range("F22").value = Range("F22").value + TB_montant.value
        If Range("F10").value > Range("B4").value Then
            MsgBox "vous dépassez le budget alloué à 'AUTRES FOURNITURES'", vbExclamation
        End If
    Case "Bibliothèque"
        Range("F10").value = Range("F10").value + TB_montant.value
        Range("F23").value = Range("F23").value + TB_montant.value
        If Range("F10").value > Range("B4").value Then
            MsgBox "vous dépassez le budget alloué à 'AUTRES FOURNITURES'", vbExclamation
        End If
    Case "Outillage"
        Range("F10").value = Range("F10").value + TB_montant.value
        Range("F24").value = Range("F24").value + TB_montant.value
        If Range("F10").value > Range("B4").value Then
            MsgBox "vous dépassez le budget alloué à 'AUTRES FOURNITURES'", vbExclamation
        End If
    Case "Autre matériel"
        Range("F10").value = Range("F10").value + TB_montant.value
        Range("F25").value = Range("F25").value + TB_montant.value
        If Range("F10").value > Range("B4").value Then
            MsgBox "vous dépassez le budget alloué à 'AUTRES FOURNITURES'", vbExclamation
        End If
    
    'RETRIBUTIONS DE TIERS
    Case "Vacataires"
        Range("F26").value = Range("F26").value + TB_montant.value
        Range("F27").value = Range("F27").value + TB_montant.value
        If Range("F26").value > Range("B5").value Then
            MsgBox "vous dépassez le budget alloué à 'RETRIBUTIONS DE TIERS'", vbExclamation
        End If
    Case "Conférenciers, autres intervenants"
        Range("F26").value = Range("F26").value + TB_montant.value
        Range("F28").value = Range("F28").value + TB_montant.value
        If Range("F26").value > Range("B5").value Then
            MsgBox "vous dépassez le budget alloué à 'RETRIBUTIONS DE TIERS'", vbExclamation
        End If
    Case "Frais de formation continuée personnel Helha"
        Range("F26").value = Range("F26").value + TB_montant.value
        Range("F29").value = Range("F29").value + TB_montant.value
        If Range("F26").value > Range("B5").value Then
            MsgBox "vous dépassez le budget alloué à 'RETRIBUTIONS DE TIERS'", vbExclamation
        End If
    Case "Cotisations, affiliations aux organismes"
        Range("F26").value = Range("F26").value + TB_montant.value
        Range("F30").value = Range("F30").value + TB_montant.value
        If Range("F26").value > Range("B5").value Then
            MsgBox "vous dépassez le budget alloué à 'RETRIBUTIONS DE TIERS'", vbExclamation
        End If
        
    'INFORMATIONS, PUBLICITES, RECEPTIONS, ACTIVITES PEDAGOGIQUES
    Case "Publicité, salons"
        Range("F31").value = Range("F31").value + TB_montant.value
        Range("F32").value = Range("F32").value + TB_montant.value
        If Range("F31").value > Range("B6").value Then
            MsgBox "vous dépassez le budget alloué à 'INFORMATIONS, PUBLICITES, RECEPTIONS, ACTIVITES PEDAGOGIQUES'", vbExclamation
        End If
    Case "Autres publications"
        Range("F31").value = Range("F31").value + TB_montant.value
        Range("F33").value = Range("F33").value + TB_montant.value
        If Range("F31").value > Range("B6").value Then
            MsgBox "vous dépassez le budget alloué à 'INFORMATIONS, PUBLICITES, RECEPTIONS, ACTIVITES PEDAGOGIQUES'", vbExclamation
        End If
    Case "Réceptions, réunions"
        Range("F31").value = Range("F31").value + TB_montant.value
        Range("F34").value = Range("F34").value + TB_montant.value
        If Range("F31").value > Range("B6").value Then
            MsgBox "vous dépassez le budget alloué à 'INFORMATIONS, PUBLICITES, RECEPTIONS, ACTIVITES PEDAGOGIQUES'", vbExclamation
        End If
    Case "Cadeaux hors jurys"
        Range("F31").value = Range("F31").value + TB_montant.value
        Range("F35").value = Range("F35").value + TB_montant.value
        If Range("F31").value > Range("B6").value Then
            MsgBox "vous dépassez le budget alloué à 'INFORMATIONS, PUBLICITES, RECEPTIONS, ACTIVITES PEDAGOGIQUES'", vbExclamation
        End If
    Case "Fleurs"
        Range("F31").value = Range("F31").value + TB_montant.value
        Range("F36").value = Range("F36").value + TB_montant.value
        If Range("F31").value > Range("B6").value Then
            MsgBox "vous dépassez le budget alloué à 'INFORMATIONS, PUBLICITES, RECEPTIONS, ACTIVITES PEDAGOGIQUES'", vbExclamation
        End If
    Case "Frais de réception des jurys"
        Range("F31").value = Range("F31").value + TB_montant.value
        Range("F37").value = Range("F37").value + TB_montant.value
        If Range("F31").value > Range("B6").value Then
            MsgBox "vous dépassez le budget alloué à 'INFORMATIONS, PUBLICITES, RECEPTIONS, ACTIVITES PEDAGOGIQUES'", vbExclamation
        End If
    Case "Animations pédagogiques, visites et voyages culturels"
        Range("F31").value = Range("F31").value + TB_montant.value
        Range("F38").value = Range("F38").value + TB_montant.value
        If Range("F31").value > Range("B6").value Then
            MsgBox "vous dépassez le budget alloué à 'INFORMATIONS, PUBLICITES, RECEPTIONS, ACTIVITES PEDAGOGIQUES'", vbExclamation
        End If
    Case "Frais séjour échanges nationaux et internationaux"
        Range("F31").value = Range("F31").value + TB_montant.value
        Range("F39").value = Range("F39").value + TB_montant.value
        If Range("F31").value > Range("B6").value Then
            MsgBox "vous dépassez le budget alloué à 'INFORMATIONS, PUBLICITES, RECEPTIONS, ACTIVITES PEDAGOGIQUES'", vbExclamation
        End If
    Case "Horeca"
        Range("F31").value = Range("F31").value + TB_montant.value
        Range("F40").value = Range("F40").value + TB_montant.value
        If Range("F31").value > Range("B6").value Then
            MsgBox "vous dépassez le budget alloué à 'INFORMATIONS, PUBLICITES, RECEPTIONS, ACTIVITES PEDAGOGIQUES'", vbExclamation
        End If

    'ASSURANCES ET DEPLACEMENTS
    Case "Assurances"
        Range("F41").value = Range("F41").value + TB_montant.value
        Range("F42").value = Range("F42").value + TB_montant.value
        If Range("F41").value > Range("B7").value Then
            MsgBox "vous dépassez le budget alloué à 'ASSURANCES ET DEPLACEMENTS'", vbExclamation
        End If
    Case "Déplacements missions"
        Range("F41").value = Range("F41").value + TB_montant.value
        Range("F43").value = Range("F43").value + TB_montant.value
        If Range("F41").value > Range("B7").value Then
            MsgBox "vous dépassez le budget alloué à 'ASSURANCES ET DEPLACEMENTS'", vbExclamation
        End If
    Case "Déplacements visites de stage"
        Range("F41").value = Range("F41").value + TB_montant.value
        Range("F44").value = Range("F44").value + TB_montant.value
        If Range("F41").value > Range("B7").value Then
            MsgBox "vous dépassez le budget alloué à 'ASSURANCES ET DEPLACEMENTS'", vbExclamation
        End If
    Case "Déplacements étudiants"
        Range("F41").value = Range("F41").value + TB_montant.value
        Range("F45").value = Range("F45").value + TB_montant.value
        If Range("F41").value > Range("B7").value Then
            MsgBox "vous dépassez le budget alloué à 'ASSURANCES ET DEPLACEMENTS'", vbExclamation
        End If
    Case "Déplacements de tiers"
        Range("F41").value = Range("F41").value + TB_montant.value
        Range("F46").value = Range("F46").value + TB_montant.value
        If Range("F41").value > Range("B7").value Then
            MsgBox "vous dépassez le budget alloué à 'ASSURANCES ET DEPLACEMENTS'", vbExclamation
        End If
    Case "Déplacements domicile-lieu de travail"
        Range("F41").value = Range("F41").value + TB_montant.value
        Range("F47").value = Range("F47").value + TB_montant.value
        If Range("F41").value > Range("B7").value Then
            MsgBox "vous dépassez le budget alloué à 'ASSURANCES ET DEPLACEMENTS'", vbExclamation
        End If
    Case "Déplacements encadrement des étudiants"
        Range("F41").value = Range("F41").value + TB_montant.value
        Range("F48").value = Range("F48").value + TB_montant.value
        If Range("F41").value > Range("B7").value Then
            MsgBox "vous dépassez le budget alloué à 'ASSURANCES ET DEPLACEMENTS'", vbExclamation
        End If
        
    'AUTRES
    Case Else
        Range("F49").value = Range("F49").value + TB_montant.value
        Range("F50").value = Range("F50").value + TB_montant.value
        If Range("F49").value > Range("B8").value Then
            MsgBox "vous dépassez le budget alloué à 'AUTRES'", vbExclamation
        End If
End Select
End Sub
