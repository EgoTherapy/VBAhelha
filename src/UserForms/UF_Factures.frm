VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_Factures 
   Caption         =   "Factures"
   ClientHeight    =   10260
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   17805
   OleObjectBlob   =   "UF_Factures.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UF_Factures"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Text

Dim bddFeuil1, plageFeuil2, critere

Private Sub CB_annee_Change()
Dim nr As Integer

If WorksheetExists(CB_annee.value) Or CB_annee.value = "Toutes" Then
    If CB_annee.value = "Toutes" Then
        UpdateFactList Worksheets("Factures")
    Else
        UpdateFactList Worksheets(CB_annee.value)
    End If
        CB_searchCol.Enabled = True
    Else
    If CB_annee.value <> "Toutes" Then
            MsgBox "Aucune facture n'a été encodée pour cette année", vbExclamation
    End If
 End If
   
End Sub

Private Sub CB_searchCol_Change()
    TB_searchKeyWords.Enabled = True
    
End Sub

Private Sub CMB_AddFacture_Click()
    UF_NewFacture.Show
End Sub

Private Sub CMB_BackMenu_Click()
    UF_Factures.Hide
End Sub

Private Sub CMB_PDF_Click()
    If CB_annee.value = "Toutes" Or CB_annee.value = "" Then
        ExportAsPDF "Factures"
    Else
        ExportAsPDF CB_annee.value
        
    End If
End Sub

Private Sub CMB_Print_Click()
    If CB_annee.value = "Toutes" Or CB_annee.value = "" Then
        PrintSheet "Factures"
    Else
        PrintSheet CB_annee.value
        
    End If
End Sub

Private Sub CMB_search_Click()
    SearchFact
End Sub

Private Sub TB_searchKeyWords_AfterUpdate()
    If CB_searchCol.value = "Montant" Then
        If (IsNumeric(TB_searchKeyWords.value) = True) Then
            TB_searchKeyWords.value = Replace(TB_searchKeyWords.value, ",", ".")
        Else
            MsgBox "Veuillez entrer une valeur numerique !", vbCritical
            TB_searchKeyWords.value = ""
        End If
    End If
End Sub

Private Sub UserForm_Initialize()
        UpdateCB SheetAnnees, CB_annee
        
        'Nombre de colonnes dans la ListBox
        LB_factures.ColumnCount = 10
        'Largeur des colonnes de la ListBox
        LB_factures.ColumnWidths = "50;60;70;170;200;200;200;110"
         'initialisation : affichage de toutes les factures
        UpdateFactList Worksheets("Factures")
        
        CB_searchCol.Enabled = False
        TB_searchKeyWords.Enabled = False
        
        Sheets("ListeFactureType").Activate
        CB_searchCol.List = Application.Transpose(Range("A1:J1"))
        

End Sub

Private Sub UpdateFactList(sheet As Worksheet)
    LB_factures = Empty
    Dim nr As Integer
   
    nr = GetSheetNumRows(sheet)
    LB_factures.ColumnHeads = True
    LB_factures.RowSource = sheet.name & "!A2:J" & nr
    
    

End Sub

Private Sub SearchFact()
Dim tmp As String

If CB_annee.value = "Toutes" Or CB_annee.value = "" Then
    tmp = "Factures"
Else
    tmp = CB_annee.value
End If

Dim nr As Integer
    nr = GetSheetNumRows(Worksheets(tmp))
    sheetFiltrage.Cells.Clear
        sheetFiltrage.[P1] = CB_searchCol.value
    If CB_searchCol.value <> "Montant" And CB_searchCol <> "Date" Then
        sheetFiltrage.[P2] = Me.TB_searchKeyWords.value & "*"
    Else
        sheetFiltrage.[P2] = Me.TB_searchKeyWords.value
        Select Case CB_searchCol.value
        Case "Montant"
            sheetFiltrage.[P2].NumberFormat = "General"
        Case "Date"
            sheetFiltrage.[P2].value = CDate(sheetFiltrage.[P2].value)
        End Select
    End If
    Sheets(tmp).Range("A1:J" & nr).AdvancedFilter Action:=xlFilterCopy, _
    criteriarange:=sheetFiltrage.Range("P1:P2"), _
    copytorange:=sheetFiltrage.[A1], Unique:=False
    
    If sheetFiltrage.[A1].CurrentRegion.Rows.Count > 1 Then
        Set plageFeuil2 = sheetFiltrage.[A1].CurrentRegion.Offset(1).Resize(sheetFiltrage.[A1].CurrentRegion.Rows.Count - 1)
        Me.LB_factures.RowSource = plageFeuil2.Address(external:=True)
    End If
End Sub


