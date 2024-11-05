VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_ComparaisonBudget 
   Caption         =   "Comparer les budgets"
   ClientHeight    =   11070
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   15225
   OleObjectBlob   =   "UF_ComparaisonBudget.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UF_ComparaisonBudget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CB_annee1_Change()
    Worksheets("comparaisonBudget").Range("B1").value = CB_annee1.value
    Worksheets("comparaisonBudget").Range("B2:B8").value = Worksheets("Budget" & CB_annee1.value).Range("B2:B8").value
    LB_comp.RowSource = Worksheets("comparaisonBudget").name & "!A2:C8"
    ShowGraph
End Sub

Private Sub CB_annee2_Change()
    Worksheets("comparaisonBudget").Range("C1").value = CB_annee2.value
    Worksheets("comparaisonBudget").Range("C2:C8").value = Worksheets("Budget" & CB_annee2.value).Range("B2:B8").value
    LB_comp.RowSource = Worksheets("comparaisonBudget").name & "!A2:C8"
    ShowGraph
End Sub

Private Sub CMB_back_Click()
    UF_ComparaisonBudget.Hide
End Sub

Private Sub CMB_Mail_Click()
    SendByMail ("ComparaisonBudget")
End Sub

Private Sub CMB_PDF_Click()
    Worksheets("ComparaisonBudget").Range("A1").Select
    
    ExportAsPDF ("ComparaisonBudget")
End Sub

Private Sub CMB_Print_Click()
    PrintSheet ("ComparaisonBudget")
End Sub

Private Sub UserForm_Initialize()

    'Nombre de colonnes dans la ListBox
    LB_comp.ColumnCount = 3
    'Largeur des colonnes de la ListBox
    LB_comp.ColumnWidths = "255;100;100"
    'titres de colonnes
    LB_comp.ColumnHeads = True
    
    Worksheets("ComparaisonBudget").Range("B1:B8").value = ""
    Worksheets("ComparaisonBudget").Range("C1:C8").value = ""
    
    UpdateCB SheetAnnees, CB_annee1
    UpdateCB SheetAnnees, CB_annee2
    
End Sub

Private Sub ShowGraph()

    Set CurrentChart = Sheets("ComparaisonBudget").ChartObjects(1).Chart
    Sheets("ComparaisonBudget").ChartObjects("GraphComp").Activate
    ActiveChart.ChartArea.Select
    
    fname = ThisWorkbook.Path & "\temp.gif"
    CurrentChart.Export Filename:=fname, FilterName:="GIF"

    IMG_graph.Picture = LoadPicture(fname)
    
    'Suppression du fichier
    Kill fname
End Sub
