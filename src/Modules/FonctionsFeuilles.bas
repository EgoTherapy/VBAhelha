Attribute VB_Name = "FonctionsFeuilles"

'Fonction pour trouver le nombre d'élément sur une feuille
Function GetSheetNumRows(sheet As Worksheet, Optional column As String = "A") As Integer
    GetSheetNumRows = sheet.Range(column & Rows.Count).End(xlUp).Row
End Function

'Fonction pour vérifier si une feuille existe dans le classeur
Function WorksheetExists(SheetName As String) As Boolean
    
Dim TempSheetName As String

TempSheetName = UCase(SheetName)

WorksheetExists = False
    
For Each sheet In Worksheets
    If TempSheetName = UCase(sheet.name) Then
        WorksheetExists = True
        Exit Function
    End If
Next sheet

End Function

'Procédure pour créer les feuilles relatives aux factures d'une année
Sub CreateSheetsFact(sheet1 As String)
    'création de la feuille de l'année et du bilan de l'année
    Sheets.Add.name = sheet1
    Sheets.Add.name = "Budget" & sheet1
    Worksheets(sheet1).Visible = xlSheetHidden
    Worksheets("Budget" & sheet1).Visible = xlSheetHidden

    'copier les informations de la Liste de facture type et les coller dans la feuille concernée
    Worksheets("ListeFactureType").Activate
    Range("A1:J1").Select
    Range("A1:J1").Copy
    
    Worksheets(sheet1).Activate
    Range("A1:J1").Select
    ActiveSheet.Paste
    Columns("C:C").Select
    Selection.NumberFormat = "#,##0.00 _€"
    ActiveSheet.EnableSelection = xlNoSelection
    Application.CutCopyMode = False
    
    'copier les informations de la feuille de Budget type et les coller dans la feuille concernée
    Worksheets("TypeBudget").Activate
    Range("A1:F50").Select
    Range("A1:F50").Copy
    
    Worksheets("Budget" & sheet1).Activate
    Range("A1:F50").Select
    ActiveSheet.Paste
    Columns("F:F").Select
    Selection.NumberFormat = "#,##0.00 _€"
    ActiveSheet.EnableSelection = xlNoSelection
    Application.CutCopyMode = False
    
    CreerGraph sheet1
End Sub

'Procédure pour ajouter une facture, a une feuille d'une année spécifique
Sub AddFacture(fact As Facture, sheet1 As String)
    
    
    'Nombre de factures
    Dim numRows As Integer
    numRows = GetSheetNumRows(Worksheets(sheet1)) + 1
    
    'ajout de la facture sur la feuille de l'année concernée
    With Worksheets(sheet1)
        .Cells(numRows, 1) = fact.num
        .Cells(numRows, 2) = fact.dateFact
        .Cells(numRows, 3) = fact.montant
        .Cells(numRows, 4) = fact.Fournisseur
        .Cells(numRows, 5) = fact.categorieFrais
        .Cells(numRows, 6) = fact.typeFrais
        .Cells(numRows, 7) = fact.objet
        .Cells(numRows, 8) = fact.concerne
        .Cells(numRows, 9) = fact.ens
        .Cells(numRows, 10) = fact.fichier
        'Tri sur les dates
        .Range("A2:J" & numRows).Sort Key1:=.Range("B2"), Order1:=xlAscending, _
            Header:=xlGuess, OrderCustom:=1, MatchCase _
            :=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal, _
            DataOption2:=xlSortNormal
    End With
    
    numRows = GetSheetNumRows(SheetFactures) + 1
    'ajout de la facture sur la feuille regroupant toutes les factures (toutes années confondue
        With SheetFactures
        .Cells(numRows, 1) = fact.num
        .Cells(numRows, 2) = fact.dateFact
        .Cells(numRows, 3) = fact.montant
        .Cells(numRows, 4) = fact.Fournisseur
        .Cells(numRows, 5) = fact.categorieFrais
        .Cells(numRows, 6) = fact.typeFrais
        .Cells(numRows, 7) = fact.objet
        .Cells(numRows, 8) = fact.concerne
        .Cells(numRows, 9) = fact.ens
        .Cells(numRows, 10) = fact.fichier
        'Tri sur les dates
        .Range("A2:J" & numRows).Sort Key1:=.Range("B2"), Order1:=xlAscending, _
            Header:=xlGuess, OrderCustom:=1, MatchCase _
            :=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal, _
            DataOption2:=xlSortNormal
    End With
    
    
    'ajout de la facture sur la feuille de l'enseignant concerné (si la feuille existe)
    If WorksheetExists(Replace(fact.ens, " ", "")) Then
        numRows = GetSheetNumRows(Worksheets(Replace(fact.ens, " ", ""))) + 1
        With Worksheets(Replace(fact.ens, " ", ""))
        .Cells(numRows, 1) = fact.num
        .Cells(numRows, 2) = fact.dateFact
        .Cells(numRows, 3) = fact.montant
        .Cells(numRows, 4) = fact.Fournisseur
        .Cells(numRows, 5) = fact.categorieFrais
        .Cells(numRows, 6) = fact.typeFrais
        .Cells(numRows, 7) = fact.objet
        .Cells(numRows, 8) = fact.concerne
        .Cells(numRows, 9) = fact.ens
        .Cells(numRows, 10) = fact.fichier
        'Tri sur les dates
        .Range("A2:J" & numRows).Sort Key1:=.Range("B2"), Order1:=xlAscending, _
            Header:=xlGuess, OrderCustom:=1, MatchCase _
            :=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal, _
            DataOption2:=xlSortNormal
        End With
    End If
    
    
End Sub

'Procédure pour generer une année d'activité
Sub AddYear(year As String)
    'Nombre d'enseignants
    Dim numRows As Integer
    numRows = GetSheetNumRows(SheetAnnees) + 1
    
    
    
    'Création de l'année
    With SheetAnnees
        .Cells(numRows, 1) = year
        
        'Tri sur les années
        .Range("A3:A" & numRows).Sort Key1:=.Range("A3"), Order1:=xlAscending, _
            Header:=xlGuess, OrderCustom:=1, MatchCase _
            :=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal, _
            DataOption2:=xlSortNormal
    End With
End Sub

'Méthode pour mettre à jour les combobox
Sub UpdateCB(sheet As Worksheet, comb As ComboBox)
    
     comb = Empty
    Dim nr As Integer
    
    nr = GetSheetNumRows(sheet)
    
    comb.RowSource = sheet.name & "!A2:A" & nr

End Sub

'méthode pour verifier si l'adresse e-mail est valide (de bon format)
Function IsValidEmail(sEmailAddress As String) As Boolean
    
    Dim sEmailPattern As String
    Dim oRegEx As Object
    Dim bReturn As Boolean
    
    'Utilise les expressions régulières suivantes
    sEmailPattern = "^\w+([\.-]?\w+)*@\w+([\.-]?\w+)*(\.\w{2,3})+$" 'or
    sEmailPattern = "^([a-zA-Z0-9_\-\.]+)@[a-z0-9-]+(\.[a-z0-9-]+)*(\.[a-z]{2,3})$"
    
    'Créer un objet d'expression régulière
    Set oRegEx = CreateObject("VBScript.RegExp")
    oRegEx.Global = True
    oRegEx.IgnoreCase = True
    oRegEx.Pattern = sEmailPattern
    bReturn = False
    
    'Vérifier si l'email correspond au modèle regex
    If oRegEx.Test(sEmailAddress) Then
        bReturn = True
    Else
        bReturn = False
    End If

    'Retourner le résultat de la validation
    IsValidEmail = bReturn
End Function

'Procédure pour créer un graph sur la feuille des budgets
Sub CreerGraph(year As String)
    'déclaration du graph
    Dim chrt As ChartObject
 
    'les propriétés du graph
    Set chrt = Sheets("Budget" & year).ChartObjects.Add(Left:=0, Width:=492, Top:=150, Height:=288)
    chrt.Chart.SetSourceData Source:=Sheets("Budget" & year).Range("A2:B8")
    chrt.Chart.ChartType = xlPie
    
    'définir l'automatisation du titre du graph
    With chrt
        .Chart.HasTitle = True
        .Chart.ChartTitle.Text = "Budget previsionnel en " & year
    End With

    
End Sub

'Procédure pour exporter une feuille en pdf
Sub ExportAsPDF(name As String)
    

    Dim iVis As XlSheetVisibility

    'definir les propriétés de la page, afin que le résultat tienne en une page
    With Worksheets(name).PageSetup
        
        .CenterVertically = False
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .BottomMargin = 0
        .TopMargin = 0
        .RightMargin = 0
        .LeftMargin = 0
    End With
    
    
    'ici on va rendre la page visible afin de generer le pdf puis la remettre invisible
    With Worksheets(name)
        .Columns("A:Z").AutoFit
        
        iVis = .Visible
        .Visible = xlSheetVisible
        'exporter au format PDF :
        .ExportAsFixedFormat Type:=xlTypePDF, _
                             Filename:=Application.ActiveWorkbook.Path & "\" & name, _
                             Quality:=xlQualityStandard, _
                             IncludeDocProperties:=True, _
                             IgnorePrintAreas:=False, _
                             OpenAfterPublish:=True
        .Visible = xlSheetHidden
        
    End With

End Sub

'Procédure pour imprimer une feuille
Sub PrintSheet(name As String)
    

    Dim iVis As XlSheetVisibility

    'definir les propriétés de la page, afin que le résultat tienne en une page
    With Worksheets(name).PageSetup
        
        .CenterVertically = False
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .BottomMargin = 0
        .TopMargin = 0
        .RightMargin = 0
        .LeftMargin = 0
    End With
    
    
    'ici on va rendre la page visible afin de pouvoir l'imprimer
    With Worksheets(name)
        .Columns("A:Z").AutoFit
        
        iVis = .Visible
        .Visible = xlSheetVisible
        .PrintOut 'impression de la page
        .Visible = xlSheetHidden 'rendre la page à nouveau invisible
        
    End With

End Sub

