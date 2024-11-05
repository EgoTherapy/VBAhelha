Attribute VB_Name = "Dictionaires"
Global ensDictionary As Object
Global factDictionary As Object
Global fournDictionary As Object
'Methode d'initialisation dictionnaire des enseignants
Sub ComputeEnsDictionary()
    'Création du dictionnaire
    Set ensDictionary = CreateObject("Scripting.Dictionary")

    Dim x As Integer
    Dim e As Enseignant
    Dim numRows As Integer
    
    Application.ScreenUpdating = False
    
    With SheetEnseignants
        'Nombre d' enseignants
        numRows = GetSheetNumRows(SheetEnseignants)
        
        'Création de l'enseignant
        For x = 2 To numRows
            Set e = New Enseignant
            With e
                .NomPrenom = SheetEnseignants.Cells(x, 1)
            End With
            
            'Ajout de l'enseignant au dictionnaire
            ensDictionary.Add Key:=e.NomPrenom, Item:=e
        Next
    End With
    
    Application.ScreenUpdating = True
End Sub
'Methode d'initialisation dictionnaire des factures
Sub computeFactDictionary(sheetFact As String)
    'creation dictionaire
    Set factDictionary = CreateObject("Scripting.Dictionary")
    
    Dim x As Integer
    Dim f As Facture
    Dim numRows As Integer
    
    Application.ScreenUpdating = False
    
    With Worksheets(sheetFact)
        'nombre de factures
        numRows = GetSheetNumRows(Worksheets(sheetFact))
        
        'création de la facture
        For x = 2 To numRows
            Set f = New Facture
            With f
                .num = Worksheets(sheetFact).Cells(x, 1)
                .dateFact = Worksheets(sheetFact).Cells(x, 2)
                .montant = Worksheets(sheetFact).Cells(x, 3)
                .Fournisseur = Worksheets(sheetFact).Cells(x, 4)
                .categorieFrais = Worksheets(sheetFact).Cells(x, 5)
                .typeFrais = Worksheets(sheetFact).Cells(x, 6)
                .objet = Worksheets(sheetFact).Cells(x, 7)
                .concerne = Worksheets(sheetFact).Cells(x, 8)
                .ens = Worksheets(sheetFact).Cells(x, 9)
                .fichier = Worksheets(sheetFact).Cells(x, 10)
            End With
            
            'Ajout de la facture au dictionaire
            factDictionary.Add Key:=f.num, Item:=f
        Next
    End With
    
    Application.ScreenUpdating = True
End Sub
'Methode d'initialisation dictionnaire des fournisseurs
Sub computeFournDictionary()
    'creation dictionaire
    Set fournDictionary = CreateObject("Scripting.Dictionary")

    Dim x As Integer
    Dim f As Fournisseur
    Dim numRows As Integer
    
    Application.ScreenUpdating = False
    
    With SheetFournisseurs
        'nombre de fournisseurs
        numRows = GetSheetNumRows(SheetFournisseurs)
        
        'création du fournisseur
        For x = 2 To numRows
            Set f = New Fournisseur
            With f
                .societe = SheetFournisseurs.Cells(x, 1)
                .telephone = SheetFournisseurs.Cells(x, 2)
                .mail = SheetFournisseurs.Cells(x, 3)
                .domaine = SheetFournisseurs.Cells(x, 4)
            End With
            
            'Ajout du fournisseur au dictionaire
            fournDictionary.Add Key:=f.societe, Item:=f
        Next
    End With
    
    Application.ScreenUpdating = True
End Sub
