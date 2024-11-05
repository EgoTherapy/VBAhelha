VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Facture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'déclaration des attributs de la classe "Facture"
Private m_num As String
Private m_dateFact As Date
Private m_montant As Double
Private m_fournisseur As String
Private m_categorieFrais As String
Private m_typeFrais As String
Private m_objet As String
Private m_concerne As String
Private m_ens As String
Private m_fichier As String

'getters
Property Get num() As String
    num = m_num
End Property

Property Get dateFact() As Date
    dateFact = m_dateFact
End Property

Property Get montant() As Double
    montant = m_montant
End Property

Property Get Fournisseur() As String
    Fournisseur = m_fournisseur
End Property

Property Get categorieFrais() As String
    categorieFrais = m_categorieFrais
End Property

Property Get typeFrais() As String
    typeFrais = m_typeFrais
End Property
Property Get objet() As String
    objet = m_objet
End Property
Property Get concerne() As String
    concerne = m_concerne
End Property
Property Get ens() As String
    ens = m_ens
End Property
Property Get fichier() As String
    fichier = m_fichier
End Property

'setters
Property Let num(value As String)
    m_num = value
End Property

Property Let dateFact(value As Date)
    m_dateFact = value
End Property

Property Let montant(value As Double)
    m_montant = value
End Property

Property Let Fournisseur(value As String)
    m_fournisseur = value
End Property
Property Let categorieFrais(value As String)
    m_categorieFrais = value
End Property
Property Let typeFrais(value As String)
    m_typeFrais = value
End Property
Property Let objet(value As String)
    m_objet = value
End Property
Property Let concerne(value As String)
    m_concerne = value
End Property
Property Let ens(value As String)
    m_ens = value
End Property

Property Let fichier(value As String)
    m_fichier = value
End Property
