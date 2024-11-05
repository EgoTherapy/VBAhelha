VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Fournisseur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'déclaration des attributs de la classe "Fournisseur"
Private m_societe As String
Private m_telephone As String
Private m_mail As String
Private m_domaine As String


'getters
Property Get societe() As String
    societe = m_societe
End Property

Property Get telephone() As String
    telephone = m_telephone
End Property

Property Get mail() As String
    mail = m_mail
End Property

Property Get domaine() As String
    domaine = m_domaine
End Property

'setters
Property Let societe(value As String)
    m_societe = value
End Property

Property Let telephone(value As String)
    m_telephone = value
End Property

Property Let mail(value As String)
    m_mail = value
End Property

Property Let domaine(value As String)
    m_domaine = value
End Property





