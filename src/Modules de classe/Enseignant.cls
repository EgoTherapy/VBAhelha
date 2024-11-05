VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Enseignant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'déclaration des attributs de la classe "Enseignants"
Private m_nomPrenom As String


'Getters & setters

Property Get NomPrenom() As String
    NomPrenom = m_nomPrenom
End Property

Property Let NomPrenom(value As String)
    m_nomPrenom = value
End Property



