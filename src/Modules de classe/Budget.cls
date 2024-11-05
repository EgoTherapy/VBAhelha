VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Budget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'declaration des attributs de la classe "budget"
Private m_entretiens As Single
Private m_telecom As Single
Private m_autresFourn As Single
Private m_retrib As Single
Private m_infos As Single
Private m_assurances As Single
Private m_autres As Single

'getters
Property Get entretiens() As Single
    entretiens = m_entretiens
End Property

Property Get telecom() As Single
    telecom = m_telecom
End Property

Property Get autresFourn() As Single
    autresFourn = m_autresFourn
End Property

Property Get retrib() As Single
    retrib = m_retrib
End Property

Property Get infos() As Single
    infos = m_infos
End Property

Property Get assurances() As Single
    assurances = m_assurances
End Property

Property Get autres() As Single
   autres = m_autres
End Property


'setters
Property Let entretiens(value As Single)
    m_entretiens = value
End Property

Property Let telecom(value As Single)
    m_telecom = value
End Property

Property Let autresFourn(value As Single)
    m_autresFourn = value
End Property

Property Let retrib(value As Single)
    m_retrib = value
End Property

Property Let infos(value As Single)
    m_infos = value
End Property

Property Let assurances(value As Single)
    m_assurances = value
End Property

Property Let autres(value As Single)
    m_autres = value
End Property





