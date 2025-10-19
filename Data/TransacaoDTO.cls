VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TransacaoDTO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pId_Transacao     As Long

Private pNumero_Cartao    As String

Private pValor_Transacao  As Currency

Private pData_Transacao   As Date

Private pDescricao        As String

Private pStatus_Transacao As String

Public Property Get Id_Transacao() As Long
    Id_Transacao = pId_Transacao
End Property

Public Property Let Id_Transacao(ByVal v As Long)
    pId_Transacao = v
End Property

Public Property Get Numero_Cartao() As String
    Numero_Cartao = pNumero_Cartao
End Property

Public Property Let Numero_Cartao(ByVal v As String)
    pNumero_Cartao = Trim(v)
End Property

Public Property Get Valor_Transacao() As Currency
    Valor_Transacao = pValor_Transacao
End Property

Public Property Let Valor_Transacao(ByVal v As Currency)
    pValor_Transacao = v
End Property

Public Property Get Data_Transacao() As Date
    Data_Transacao = pData_Transacao
End Property

Public Property Let Data_Transacao(ByVal v As Date)
    pData_Transacao = v
End Property

Public Property Get descricao() As String
    descricao = pDescricao
End Property

Public Property Let descricao(ByVal v As String)
    pDescricao = Trim(v)
End Property

Public Property Get Status_Transacao() As String
    Status_Transacao = pStatus_Transacao
End Property

Public Property Let Status_Transacao(ByVal v As String)
    pStatus_Transacao = Trim(v)
End Property

