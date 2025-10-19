VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TransacaoBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private objTransacaoDao As TransacaoDAO

Private Sub Class_Initialize()

        '<EhHeader>
        On Error GoTo Class_Initialize_Err

        '</EhHeader>

100     Set objTransacaoDao = New TransacaoDAO

        '<EhFooter>
        Exit Sub

Class_Initialize_Err:
        RegistrarErro "TransacaoBO", Err.Description
        Err.Raise vbObjectError + 100, "CobuccioCard.TransacaoBO.Class_Initialize", "TransacaoBO component failure"

        '</EhFooter>

End Sub

Public Function Listar(Optional ByVal filtro As String = "") As ADODB.Recordset

        '<EhHeader>
        On Error GoTo Listar_Err

        '</EhHeader>

100     Set Listar = objTransacaoDao.Listar(filtro)

        '<EhFooter>
        Exit Function

Listar_Err:
        RegistrarErro "TransacaoBO", Err.Description
        Err.Raise vbObjectError + 100, "CobuccioCard.TransacaoBO.Listar", "TransacaoBO component failure"

        '</EhFooter>

End Function

Public Function Salvar(ByVal objTransacaoDto As TransacaoDTO) As Boolean

        '<EhHeader>
        On Error GoTo Salvar_Err

        '</EhHeader>

100     If Len(objTransacaoDto.Numero_Cartao) <> 16 Or Not IsNumeric(objTransacaoDto.Numero_Cartao) Then
102         MsgBox "Número do cartão deve conter 16 dígitos numéricos.", vbExclamation

            Exit Function

        End If

104     If objTransacaoDto.Valor_Transacao <= 0 Then
106         MsgBox "Valor da transação deve ser positivo.", vbExclamation

            Exit Function

        End If

108     If Len(objTransacaoDto.Status_Transacao) = 0 Then
110         MsgBox "Informe o status da transação.", vbExclamation

            Exit Function

        End If

112     If objTransacaoDto.Id_Transacao = 0 Then
114         Salvar = objTransacaoDao.Inserir(objTransacaoDto)
        Else
116         Salvar = objTransacaoDao.Atualizar(objTransacaoDto)
        End If

        '<EhFooter>
        Exit Function

Salvar_Err:
        RegistrarErro "TransacaoBO", Err.Description
        Err.Raise vbObjectError + 100, "CobuccioCard.TransacaoBO.Salvar", "TransacaoBO component failure"
        Salvar = False
        '</EhFooter>

End Function

Public Function Excluir(ByVal idTransacao As Long) As Boolean

        '<EhHeader>
        On Error GoTo Excluir_Err

        '</EhHeader>

100     Excluir = objTransacaoDao.Excluir(idTransacao)

        '<EhFooter>
        Exit Function

Excluir_Err:
        RegistrarErro "TransacaoBO", Err.Description
        Err.Raise vbObjectError + 100, "CobuccioCard.TransacaoBO.Excluir", "TransacaoBO component failure"
        Excluir = False

        '</EhFooter>

End Function

Public Function ObterTransacoesCategorizadas(ByVal dataInicial As Date, _
                                             ByVal dataFinal As Date) As ADODB.Recordset

        '<EhHeader>
        On Error GoTo ObterTransacoesCategorizadas_Err

        '</EhHeader>

100     If Not IsDate(dataInicial) Or Not IsDate(dataFinal) Then
102         MsgBox "Datas inválidas para consulta.", vbExclamation

            Exit Function

        End If

104     If dataInicial > dataFinal Then
106         MsgBox "A data inicial não pode ser maior que a final.", vbExclamation

            Exit Function

        End If

108     Set ObterTransacoesCategorizadas = objTransacaoDao.ObterTransacoesCategorizadas(dataInicial, dataFinal)

        Exit Function

        '<EhFooter>
        Exit Function

ObterTransacoesCategorizadas_Err:
        RegistrarErro "TransacaoBO", Err.Description
        Err.Raise vbObjectError + 100, "CobuccioCard.TransacaoBO.ObterTransacoesCategorizadas", "TransacaoBO component failure"
        Set ObterTransacoesCategorizadas = Nothing

        '</EhFooter>

End Function

Public Function ObterResumoFinanceiro() As ADODB.Recordset

        '<EhHeader>
        On Error GoTo ObterResumoFinanceiro_Err

        '</EhHeader>

100     Set ObterResumoFinanceiro = objTransacaoDao.ObterResumoFinanceiro

        '<EhFooter>
        Exit Function

ObterResumoFinanceiro_Err:
        RegistrarErro "TransacaoBO", Err.Description
        Err.Raise vbObjectError + 100, "CobuccioCard.TransacaoBO.ObterResumoFinanceiro", "TransacaoBO component failure"
        Set ObterResumoFinanceiro = Nothing

        '</EhFooter>

End Function

