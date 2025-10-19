VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TransacaoDAO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function Inserir(ByVal objTransacaoDto As TransacaoDTO) As Boolean

        '<EhHeader>
        On Error GoTo Inserir_Err

        '</EhHeader>

        Dim sql As String

        Dim cmd As ADODB.Command

100     If Not OpenConnection() Then Exit Function

102     sql = "INSERT INTO Transacoes " & "(Numero_Cartao, Valor_Transacao, Data_Transacao, Descricao, Status_Transacao) " & "VALUES (?, ?, ?, ?, ?)"

104     Set cmd = New ADODB.Command

106     With cmd
108         .ActiveConnection = cn
110         .CommandType = adCmdText
112         .CommandText = sql
114         .Parameters.Append .CreateParameter("p1", adVarChar, adParamInput, 16, objTransacaoDto.Numero_Cartao)
116         .Parameters.Append .CreateParameter("p2", adCurrency, adParamInput, , objTransacaoDto.Valor_Transacao)
118         .Parameters.Append .CreateParameter("p3", adDate, adParamInput, , objTransacaoDto.Data_Transacao)
120         .Parameters.Append .CreateParameter("p4", adVarChar, adParamInput, 255, objTransacaoDto.descricao)
122         .Parameters.Append .CreateParameter("p5", adVarChar, adParamInput, 20, objTransacaoDto.Status_Transacao)
124         .Execute
        End With

126     Inserir = True

        '<EhFooter>
        Exit Function

Inserir_Err:

        RegistrarErro "TransacaoDAO", Err.Description
        Err.Raise vbObjectError + 100, "CobuccioCard.TransacaoDAO.Inserir", "TransacaoDAO component failure"
        Inserir = False
        
        '</EhFooter>

End Function

Public Function Atualizar(ByVal objTransacaoDto As TransacaoDTO) As Boolean

        '<EhHeader>
        On Error GoTo Atualizar_Err

        '</EhHeader>

        Dim cmd As ADODB.Command

        Dim sql As String

        Dim rs  As ADODB.Recordset

100     If Not OpenConnection() Then Exit Function

102     sql = "SELECT Status_Transacao FROM Transacoes WHERE Id_Transacao=" & objTransacaoDto.Id_Transacao
104     Set rs = cn.Execute(sql)

106     If Not rs.EOF Then
108         If rs!Status_Transacao = "Aprovada" Then
110             MsgBox "Transações aprovadas não podem ser alteradas.", vbExclamation
112             rs.Close
114             Atualizar = False

                Exit Function

            End If
        End If

116     rs.Close

118     sql = "UPDATE Transacoes SET " & "Numero_Cartao=?, Valor_Transacao=?, Data_Transacao=?, Descricao=?, Status_Transacao=? " & "WHERE Id_Transacao=?"

120     Set cmd = New ADODB.Command

122     With cmd
124         .ActiveConnection = cn
126         .CommandType = adCmdText
128         .CommandText = sql
130         .Parameters.Append .CreateParameter("p1", adVarChar, adParamInput, 16, objTransacaoDto.Numero_Cartao)
132         .Parameters.Append .CreateParameter("p2", adCurrency, adParamInput, , objTransacaoDto.Valor_Transacao)
134         .Parameters.Append .CreateParameter("p3", adDate, adParamInput, , objTransacaoDto.Data_Transacao)
136         .Parameters.Append .CreateParameter("p4", adVarChar, adParamInput, 255, objTransacaoDto.descricao)
138         .Parameters.Append .CreateParameter("p5", adVarChar, adParamInput, 20, objTransacaoDto.Status_Transacao)
140         .Parameters.Append .CreateParameter("p6", adInteger, adParamInput, , objTransacaoDto.Id_Transacao)
142         .Execute
        End With

144     Atualizar = True

        '<EhFooter>
        Exit Function

Atualizar_Err:
        RegistrarErro "TransacaoDAO", Err.Description
        Err.Raise vbObjectError + 100, "CobuccioCard.TransacaoDAO.Atualizar", "TransacaoDAO component failure"
        Atualizar = False

        '</EhFooter>

End Function

Public Function Excluir(ByVal idTransacao As Long) As Boolean

        '<EhHeader>
        On Error GoTo Excluir_Err

        '</EhHeader>

        Dim sql As String

100     If Not OpenConnection() Then Exit Function

102     If MsgBox("Confirma exclusão da transação " & idTransacao & "?", vbYesNo + vbQuestion) = vbNo Then

            Exit Function

        End If

104     sql = "DELETE FROM Transacoes WHERE Id_Transacao=" & idTransacao
106     cn.Execute sql

108     Excluir = True

        '<EhFooter>
        Exit Function

Excluir_Err:
        RegistrarErro "TransacaoDAO", Err.Description
        Err.Raise vbObjectError + 100, "CobuccioCard.TransacaoDAO.Excluir", "TransacaoDAO component failure"
        Excluir = False

        '</EhFooter>

End Function

Public Function Listar(Optional ByVal filtro As String = "") As ADODB.Recordset

        '<EhHeader>
        On Error GoTo Listar_Err

        '</EhHeader>

        Dim rs  As New ADODB.Recordset

        Dim sql As String

100     If Not OpenConnection() Then Exit Function

102     sql = "SELECT Id_Transacao, Numero_Cartao, Valor_Transacao, Data_Transacao, Descricao, Status_Transacao " & "FROM Transacoes"

104     If filtro <> "" Then sql = sql & " WHERE " & filtro
106     sql = sql & " ORDER BY Id_Transacao ASC"

108     rs.Open sql, cn, adOpenStatic, adLockReadOnly
110     Set Listar = rs

        '<EhFooter>
        Exit Function

Listar_Err:
        RegistrarErro "TransacaoDAO", Err.Description
        Err.Raise vbObjectError + 100, "CobuccioCard.TransacaoDAO.Listar", "TransacaoDAO component failure"
        Set Listar = Nothing
        '</EhFooter>

End Function

Public Function ObterTransacoesCategorizadas(ByVal dataInicial As Date, _
                                             ByVal dataFinal As Date) As ADODB.Recordset

        '<EhHeader>
        On Error GoTo ObterTransacoesCategorizadas_Err

        '</EhHeader>

        Dim rs  As New ADODB.Recordset

        Dim cmd As ADODB.Command

        Dim sql As String

100     If Not OpenConnection() Then Exit Function

102     sql = "SELECT * FROM fn_TransacoesCategorizadas(?, ?)"

104     Set cmd = New ADODB.Command

106     With cmd
108         .ActiveConnection = cn
110         .CommandType = adCmdText
112         .CommandText = sql
114         .Parameters.Append .CreateParameter("p1", adDate, adParamInput, , dataInicial)
116         .Parameters.Append .CreateParameter("p2", adDate, adParamInput, , dataFinal)
118         Set rs = .Execute
        End With

120     Set ObterTransacoesCategorizadas = rs

        '<EhFooter>
        Exit Function

ObterTransacoesCategorizadas_Err:
        RegistrarErro "TransacaoDAO", Err.Description
        Err.Raise vbObjectError + 100, "CobuccioCard.TransacaoDAO.ObterTransacoesCategorizadas", "TransacaoDAO component failure"
        Set ObterTransacoesCategorizadas = Nothing
        '</EhFooter>

End Function

Public Function ObterResumoFinanceiro() As ADODB.Recordset

        '<EhHeader>
        On Error GoTo ObterResumoFinanceiro_Err

        '</EhHeader>

        Dim rs  As New ADODB.Recordset

        Dim sql As String

100     If Not OpenConnection() Then Exit Function

102     sql = "SELECT * FROM vw_ResumoFinanceiro"

104     rs.Open sql, cn, adOpenStatic, adLockReadOnly
106     Set ObterResumoFinanceiro = rs

        '<EhFooter>
        Exit Function

ObterResumoFinanceiro_Err:
        RegistrarErro "TransacaoDAO", Err.Description
        Err.Raise vbObjectError + 100, "CobuccioCard.TransacaoDAO.ObterResumoFinanceiro", "TransacaoDAO component failure"
        Set ObterResumoFinanceiro = Nothing

        '</EhFooter>

End Function

