Attribute VB_Name = "modGerarRegistro"
Option Explicit

Private descricoes() As String
Private status() As String

Public Sub InicializarListas()
    descricoes = Split("CINEMA,UBER,SORVETERIA,RESTAURANTE,GASOLINA,LOJA ONLINE,ACADEMIA,SUPERMERCADO,FARMÁCIA,FAST FOOD," & _
       "PIZZARIA,PEDÁGIO,ESTACIONAMENTO,LOJA DE ROUPAS,BAR,ASSINATURA STREAMING,PARQUE,CARTÃO TRANSPORTE," & _
       "POSTO DE GASOLINA,DELIVERY,SALÃO DE BELEZA,PET SHOP,LIVRARIA,JOGOS ONLINE,VIAGEM,HOTEL,PASSAGEM AÉREA," & _
       "ALUGUEL DE CARRO,TAXI,CAFETERIA,CONVENIÊNCIA,ELETRODOMÉSTICOS,ELETRÔNICOS,LOJA DE CALÇADOS,SHOW,TEATRO", ",")
                       
    status = Split("Aprovada,Pendente,Cancelada", ",")
End Sub

Public Sub GerarTransacoes(Optional ByVal quantidade As Long = 1000)

        '<EhHeader>
        On Error GoTo GerarTransacoes_Err

        '</EhHeader>
    
        Dim i               As Long

        Dim numeroCartao    As String

        Dim valor           As Double

        Dim dataTransacao   As Date

        Dim descricao       As String

        Dim statusTransacao As String

        Dim rndIdx          As Long

        Dim objTransacaoDto As TransacaoDTO

        Dim objTransacaoBO  As TransacaoBO
    
100     Call InicializarListas
102     Randomize
    
104     Set objTransacaoBO = New TransacaoBO
    
106     For i = 1 To quantidade
108         numeroCartao = GerarNumeroCartao()
110         valor = GerarValor()
112         dataTransacao = GerarData2025()
114         rndIdx = Int((UBound(descricoes) + 1) * Rnd)
116         descricao = descricoes(rndIdx)
118         rndIdx = Int((UBound(status) + 1) * Rnd)
120         statusTransacao = status(rndIdx)
        
122         Set objTransacaoDto = New TransacaoDTO

124         With objTransacaoDto
126             .Numero_Cartao = numeroCartao
128             .Valor_Transacao = valor
130             .Data_Transacao = dataTransacao
132             .descricao = descricao
134             .Status_Transacao = statusTransacao
            End With
        
136         If Not objTransacaoBO.Salvar(objTransacaoDto) Then
138             MsgBox "Falha ao salvar transação " & i, vbExclamation
            End If

140     Next i

142     MsgBox quantidade & " transações geradas com sucesso!", vbInformation

        Exit Sub

        '<EhFooter>
        Exit Sub

GerarTransacoes_Err:
        RegistrarErro "modGerarRelatorio", Err.Description
        MsgBox Err.Description & vbCrLf & "in CobuccioCard.modGerarRegistro.GerarTransacoes " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Sub

Private Function GerarNumeroCartao() As String

        '<EhHeader>
        On Error GoTo GerarNumeroCartao_Err

        '</EhHeader>

        Dim i      As Integer

        Dim numero As String

100     numero = ""

102     For i = 1 To 16
104         numero = numero & CStr(Int(Rnd * 10))
106     Next i

108     GerarNumeroCartao = numero

        '<EhFooter>
        Exit Function

GerarNumeroCartao_Err:
        RegistrarErro "modGerarRelatorio", Err.Description
        MsgBox Err.Description & vbCrLf & "in CobuccioCard.modGerarRegistro.GerarNumeroCartao " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Function

Private Function GerarValor() As Double

        '<EhHeader>
        On Error GoTo GerarValor_Err

        '</EhHeader>

100     GerarValor = Round((Rnd * 8990) + 10, 2)

        '<EhFooter>
        Exit Function

GerarValor_Err:
        RegistrarErro "modGerarRelatorio", Err.Description
        MsgBox Err.Description & vbCrLf & "in CobuccioCard.modGerarRegistro.GerarValor " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Function

Private Function GerarData2025() As Date

        '<EhHeader>
        On Error GoTo GerarData2025_Err

        '</EhHeader>

        Dim ano  As Integer, mes As Integer, dia As Integer

        Dim hora As Integer, minuto As Integer, segundo As Integer

100     ano = 2025
102     mes = Int((12 * Rnd) + 1)
104     dia = Int((DiasNoMes(ano, mes) * Rnd) + 1)
106     hora = Int((24 * Rnd))
108     minuto = Int((60 * Rnd))
110     segundo = Int((60 * Rnd))
112     GerarData2025 = DateSerial(ano, mes, dia) + TimeSerial(hora, minuto, segundo)

        '<EhFooter>
        Exit Function

GerarData2025_Err:
        RegistrarErro "modGerarRelatorio", Err.Description
        MsgBox Err.Description & vbCrLf & "in CobuccioCard.modGerarRegistro.GerarData2025 " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Function

Private Function DiasNoMes(ByVal ano As Integer, ByVal mes As Integer) As Integer
    DiasNoMes = Day(DateSerial(ano, mes + 1, 0))
End Function


