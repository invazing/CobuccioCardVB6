VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Form_PsqTransacoes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pesquisa de Transações"
   ClientHeight    =   9615
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12375
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form_PsqTransacoes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9615
   ScaleWidth      =   12375
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmeOutros 
      Height          =   735
      Left            =   0
      TabIndex        =   17
      Top             =   8880
      Width           =   12375
      Begin VB.CommandButton btnExport 
         Caption         =   "Exportar"
         Height          =   360
         Left            =   9120
         TabIndex        =   20
         Top             =   240
         Width           =   990
      End
      Begin VB.CommandButton btnDeletar 
         Caption         =   "&Deletar"
         Height          =   360
         Left            =   10200
         TabIndex        =   19
         Top             =   240
         Width           =   990
      End
      Begin VB.CommandButton btnEditar 
         Caption         =   "&Editar"
         Height          =   360
         Left            =   11280
         TabIndex        =   18
         Top             =   240
         Width           =   990
      End
   End
   Begin VB.Frame FrmeDataGridView 
      Height          =   7935
      Left            =   0
      TabIndex        =   8
      Top             =   960
      Width           =   12375
      Begin MSFlexGridLib.MSFlexGrid gridTransacoes 
         Height          =   7095
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   12135
         _ExtentX        =   21405
         _ExtentY        =   12515
         _Version        =   393216
      End
      Begin VB.CommandButton btnProximo 
         Caption         =   ">>>"
         Height          =   360
         Left            =   11280
         TabIndex        =   10
         Top             =   7440
         Width           =   990
      End
      Begin VB.CommandButton btnAnterior 
         Caption         =   "<<<"
         Height          =   360
         Left            =   10200
         TabIndex        =   9
         Top             =   7440
         Width           =   990
      End
      Begin VB.Label lblIndex 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Páginas:   0"
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   7440
         Width           =   840
      End
   End
   Begin VB.Frame FrmeDetalhePesquisa 
      Caption         =   "Dados da Consulta"
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12375
      Begin VB.CommandButton btnPesquisa 
         Caption         =   "Consultar"
         Height          =   480
         Left            =   11280
         TabIndex        =   11
         Top             =   360
         Width           =   990
      End
      Begin VB.ComboBox cmbStatus 
         Height          =   315
         Left            =   9720
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   480
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker dtFim 
         Height          =   315
         Left            =   8160
         TabIndex        =   6
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Format          =   145817601
         CurrentDate     =   45948
      End
      Begin MSComCtl2.DTPicker dtInicio 
         Height          =   315
         Left            =   6600
         TabIndex        =   5
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Format          =   145817601
         CurrentDate     =   45948
      End
      Begin VB.TextBox txtDescricao 
         Height          =   315
         Left            =   2640
         MaxLength       =   100
         TabIndex        =   4
         Top             =   480
         Width           =   3855
      End
      Begin VB.TextBox txtNumCartao 
         Height          =   315
         Left            =   240
         MaxLength       =   16
         TabIndex        =   2
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
         Height          =   195
         Left            =   9720
         TabIndex        =   15
         Top             =   240
         Width           =   465
      End
      Begin VB.Label lblFim 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fim"
         Height          =   195
         Left            =   8160
         TabIndex        =   14
         Top             =   240
         Width           =   240
      End
      Begin VB.Label lblInicio 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data Inicio"
         Height          =   195
         Left            =   6600
         TabIndex        =   13
         Top             =   240
         Width           =   765
      End
      Begin VB.Label lblDescricao 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descricao"
         Height          =   195
         Left            =   2640
         TabIndex        =   3
         Top             =   240
         Width           =   690
      End
      Begin VB.Label lblNumCartao 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Num cartao"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   825
      End
   End
End
Attribute VB_Name = "Form_PsqTransacoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private paginaAtual As Long
Private tamanhoPagina As Long

Private Sub btnDeletar_Click()

        '<EhHeader>
        On Error GoTo btnDeletar_Click_Err

        '</EhHeader>
        
        Dim idTransacao    As Long

        Dim objTransacaoBO As New TransacaoBO
    
100     idTransacao = ObterIdSelecionado(gridTransacoes)
    
102     If idTransacao = 0 Then
104         MsgBox "Selecione uma transação para excluir.", vbExclamation

            Exit Sub

        End If
    
106     If MsgBox("Deseja realmente excluir a transação ID: " & idTransacao & "?", vbYesNo + vbQuestion) = vbNo Then

            Exit Sub

        End If
    
108     If objTransacaoBO.Excluir(idTransacao) Then
110         MsgBox "Transação excluída com sucesso!", vbInformation
112         CarregarTransacoes
        Else
114         MsgBox "Erro ao excluir a transação.", vbCritical
        End If
    
116     Set objTransacaoBO = Nothing

        '<EhFooter>
        Exit Sub

btnDeletar_Click_Err:
        RegistrarErro "Form_PsqTransacoes", Err.Description
        MsgBox Err.Description & vbCrLf & "in CobuccioCard.Form_PsqTransacoes.btnDeletar_Click " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Sub

Private Sub btnEditar_Click()

        '<EhHeader>
        On Error GoTo btnEditar_Click_Err

        '</EhHeader>
        
        Dim idTransacao     As Long

        Dim statusTransacao As String

100     idTransacao = ObterIdSelecionado(gridTransacoes)
    
102     If idTransacao = 0 Then
104         MsgBox "Selecione uma transação para editar.", vbExclamation

            Exit Sub

        End If

106     statusTransacao = gridTransacoes.TextMatrix(gridTransacoes.Row, 5)

108     If UCase(Trim(statusTransacao)) = "APROVADA" Then
110         MsgBox "Transações aprovadas não podem ser alteradas.", vbExclamation

            Exit Sub

        End If

112     Form_CadTransacoes.CarregarModoEdicao idTransacao
114     Form_CadTransacoes.Show vbModal

116     CarregarTransacoes

        '<EhFooter>
        Exit Sub

btnEditar_Click_Err:
        RegistrarErro "Form_PsqTransacoes", Err.Description
        MsgBox Err.Description & vbCrLf & "in CobuccioCard.Form_PsqTransacoes.btnEditar_Click " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Sub

Private Sub btnExport_Click()

        '<EhHeader>
        On Error GoTo btnExport_Click_Err

        '</EhHeader>
        
        Dim objTransacaoBO As New TransacaoBO

        Dim rsExport       As ADODB.Recordset

        Dim xlApp          As Object, xlBook As Object, xlSheet As Object

        Dim i              As Long, j As Long

100     Set rsExport = objTransacaoBO.Listar(ConstruirFiltro)

102     If rsExport Is Nothing Then Exit Sub
104     If rsExport.EOF Then
106         MsgBox "Nenhum dado para exportar.", vbInformation

            Exit Sub

        End If

108     If MsgBox("Deseja exportar o relatório para o Excel? O processo pode demorar dependendo da quantidade de registros.", vbYesNo + vbQuestion, "Exportar para Excel") = vbNo Then

            Exit Sub
        
        End If

110     Set xlApp = CreateObject("Excel.Application")
112     Set xlBook = xlApp.Workbooks.Add
114     Set xlSheet = xlBook.Sheets(1)

116     For j = 0 To rsExport.Fields.Count - 1
118         xlSheet.Cells(1, j + 1).Value = rsExport.Fields(j).Name
120         xlSheet.Cells(1, j + 1).Font.Bold = True
122     Next j

124     i = 2

126     Do Until rsExport.EOF

128         For j = 0 To rsExport.Fields.Count - 1
130             xlSheet.Cells(i, j + 1).Value = rsExport.Fields(j).Value
132         Next j

134         rsExport.MoveNext
136         i = i + 1
        Loop

138     xlSheet.Columns("A:" & Chr(65 + rsExport.Fields.Count - 1)).AutoFit

140     xlApp.Visible = True

142     rsExport.Close
144     Set rsExport = Nothing
146     Set xlSheet = Nothing
148     Set xlBook = Nothing
150     Set xlApp = Nothing
152     Set objTransacaoBO = Nothing

        '<EhFooter>
        Exit Sub

btnExport_Click_Err:
        RegistrarErro "Form_PsqTransacoes", Err.Description
        MsgBox Err.Description & vbCrLf & "in CobuccioCard.Form_PsqTransacoes.btnExport_Click " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_Load()

        '<EhHeader>
        On Error GoTo Form_Load_Err

        '</EhHeader>
              
100     paginaAtual = 1
102     tamanhoPagina = 100
    
104     ConfigurarGridTransacoes gridTransacoes
106     AjustarLarguraAutomatica gridTransacoes
    
108     PopularStatus
    
110     If Not OpenConnection() Then Exit Sub

        '<EhFooter>
        Exit Sub

Form_Load_Err:
        RegistrarErro "Form_PsqTransacoes", Err.Description
        MsgBox Err.Description & vbCrLf & "in CobuccioCard.Form_PsqTransacoes.Form_Load " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Sub

Private Sub PopularStatus()

        '<EhHeader>
        On Error GoTo PopularStatus_Err

        '</EhHeader>
        
100     cmbStatus.Clear
    
102     cmbStatus.AddItem "Todos"
104     cmbStatus.AddItem "Pendente"
106     cmbStatus.AddItem "Aprovada"
108     cmbStatus.AddItem "Cancelada"
    
110     cmbStatus.ListIndex = 0

        '<EhFooter>
        Exit Sub

PopularStatus_Err:
        RegistrarErro "Form_PsqTransacoes", Err.Description
        MsgBox Err.Description & vbCrLf & "in CobuccioCard.Form_PsqTransacoes.PopularStatus " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_Resize()
    AjustarLarguraAutomatica gridTransacoes
End Sub

Private Sub btnAnterior_Click()

        '<EhHeader>
        On Error GoTo btnAnterior_Click_Err

        '</EhHeader>

100     If paginaAtual > 1 Then
102         paginaAtual = paginaAtual - 1
104         CarregarTransacoes
        End If

        '<EhFooter>
        Exit Sub

btnAnterior_Click_Err:
        RegistrarErro "Form_PsqTransacoes", Err.Description
        MsgBox Err.Description & vbCrLf & "in CobuccioCard.Form_PsqTransacoes.btnAnterior_Click " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Sub

Private Sub btnPesquisa_Click()

        '<EhHeader>
        On Error GoTo btnPesquisa_Click_Err

        '</EhHeader>

100     paginaAtual = 1
102     CarregarTransacoes

        '<EhFooter>
        Exit Sub

btnPesquisa_Click_Err:
        RegistrarErro "Form_PsqTransacoes", Err.Description
        MsgBox Err.Description & vbCrLf & "in CobuccioCard.Form_PsqTransacoes.btnPesquisa_Click " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Sub

Private Sub btnProximo_Click()

        '<EhHeader>
        On Error GoTo btnProximo_Click_Err

        '</EhHeader>

100     paginaAtual = paginaAtual + 1
102     CarregarTransacoes

        '<EhFooter>
        Exit Sub

btnProximo_Click_Err:
        MsgBox Err.Description & vbCrLf & "in CobuccioCard.Form_PsqTransacoes.btnProximo_Click " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Sub

Private Sub gridTransacoes_DblClick()

        '<EhHeader>
        On Error GoTo gridTransacoes_DblClick_Err

        '</EhHeader>

        Dim idTransacao As Long

100     idTransacao = ObterIdSelecionado(gridTransacoes)
    
102     If idTransacao > 0 Then
104         MsgBox "Editar transação ID: " & idTransacao, vbInformation
        End If

        '<EhFooter>
        Exit Sub

gridTransacoes_DblClick_Err:
        RegistrarErro "Form_PsqTransacoes", Err.Description
        MsgBox Err.Description & vbCrLf & "in CobuccioCard.Form_PsqTransacoes.gridTransacoes_DblClick " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Sub

Private Sub CarregarTransacoes()

        '<EhHeader>
        On Error GoTo CarregarTransacoes_Err

        '</EhHeader>
        
        Dim objTransacaoBO As New TransacaoBO

        Dim rsTemp         As ADODB.Recordset

        Dim filtro         As String
    
100     If Not OpenConnection() Then
102         MsgBox "Erro de conexão com o banco", vbCritical

            Exit Sub

        End If
    
104     LimparGrid gridTransacoes
    
106     filtro = ConstruirFiltro()
    
108     Set rsTemp = objTransacaoBO.Listar(filtro)
    
110     If Not ValidarRecordset(rsTemp) Then
112         Set rsTemp = Nothing
114         Set objTransacaoBO = Nothing

            Exit Sub

        End If
    
116     PreencherGridComPaginacao rsTemp
    
118     rsTemp.Close
120     Set rsTemp = Nothing
122     Set objTransacaoBO = Nothing

        '<EhFooter>
        Exit Sub

CarregarTransacoes_Err:
        RegistrarErro "Form_PsqTransacoes", Err.Description
        MsgBox Err.Description & vbCrLf & "in CobuccioCard.Form_PsqTransacoes.CarregarTransacoes " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Sub

Private Function ConstruirFiltro() As String

        '<EhHeader>
        On Error GoTo ConstruirFiltro_Err

        '</EhHeader>
        
        Dim filtro As String

100     filtro = ""
    
102     If Trim(txtDescricao.Text) <> "" Then
104         filtro = "Descricao LIKE '%" & Trim(txtDescricao.Text) & "%'"
        End If
    
106     If Trim(txtNumCartao.Text) <> "" Then
108         If filtro <> "" Then filtro = filtro & " AND "
110         filtro = filtro & "Numero_Cartao LIKE '%" & Trim(txtNumCartao.Text) & "%'"
        End If
    
112     If cmbStatus.ListIndex >= 1 Then
114         If filtro <> "" Then filtro = filtro & " AND "
116         filtro = filtro & "Status_Transacao='" & cmbStatus.Text & "'"
        End If
    
118     If IsDate(dtInicio.Value) Then
120         If filtro <> "" Then filtro = filtro & " AND "
122         filtro = filtro & "Data_Transacao>='" & Format(CDate(dtInicio.Value), "yyyy-mm-dd") & " 00:00:00'"
        End If
    
124     If IsDate(dtFim.Value) Then
126         If filtro <> "" Then filtro = filtro & " AND "
128         filtro = filtro & "Data_Transacao<='" & Format(CDate(dtFim.Value), "yyyy-mm-dd") & " 23:59:59'"
        End If
    
130     ConstruirFiltro = filtro

        '<EhFooter>
        Exit Function

ConstruirFiltro_Err:
        RegistrarErro "Form_PsqTransacoes", Err.Description
        MsgBox Err.Description & vbCrLf & "in CobuccioCard.Form_PsqTransacoes.ConstruirFiltro " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Function

Private Function ValidarRecordset(ByRef rs As ADODB.Recordset) As Boolean

        '<EhHeader>
        On Error GoTo ValidarRecordset_Err

        '</EhHeader>
         
100     ValidarRecordset = False
    
102     If rs Is Nothing Then
104         MsgBox "Erro ao carregar transações", vbCritical

            Exit Function

        End If
    
106     If rs.EOF Then
108         MsgBox "Nenhuma transação encontrada", vbInformation
110         btnAnterior.Enabled = False
112         btnProximo.Enabled = False

            Exit Function

        End If
    
114     ValidarRecordset = True

        '<EhFooter>
        Exit Function

ValidarRecordset_Err:
        RegistrarErro "Form_PsqTransacoes", Err.Description
        MsgBox Err.Description & vbCrLf & "in CobuccioCard.Form_PsqTransacoes.ValidarRecordset " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Function

Private Sub PreencherGridComPaginacao(ByRef rs As ADODB.Recordset)

        '<EhHeader>
        On Error GoTo PreencherGridComPaginacao_Err

        '</EhHeader>
        
        Dim offset   As Long

        Dim contador As Long

100     offset = (paginaAtual - 1) * tamanhoPagina

102     If offset > 0 And Not rs.EOF Then
104         rs.Move offset
        End If

106     contador = 0

108     Do While Not rs.EOF And contador < tamanhoPagina
110         AdicionarLinhaGrid gridTransacoes, rs!Id_Transacao, rs!Numero_Cartao, rs!Valor_Transacao, rs!Data_Transacao, rs!descricao & "", rs!Status_Transacao
        
112         rs.MoveNext
114         contador = contador + 1
        Loop

116     AtualizarControlesNavegacao Not rs.EOF

        '<EhFooter>
        Exit Sub

PreencherGridComPaginacao_Err:
        RegistrarErro "Form_PsqTransacoes", Err.Description
        MsgBox Err.Description & vbCrLf & "in CobuccioCard.Form_PsqTransacoes.PreencherGridComPaginacao " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Sub

Private Sub AtualizarControlesNavegacao(ByVal temProximaPagina As Boolean)

        '<EhHeader>
        On Error GoTo AtualizarControlesNavegacao_Err

        '</EhHeader>
       
100     btnAnterior.Enabled = (paginaAtual > 1)
102     btnProximo.Enabled = temProximaPagina

104     lblIndex.Caption = "Página " & paginaAtual & " - " & (gridTransacoes.Rows - 1) & " registro(s)"

        '<EhFooter>
        Exit Sub

AtualizarControlesNavegacao_Err:
        RegistrarErro "Form_PsqTransacoes", Err.Description
        MsgBox Err.Description & vbCrLf & "in CobuccioCard.Form_PsqTransacoes.AtualizarControlesNavegacao " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Sub

