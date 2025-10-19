VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form_CadTransacoes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro Transações"
   ClientHeight    =   2055
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5535
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form_Cad_Transacoes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   5535
   StartUpPosition =   3  'Windows Default
   Begin MSComCtl2.DTPicker dtData 
      Height          =   315
      Left            =   1800
      TabIndex        =   11
      Top             =   1080
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy HH:mm"
      Format          =   146407425
      CurrentDate     =   45948
   End
   Begin VB.CommandButton btnCancelar 
      Caption         =   "&Cancelar"
      Height          =   360
      Left            =   3360
      TabIndex        =   10
      Top             =   1560
      Width           =   990
   End
   Begin VB.CommandButton btnSalvar 
      Caption         =   "&Salvar"
      Height          =   360
      Left            =   4440
      TabIndex        =   9
      Top             =   1560
      Width           =   990
   End
   Begin VB.ComboBox cmbStatus 
      Height          =   315
      Left            =   3840
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1080
      Width           =   1575
   End
   Begin VB.TextBox txtValor 
      Height          =   315
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   1335
   End
   Begin VB.TextBox txtDescricao 
      Height          =   315
      Left            =   2160
      MaxLength       =   100
      TabIndex        =   1
      Top             =   480
      Width           =   3255
   End
   Begin VB.TextBox txtNumCartao 
      Height          =   315
      Left            =   240
      MaxLength       =   16
      TabIndex        =   0
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
      Height          =   195
      Left            =   3840
      TabIndex        =   8
      Top             =   840
      Width           =   465
   End
   Begin VB.Label lblData 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dt data"
      Height          =   195
      Left            =   1800
      TabIndex        =   7
      Top             =   840
      Width           =   540
   End
   Begin VB.Label lblValor 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor"
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Top             =   840
      Width           =   360
   End
   Begin VB.Label lblDescricao 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Descrição"
      Height          =   195
      Left            =   2160
      TabIndex        =   5
      Top             =   240
      Width           =   690
   End
   Begin VB.Label lblNumCartao 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Num cartão"
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   825
   End
End
Attribute VB_Name = "Form_CadTransacoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private objTransacaoBO  As New TransacaoBO

Private objTransacaoDto As New TransacaoDTO

Private ModoEdicao      As Boolean

Private idEdicao        As Long

Private Sub btnCancelar_Click()

        '<EhHeader>
        On Error GoTo btnCancelar_Click_Err

        '</EhHeader>

100     LimparCampos
102     Unload Me

        '<EhFooter>
        Exit Sub

btnCancelar_Click_Err:
        RegistrarErro "Form_CadTransacoes", Err.Description
        MsgBox Err.Description & vbCrLf & "in CobuccioCard.Form_CadTransacoes.btnCancelar_Click " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Sub

Public Sub CarregarModoEdicao(ByVal idTransacao As Long)

        '<EhHeader>
        On Error GoTo CarregarModoEdicao_Err

        '</EhHeader>
        Dim rs As ADODB.Recordset
    
100     ModoEdicao = True
102     idEdicao = idTransacao
    
104     Set rs = objTransacaoBO.Listar("Id_Transacao=" & idTransacao)
    
106     If rs Is Nothing Or rs.EOF Then
108         MsgBox "Transação não encontrada.", vbCritical

            Exit Sub

        End If
    
110     With rs
112         txtNumCartao.Text = !Numero_Cartao & ""
114         txtValor.Text = FormatNumber(!Valor_Transacao, 2)
116         dtData.Value = !Data_Transacao
118         txtDescricao.Text = !descricao & ""
        
            Dim i As Integer

120         For i = 0 To cmbStatus.ListCount - 1

122             If cmbStatus.List(i) = Trim(!Status_Transacao) Then
124                 cmbStatus.ListIndex = i

                    Exit For

                End If

126         Next i
        
128         If UCase(Trim(!Status_Transacao)) = "APROVADA" Then
130             MsgBox "Transações aprovadas não podem ser alteradas.", vbExclamation
132             rs.Close
134             Set rs = Nothing
136             Unload Me

                Exit Sub

            End If

        End With
    
138     rs.Close
140     Set rs = Nothing
    
142     Me.Caption = "Editar Transação - ID: " & idTransacao

        '<EhFooter>
        Exit Sub

CarregarModoEdicao_Err:
        RegistrarErro "Form_CadTransacoes", Err.Description
        MsgBox Err.Description & vbCrLf & "in CobuccioCard.Form_CadTransacoes.CarregarModoEdicao " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Sub

Private Sub btnSalvar_Click()

        '<EhHeader>
        On Error GoTo btnSalvar_Click_Err

        '</EhHeader>

100     If Not ValidarCampos() Then Exit Sub
    
102     With objTransacaoDto

104         If ModoEdicao Then
106             .Id_Transacao = idEdicao
            Else
108             .Id_Transacao = 0
            End If
        
110         .Numero_Cartao = txtNumCartao.Text
112         .Valor_Transacao = CCur(txtValor.Text)
114         .Data_Transacao = dtData.Value
116         .descricao = txtDescricao.Text
118         .Status_Transacao = cmbStatus.Text
        End With
    
120     If ModoEdicao Then
    
122         If objTransacaoBO.Salvar(objTransacaoDto) Then
124             MsgBox "Transação atualizada com sucesso!", vbInformation
126             Unload Me
            End If

        Else

128         If objTransacaoBO.Salvar(objTransacaoDto) Then
130             MsgBox "Transação salva com sucesso!", vbInformation
132             Unload Me
            End If
        End If

        '<EhFooter>
        Exit Sub

btnSalvar_Click_Err:
        RegistrarErro "Form_CadTransacoes", Err.Description
        MsgBox Err.Description & vbCrLf & "in CobuccioCard.Form_CadTransacoes.btnSalvar_Click " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Sub

Private Function ValidarCampos() As Boolean

        '<EhHeader>
        On Error GoTo ValidarCampos_Err

        '</EhHeader>

100     ValidarCampos = False

102     If Len(txtNumCartao.Text) <> 16 Or Not IsNumeric(txtNumCartao.Text) Then
104         MsgBox "Número do cartão inválido. Deve ter 16 dígitos.", vbExclamation
106         txtNumCartao.SetFocus

            Exit Function

        End If

108     If Not IsNumeric(txtValor.Text) Or CCur(txtValor.Text) <= 0 Then
110         MsgBox "Informe um valor válido para a transação.", vbExclamation
112         txtValor.SetFocus

            Exit Function

        End If

114     If Trim(txtDescricao.Text) = "" Then
116         MsgBox "Informe a descrição da transação.", vbExclamation
118         txtDescricao.SetFocus

            Exit Function

        End If

120     If cmbStatus.ListIndex = -1 Then
122         MsgBox "Selecione o status da transação.", vbExclamation
124         cmbStatus.SetFocus

            Exit Function

        End If

126     If dtData.Value = 0 Then
128         MsgBox "Informe a data da transação.", vbExclamation
130         dtData.SetFocus

            Exit Function

        End If

132     ValidarCampos = True

        '<EhFooter>
        Exit Function

ValidarCampos_Err:
        RegistrarErro "Form_CadTransacoes", Err.Description
        MsgBox Err.Description & vbCrLf & "in CobuccioCard.Form_CadTransacoes.ValidarCampos " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Function

Private Sub LimparCampos()

        '<EhHeader>
        On Error GoTo LimparCampos_Err

        '</EhHeader>

100     txtNumCartao.Text = ""
102     txtValor.Text = ""
104     txtDescricao.Text = ""
106     cmbStatus.ListIndex = -1
108     dtData.Value = Date
110     txtNumCartao.SetFocus

        '<EhFooter>
        Exit Sub

LimparCampos_Err:
        RegistrarErro "Form_CadTransacoes", Err.Description
        MsgBox Err.Description & vbCrLf & "in CobuccioCard.Form_CadTransacoes.LimparCampos " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Sub

Private Sub PopularStatus()

        '<EhHeader>
        On Error GoTo PopularStatus_Err

        '</EhHeader>

100     cmbStatus.Clear
    
102     cmbStatus.AddItem "Pendente"
104     cmbStatus.AddItem "Aprovada"
106     cmbStatus.AddItem "Cancelada"
    
108     cmbStatus.ListIndex = 0
    
        '<EhFooter>
        Exit Sub

PopularStatus_Err:
        RegistrarErro "Form_CadTransacoes", Err.Description
        MsgBox Err.Description & vbCrLf & "in CobuccioCard.Form_CadTransacoes.PopularStatus " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_Load()
    PopularStatus
End Sub

