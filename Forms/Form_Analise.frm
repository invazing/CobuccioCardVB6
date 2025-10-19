VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Form_Analise 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Analise - View e TVF"
   ClientHeight    =   10020
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13965
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form_Analise.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10020
   ScaleWidth      =   13965
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FrmeDataGridView 
      Height          =   9015
      Left            =   0
      TabIndex        =   6
      Top             =   960
      Width           =   13935
      Begin MSFlexGridLib.MSFlexGrid gridResumoFinanceiro 
         Height          =   4335
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   13695
         _ExtentX        =   24156
         _ExtentY        =   7646
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid gridAnaliseFinanceira 
         Height          =   4335
         Left            =   120
         TabIndex        =   8
         Top             =   4560
         Width           =   13695
         _ExtentX        =   24156
         _ExtentY        =   7646
         _Version        =   393216
      End
   End
   Begin VB.Frame FrmeDetalhePesquisa 
      Caption         =   "Dados da Consulta"
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13935
      Begin VB.CommandButton btnPesquisa 
         Caption         =   "Consultar"
         Height          =   480
         Left            =   12720
         TabIndex        =   1
         Top             =   240
         Width           =   990
      End
      Begin MSComCtl2.DTPicker dtFim 
         Height          =   315
         Left            =   1920
         TabIndex        =   2
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Format          =   146341889
         CurrentDate     =   45948
      End
      Begin MSComCtl2.DTPicker dtInicio 
         Height          =   315
         Left            =   360
         TabIndex        =   3
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Format          =   146341889
         CurrentDate     =   45948
      End
      Begin VB.Label lblInicio 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data Inicio"
         Height          =   195
         Left            =   360
         TabIndex        =   5
         Top             =   240
         Width           =   765
      End
      Begin VB.Label lblFim 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fim"
         Height          =   195
         Left            =   1920
         TabIndex        =   4
         Top             =   240
         Width           =   240
      End
   End
End
Attribute VB_Name = "Form_Analise"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnPesquisa_Click()

        '<EhHeader>
        On Error GoTo btnPesquisa_Click_Err

        '</EhHeader>

        Dim objTransacaoBO As New TransacaoBO

        Dim rsResumo       As ADODB.Recordset

        Dim rsAnalise      As ADODB.Recordset

100     LimparGrid gridResumoFinanceiro
102     LimparGrid gridAnaliseFinanceira

104     Set rsResumo = objTransacaoBO.ObterTransacoesCategorizadas(CDate(dtInicio.Value), CDate(dtFim.Value))

106     PreencherGridComRecordset gridResumoFinanceiro, rsResumo

108     Set rsAnalise = objTransacaoBO.ObterResumoFinanceiro()
110     PreencherGridComRecordset gridAnaliseFinanceira, rsAnalise

112     If Not rsResumo Is Nothing Then rsResumo.Close
114     If Not rsAnalise Is Nothing Then rsAnalise.Close
116     Set rsResumo = Nothing
118     Set rsAnalise = Nothing
120     Set objTransacaoBO = Nothing

        Exit Sub

        '<EhFooter>
        Exit Sub

btnPesquisa_Click_Err:
        RegistrarErro "Form_Analise", Err.Description
        MsgBox Err.Description & vbCrLf & "in CobuccioCard.Form_Analise.btnPesquisa_Click " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_Load()

        '<EhHeader>
        On Error GoTo Form_Load_Err

        '</EhHeader>

100     ConfigurarGridResumoFinanceiro gridResumoFinanceiro
102     ConfigurarGridAnaliseFinanceira gridAnaliseFinanceira

104     AjustarLarguraAutomatica gridResumoFinanceiro
106     AjustarLarguraAutomatica gridAnaliseFinanceira

108     dtInicio.Value = DateAdd("d", -7, Date)
110     dtFim.Value = Date

112     If Not OpenConnection() Then
114         MsgBox "Erro ao conectar ao banco de dados.", vbCritical

            Exit Sub

        End If

        Exit Sub

Form_Load_Err:
        RegistrarErro "Form_Analise", Err.Description
        MsgBox Err.Description & vbCrLf & "in CobuccioCard.Form_Analise.Form_Load " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_Resize()
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>
    
    AjustarLarguraAutomatica gridResumoFinanceiro
    AjustarLarguraAutomatica gridAnaliseFinanceira
End Sub



