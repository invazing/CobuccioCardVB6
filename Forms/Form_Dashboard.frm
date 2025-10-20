VERSION 5.00
Begin VB.Form Form_Dashboard 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CobuccioCard - Grupo Adriano Cobuccio"
   ClientHeight    =   4005
   ClientLeft      =   150
   ClientTop       =   795
   ClientWidth     =   7980
   DrawStyle       =   1  'Dash
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form_Dashboard.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   7980
   StartUpPosition =   3  'Windows Default
   Begin VB.Image imgDashboard 
      DragMode        =   1  'Automatic
      Height          =   4000
      Left            =   0
      Picture         =   "Form_Dashboard.frx":446A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8000
   End
   Begin VB.Menu MenuCadastro 
      Caption         =   "Cadastro"
      Index           =   1
      NegotiatePosition=   1  'Left
   End
   Begin VB.Menu MenuGerenciamento 
      Caption         =   "Gerenciamento"
      Index           =   2
      NegotiatePosition=   1  'Left
   End
   Begin VB.Menu MenuDados 
      Caption         =   "Dados"
      Index           =   3
      NegotiatePosition=   2  'Middle
   End
   Begin VB.Menu MenuAnalise 
      Caption         =   "Analise"
      Index           =   4
      NegotiatePosition=   2  'Middle
   End
End
Attribute VB_Name = "Form_Dashboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()

        '<EhHeader>
        On Error GoTo Form_Load_Err

        '</EhHeader>

100     If Not OpenConnection() Then
102         MsgBox "Falha na conexão com o banco de dados.", vbCritical
104         Unload Me
        End If

        '<EhFooter>
        Exit Sub

Form_Load_Err:
        RegistrarErro "Form_Dashboard", Err.Description
        MsgBox Err.Description & vbCrLf & "in CobuccioCard.Form_Dashboard.Form_Load " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Sub

Private Sub MenuAnalise_Click(Index As Integer)

        '<EhHeader>
        On Error GoTo MenuAnalise_Click_Err

        '</EhHeader>

100     Form_Analise.Show

        '<EhFooter>
        Exit Sub

MenuAnalise_Click_Err:
        RegistrarErro "Form_Dashboard", Err.Description
        MsgBox Err.Description & vbCrLf & "in CobuccioCard.Form_Dashboard.MenuAnalise_Click " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Sub

Private Sub MenuCadastro_Click(Index As Integer)

        '<EhHeader>
        On Error GoTo MenuCadastro_Click_Err

        '</EhHeader>

100     Form_CadTransacoes.Show

        '<EhFooter>
        Exit Sub

MenuCadastro_Click_Err:
        RegistrarErro "Form_Dashboard", Err.Description
        MsgBox Err.Description & vbCrLf & "in CobuccioCard.Form_Dashboard.MenuCadastro_Click " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Sub

Private Sub MenuDados_Click(Index As Integer)

        '<EhHeader>
        On Error GoTo MenuDados_Click_Err

        '</EhHeader>
    
100     If MsgBox("Deseja gerar 1000 registros de transações?", vbYesNo + vbQuestion, "Gerar Transações") = vbYes Then
        
102         Call GerarTransacoes(1000)

        End If
        
        '<EhFooter>
        Exit Sub

MenuDados_Click_Err:
        RegistrarErro "Form_Dashboard", Err.Description
        MsgBox Err.Description & vbCrLf & "in CobuccioCard.Form_Dashboard.MenuDados_Click " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Sub

Private Sub MenuGerenciamento_Click(Index As Integer)

        '<EhHeader>
        On Error GoTo MenuGerenciamento_Click_Err

        '</EhHeader>

100     Form_PsqTransacoes.Show

        '<EhFooter>
        Exit Sub

MenuGerenciamento_Click_Err:
        RegistrarErro "Form_Dashboard", Err.Description
        MsgBox Err.Description & vbCrLf & "in CobuccioCard.Form_Dashboard.MenuGerenciamento_Click " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Sub
