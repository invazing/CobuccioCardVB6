Attribute VB_Name = "modLogHelper"
Option Explicit

Public Sub RegistrarErro(ByVal metodo As String, ByVal descricaoErro As String)

        '<EhHeader>
        On Error GoTo RegistrarErro_Err

        '</EhHeader>
  
        Dim caminhoLog As String

        Dim arquivoNum As Integer
    
100     caminhoLog = App.Path

102     If Right(caminhoLog, 1) <> "\" Then caminhoLog = caminhoLog & "\"
104     caminhoLog = caminhoLog & "cobucciocard.log"
    
106     arquivoNum = FreeFile
    
108     Open caminhoLog For Append As #arquivoNum
    
110     Print #arquivoNum, "[" & Format(Now, "yyyy-mm-dd HH:nn:ss") & "] ERRO em " & metodo
112     Print #arquivoNum, "Mensagem: " & descricaoErro
114     Print #arquivoNum, String(80, "-")
    
116     Close #arquivoNum
    
        '<EhFooter>
        Exit Sub

RegistrarErro_Err:
        MsgBox Err.Description & vbCrLf & "in CobuccioCard.modLogHelper.RegistrarErro " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Sub
