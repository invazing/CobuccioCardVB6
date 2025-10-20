Attribute VB_Name = "modConnection"
Option Explicit

Public cn As ADODB.Connection

Public Function OpenConnection() As Boolean

        '<EhHeader>
        On Error GoTo OpenConnection_Err

        '</EhHeader>

100     If cn Is Nothing Then
102         Set cn = New ADODB.Connection
        End If

104     If cn.State = adStateOpen Then
106         OpenConnection = True

            Exit Function

        End If

        Dim strConn As String
    
108     strConn = "Provider=SQLOLEDB;Data Source=192.168.100.1;Initial Catalog=cobucciocard;User ID=cobuccio;Password=A@kB]vxW12u2025;TrustServerCertificate=Yes;"

110     cn.Open strConn

112     OpenConnection = True

        '<EhFooter>
        Exit Function

OpenConnection_Err:
        RegistrarErro "modConnection", Err.Description
        MsgBox Err.Description & vbCrLf & "in CobuccioCard.modConnection.OpenConnection " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"
        OpenConnection = False

        Resume Next

        '</EhFooter>

End Function

Public Sub CloseConnection()

        '<EhHeader>
        On Error GoTo CloseConnection_Err

        '</EhHeader>

        On Error Resume Next

100     If Not cn Is Nothing Then
102         If cn.State = adStateOpen Then cn.Close
        End If

104     Set cn = Nothing
    
        '<EhFooter>
        Exit Sub

CloseConnection_Err:
        RegistrarErro "modConnection", Err.Description
        MsgBox Err.Description & vbCrLf & "in CobuccioCard.modConnection.CloseConnection " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Sub

