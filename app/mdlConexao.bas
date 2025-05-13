Attribute VB_Name = "mdlConexao"
Public Conn As ADODB.Connection

Public Sub ConectarBanco()
    On Error GoTo Erro

    Set Conn = New ADODB.Connection
    Conn.ConnectionString = "Provider=SQLOLEDB;Data Source=DESKTOP-ISQAL42\SQLEXPRESS;Initial Catalog=TransacoesDB;User ID=app_user;Password=q1w2e3r4@!;"
    Conn.Open
    Exit Sub

Erro:
    MsgBox "Erro na conexão com banco: " & Err.Description, vbCritical
    End
End Sub

