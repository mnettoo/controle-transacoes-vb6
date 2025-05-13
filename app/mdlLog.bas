Attribute VB_Name = "mdlLog"
Public Sub GravarLog(msg As String)
    Dim num As Integer
    num = FreeFile

    Open App.Path & "\log_erros.txt" For Append As #num
    Print #num, Now & " - " & msg
    Close #num
End Sub
