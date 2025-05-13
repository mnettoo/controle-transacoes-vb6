Attribute VB_Name = "mdlUtil"
Public Function FormatarDataParaSQL(dataTexto As String) As String
    On Error GoTo Invalida

    Dim partes() As String
    Dim dia As Integer, mes As Integer, ano As Integer
    Dim dataMontada As String

    partes = Split(dataTexto, "/")
    If UBound(partes) <> 2 Then GoTo Invalida

    dia = Val(partes(0))
    mes = Val(partes(1))
    ano = Val(partes(2))

    If dia < 1 Or dia > 31 Then GoTo Invalida
    If mes < 1 Or mes > 12 Then GoTo Invalida
    If ano < 1900 Or ano > 2100 Then GoTo Invalida

    dataMontada = Format(dia, "00") & "/" & Format(mes, "00") & "/" & Format(ano, "0000")
    If Not IsDate(dataMontada) Then GoTo Invalida

    FormatarDataParaSQL = Format(ano, "0000") & "-" & Format(mes, "00") & "-" & Format(dia, "00")
    Exit Function

Invalida:
    FormatarDataParaSQL = ""
End Function

Public Function ValidarData(valor As String) As Boolean
    On Error GoTo Invalida
    Dim partes() As String
    partes = Split(valor, "/")

    If UBound(partes) <> 2 Then GoTo Invalida
    If Not IsNumeric(partes(0)) Or Not IsNumeric(partes(1)) Or Not IsNumeric(partes(2)) Then GoTo Invalida

    Dim dt As Date
    dt = DateSerial(CInt(partes(2)), CInt(partes(1)), CInt(partes(0)))

    ValidarData = True
    Exit Function

Invalida:
    ValidarData = False
End Function
