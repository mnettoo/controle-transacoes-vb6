VERSION 5.00
Begin VB.Form frmCadastro 
   Caption         =   "Cadastro"
   ClientHeight    =   5145
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7335
   LinkTopic       =   "Form1"
   ScaleHeight     =   5145
   ScaleWidth      =   7335
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtDescricao 
      Height          =   1695
      Left            =   360
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   1800
      Width           =   6375
   End
   Begin VB.ComboBox cmbStatus 
      Height          =   315
      Left            =   4920
      TabIndex        =   3
      Top             =   840
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   5400
      TabIndex        =   7
      Top             =   4320
      Width           =   1575
   End
   Begin VB.CommandButton cmdSalvar 
      Caption         =   "Salvar"
      Height          =   495
      Left            =   3600
      TabIndex        =   5
      Top             =   4320
      Width           =   1575
   End
   Begin VB.TextBox txtNumero 
      Height          =   375
      Left            =   1440
      MaxLength       =   16
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
   Begin VB.TextBox txtValor 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0;(0)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   1
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   840
      Width           =   2295
   End
   Begin VB.TextBox txtData 
      Height          =   375
      Left            =   4920
      TabIndex        =   1
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Status: "
      Height          =   375
      Left            =   4200
      TabIndex        =   11
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Descri��o"
      Height          =   375
      Left            =   480
      TabIndex        =   10
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label lblData 
      Caption         =   "Data da Transa��o:"
      Height          =   375
      Left            =   3360
      TabIndex        =   9
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label lblValor 
      Caption         =   "Valor:"
      Height          =   375
      Left            =   720
      TabIndex        =   8
      Top             =   840
      Width           =   615
   End
   Begin VB.Label lblNumero 
      Caption         =   "N� Cart�o:"
      Height          =   375
      Left            =   480
      TabIndex        =   6
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "frmCadastro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Modo As String
Public ID_Selecionado As Long

Private Sub cmdSalvar_Click()
    On Error GoTo TrataErro

    ' Remove espa�os indesejados
    txtNumero.Text = Trim(txtNumero.Text)
    txtValor.Text = Trim(txtValor.Text)
    txtData.Text = Trim(txtData.Text)
    txtDescricao.Text = Trim(txtDescricao.Text)
    cmbStatus.Text = Trim(cmbStatus.Text)

    ' Valida��o dos campos obrigat�rios

    If txtNumero.Text = "" Then
        MsgBox "Informe o n�mero do cart�o.", vbExclamation
        txtNumero.SetFocus
        Exit Sub
    End If

    If Len(txtNumero.Text) <> 16 Or Not IsNumeric(txtNumero.Text) Then
        MsgBox "O n�mero do cart�o deve ter 16 d�gitos num�ricos.", vbExclamation
        txtNumero.SetFocus
        Exit Sub
    End If

    If txtValor.Text = "" Or Not IsNumeric(txtValor.Text) Or CDbl(txtValor.Text) <= 0 Then
        MsgBox "Informe um valor monet�rio v�lido maior que zero.", vbExclamation
        txtValor.SetFocus
        Exit Sub
    End If

    If txtData.Text = "" Or Not ValidarData(txtData.Text) Then
        MsgBox "Informe uma data v�lida no formato DD/MM/AAAA.", vbExclamation
        txtData.SetFocus
        Exit Sub
    End If

    If txtDescricao.Text = "" Then
        MsgBox "Informe uma descri��o.", vbExclamation
        txtDescricao.SetFocus
        Exit Sub
    End If

    If cmbStatus.Text = "" Then
        MsgBox "Selecione o status da transa��o.", vbExclamation
        cmbStatus.SetFocus
        Exit Sub
    End If

    If cmbStatus.Text <> "Aprovada" And cmbStatus.Text <> "Pendente" And cmbStatus.Text <> "Cancelada" Then
        MsgBox "Status inv�lido. Escolha entre Aprovada, Pendente ou Cancelada.", vbExclamation
        cmbStatus.SetFocus
        Exit Sub
    End If

    Dim dataFormatada As String
    dataFormatada = FormatarDataParaSQL(txtData.Text)

    If dataFormatada = "" Then
        MsgBox "Erro ao formatar a data.", vbCritical
        txtData.SetFocus
        Exit Sub
    End If

    Dim valorSQL As String
    valorSQL = Replace(Replace(txtValor.Text, ".", ""), ",", ".")

    Dim sql As String
    If Modo = "Novo" Then
        sql = "INSERT INTO Transacoes (Numero_Cartao, Valor_Transacao, Data_Transacao, Descricao, Status_Transacao) VALUES (" & _
              "'" & txtNumero.Text & "'," & valorSQL & ", CONVERT(DATETIME, '" & dataFormatada & "', 120),'" & txtDescricao.Text & "','" & cmbStatus.Text & "')"
    Else
        sql = "UPDATE Transacoes SET Numero_Cartao = '" & txtNumero.Text & "', Valor_Transacao = " & valorSQL & _
              ", Data_Transacao = CONVERT(DATETIME, '" & dataFormatada & "', 120), Descricao = '" & txtDescricao.Text & "', Status_Transacao = '" & cmbStatus.Text & "' " & _
              "WHERE Id_Transacao = " & ID_Selecionado
    End If

    Conn.Execute sql
    MsgBox "Transa��o salva com sucesso!"
    Unload Me
    frmPrincipal.CarregarTransacoes
    Exit Sub

TrataErro:
    MsgBox "Erro ao salvar: " & Err.Description, vbCritical
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    cmbStatus.AddItem "Aprovada"
    cmbStatus.AddItem "Pendente"
    cmbStatus.AddItem "Cancelada"

    If Modo = "Novo" Then
        txtData.Text = Format(Date, "dd/mm/yyyy")
    Else
        ' Modo edi��o: carregar os dados do banco
        Dim rs As ADODB.Recordset
        Set rs = New ADODB.Recordset
        rs.Open "SELECT * FROM Transacoes WHERE Id_Transacao = " & ID_Selecionado, Conn, adOpenStatic, adLockReadOnly

        If Not rs.EOF Then
            txtNumero.Text = rs!Numero_Cartao
            txtValor.Text = Format(rs!Valor_Transacao, "0.00")
            txtDescricao.Text = rs!Descricao
            txtData.Text = Format(rs!Data_Transacao, "dd/mm/yyyy")
            cmbStatus.Text = rs!Status_Transacao

            If rs!Status_Transacao = "Aprovada" Then
                txtNumero.Enabled = False
                txtValor.Enabled = False
                txtDescricao.Enabled = False
                txtData.Enabled = False
                cmbStatus.Enabled = False
            End If
        End If
        rs.Close
    End If
End Sub


Private Sub txtData_GotFocus()
    If txtData.Text = "DD/MM/AAAA" Then
        txtData.Text = ""
        txtData.ForeColor = vbWindowText
    End If
End Sub

Private Sub txtData_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= 48 And KeyAscii <= 57) And KeyAscii <> 8 Then
        KeyAscii = 0
        Exit Sub
    End If

    Dim txt As String
    txt = txtData.Text

    If KeyAscii = 8 Then Exit Sub

    Select Case Len(txt)
        Case 2, 5
            txtData.Text = txt & "/"
            txtData.SelStart = Len(txtData.Text)
    End Select
End Sub

Private Sub txtData_LostFocus()
    If Trim(txtData.Text) = "" Then
        txtData.Text = "DD/MM/AAAA"
        txtData.ForeColor = vbGrayText
    ElseIf Not ValidarData(txtData.Text) Then
        MsgBox "Data inv�lida. Digite uma data real no formato DD/MM/AAAA.", vbExclamation
        txtData.SetFocus
    End If
End Sub

Private Sub txtValor_KeyPress(KeyAscii As Integer)
    If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 44) Then
        KeyAscii = 0
        Exit Sub
    End If
    If KeyAscii = 44 And InStr(txtValor.Text, ",") > 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtValor_LostFocus()
    If Trim(txtValor.Text) <> "" Then
        Dim valor As Double
        On Error GoTo invalido
        valor = CDbl(Replace(txtValor.Text, ".", ""))
        txtValor.Text = Format(valor, "0.00")
        Exit Sub
invalido:
        MsgBox "Valor inv�lido. Digite apenas n�meros e uma v�rgula.", vbExclamation
        txtValor.Text = ""
        txtValor.SetFocus
    End If
End Sub


