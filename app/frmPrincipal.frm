VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmPrincipal 
   Caption         =   "Controle de Transações"
   ClientHeight    =   6675
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12585
   LinkTopic       =   "Form1"
   ScaleHeight     =   6675
   ScaleWidth      =   12585
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog dlgSalvar 
      Left            =   8760
      Top             =   6120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdExportar 
      Caption         =   "Exportar"
      Height          =   615
      Left            =   4080
      TabIndex        =   8
      Top             =   5880
      Width           =   3255
   End
   Begin VB.Frame frFiltrar 
      Caption         =   "Filtrar"
      Height          =   1815
      Left            =   360
      TabIndex        =   10
      Top             =   240
      Width           =   11775
      Begin VB.TextBox txtNumero 
         Height          =   375
         Left            =   1560
         MaxLength       =   16
         TabIndex        =   0
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox txtValor 
         Height          =   375
         Left            =   1560
         TabIndex        =   2
         Top             =   1080
         Width           =   1695
      End
      Begin VB.TextBox txtData 
         Height          =   375
         Left            =   5640
         TabIndex        =   3
         Top             =   960
         Width           =   1695
      End
      Begin VB.ComboBox cmbStatus 
         Height          =   315
         Left            =   5640
         TabIndex        =   1
         Top             =   480
         Width           =   1695
      End
      Begin VB.CommandButton cmdFiltrar 
         Caption         =   "Filtrar"
         Height          =   495
         Left            =   7680
         TabIndex        =   4
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label lblStatus 
         Caption         =   "Satus da transação"
         Height          =   375
         Left            =   3960
         TabIndex        =   14
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label lblData 
         Caption         =   "Data da transção"
         Height          =   375
         Left            =   3960
         TabIndex        =   13
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label lblValor 
         Caption         =   "Valor:"
         Height          =   375
         Left            =   960
         TabIndex        =   12
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label lblNumero 
         Caption         =   "Cartão: "
         Height          =   375
         Left            =   960
         TabIndex        =   11
         Top             =   480
         Width           =   615
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3495
      Left            =   360
      TabIndex        =   9
      Top             =   2280
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   6165
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdExcluir 
      Caption         =   "Excluir"
      Height          =   495
      Left            =   10920
      TabIndex        =   7
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton cmdEditar 
      Caption         =   "Editar"
      Height          =   495
      Left            =   10920
      TabIndex        =   6
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton cmdNovo 
      Caption         =   "Incluir"
      Height          =   495
      Left            =   10920
      TabIndex        =   5
      Top             =   2280
      Width           =   1215
   End
End
Attribute VB_Name = "frmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ID_Selecionado As Long

Private Sub cmdEditar_Click()
    If DataGrid1.ApproxCount = 0 Then Exit Sub
    ID_Selecionado = DataGrid1.Columns(0).Value 'coluna 0 = Id_Transacao
    frmCadastro.ID_Selecionado = ID_Selecionado
    
    If Status = "Aprovada" Then
        MsgBox "Transações com status 'Aprovada' não podem ser editadas.", vbExclamation
        Exit Sub
    End If
    
    frmCadastro.Modo = "Editar"
    frmCadastro.Show vbModal
    Call CarregarTransacoes
End Sub

Private Sub cmdExcluir_Click()
    If DataGrid1.ApproxCount = 0 Then Exit Sub

    Dim id As Long
    id = DataGrid1.Columns(0).Value

    If MsgBox("Confirma exclusão da transação " & id & "?", vbYesNo + vbQuestion) = vbYes Then
        On Error GoTo TrataErro
        Conn.Execute "DELETE FROM Transacoes WHERE Id_Transacao = " & id
        Call CarregarTransacoes
    End If
    Exit Sub

TrataErro:
    Call GravarLog("Erro ao excluir: " & Err.Description)
    MsgBox "Erro ao excluir: " & Err.Description
End Sub

Private Sub cmdExportar_Click()
    On Error GoTo TrataErro

    Dim rs As ADODB.Recordset
    Set rs = DataGrid1.DataSource

    If rs Is Nothing Or rs.EOF Then
        MsgBox "Nenhum dado para exportar.", vbExclamation
        Exit Sub
    End If

    With dlgSalvar
        .DialogTitle = "Salvar como"
        .Filter = "Arquivos CSV (*.csv)|*.csv"
        .FileName = "transacoes_exportadas.csv"
        .ShowSave
    End With

    Dim caminho As String
    caminho = dlgSalvar.FileName
    If caminho = "" Then Exit Sub

    Dim arq As Integer
    arq = FreeFile

    Open caminho For Output As #arq

    Dim linha As String
    Dim i As Integer

    linha = "ID;Número do Cartão;Valor;Data;Descrição;Status"
    Print #arq, linha

    rs.MoveFirst
    Do While Not rs.EOF
        linha = ""
        For i = 0 To rs.Fields.Count - 1
            linha = linha & rs.Fields(i).Value
            If i < rs.Fields.Count - 1 Then linha = linha & ";"
        Next i
        Print #arq, linha
        rs.MoveNext
    Loop

    Close #arq

    MsgBox "Exportado com sucesso para: " & caminho, vbInformation
    Exit Sub

TrataErro:
    MsgBox "Erro ao exportar: " & Err.Description, vbCritical
End Sub


Private Sub cmdNovo_Click()
    frmCadastro.Modo = "Novo"
    frmCadastro.Show vbModal
    Call CarregarTransacoes
End Sub

Private Sub Form_Load()
    Call ConectarBanco
    Call CarregarTransacoes
    
    txtData.Text = "DD/MM/AAAA"
    txtData.ForeColor = vbGrayText

    cmbStatus.AddItem ""
    cmbStatus.AddItem "Aprovada"
    cmbStatus.AddItem "Pendente"
    cmbStatus.AddItem "Cancelada"
End Sub

Sub CarregarTransacoes(Optional ByVal filtroSQL As String = "")
    Dim rs As ADODB.Recordset
    Dim sql As String

    sql = "SELECT * FROM Transacoes WHERE 1=1 " & filtroSQL & " ORDER BY Data_Transacao DESC"

    Set rs = New ADODB.Recordset
    rs.Open sql, Conn, adOpenStatic, adLockReadOnly
    Set DataGrid1.DataSource = rs
End Sub

Private Sub cmdFiltrar_Click()
    Dim filtro As String
    filtro = ""

    If Trim(txtNumero.Text) <> "" Then
        filtro = filtro & " AND Numero_Cartao LIKE '%" & txtNumero.Text & "%'"
    End If

    If Trim(txtValor.Text) <> "" Then
        If IsNumeric(txtValor.Text) Then
            filtro = filtro & " AND Valor_Transacao = " & Replace(txtValor.Text, ",", ".")
        End If
    End If

    If Trim(cmbStatus.Text) <> "" Then
        filtro = filtro & " AND Status_Transacao = '" & cmbStatus.Text & "'"
    End If

    If Trim(txtData.Text) <> "" And Trim(txtData.Text) <> "DD/MM/AAAA" Then
        dataSQL = FormatarDataParaSQL(txtData.Text)
        If dataSQL <> "" Then
            filtro = filtro & " AND CONVERT(DATE, Data_Transacao) = '" & dataSQL & "'"
        Else
            MsgBox "Data inválida. Use o formato DD/MM/AAAA.", vbExclamation
            txtData.SetFocus
            Exit Sub
        End If
    End If

    Call CarregarTransacoes(filtro)
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
        MsgBox "Data inválida. Digite uma data real no formato DD/MM/AAAA.", vbExclamation
        txtData.SetFocus
    End If
End Sub
