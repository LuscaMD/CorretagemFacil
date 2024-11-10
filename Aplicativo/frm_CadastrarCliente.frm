VERSION 5.00
Begin VB.Form frm_CadastrarCliente 
   Caption         =   "Cadastro de Cliente"
   ClientHeight    =   5010
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5370
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5010
   ScaleWidth      =   5370
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmb_Corretor 
      Height          =   315
      ItemData        =   "frm_CadastrarCliente.frx":0000
      Left            =   1560
      List            =   "frm_CadastrarCliente.frx":0002
      TabIndex        =   13
      Top             =   2040
      Width           =   3495
   End
   Begin VB.ComboBox cmb_Cidades 
      Height          =   315
      ItemData        =   "frm_CadastrarCliente.frx":0004
      Left            =   1560
      List            =   "frm_CadastrarCliente.frx":0006
      TabIndex        =   12
      Top             =   2640
      Width           =   3495
   End
   Begin VB.CommandButton btn_Limpar 
      Caption         =   "Limpar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   578
      TabIndex        =   11
      Top             =   3720
      Width           =   2055
   End
   Begin VB.CommandButton btn_Salvar 
      Caption         =   "Salvar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2738
      TabIndex        =   10
      Top             =   3720
      Width           =   2055
   End
   Begin VB.ComboBox cmb_UF 
      Height          =   315
      ItemData        =   "frm_CadastrarCliente.frx":0008
      Left            =   1560
      List            =   "frm_CadastrarCliente.frx":000A
      TabIndex        =   7
      Top             =   3240
      Width           =   1455
   End
   Begin VB.TextBox txt_Endereco 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      MaxLength       =   50
      TabIndex        =   5
      Top             =   840
      Width           =   3495
   End
   Begin VB.TextBox txt_Nome 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1560
      MaxLength       =   50
      TabIndex        =   2
      Top             =   240
      Width           =   3495
   End
   Begin VB.TextBox txt_CPF 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      MaxLength       =   14
      TabIndex        =   1
      Top             =   1440
      Width           =   3495
   End
   Begin VB.CheckBox chk_Ativo 
      Alignment       =   1  'Right Justify
      Caption         =   "Ativo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   0
      Top             =   3240
      Width           =   1080
   End
   Begin VB.Label lbl_Corretor 
      Alignment       =   1  'Right Justify
      Caption         =   "Corretor:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   14
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label lbl_Cidade 
      Alignment       =   1  'Right Justify
      Caption         =   "Cidade:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   9
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label lbl_UF 
      Alignment       =   1  'Right Justify
      Caption         =   "UF:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   8
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Label lbl_Endereco 
      Alignment       =   1  'Right Justify
      Caption         =   "Endereco:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label lbl_Nome 
      Alignment       =   1  'Right Justify
      Caption         =   "Nome:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "CPF:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   1440
      Width           =   1095
   End
End
Attribute VB_Name = "frm_CadastrarCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_Limpar_Click()
  Me.txt_CPF = ""
  Me.txt_Endereco = ""
  Me.txt_Nome = ""
  Me.cmb_Cidades = ""
  Me.cmb_Corretor = ""
  Me.cmb_UF = ""
  Me.chk_Ativo.Value = 0
End Sub

Private Sub btn_Salvar_Click()
 Dim str_SQL As String
  
  If ValidarDados Then
    str_SQL = "INSERT INTO Cadastros VALUES('" & Me.txt_Nome & "', '" & Me.txt_CPF & "', '" & Me.txt_Endereco & "', " & Me.cmb_UF.ItemData(Me.cmb_UF.ListIndex) & ", " & Me.cmb_Cidades.ItemData(Me.cmb_Cidades.ListIndex) & ", " & Me.chk_Ativo & ", 0, '" & Me.cmb_Corretor.ItemData(Me.cmb_Corretor.NewIndex) & "') "
  
    ExecutarComando str_SQL
    
    Call btn_Limpar_Click
    MsgBox "O Cliente foi salvo com sucesso."
  End If
End Sub

Private Function ValidarDados() As Boolean
  Dim str_SQL As String
  
  ValidarDados = True
  
  If Nz(Me.txt_Nome, "") = "" Then
    ValidarDados = False
    MsgBox "Ocorreu um erro ao salvar os dados, o campo 'Nome' nao pode estar vazio."
    Exit Function
  End If
  
  If Nz(Me.txt_CPF, "") = "" Or Len(Me.txt_CPF) <> 14 Then
    ValidarDados = False
    MsgBox "Ocorreu um erro ao salvar os dados, o campo 'CPF' nao pode estar vazio e deve estar preenchido corretamente."
    Exit Function
  ElseIf Not VerificaCPF(Me.txt_CPF) Then
    ValidarDados = False
    MsgBox "Ocorreu um erro ao salvar os dados, já existe um cadastro com o 'CPF' preenchido."
    Exit Function
  End If
  
  If Nz(Me.cmb_Corretor, "") = "" Then
    ValidarDados = False
    MsgBox "Ocorreu um erro ao salvar os dados, o campo 'Corretor' nao pode estar vazio."
    Exit Function
  End If
  
  If Nz(Me.txt_Endereco, "") = "" Then
    ValidarDados = False
    MsgBox "Ocorreu um erro ao salvar os dados, o campo 'Endereco' nao pode estar vazio."
    Exit Function
  End If
  
  If Nz(Me.cmb_UF, "") = "" Then
    ValidarDados = False
    MsgBox "Ocorreu um erro ao salvar os dados, o campo 'UF' nao pode estar vazio."
    Exit Function
  End If
  
  If Nz(Me.cmb_Cidades, "") = "" Then
    ValidarDados = False
    MsgBox "Ocorreu um erro ao salvar os dados, o campo 'Cidade' nao pode estar vazio."
    Exit Function
  End If
  
End Function

Private Sub cmb_UF_Click()
  Dim recordSourceCidade As Object
  
  If Nz(Me.cmb_Cidades, "") <> "" Then
    Set recordSourceCidade = RetornarDados("SELECT * FROM Cidades WHERE fk_int_IdUF = " & Me.cmb_UF.ItemData(Me.cmb_UF.ListIndex) & " AND str_NomeCidade = '" & Me.cmb_Cidades & "'")
    
    If Not recordSourceCidade.EOF Then
      recordSourceCidade.Close
      Set recordSourceCidade = Nothing
    Else
      Call PreencherCmbCidade
    End If
  Else
    Call PreencherCmbCidade
  End If
End Sub

Private Sub PreencherCmbCidade()
  Dim recordSourceCidade As Object
  
  Me.cmb_Cidades.Clear
    
  Set recordSourceCidade = RetornarDados("SELECT * FROM Cidades WHERE fk_int_IdUF = " & Me.cmb_UF.ItemData(Me.cmb_UF.ListIndex))

  Do While Not recordSourceCidade.EOF
    Me.cmb_Cidades.AddItem recordSourceCidade!str_NomeCidade
    Me.cmb_Cidades.ItemData(Me.cmb_Cidades.NewIndex) = recordSourceCidade!pk_int_IdCidade
    recordSourceCidade.MoveNext
  Loop
  
  recordSourceCidade.Close
  Set recordSourceCidade = Nothing
End Sub

Private Sub Form_Load()
  Dim recordSource As Object
  
  Set recordSource = RetornarDados("SELECT * FROM Estados")
  
  Do While Not recordSource.EOF
    Me.cmb_UF.AddItem recordSource!str_UF
    Me.cmb_UF.ItemData(Me.cmb_UF.NewIndex) = recordSource!pk_int_IdUF
    recordSource.MoveNext
  Loop
  
  Set recordSource = RetornarDados("SELECT * FROM Cadastros WHERE bit_Ativo = 1 AND bit_Corretor = 1")
  
  Do While Not recordSource.EOF
    Me.cmb_Corretor.AddItem recordSource!str_Nome
    Me.cmb_Corretor.ItemData(Me.cmb_Corretor.NewIndex) = recordSource!int_CodCorretor
    recordSource.MoveNext
  Loop
  
  recordSource.Close
  Set recordSource = Nothing
End Sub

Private Sub txt_CPF_KeyPress(KeyAscii As Integer)
  Call VerificaMascaraCPF(Me, KeyAscii)
End Sub
