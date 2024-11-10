VERSION 5.00
Begin VB.Form frm_CadastrarCorretor 
   Caption         =   "Cadastro de Corretor"
   ClientHeight    =   4050
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5235
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   5235
   StartUpPosition =   2  'CenterScreen
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
      Left            =   600
      TabIndex        =   8
      Top             =   2040
      Width           =   1080
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
      Left            =   2670
      TabIndex        =   7
      Top             =   2640
      Width           =   2055
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
      Left            =   510
      TabIndex        =   6
      Top             =   2640
      Width           =   2055
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
      Left            =   1440
      MaxLength       =   14
      TabIndex        =   5
      Top             =   1440
      Width           =   3495
   End
   Begin VB.TextBox txt_Codigo 
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
      Left            =   1440
      MaxLength       =   4
      TabIndex        =   3
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
      Height          =   375
      Left            =   1440
      MaxLength       =   50
      TabIndex        =   1
      Top             =   240
      Width           =   3495
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
      Left            =   240
      TabIndex        =   4
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label lbl_Codigo 
      Alignment       =   1  'Right Justify
      Caption         =   "Codigo:"
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
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   1095
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
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "frm_CadastrarCorretor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_Limpar_Click()
  Me.txt_Nome = ""
  Me.txt_CPF = ""
  Me.txt_Codigo = ""
  Me.chk_Ativo.Value = 0
End Sub

Private Sub btn_Salvar_Click()
  Dim str_SQL As String
  
  If ValidarDados Then
    str_SQL = "INSERT INTO Cadastros VALUES('" & Me.txt_Nome & "', '" & Me.txt_CPF & "', " & "NULL, NULL, NULL, " & Me.chk_Ativo & ", 1, '" & Me.txt_Codigo & "') "
  
    ExecutarComando str_SQL
    
    Call btn_Limpar_Click
    MsgBox "O Corretor foi salvo com sucesso."
  End If
End Sub

Private Function ValidarDados() As Boolean
  Dim str_SQL As String
  
  ValidarDados = True
  
  If Nz(Me.txt_Nome, "") = "" Then
    ValidarDados = False
    MsgBox "Ocorreu um erro ao salvar os dados, o campo 'Nome' não pode estar vazio."
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
  
  If Nz(Me.txt_Codigo, "") = "" Then
    ValidarDados = False
    MsgBox "Ocorreu um erro ao salvar os dados, o campo 'Codigo' nao pode estar vazio."
    Exit Function
  ElseIf Nz(Me.txt_Codigo, "") <> "" And IsNumeric(Me.txt_Codigo) = False Then
    ValidarDados = False
    MsgBox "Ocorreu um erro ao salvar os dados, o campo 'Codigo' nao pode letras."
    Exit Function
  ElseIf Not VerificaCodCorretor(Me.txt_Codigo) Then
    ValidarDados = False
    MsgBox "Ocorreu um erro ao salvar os dados, já existe um cadastro com o 'Codigo' preenchido."
    Exit Function
  End If
  
End Function

Private Sub txt_Codigo_KeyPress(KeyAscii As Integer)
 Call VerificaDigitoNumerico(Me, KeyAscii)
End Sub

Private Sub txt_CPF_KeyPress(KeyAscii As Integer)
 Call VerificaMascaraCPF(Me, KeyAscii)
End Sub
