VERSION 5.00
Begin VB.Form frm_MenuPrincipal 
   Caption         =   "Menu Principal"
   ClientHeight    =   6210
   ClientLeft      =   10245
   ClientTop       =   4260
   ClientWidth     =   7725
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   7725
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btn_CadastrarCliente 
      Caption         =   "Cadastrar Cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   1215
      TabIndex        =   2
      Top             =   3120
      Width           =   5295
   End
   Begin VB.CommandButton btn_CadastrarCorretor 
      Caption         =   "Cadastrar Corretor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   1215
      TabIndex        =   1
      Top             =   1680
      Width           =   5295
   End
   Begin VB.CommandButton btn_ConsultarClientes 
      Caption         =   "Consultar Clientes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   1215
      TabIndex        =   3
      Top             =   4680
      Width           =   5295
   End
   Begin VB.Label lbl_Titulo 
      Caption         =   "MENU"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   60
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   1995
      TabIndex        =   0
      Top             =   240
      Width           =   3720
   End
End
Attribute VB_Name = "frm_MenuPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_CadastrarCliente_Click()
    frm_CadastrarCliente.Show
End Sub

Private Sub btn_CadastrarCorretor_Click()
    frm_CadastrarCorretor.Show
End Sub

Private Sub btn_ConsultarClientes_Click()
    frm_ConsultarClientes.Show
End Sub

Private Sub Form_Load()
    Me.Show
    
    Call PreencheConnetionString
End Sub
