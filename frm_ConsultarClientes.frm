VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frm_ConsultarClientes 
   Appearance      =   0  'Flat
   BackColor       =   &H80000004&
   Caption         =   "Consultar Clientes"
   ClientHeight    =   5775
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   9870
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Adodc 
      Height          =   615
      Left            =   120
      Top             =   5040
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1085
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
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
      Height          =   360
      Left            =   5520
      TabIndex        =   14
      Top             =   1800
      Width           =   2055
   End
   Begin VB.CommandButton btn_Filtrar 
      Caption         =   "Filtrar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   7680
      TabIndex        =   13
      Top             =   1800
      Width           =   2055
   End
   Begin VB.TextBox txt_CodCorretor 
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
      Left            =   7320
      MaxLength       =   4
      TabIndex        =   11
      Top             =   720
      Width           =   2415
   End
   Begin VB.TextBox txt_NomeCorretor 
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
      Left            =   2280
      MaxLength       =   50
      TabIndex        =   9
      Top             =   720
      Width           =   2415
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
      Left            =   8640
      TabIndex        =   4
      Top             =   1200
      Width           =   1080
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
      Height          =   420
      Left            =   7320
      MaxLength       =   14
      TabIndex        =   3
      Top             =   240
      Width           =   2415
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
      Left            =   2280
      MaxLength       =   50
      TabIndex        =   2
      Top             =   240
      Width           =   2415
   End
   Begin VB.ComboBox cmb_UF 
      Height          =   315
      ItemData        =   "frm_ConsultarClientes.frx":0000
      Left            =   6600
      List            =   "frm_ConsultarClientes.frx":0002
      TabIndex        =   1
      Top             =   1200
      Width           =   1455
   End
   Begin VB.ComboBox cmb_Cidades 
      Height          =   315
      ItemData        =   "frm_ConsultarClientes.frx":0004
      Left            =   2280
      List            =   "frm_ConsultarClientes.frx":0006
      TabIndex        =   0
      Top             =   1200
      Width           =   2415
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "frm_ConsultarClientes.frx":0008
      Height          =   3375
      Left            =   120
      TabIndex        =   15
      Top             =   2280
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   5953
      _Version        =   393216
      AllowUpdate     =   0   'False
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
   Begin VB.Label lbl_CodCorretor 
      Alignment       =   1  'Right Justify
      Caption         =   "Codigo Corretor:"
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
      Left            =   4800
      TabIndex        =   12
      Top             =   720
      Width           =   2415
   End
   Begin VB.Label lbl_NomeCorretor 
      Alignment       =   1  'Right Justify
      Caption         =   "Nome Corretor:"
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
      TabIndex        =   10
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label lbl_CpfCliente 
      Alignment       =   1  'Right Justify
      Caption         =   "CPF Cliente:"
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
      Left            =   5520
      TabIndex        =   8
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label lbl_Nome 
      Alignment       =   1  'Right Justify
      Caption         =   "Nome Cliente:"
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
      TabIndex        =   7
      Top             =   240
      Width           =   1815
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
      Left            =   5400
      TabIndex        =   6
      Top             =   1200
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
      Left            =   1080
      TabIndex        =   5
      Top             =   1200
      Width           =   1095
   End
End
Attribute VB_Name = "frm_ConsultarClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btn_Filtrar_Click()
  Dim str_SQL As String
  Dim str_SQLWhere As String
  Dim recordSource As Object
  
  Me.Adodc.ConnectionString = pub_str_ConnectionString
  
  str_SQL = "SELECT " & _
              "'Id' = a.pk_int_Cadastro, " & _
              "'Nome Cliente' = a.str_Nome, " & _
              "'CPF' = a.str_CPF, " & _
              "'Ativo' = CASE WHEN a.bit_Ativo = 1 THEN 'Sim' ELSE 'Nao' END, " & _
              "'Nome Corretor' = a.str_Nome, " & _
              "'Cod. Corretor' = a.int_CodCorretor, " & _
              "'UF' = c.str_UF, " & _
              "'Cidade' = d.str_NomeCidade, " & _
              "'X' = 'X' " & _
            "FROM " & _
              "Cadastros a WITH(NOLOCK) " & _
                "INNER JOIN " & _
              "Cadastros b WITH(NOLOCK) " & _
                  "ON a.int_CodCorretor = b.int_CodCorretor " & _
                  "AND b.bit_Corretor = 1 " & _
                "INNER JOIN " & _
              "Estados c WITH(NOLOCK) " & _
                  "ON a.fk_int_IdUF = c.pk_int_IdUF " & _
                "INNER JOIN " & _
              "Cidades d WITH(NOLOCK) " & _
                  "ON c.pk_int_IdUF = d.fk_int_IdUF " & _
                  "AND a.fk_int_IdCidade = d.pk_int_IdCidade "
  
  If Nz(Me.txt_Nome, "") <> "" Then
    If Nz(str_SQLWhere, "") = "" Then
      str_SQLWhere = "WHERE a.str_Nome = '" & Me.txt_Nome & "' "
    Else
      str_SQLWhere = str_SQLWhere & "AND a.str_Nome = '" & Me.txt_Nome & "' "
    End If
  End If
  
  If Nz(Me.txt_CPF, "") <> "" Then
    If Nz(str_SQLWhere, "") = "" Then
      str_SQLWhere = "WHERE a.str_CPF = '" & Me.txt_CPF & "' "
    Else
      str_SQLWhere = str_SQLWhere & "AND a.str_CPF = '" & Me.txt_CPF & "' "
    End If
  End If
  
  If Nz(Me.txt_NomeCorretor, "") <> "" Then
    If Nz(str_SQLWhere, "") = "" Then
      str_SQLWhere = "WHERE b.str_Nome = '" & Me.txt_NomeCorretor & "' "
    Else
      str_SQLWhere = str_SQLWhere & "AND b.str_Nome = '" & Me.txt_NomeCorretor & "' "
    End If
  End If
  
  If Nz(Me.txt_CodCorretor, "") <> "" Then
    If Nz(str_SQLWhere, "") = "" Then
      str_SQLWhere = "WHERE b.int_CodCorretor = '" & Me.txt_CodCorretor & "' "
    Else
      str_SQLWhere = str_SQLWhere & "AND b.int_CodCorretor = '" & Me.txt_CodCorretor & "' "
    End If
  End If
  
  If Nz(Me.cmb_Cidades, "") <> "" Then
    If Nz(str_SQLWhere, "") = "" Then
      str_SQLWhere = "WHERE d.str_NomeCidade = '" & Me.cmb_Cidades & "' "
    Else
      str_SQLWhere = str_SQLWhere & "AND d.str_NomeCidade = '" & Me.cmb_Cidades & "' "
    End If
  End If
  
  If Nz(Me.cmb_UF, "") <> "" Then
    If Nz(str_SQLWhere, "") = "" Then
      str_SQLWhere = "WHERE c.str_UF = '" & Me.cmb_UF & "' "
    Else
      str_SQLWhere = str_SQLWhere & "AND c.str_UF = '" & Me.cmb_UF & "' "
    End If
  End If
  
  If Nz(Me.chk_Ativo, 0) <> 0 Then
    If Nz(str_SQLWhere, "") = "" Then
      str_SQLWhere = "WHERE a.bit_Ativo = '" & Me.chk_Ativo
    Else
      str_SQLWhere = str_SQLWhere & "AND a.bit_Ativo = " & Me.chk_Ativo
    End If
  End If
  
  Me.Adodc.recordSource = str_SQL & str_SQLWhere
  Me.Adodc.Refresh
End Sub

Private Sub btn_Limpar_Click()
  Me.txt_CPF = ""
  Me.txt_Nome = ""
  Me.txt_NomeCorretor = ""
  Me.cmb_Cidades = ""
  Me.txt_CodCorretor = ""
  Me.cmb_UF = ""
  Me.chk_Ativo.Value = 0
End Sub

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

Private Sub DataGrid2_Click()
  If DataGrid2.Col = DataGrid2.Columns.Count - 1 Then
    
    If MsgBox("Deseja excluir este registro?", vbYesNo + vbQuestion, "Confirmacao") = vbYes Then
        Call ExecutarComando("DELETE FROM Cadastros WHERE pk_int_Cadastro = " & DataGrid2.Columns(0))
        
        'Me.DataGrid2.Refresh
        Me.Adodc.Refresh
      End If
  End If
End Sub

Private Sub Form_Load()
  Dim recordSource As Object
  
  Set recordSource = RetornarDados("SELECT * FROM Estados")
  
  Do While Not recordSource.EOF
    Me.cmb_UF.AddItem recordSource!str_UF
    Me.cmb_UF.ItemData(Me.cmb_UF.NewIndex) = recordSource!pk_int_IdUF
    recordSource.MoveNext
  Loop
End Sub

Private Sub txt_CodCorretor_KeyPress(KeyAscii As Integer)
  Call VerificaDigitoNumerico(Me, KeyAscii)
End Sub

Private Sub txt_CPF_KeyPress(KeyAscii As Integer)
  Call VerificaMascaraCPF(Me, KeyAscii)
End Sub
