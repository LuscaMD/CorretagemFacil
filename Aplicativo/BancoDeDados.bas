Attribute VB_Name = "BancoDeDados"
Public SQL As ADODB.Connection
Public pub_str_ConnectionString As String

Public Sub PreencheConnetionString()
  pub_str_ConnectionString = "Provider=SQLOLEDB;Data Source=DESKTOP-KULQ16T\SQLEXPRESS;Initial Catalog=ViceriSeidor;Integrated Security=SSPI;"
End Sub


Public Sub AbrirConexao()
  On Error GoTo Erro
  
  If SQL Is Nothing Then
    Set SQL = New ADODB.Connection
  
    SQL.ConnectionString = pub_str_ConnectionString
    SQL.Open
  End If
  Exit Sub
  
Erro:
    MsgBox "Erro ao abrir conexão: " & Err.Description
End Sub

Public Sub FecharConexao()
  On Error Resume Next
  
  If Not SQL Is Nothing Then
    SQL.Close
    Set SQL = Nothing
  End If
End Sub

Public Sub ExecutarComando(ByVal str_SQL As String)
  On Error GoTo Erro
  
  Call AbrirConexao
  Call IniciarTransacao
  
  SQL.Execute str_SQL
  Call CommitTransacao
  
  Call FecharConexao
  Exit Sub
Erro:
  Call RollbackTransacao
  MsgBox "Erro ao executar o comando SQL: " & Err.Description
End Sub

Public Sub IniciarTransacao()
    On Error GoTo Erro
    
    SQL.BeginTrans
    Exit Sub
Erro:
    MsgBox "Erro ao iniciar a transação: " & Err.Description
End Sub

Public Sub CommitTransacao()
    On Error GoTo Erro
    
    SQL.CommitTrans
    Exit Sub
    
Erro:
    MsgBox "Erro ao fazer commit da transação: " & Err.Description
    Call RollbackTransacao
End Sub

Public Sub RollbackTransacao()
    On Error Resume Next
    SQL.RollbackTrans
End Sub

Public Function RetornarDados(str_SQL As String) As Object
  Dim recordSource As New ADODB.Recordset
  
  On Error GoTo Erro
  
  Call AbrirConexao
  
  recordSource.Open str_SQL, SQL, adOpenStatic, adLockReadOnly
  
  Set RetornarDados = recordSource
  
  Exit Function
Erro:
  Call RollbackTransacao
  MsgBox "Erro ao executar o comando SQL: " & Err.Description
End Function
