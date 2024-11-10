Attribute VB_Name = "FuncoesGerais"
Public Function Nz(ByVal valorOriginal As Variant, Optional valorSeNulo As Variant = "") As Variant
  If IsNull(valorOriginal) Then
    Nz = valorSeNulo
  Else
    Nz = valorOriginal
  End If
End Function

Public Sub VerificaMascaraCPF(Form As Object, ByRef KeyAscii As Integer)
  If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
    KeyAscii = 0
    Exit Sub
  End If

  If (Len(Form.txt_CPF.Text) = 3 Or Len(Form.txt_CPF.Text) = 7) And KeyAscii <> 8 Then
    Form.txt_CPF.Text = Form.txt_CPF.Text & "."
    Form.txt_CPF.SelStart = Len(Form.txt_CPF.Text)
  ElseIf Len(Form.txt_CPF.Text) = 11 And KeyAscii <> 8 Then
    Form.txt_CPF.Text = Form.txt_CPF.Text & "-"
    Form.txt_CPF.SelStart = Len(Form.txt_CPF.Text)
  End If
End Sub

Public Sub VerificaDigitoNumerico(Form As Object, ByRef KeyAscii As Integer)
  If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
    KeyAscii = 0
    Exit Sub
  End If
End Sub

Public Function VerificaCPF(str_CPF As String) As Boolean
  Dim recordSourceCPF As Object
    
  Set recordSourceCPF = RetornarDados("SELECT * FROM Cadastros WHERE str_CPF = '" & str_CPF & "'")

  If recordSourceCPF.EOF Then
    VerificaCPF = True
  Else
    VerificaCPF = False
  End If
  
  recordSourceCPF.Close
  Set recordSourceCPF = Nothing
End Function

Public Function VerificaCodCorretor(int_CodCorretor As String) As Boolean
  Dim recordSourceCodCorretor As Object
    
  Set recordSourceCodCorretor = RetornarDados("SELECT * FROM Cadastros WHERE int_CodCorretor = " & int_CodCorretor)

  If recordSourceCodCorretor.EOF Then
    VerificaCodCorretor = True
  Else
    VerificaCodCorretor = False
  End If
  
  recordSourceCodCorretor.Close
  Set recordSourceCodCorretor = Nothing
End Function
