_________________________________________________________________________________________________________________________________________________________________________________
Separado por tópico estão os blocos de código mais usados para se criar relação de arquitetura de software.

ERROR HANDLING

  
    'podemos declarar uma só subrotina que executa várias funções lateriais
    'isso acelera o processo de execução do código conjuntamente com a declaração das variáveis
    'em tipos básicos ao invés de objetos específicos
    'por exemplo usar um Variant ao invés de Object
    'permite que erros não sejam alertados, com isso o código fica mais "silecioso"
    'ter esse logger e salvar num arquivo ajuda na criação de uma aplicação consistente em VBA.
    'é fácil abstrair essa manipulação da linguagem para outras linguagens de programação.
    
      Public Sub ErrorFilter()
      On Error GoTo CapturaErro
      If Err.Number = 0 Then
          menu
      End If
      Err.Clear
      Exit Sub
      CapturaErro:
      Select Case Err.Number
          Case Else
              Debug.Print _
              "erro: " & CStr(Err.Number & " " & Err.Description)
      End Select
      Resume Next
      End Sub
      
      Function menu()
     
      Dim j As Integer
       
      Dim z As Object 'change to Variant to fix some literal abstraction parse.

      j = z

      For i = 0 To 10
          Debug.Print "ok"
      Next
      End Function




_________________________________________________________________________________________________________________________________________________________________________________
