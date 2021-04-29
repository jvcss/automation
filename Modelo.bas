Attribute VB_Name = "Modelo"
#If VBA7 Then
    Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
#Else
    Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Public oAutomation As New CUIAutomation
'API
'
'        UIAutomationClient .
'
'
Public AppObj As UIAutomationClient.IUIAutomationElement

Public AppObjGrade As UIAutomationClient.IUIAutomationElement
Public AppObjPainelAcao As UIAutomationClient.IUIAutomationElement

Public elementClean As UIAutomationClient.IUIAutomationElement
Public MyElement1 As UIAutomationClient.IUIAutomationElement
Public MyElement2 As UIAutomationClient.IUIAutomationElement
Public MyElement3 As UIAutomationClient.IUIAutomationElement

Public o_ValuePattern As UIAutomationClient.IUIAutomationValuePattern
Public o_InvokePattern As UIAutomationClient.IUIAutomationInvokePattern
Public o_LegacyAccessiblePattern As UIAutomationClient.IUIAutomationLegacyIAccessiblePattern


Public Enum ConditionLegacyPattern
   CurrentDefaultAction
   CurrentState
   CurrentDescription
   CurrentValue
   CurrentName
End Enum

Public Enum Condition
   eUIA_NamePropertyId
   eUIA_AutomationIdPropertyId
   eUIA_ClassNamePropertyId
   eUIA_LocalizedControlTypePropertyId
End Enum

Public Function Clear(c As UIAutomationClient.IUIAutomationElement)
    If Not c Is Nothing Then
        Set c = Nothing
    End If
End Function

Public Function show(ParamArray Arr() As Variant) As String
    Dim N As Long
    Dim finalStr As String
    For N = LBound(Arr) To UBound(Arr)
        finalStr = finalStr & " " & Arr(N)
    Next N
    Debug.Print finalStr
End Function

Public Function Pause(val As Integer)
        val = val * 1000
        Sleep val
       Debug.Print Now()
End Function

Public Sub test()

    UF = "A" + CStr(ActiveCell.Row)
    Movel = "B" + CStr(ActiveCell.Row)
    OS = "C" + CStr(ActiveCell.Row)
    'show Range(UF).Value, Range(Movel).Value, Range(OS).Value

    Set AppObj = oAutomation.GetRootElement.FindFirst(TreeScope_Children, PropCondition(oAutomation, eUIA_AutomationIdPropertyId, "Form_Perfilacao_Outros"))

    If AppObj Is Nothing Then
        MsgBox "ALERTA CORP MAIL N�O EST� ABERTO NA TELA DE PERFILAR OUTROS PROTOCOLOS - PERSONALIZADO"
    Exit Sub
    End If

    Set AppObj = AppObj.FindFirst(TreeScope_Children, PropCondition(oAutomation, eUIA_AutomationIdPropertyId, "TableLayoutPanel1"))

    Set AppObj = AppObj.FindFirst(TreeScope_Children, PropCondition(oAutomation, eUIA_AutomationIdPropertyId, "GroupBox1"))

    Set AppObj = AppObj.FindFirst(TreeScope_Children, PropCondition(oAutomation, eUIA_AutomationIdPropertyId, "TableLayoutPanel5"))
    

    Set AppObjPainelAcao = AppObj.FindFirst(TreeScope_Children, PropCondition(oAutomation, eUIA_AutomationIdPropertyId, "TableLayoutPanel9"))

    Set AppObjPainelAcao = AppObjPainelAcao.FindFirst(TreeScope_Children, PropCondition(oAutomation, eUIA_AutomationIdPropertyId, "PainelAcaoBO"))
    

    Set AppObjGrade = AppObj.FindFirst(TreeScope_Children, PropCondition(oAutomation, eUIA_AutomationIdPropertyId, "TableLayoutPanel2"))

    Set AppObjGrade = AppObjGrade.FindFirst(TreeScope_Children, PropCondition(oAutomation, eUIA_AutomationIdPropertyId, "TabelaDadosManual"))


    'A��O   ---

    'Clear elementClean

    'root-> AppObj   |   Painel Acao-> AppObjGrade   |   Grade Filtro-> AppObjPainelAcao

    'Call Search(AppObj, eUIA_AutomationIdPropertyId, "EmailTextBox", elementClean)

    'show " >> element FOUND: ", elementClean.CurrentAutomationId

    'Set o_LegacyAccessiblePattern = elementClean.GetCurrentPattern(UIA_LegacyIAccessiblePatternId)

    'o_LegacyAccessiblePattern.SetValue (protocolo)

    '1�

    ' A��O : NOVA SOLICITA��O

    Clear elementClean
    'root-> AppObj   |   Painel Acao-> AppObjGrade   |   Grade Filtro-> AppObjPainelAcao

    Call Search(AppObj, eUIA_AutomationIdPropertyId, "NovaSolicitacaoButton", elementClean)
    
    show ">>>>>>:   ", AppObj.CurrentAutomationId


    Set o_InvokePattern = elementClean.GetCurrentPattern(UIAutomationClient.UIA_InvokePatternId)

    If elementClean.CurrentIsEnabled Then

        o_InvokePattern.Invoke

    End If


      'A��O COMBO BOX  Dados - Servi�o
      '
      Clear elementClean
      Clear MyElement2
      Clear MyElement1
      Clear MyElement3
      'root-> AppObj   |   Painel Acao-> AppObjGrade   |   Grade Filtro-> AppObjPainelAcao
      '
      Call Search(AppObj, eUIA_AutomationIdPropertyId, "EquipePersonalizadoComboBox", elementClean)
      Call Search(elementClean, eUIA_LocalizedControlTypePropertyId, "text", MyElement2)
      Call Search(elementClean, eUIA_NamePropertyId, "Abrir", MyElement1)
      elementClean.SetFocus
      
      Set o_LegacyAccessiblePattern = elementClean.GetCurrentPattern(UIA_LegacyIAccessiblePatternId)
      '
      o_LegacyAccessiblePattern.SetValue (StoragePlan.Range("C4").Value)
      
      
      
      Set o_LegacyAccessiblePattern = MyElement2.GetCurrentPattern(UIA_LegacyIAccessiblePatternId)
      o_LegacyAccessiblePattern.SetValue (StoragePlan.Range("C4").Value)

      MyElement1.SetFocus
      Set o_InvokePattern = MyElement1.GetCurrentPattern( _
      UIAutomationClient.UIA_InvokePatternId)
      o_InvokePattern.Invoke
      
      Clear MyElement2
      ' OPENNED DIALOG ELEMENT
      Call Search(elementClean, eUIA_ClassNamePropertyId, "ComboLBox", MyElement2)
      
      Call Search(MyElement2, eUIA_NamePropertyId, StoragePlan.Range("C4").Value, MyElement3)
      Set o_InvokePattern = Nothing
      Set o_InvokePattern = MyElement3.GetCurrentPattern(UIAutomationClient.UIA_InvokePatternId)
      o_InvokePattern.Invoke
      'show
      'Set o_LegacyAccessiblePattern = MyElement3.GetCurrentPattern(UIA_LegacyIAccessiblePatternId)
      'o_LegacyAccessiblePattern.SetValue (StoragePlan.Range("C4").Value)
      
      
      
      
      
      
      
      
      
      
      
      
      
      'A��O COMBO BOX  TT
      '
      Clear elementClean
      Clear MyElement1
      Clear MyElement2
      Clear MyElement3
      
      Call Search(AppObj, eUIA_AutomationIdPropertyId, StoragePlan.Range("B5").Value, elementClean)
      Call Search(elementClean, eUIA_NamePropertyId, "Abrir", MyElement1)
      Call Search(elementClean, eUIA_LocalizedControlTypePropertyId, "text", MyElement2)
      '
      elementClean.SetFocus
      Set o_LegacyAccessiblePattern = elementClean.GetCurrentPattern(UIA_LegacyIAccessiblePatternId)
      o_LegacyAccessiblePattern.SetValue (StoragePlan.Range("C5").Value)
      '
      Set o_InvokePattern = MyElement1.GetCurrentPattern(UIAutomationClient.UIA_InvokePatternId)
      o_InvokePattern.Invoke
      '
      
      elementClean.SetFocus
      MyElement2.SetFocus
      ' OPENNED DIALOG ELEMENT
      Call Search(elementClean, eUIA_ClassNamePropertyId, "ComboLBox", _
      MyElement2)
      
      Call Search(MyElement2, eUIA_NamePropertyId, StoragePlan.Range("C5").Value, MyElement3)
      Set o_InvokePattern = Nothing
      Set o_InvokePattern = MyElement3.GetCurrentPattern(UIAutomationClient.UIA_InvokePatternId)
      o_InvokePattern.Invoke
      
      Set o_LegacyAccessiblePattern = MyElement3.GetCurrentPattern(UIA_LegacyIAccessiblePatternId)
      o_LegacyAccessiblePattern.SetValue (StoragePlan.Range("C5").Value)
      
      
      
      
      
      
      
      'A��O COMBO BOX Concl�do
      '
      Clear elementClean
      Clear MyElement1
      Clear MyElement2
      Clear MyElement3
      Set o_InvokePattern = Nothing
      
      Call Search(AppObj, eUIA_AutomationIdPropertyId, _
      "StatusComboBox", elementClean)
      Call Search(elementClean, eUIA_NamePropertyId, _
      "Abrir", MyElement1)

      'StatusComboBox
      elementClean.SetFocus
      Set o_LegacyAccessiblePattern = elementClean.GetCurrentPattern( _
      UIA_LegacyIAccessiblePatternId)
      o_LegacyAccessiblePattern.SetValue (StoragePlan.Range("C6").Value)
      
      'AbrirButtton
      
      Set o_InvokePattern = MyElement1.GetCurrentPattern( _
      UIAutomationClient.UIA_InvokePatternId)
      o_InvokePattern.Invoke
      'show elementClean.CurrentAutomationId
      
      ' OPENNED DIALOG ELEMENT
      Call Search(elementClean, eUIA_ClassNamePropertyId, "ComboLBox", _
      MyElement2)
      
      Call Search(MyElement2, eUIA_NamePropertyId, StoragePlan.Range("C6").Value, _
      MyElement3)
      
      Set o_InvokePattern = Nothing
      Set o_InvokePattern = MyElement3.GetCurrentPattern( _
      UIAutomationClient.UIA_InvokePatternId)
      o_InvokePattern.Invoke
      'show MyElement3.CurrentLocalizedControlType
      Set o_LegacyAccessiblePattern = MyElement3.GetCurrentPattern( _
      UIA_LegacyIAccessiblePatternId)
      o_LegacyAccessiblePattern.SetValue (StoragePlan.Range("C6").Value)


      
      




              'A��O COMBO BOX Solicita��o
      
      Clear elementClean
      Clear MyElement1
      Clear MyElement2
      Clear MyElement3
      Set o_InvokePattern = Nothing
      
      Call Search(AppObj, eUIA_AutomationIdPropertyId, StoragePlan.Range("B7").Value, elementClean)
      Call Search(elementClean, eUIA_NamePropertyId, "Abrir", MyElement1)
      'show "ENCONTRADO= ABRIR", MyElement1.CurrentName
      elementClean.SetFocus
      Set o_LegacyAccessiblePattern = elementClean.GetCurrentPattern(UIA_LegacyIAccessiblePatternId)
      o_LegacyAccessiblePattern.SetValue (StoragePlan.Range("C7").Value)
      elementClean.SetFocus
      
      Set o_InvokePattern = MyElement1.GetCurrentPattern(UIAutomationClient.UIA_InvokePatternId)
      o_InvokePattern.Invoke
      
      Call Search(elementClean, eUIA_ClassNamePropertyId, "ComboLBox", MyElement2)
      Call Search(MyElement2, eUIA_NamePropertyId, StoragePlan.Range("C7").Value, MyElement3)
      
      Set o_InvokePattern = Nothing
      Set o_InvokePattern = MyElement3.GetCurrentPattern(UIAutomationClient.UIA_InvokePatternId)
      o_InvokePattern.Invoke
      
      
      Set o_LegacyAccessiblePattern = MyElement3.GetCurrentPattern(UIA_LegacyIAccessiblePatternId)
      o_LegacyAccessiblePattern.SetValue (StoragePlan.Range("C7").Value)
      
      
      'show elementClean.CurrentClassName
      'show MyElement1.CurrentName
      'show MyElement2.CurrentClassName
      'show MyElement3.CurrentClassName


      
      
      
      
      
      
      

      
      
      
      
      
      
      
      'A��O COMBO BOX UF GRADE
      
      Clear elementClean
      
      'root-> AppObj   |   Painel Acao-> AppObjGrade   |   Grade Filtro-> AppObjPainelAcao
      
      Call Search(AppObj, eUIA_NamePropertyId, "UF Linha 0", elementClean)
      
      'show ">> element FOUND: ", elementClean.CurrentName
      
      
      
      
      Set o_LegacyAccessiblePattern = elementClean.GetCurrentPattern(UIA_LegacyIAccessiblePatternId)
      '
      'buscar elemento na ATUAL LINHA NA COLUNA A TEMOS ELEMENTO UF
      '
      o_LegacyAccessiblePattern.SetValue (Range(UF).Value)
      
                      '
              'A��O COMBO BOX Regiao GRADE
              '
              '
      'SE UF = RS,SC,PR,MS,TO,GO,MT,RO,AC = R2
      'SEN�O = AM,RR,AP,PA,MA,CE,RN,PB,PE,AL,SE,BA,MG
              'ES,SP,PI,RJ
      Select Case Range(UF).Value
          Case "RS", "SC", "PR", "MS", "TO", "GO", "MT", "RO", "AC"

              Clear elementClean
              '
              'root-> AppObj   |   Painel Acao-> AppObjGrade   |   Grade Filtro-> AppObjPainelAcao
              '
              Call Search(AppObj, eUIA_NamePropertyId, "Regiao Linha 0", elementClean)
              
              'show ">> element FOUND: ", elementClean.CurrentName
              '
              '
              '
              '
              Set o_LegacyAccessiblePattern = elementClean.GetCurrentPattern(UIA_LegacyIAccessiblePatternId)
              '
              'ATUAL LINHA NA COLUNA A TEMOS ELEMENTO REGI�O
              
              '
              o_LegacyAccessiblePattern.SetValue ("R2")
          Case "AM", "RR", "AP", "PA", "MA", "CE", "RN", "PB", "PE", "AL" _
           , "SE", "BA", "MG", "ES", "SP", "PI", "RJ"
              Clear elementClean
              '
              'root-> AppObj   |   Painel Acao-> AppObjGrade   |   Grade Filtro-> AppObjPainelAcao
              '
              Call Search(AppObj, eUIA_NamePropertyId, "Regiao Linha 0", elementClean)
              
              'show ">> element FOUND: ", elementClean.CurrentName
              '
              '
              '
              '
              Set o_LegacyAccessiblePattern = elementClean.GetCurrentPattern(UIA_LegacyIAccessiblePatternId)
              '
              'ATUAL LINHA NA COLUNA A TEMOS ELEMENTO REGI�O
              
              '
              o_LegacyAccessiblePattern.SetValue ("R1")
      End Select
      
      
      
      
      
      'A��O COMBO BOX CNPJ GRADE
      '
      Clear elementClean
      '
      'root-> AppObj   |   Painel Acao-> AppObjGrade   |   Grade Filtro-> AppObjPainelAcao
      '
      Call Search(AppObj, eUIA_NamePropertyId, "CNPJ Linha 0", elementClean)
      
      'show ">> element FOUND: ", elementClean.CurrentName
      '
      '
      '
      '
      Set o_LegacyAccessiblePattern = elementClean.GetCurrentPattern(UIA_LegacyIAccessiblePatternId)
      '
      'ATUAL LINHA NA COLUNA A TEMOS ELEMENTO REGI�O
      o_LegacyAccessiblePattern.SetValue (StoragePlan.Range("C10").Value)
      
      
      
      'A��O COMBO BOX PRODUTO GRADE
              '
      Clear elementClean
      '
      'root-> AppObj   |   Painel Acao-> AppObjGrade   |   Grade Filtro-> AppObjPainelAcao
      '
      Call Search(AppObj, eUIA_NamePropertyId, "Produto Linha 0", elementClean)
      
      'show ">> element FOUND: ", elementClean.CurrentName
      '
      '
      '
      '
      Set o_LegacyAccessiblePattern = elementClean.GetCurrentPattern(UIA_LegacyIAccessiblePatternId)
      '
      'ATUAL LINHA NA COLUNA A TEMOS ELEMENTO REGI�O
      o_LegacyAccessiblePattern.SetValue (StoragePlan.Range("C11").Value)
      
      
      
      
      
      
      '''''''''''''''''''''A��O COMBO BOX TIPO SOLICITA��O GRADE
      
      '
      
      Clear elementClean
      '
      'root-> AppObj   |   Painel Acao-> AppObjGrade   |   Grade Filtro-> AppObjPainelAcao
      '
      Call Search(AppObj, eUIA_NamePropertyId, "Tipo Solicita��o Linha 0", elementClean)
      
      'show ">> element FOUND: ", elementClean.CurrentName
      '
      '
      '
      '
      Set o_LegacyAccessiblePattern = elementClean.GetCurrentPattern(UIA_LegacyIAccessiblePatternId)
      '
      'ATUAL LINHA NA COLUNA A TEMOS ELEMENTO REGI�O
      o_LegacyAccessiblePattern.SetValue (StoragePlan.Range("C12").Value)
      elementClean.SetFocus
      
    '  Set o_LegacyAccessiblePattern = elementClean.GetCurrentPattern(UIA_LegacyIAccessiblePatternId)
     ' show o_LegacyAccessiblePattern.CurrentValue
      
      
      
      
      
      
      '''''''''''''''''''''A��O COMBO BOX MOTIVO GRADE

      
      Clear elementClean
      Clear MyElement1
      Clear MyElement2
      Clear MyElement3
      Set o_InvokePattern = Nothing
      Set o_LegacyAccessiblePattern = Nothing
      '
      Call Search(AppObj, eUIA_NamePropertyId, StoragePlan.Range("B13").Value, elementClean)
      elementClean.SetFocus
      '
      Set o_LegacyAccessiblePattern = elementClean.GetCurrentPattern(UIA_LegacyIAccessiblePatternId)
      '
      o_LegacyAccessiblePattern.SetValue (StoragePlan.Range("C13").Value)
      '
      
      '
      'Set o_InvokePattern = elementClean.GetCurrentPattern(UIAutomationClient.UIA_InvokePatternId)
      'elementClean.SetFocus
      'o_LegacyAccessiblePattern.DoDefaultAction
      'o_InvokePattern.Invoke
      
      '  ***DIALOG ELEMENT ACTION***
      '
      'Call Search(elementClean, eUIA_ClassNamePropertyId, "ComboLBox", MyElement2)
      '
      'Call Search(MyElement2, eUIA_NamePropertyId, StoragePlan.Range("C13").Value, MyElement3)
      '
      'Set o_InvokePattern = Nothing
      'Set o_InvokePattern = MyElement3.GetCurrentPattern(UIAutomationClient.UIA_InvokePatternId)
      'o_InvokePattern.Invoke
      '
      'Set o_LegacyAccessiblePattern = MyElement3.GetCurrentPattern(UIA_LegacyIAccessiblePatternId)
      'o_LegacyAccessiblePattern.SetValue (StoragePlan.Range("C6").Value)
      
      
      
      
      '''''''''''''''''''''A��O COMBO BOX TIPO SOLICITA��O GRADE  2�vez
      
      '
      
      Clear elementClean
      '
      'root-> AppObj   |   Painel Acao-> AppObjGrade   |   Grade Filtro-> AppObjPainelAcao
      '
      Call Search(AppObj, eUIA_NamePropertyId, "Tipo Solicita��o Linha 0", elementClean)
      
      'show ">> element FOUND: ", elementClean.CurrentName
      '
      '
      '
      '
      Set o_LegacyAccessiblePattern = elementClean.GetCurrentPattern(UIA_LegacyIAccessiblePatternId)
      '
      'ATUAL LINHA NA COLUNA A TEMOS ELEMENTO REGI�O
      o_LegacyAccessiblePattern.SetValue (StoragePlan.Range("C12").Value)
      elementClean.SetFocus
      
    '  Set o_LegacyAccessiblePattern = elementClean.GetCurrentPattern(UIA_LegacyIAccessiblePatternId)
     ' show o_LegacyAccessiblePattern.CurrentValue
      
      
      
      
      
      
      
      
      
      '''''''''''''''''''''''''''''''A��O COMBO BOX OS GERADA / TT GRADE
      '
      Clear elementClean
      '
      'root-> AppObj   |   Painel Acao-> AppObjGrade   |   Grade Filtro-> AppObjPainelAcao
      '
      Call Search(AppObj, eUIA_NamePropertyId, StoragePlan.Range("B14").Value, elementClean)
      
      'show ">> element FOUND: ", elementClean.CurrentName
      '
      '
      '
      '
      Set o_LegacyAccessiblePattern = elementClean.GetCurrentPattern(UIA_LegacyIAccessiblePatternId)
      '
      'ATUAL LINHA NA COLUNA A TEMOS ELEMENTO REGI�O
      o_LegacyAccessiblePattern.SetValue (Range(OS).Value)
      
      
      
      '''''''''''''''''''''''''''A��O COMBO BOX QTD. GRADE
      Clear elementClean
      '
      'root-> AppObj   |   Painel Acao-> AppObjGrade   |   Grade Filtro-> AppObjPainelAcao
      '
      Call Search(AppObj, eUIA_NamePropertyId, StoragePlan.Range("B15").Value, elementClean)
      
      'show ">> element FOUND: ", elementClean.CurrentName
      '
      '
      '
      '
      Set o_LegacyAccessiblePattern = elementClean.GetCurrentPattern(UIA_LegacyIAccessiblePatternId)
      '
      'ATUAL LINHA NA COLUNA A TEMOS ELEMENTO***************************
      o_LegacyAccessiblePattern.SetValue ("")
      elementClean.SetFocus
      
      
      
      
      '''''''''''''''''''''''''''''''''A��O COMBO BOX RESPONSAVEL GRADE

      Clear elementClean
      '
      'root-> AppObj   |   Painel Acao-> AppObjGrade   |   Grade Filtro-> AppObjPainelAcao
      '
      Call Search(AppObj, eUIA_NamePropertyId, StoragePlan.Range("B16").Value, elementClean)
      
      'show ">> element FOUND: ", elementClean.CurrentName
      '
      '
      '
      '
      Set o_LegacyAccessiblePattern = elementClean.GetCurrentPattern(UIA_LegacyIAccessiblePatternId)
      '
      'ATUAL LINHA NA COLUNA A TEMOS ELEMENTO REGI�O
      o_LegacyAccessiblePattern.SetValue (StoragePlan.Range("C16").Value)
      
      
      
      
      
      
      
      
      
      '2�
      '
      ' A��O: PROTOCOLO ORIGEM
      '
      Clear elementClean
      '
      'root-> AppObj   |   Painel Acao-> AppObjGrade   |   Grade Filtro-> AppObjPainelAcao
      '
      Call Search(AppObj, eUIA_AutomationIdPropertyId, StoragePlan.Range("B1").Value, elementClean)
      
      'show ">>>>>>TEXTO Protocolo:    ", elementClean.CurrentAutomationId
      '
      '
      'BOTAO TEXTO EMAIL CLIENTE
      '
      Set o_LegacyAccessiblePattern = elementClean.GetCurrentPattern(UIA_LegacyIAccessiblePatternId)
      
      o_LegacyAccessiblePattern.SetValue (StoragePlan.Range("C1").Value)
      
      
      
      
      
      
      '3�

      'A��O: EMAIL CLIENTE

      Clear elementClean

      Call Search(AppObj, eUIA_AutomationIdPropertyId, StoragePlan.Range("B2").Value, elementClean)

      'BOTAO EMAIL CLIENTE

      Set o_LegacyAccessiblePattern = elementClean.GetCurrentPattern(UIA_LegacyIAccessiblePatternId)
      '
      o_LegacyAccessiblePattern.SetValue (StoragePlan.Range("C2").Value)


      '4� comentario

      'A��O: COMENTARIO

      Clear elementClean

      Call Search(AppObjPainelAcao, eUIA_AutomationIdPropertyId, StoragePlan.Range("B3").Value, elementClean)

      'TEXTO COMENTARIO

      Set o_LegacyAccessiblePattern = elementClean.GetCurrentPattern(UIA_LegacyIAccessiblePatternId)

      o_LegacyAccessiblePattern.SetValue (StoragePlan.Range("C3").Value)

      'A��O ATIVAR GRADE
      Clear elementClean
      Clear MyElement1
      Clear MyElement2
      Clear MyElement3
      Set o_InvokePattern = Nothing

      'root-> AppObj   |   Painel Acao-> AppObjGrade   |   Grade Filtro-> AppObjPainelAcao

      Call Search(AppObj, eUIA_NamePropertyId, StoragePlan.Range("B15").Value, elementClean)

      Set o_LegacyAccessiblePattern = elementClean.GetCurrentPattern(UIA_LegacyIAccessiblePatternId)
      elementClean.SetFocus

      o_LegacyAccessiblePattern.SetValue ("")

      Set o_InvokePattern = elementClean.GetCurrentPattern(UIAutomationClient.UIA_InvokePatternId)
      o_InvokePattern.Invoke

      Call Search(AppObjGrade, eUIA_NamePropertyId, "Controle de Edi��o", MyElement1)

      MyElement1.SetFocus
      Pause 1
      SendKeys ("1")

End Sub


    




'

Public Function PropCondition(UIAutomation As CUIAutomation, Prop As Condition, Requirement As String) As UIAutomationClient.IUIAutomationCondition
        Select Case Prop
            Case 0
                Set PropCondition = UIAutomation.CreatePropertyCondition(UIAutomationClient.UIA_NamePropertyId, Requirement)
            Case 1
                Set PropCondition = UIAutomation.CreatePropertyCondition(UIAutomationClient.UIA_AutomationIdPropertyId, Requirement)
            Case 2
                Set PropCondition = UIAutomation.CreatePropertyCondition(UIAutomationClient.UIA_ClassNamePropertyId, Requirement)
            Case 3
                Set PropCondition = UIAutomation.CreatePropertyCondition(UIAutomationClient.UIA_LocalizedControlTypePropertyId, Requirement)
        End Select
End Function


Public Function Search(ByVal obj As UIAutomationClient.IUIAutomationElement, typed As Condition, strFinalElemSearch As String, _
ByRef elem As UIAutomationClient.IUIAutomationElement) As UIAutomationClient.IUIAutomationElement
        On Error Resume Next
        Dim ended As Boolean
        ended = False
        Dim walker As UIAutomationClient.IUIAutomationTreeWalker
        Dim element1 As UIAutomationClient.IUIAutomationElementArray
        Dim element2 As UIAutomationClient.IUIAutomationElement
        
        Set walker = oAutomation.ControlViewWalker
        Dim condition1 As UIAutomationClient.IUIAutomationCondition
        Set condition1 = oAutomation.CreateTrueCondition
        Set element1 = obj.FindAll(TreeScope_Children, condition1)
        
        'wait execution to do other tasks
        DoEvents
        If element1.Length <> 0 Then
                Set element2 = obj.FindFirst(TreeScope_Children, condition1)
        End If
        Do While Not element2 Is Nothing
            'verificar como tratar elemento atual
            Dim oPattern As UIAutomationClient.IUIAutomationLegacyIAccessiblePattern
            Set oPattern = element2.GetCurrentPattern(UIA_LegacyIAccessiblePatternId)
            show element2.CurrentAutomationId
            'se pattern � o certo, temos mais um elemento pra adicionar aqui
            'try catch bloc to
            Select Case typed
                Case eUIA_NamePropertyId
                    If StrComp(element2.CurrentName, strFinalElemSearch) = 0 Then
                        'show "result CurrentLocalizedControlType $$ ", strFinalElemSearch
                        ended = True
                        Set elem = element2
                        Set Search = elem
                        Exit Function
                    End If
                    
                Case eUIA_AutomationIdPropertyId
                    If StrComp(element2.CurrentAutomationId, strFinalElemSearch) = 0 Then
                        show "result CurrentLocalizedControlType $$ ", strFinalElemSearch
                        ended = True
                        Set elem = element2
                        Set Search = elem
                        Exit Function
                    End If
                    
                Case eUIA_ClassNamePropertyId
                    If StrComp(element2.CurrentClassName, strFinalElemSearch) = 0 Then
                        'show "result CurrentLocalizedControlType $$ ", strFinalElemSearch
                        ended = True
                        Set elem = element2
                        Set Search = elem
                        Exit Function
                    End If
                    
                Case eUIA_LocalizedControlTypePropertyId
                    If StrComp(element2.CurrentLocalizedControlType, strFinalElemSearch) = 0 Then
                        'show "result CurrentLocalizedControlType $$ ", strFinalElemSearch
                        ended = True
                        Set elem = element2
                        Set Search = elem
                        Exit Function
                    End If
            End Select
            
            'n�o � o elemento que pedi pq saiu do SELECT SEM achar o     elem     ENT�O
            'next child  pra frente busca elemento interno
            Search element2, typed, strFinalElemSearch, elem

            show "n�o � o elemento que pedi : ", element2.CurrentAutomationId, element2.CurrentAutomationId
            
            If Not elem Is Nothing Then Exit Function
            
            Set element2 = walker.GetNextSiblingElement(element2)
            'andando pra traz
 
        Loop
        If ended Then Exit Function
End Function






Public Function SearchPattern(ByVal obj As UIAutomationClient.IUIAutomationElement, _
    typed As ConditionLegacyPattern, strFinalElemSearch As String, _
    ByRef elem As UIAutomationClient.IUIAutomationLegacyIAccessiblePattern) As UIAutomationClient.IUIAutomationLegacyIAccessiblePattern
        On Error Resume Next
        Dim ended As Boolean
        ended = False
        Dim walker As UIAutomationClient.IUIAutomationTreeWalker
        Dim element1 As UIAutomationClient.IUIAutomationElementArray
        Dim element2 As UIAutomationClient.IUIAutomationElement
        Dim oPattern As UIAutomationClient.IUIAutomationLegacyIAccessiblePattern
        
        Set walker = oAutomation.ControlViewWalker
        Dim condition1 As UIAutomationClient.IUIAutomationCondition
        Set condition1 = oAutomation.CreateTrueCondition
        Set element1 = obj.FindAll(TreeScope_Descendants, condition1)
        
        'aguarda execu��o para q o pc possa fazer outras tarefas
        DoEvents
        If element1.Length <> 0 Then
                Set element2 = obj.FindFirst(TreeScope_Children, condition1)
        End If
        Do While Not element2 Is Nothing
            'verificar como tratar elemento atual
            
            Set oPattern = element2.GetCurrentPattern(UIA_LegacyIAccessiblePatternId)
            'show "CLASS:", element2.CurrentClassName, _
            " autoID:", element2.CurrentAutomationId, _
            " LocalizedControlType:", element2.CurrentLocalizedControlType
            
        'o_LegacyAccessiblePattern.
        'CurrentDefaultAction
        'CurrentState
        'CurrentDescription
        'CurrentValue
        'CurrentName
        
            'se pattern � o certo, temos mais um elemento pra adicionar aqui
            'try catch bloc to
            Select Case typed
                Case CurrentValue
                    If StrComp(oPattern.CurrentValue, strFinalElemSearch) = 0 Then
                        'show "result CurrentLocalizedControlType $$ ", strFinalElemSearch
                        ended = True
                        Set elem = oPattern
                        Set SearchPattern = elem
                        Exit Do
                    End If
                    
                Case CurrentDescription
                    If StrComp(oPattern.CurrentDescription, strFinalElemSearch) = 0 Then
                        'show "result CurrentLocalizedControlType $$ ", strFinalElemSearch
                        ended = True
                        Set elem = oPattern
                        'Set SearchPattern = elem
                        Exit Do
                    End If
                    
                Case CurrentState
                    If StrComp(oPattern.CurrentChildId, strFinalElemSearch) = 0 Then
                        'show "result CurrentLocalizedControlType $$ ", strFinalElemSearch
                        ended = True
                        Set elem = oPattern
                        'Set Search = elem
                        Exit Do
                    End If
                    
                Case CurrentDefaultAction
                    If StrComp(oPattern.CurrentDefaultAction, strFinalElemSearch) = 0 Then
                        'show "result CurrentLocalizedControlType $$ ", strFinalElemSearch
                        ended = True
                        Set elem = oPattern
                        'Set Search = elem
                        Exit Do
                    End If
                    
                Case CurrentName
                    If StrComp(oPattern.CurrentName, strFinalElemSearch) = 0 Then
                        'show " autoID:", element2.CurrentAutomationId
                        ended = True
                        Set elem = oPattern
                        Set SearchPattern = elem
                        Exit Do
                    End If
            End Select
             
            'n�o � o elemento que pedi pq saiu do SELECT SEM achar o     elem     ENT�O
            'andando pra frente busca elemento interno
            SearchPattern element2, typed, strFinalElemSearch, elem

       '     show "n�o � o elemento que pedi : ", element2.CurrentControlType, elem.CurrentName
            'case not empty we found the SPECIFIC PATTERN
            If Not elem Is Nothing Then Exit Function
            
            Set element2 = walker.GetNextSiblingElement(element2)
            'andando pra traz
 
        Loop
        If ended Then Exit Function
End Function












Public Function Get_RootElement(strWindowName As String, cond As Condition) As UIAutomationClient.IUIAutomationElement
        Dim walker As UIAutomationClient.IUIAutomationTreeWalker
        Dim element As UIAutomationClient.IUIAutomationElement

        Set walker = oAutomation.ControlViewWalker
        Set element = walker.GetFirstChildElement(oAutomation.GetRootElement)

        Do While Not element Is Nothing
            'show ">>>>>>element CurrentName: ", element.CurrentName
            'StrComp return 0 � if strings are equal;
            If StrComp(element.CurrentAutomationId, strWindowName, vbBinaryCompare) = 0 Then
                Set Get_RootElement = element
                Exit Function
            End If
            If StrComp(element.CurrentName, strWindowName, vbBinaryCompare) = 0 Then
                Set Get_RootElement = element
                Exit Function
            End If
            If StrComp(element.CurrentClassName, strWindowName, vbBinaryCompare) = 0 Then
                Set Get_RootElement = element
                Exit Function
            End If
            If StrComp(element.CurrentControlType, strWindowName, vbBinaryCompare) = 0 Then
                Set Get_RootElement = element
                Exit Function
            End If

            Set element = walker.GetNextSiblingElement(element)
        Loop
End Function
 
