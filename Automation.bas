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

Public o_ElementListPattern As UIAutomationClient.IUIAutomationElementArray



Public o_ValuePattern As UIAutomationClient.IUIAutomationValuePattern
Public o_InvokePattern As UIAutomationClient.IUIAutomationInvokePattern
Public o_LegacyAccessiblePattern As UIAutomationClient.IUIAutomationLegacyIAccessiblePattern


Public protocolo As String
Public emailcliente As String
Public cnpjcliente As String
Public comentario As String
                    'o_LegacyAccessiblePattern.
        'CurrentDefaultAction
        'CurrentState
        'CurrentDescription
        'CurrentValue
        'CurrentName
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
        'c = Nothing
       '
End Function

Public Function show(ParamArray Arr() As Variant) As String
    'Dim N As Variant
    Dim N As Long
    Dim finalStr As String
    'do menor valor dentro da array até o maior valor
    For N = LBound(Arr) To UBound(Arr)
        finalStr = finalStr & " " & Arr(N)
    Next N
    Debug.Print finalStr
End Function

Public Function Pause(val As Integer)
        val = val * 1000
        Sleep val
    '    newHour = Hour(Now())
    '   newMinute = Minute(Now())
    '    newSecond = Second(Now()) + val
    '    waitTime = TimeSerial(newHour, newMinute, newSecond)
    '    Application.Wait waitTime
       Debug.Print Now()
End Function


        '
        '
        '
        'TIPO DE VARIAVEL QUE RECEBE CAMPO DE TEXTO
        '
        'Set o_InvokePattern = elementClean.GetCurrentPattern(UIAutomationClient.UIA_InvokePatternId)
        '
        '
        '
        'TIPO DE VARIAVEL QUE RECEBE DADOS LEGACY DE ELEMENTOS
        '
        'Set o_LegacyAccessiblePattern = elementClean.GetCurrentPattern(UIA_LegacyIAccessiblePatternId)
        '
        '
        '
        'TIPO DE VARIAVEL QUE RECEBE BOTÃO
        '
        'Set o_InvokePattern = elementClean.GetCurrentPropertyValue(50000)
Public Sub Test()
        protocolo = "202100016611927"
        emailcliente = "não tem"
        cnpjcliente = "06649157000129"
        comentario = "retirar: Envio de Oi Torpedo Recepção de Sons & Imagens e Bate Papo Recepção de Mensagens da Oi"
     CNPJ = StoragePlan.Range("C12").Value
     UF = "A" + CStr(ActiveCell.Row)
     Movel = "B" + CStr(ActiveCell.Row)
     OS = "C" + CStr(ActiveCell.Row)
     'show Range(UF).Value, Range(Movel).Value, Range(OS).Value
     
     
        'acessa PAINEL AÇÃO BO. PARA FILTRO DE AÇÃO ESPECIFICA
        '
        '
        Set AppObj = oAutomation.GetRootElement.FindFirst( _
        TreeScope_Children, _
        PropCondition(oAutomation, _
        eUIA_AutomationIdPropertyId, "Form_Perfilacao_Outros"))

        If AppObj Is Nothing Then
            MsgBox "ALERTA CORP MAIL NÃO ESTÁ ABERTO NA TELA DE PERFILAR OUTROS PROTOCOLOS - PERSONALIZADO"
            Exit Sub
        End If

        Set AppObj = AppObj.FindFirst( _
        TreeScope_Children, _
        PropCondition(oAutomation, _
        eUIA_AutomationIdPropertyId, "GroupBox3"))

        Set AppObj = AppObj.FindFirst( _
        TreeScope_Children, _
        PropCondition(oAutomation, _
        eUIA_AutomationIdPropertyId, "TableLayoutPanel1"))

        Set AppObj = AppObj.FindFirst( _
        TreeScope_Children, _
        PropCondition(oAutomation, _
        eUIA_AutomationIdPropertyId, "GroupBox1"))
        
        'root
        Set AppObj = AppObj.FindFirst( _
        TreeScope_Children, _
        PropCondition(oAutomation, _
        eUIA_AutomationIdPropertyId, "TableLayoutPanel5"))
        
        
        'PainelAcaoBO
        Set AppObjPainelAcao = AppObj.FindFirst(TreeScope_Children, PropCondition(oAutomation, _
        eUIA_AutomationIdPropertyId, "TableLayoutPanel9"))

        Set AppObjPainelAcao = AppObjPainelAcao.FindFirst(TreeScope_Children, PropCondition(oAutomation, _
        eUIA_AutomationIdPropertyId, "PainelAcaoBO"))
        
        'GRADE
        Set AppObjGrade = AppObj.FindFirst(TreeScope_Children, PropCondition(oAutomation, eUIA_AutomationIdPropertyId, "TableLayoutPanel2"))

        Set AppObjGrade = AppObjGrade.FindFirst(TreeScope_Children, PropCondition(oAutomation, _
        eUIA_AutomationIdPropertyId, "TabelaDadosManual"))

        
        '
        '
        '
        '
        '
        
        '
        '
        '
        '
        '
        '
        'AÇÃO   ---
        '
        'Clear elementClean
        '
        'root-> AppObj   |   Painel Acao-> AppObjGrade   |   Grade Filtro-> AppObjPainelAcao
        '
        'Call Search(AppObj, eUIA_AutomationIdPropertyId, "EmailTextBox", elementClean)
        '
        'show " >> element FOUND: ", elementClean.CurrentAutomationId
        '
        '
        'Set o_LegacyAccessiblePattern = elementClean.GetCurrentPattern(UIA_LegacyIAccessiblePatternId)
        '
        'o_LegacyAccessiblePattern.SetValue (protocolo)
        '
        '
        '
        '
        '
        '
        '
        '
        '
        '
        
        '1º
        '
        ' AÇÃO : NOVA SOLICITAÇÃO
        '
        Clear elementClean
        'root-> AppObj   |   Painel Acao-> AppObjGrade   |   Grade Filtro-> AppObjPainelAcao
        '
        Call Search(AppObj, eUIA_AutomationIdPropertyId, "NovaSolicitacaoButton", elementClean)
        
        'show ">>>>>>BOTAO Nova Solicitacao:   ", elementClean.CurrentAutomationId
        '
        '
        Set o_InvokePattern = elementClean.GetCurrentPattern(UIAutomationClient.UIA_InvokePatternId)
        '
        If Not elementClean.CurrentIsEnabled Then
        
         o_InvokePattern.Invoke
        
        End If
        
        
        
        
        
        
        
        
        
        '
        'AÇÃO COMBO BOX  Dados - Serviço
        '
        Clear elementClean
        Clear MyElement2
        
        'root-> AppObj   |   Painel Acao-> AppObjGrade   |   Grade Filtro-> AppObjPainelAcao
        '
        Call Search(AppObj, eUIA_AutomationIdPropertyId, "EquipePersonalizadoComboBox", elementClean)
        
        '
        Set o_LegacyAccessiblePattern = elementClean.GetCurrentPattern(UIA_LegacyIAccessiblePatternId)
        '
        elementClean.SetFocus
        
        o_LegacyAccessiblePattern.SetValue (StoragePlan.Range("C4").Value)
        
        Call Search(elementClean, eUIA_LocalizedControlTypePropertyId, "text", MyElement2)
        elementClean.SetFocus
        MyElement2.SetFocus
        o_LegacyAccessiblePattern.DoDefaultAction
        o_LegacyAccessiblePattern.Select (4)
            Set o_ValuePattern = Nothing
        Set o_ValuePattern = MyElement2.GetCurrentPattern(UIAutomationClient.UIA_ValuePatternId)
        o_ValuePattern.SetValue (StoragePlan.Range("C4").Value)
        





        '
        'AÇÃO COMBO BOX  TT
        '
        Clear elementClean
        Clear MyElement2
        'root-> AppObj   |   Painel Acao-> AppObjGrade   |   Grade Filtro-> AppObjPainelAcao
        '
        Call Search(AppObj, eUIA_AutomationIdPropertyId, "FonteDaPerfilacaoComboBox", elementClean)
        '
        Set o_LegacyAccessiblePattern = elementClean.GetCurrentPattern(UIA_LegacyIAccessiblePatternId)
        elementClean.SetFocus
        '
        o_LegacyAccessiblePattern.SetValue (StoragePlan.Range("C5").Value)
        'show elementClean.CurrentLocalizedControlType
        Call Search(elementClean, eUIA_LocalizedControlTypePropertyId, "text", MyElement2)
        elementClean.SetFocus
        MyElement2.SetFocus
        o_LegacyAccessiblePattern.DoDefaultAction
        o_LegacyAccessiblePattern.Select (3)
            Set o_ValuePattern = Nothing
        Set o_ValuePattern = MyElement2.GetCurrentPattern(UIAutomationClient.UIA_ValuePatternId)
        o_ValuePattern.SetValue (StoragePlan.Range("C5").Value)
        '
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        'AÇÃO COMBO BOX Conclído
        '
        Clear elementClean
        Clear MyElement1
        Clear MyElement2
        Clear MyElement3
        Set o_InvokePattern = Nothing
        '
        'root-> AppObj   |   Painel Acao-> AppObjGrade   |   Grade Filtro-> AppObjPainelAcao
        '
        
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
        
        Call Search(MyElement2, eUIA_NamePropertyId, "Concluído", _
        MyElement3)
        Set o_InvokePattern = Nothing
        Set o_InvokePattern = MyElement3.GetCurrentPattern( _
        UIAutomationClient.UIA_InvokePatternId)
        o_InvokePattern.Invoke
        'show MyElement3.CurrentLocalizedControlType
        Set o_LegacyAccessiblePattern = MyElement3.GetCurrentPattern( _
        UIA_LegacyIAccessiblePatternId)
        o_LegacyAccessiblePattern.SetValue (StoragePlan.Range("C6").Value)


        
        




        '
        'AÇÃO COMBO BOX Solicitação
        '
        'Pause 1
        Clear elementClean
        Dim MyElement4 As UIAutomationClient.IUIAutomationElement
        Clear MyElement2
        Clear MyElement3
        Set o_InvokePattern = Nothing
        
        Call Search(AppObj, eUIA_AutomationIdPropertyId, _
            "MotivoComboBox", elementClean)

        Call Search(elementClean, eUIA_NamePropertyId, "Abrir", MyElement4)
        
        elementClean.SetFocus
        Set o_LegacyAccessiblePattern = elementClean.GetCurrentPattern( _
            UIA_LegacyIAccessiblePatternId)
        o_LegacyAccessiblePattern.SetValue (StoragePlan.Range("C7").Value)
        
        'Set o_InvokePattern = MyElement4.GetCurrentPattern( _
        UIAutomationClient.UIA_InvokePatternId)
        'o_InvokePattern.Invoke
        'show MyElement1.CurrentClassName, elementClean.CurrentClassName
        'MyElement4 = Nothing
        'ComboLBox
        
        
        Call Search(elementClean, eUIA_ClassNamePropertyId, "ComboLBox", MyElement2)
        Call Search(MyElement2, eUIA_NamePropertyId, "Solicitação", _
        MyElement3)
        
        Set o_InvokePattern = Nothing
        Set o_InvokePattern = MyElement3.GetCurrentPattern( _
            UIAutomationClient.UIA_InvokePatternId)
        o_InvokePattern.Invoke
        'show MyElement2.CurrentClassName
        Set o_ValuePattern = Nothing
        Set o_ValuePattern = elementClean.GetCurrentPattern(UIAutomationClient.UIA_ValuePatternId)
        o_ValuePattern.SetValue (StoragePlan.Range("C7").Value)

        'show elementClean.CurrentLocalizedControlType


        
        
        
        
        
        
        

        
        
        
        
        
        
        
        'AÇÃO COMBO BOX UF GRADE
        '
        '
        Clear elementClean
        
        'root-> AppObj   |   Painel Acao-> AppObjGrade   |   Grade Filtro-> AppObjPainelAcao
        '
        Call Search(AppObj, eUIA_NamePropertyId, "UF Linha 0", elementClean)
        
        'show ">> element FOUND: ", elementClean.CurrentName
        '
        '
        '
        '
        Set o_LegacyAccessiblePattern = elementClean.GetCurrentPattern(UIA_LegacyIAccessiblePatternId)
        '
        'buscar elemento na ATUAL LINHA NA COLUNA A TEMOS ELEMENTO UF
        '
        o_LegacyAccessiblePattern.SetValue (Range(UF).Value)
        
                        '
                'AÇÃO COMBO BOX Regiao GRADE
                '
                '
        'SE UF = RS,SC,PR,MS,TO,GO,MT,RO,AC = R2
        'SENÃO = AM,RR,AP,PA,MA,CE,RN,PB,PE,AL,SE,BA,MG
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
                'ATUAL LINHA NA COLUNA A TEMOS ELEMENTO REGIÃO
                
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
                'ATUAL LINHA NA COLUNA A TEMOS ELEMENTO REGIÃO
                
                '
                o_LegacyAccessiblePattern.SetValue ("R1")
        End Select
        
        
        
        
        
        'AÇÃO COMBO BOX CNPJ GRADE
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
        'ATUAL LINHA NA COLUNA A TEMOS ELEMENTO REGIÃO
        o_LegacyAccessiblePattern.SetValue (StoragePlan.Range("C10").Value)
        
        
        
        'AÇÃO COMBO BOX PRODUTO GRADE
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
        'ATUAL LINHA NA COLUNA A TEMOS ELEMENTO REGIÃO
        o_LegacyAccessiblePattern.SetValue (StoragePlan.Range("C11").Value)
        
        
        
        
        
        
        '''''''''''''''''''''AÇÃO COMBO BOX TIPO SOLICITAÇÃO GRADE
        
        '
        
        Clear elementClean
        '
        'root-> AppObj   |   Painel Acao-> AppObjGrade   |   Grade Filtro-> AppObjPainelAcao
        '
        Call Search(AppObj, eUIA_NamePropertyId, "Tipo Solicitação Linha 0", elementClean)
        
        'show ">> element FOUND: ", elementClean.CurrentName
        '
        '
        '
        '
        Set o_LegacyAccessiblePattern = elementClean.GetCurrentPattern(UIA_LegacyIAccessiblePatternId)
        '
        'ATUAL LINHA NA COLUNA A TEMOS ELEMENTO REGIÃO
        o_LegacyAccessiblePattern.SetValue (StoragePlan.Range("C12").Value)
        
      '  Set o_LegacyAccessiblePattern = elementClean.GetCurrentPattern(UIA_LegacyIAccessiblePatternId)
       ' show o_LegacyAccessiblePattern.CurrentValue
        
        
        
        
        
        
        '''''''''''''''''''''AÇÃO COMBO BOX MOTIVO GRADE

        
        Clear elementClean
        Clear MyElement1
        Set o_InvokePattern = Nothing
        '
        'root-> AppObj   |   Painel Acao-> AppObjGrade   |   Grade Filtro-> AppObjPainelAcao
        '
        Call Search(AppObj, eUIA_NamePropertyId, "Motivo Linha 0", elementClean)
       
        elementClean.SetFocus
        '
        Set o_LegacyAccessiblePattern = elementClean.GetCurrentPattern(UIA_LegacyIAccessiblePatternId)
        '
        '
        o_LegacyAccessiblePattern.SetValue (StoragePlan.Range("C13").Value)
        '
        '
        Set o_InvokePattern = elementClean.GetCurrentPattern( _
        UIAutomationClient.UIA_InvokePatternId)
        o_InvokePattern.Invoke
        Pause 5
        Set MyElement1 = elementClean.FindFirst( _
        TreeScope_Children, _
        PropCondition(oAutomation, _
        eUIA_NamePropertyId, StoragePlan.Range("C13").Value))
        
        show elementClean.CurrentHelpText
        'Call Search(elementClean, eUIA_NamePropertyId, StoragePlan.Range("C13").Value, MyElement1)
        Set o_InvokePattern = Nothing
        Set o_InvokePattern = MyElement1.GetCurrentPattern( _
        UIAutomationClient.UIA_InvokePatternId)
        o_InvokePattern.Invoke
        
        
        
        
        
        
        
        
        
        '''''''''''''''''''''''''''''''AÇÃO COMBO BOX OS GERADA / TT GRADE
        '
        Clear elementClean
        '
        'root-> AppObj   |   Painel Acao-> AppObjGrade   |   Grade Filtro-> AppObjPainelAcao
        '
        Call Search(AppObj, eUIA_NamePropertyId, "OS Gerada / TT Linha 0", elementClean)
        
        'show ">> element FOUND: ", elementClean.CurrentName
        '
        '
        '
        '
        Set o_LegacyAccessiblePattern = elementClean.GetCurrentPattern(UIA_LegacyIAccessiblePatternId)
        '
        'ATUAL LINHA NA COLUNA A TEMOS ELEMENTO REGIÃO
        o_LegacyAccessiblePattern.SetValue (Range(OS).Value)
        
        
        
        '''''''''''''''''''''''''''AÇÃO COMBO BOX QTD. GRADE
        Clear elementClean
        '
        'root-> AppObj   |   Painel Acao-> AppObjGrade   |   Grade Filtro-> AppObjPainelAcao
        '
        Call Search(AppObj, eUIA_NamePropertyId, "Qtd. Linha 0", elementClean)
        
        'show ">> element FOUND: ", elementClean.CurrentName
        '
        '
        '
        '
        Set o_LegacyAccessiblePattern = elementClean.GetCurrentPattern(UIA_LegacyIAccessiblePatternId)
        '
        'ATUAL LINHA NA COLUNA A TEMOS ELEMENTO REGIÃO
        o_LegacyAccessiblePattern.SetValue ("1")
        
        
        
        
        
        '''''''''''''''''''''''''''''''''AÇÃO COMBO BOX RESPONSAVEL GRADE

        Clear elementClean
        '
        'root-> AppObj   |   Painel Acao-> AppObjGrade   |   Grade Filtro-> AppObjPainelAcao
        '
        Call Search(AppObj, eUIA_NamePropertyId, "Resp. Tratamento Linha 0", elementClean)
        
        'show ">> element FOUND: ", elementClean.CurrentName
        '
        '
        '
        '
        Set o_LegacyAccessiblePattern = elementClean.GetCurrentPattern(UIA_LegacyIAccessiblePatternId)
        '
        'ATUAL LINHA NA COLUNA A TEMOS ELEMENTO REGIÃO
        o_LegacyAccessiblePattern.SetValue (StoragePlan.Range("C16").Value)
        
        
        
        
        
        
        
        
        
        '2º
        '
        ' AÇÃO: PROTOCOLO ORIGEM
        '
        Clear elementClean
        '
        'root-> AppObj   |   Painel Acao-> AppObjGrade   |   Grade Filtro-> AppObjPainelAcao
        '
        Call Search(AppObj, eUIA_AutomationIdPropertyId, "ProtocoloTextBox", elementClean)
        
        'show ">>>>>>TEXTO Protocolo:    ", elementClean.CurrentAutomationId
        '
        '
        'BOTAO TEXTO EMAIL CLIENTE
        '
        Set o_LegacyAccessiblePattern = elementClean.GetCurrentPattern(UIA_LegacyIAccessiblePatternId)
        
        o_LegacyAccessiblePattern.SetValue (StoragePlan.Range("C1").Value)
        
        
        
        
        
        
        '3º
        
        '
        '
        'AÇÃO: EMAIL CLIENTE
        '
        Clear elementClean
        '
        'root-> AppObj   |   Painel Acao-> AppObjGrade   |   Grade Filtro-> AppObjPainelAcao
        '
        Call Search(AppObj, eUIA_AutomationIdPropertyId, "EmailTextBox", elementClean)
        
        'show " >>TEXTO Email Cliente: " & elementClean.CurrentAutomationId
        '
        '
        'BOTAO EMAIL CLIENTE
        '
        Set o_LegacyAccessiblePattern = elementClean.GetCurrentPattern(UIA_LegacyIAccessiblePatternId)
        '
        o_LegacyAccessiblePattern.SetValue (StoragePlan.Range("C2").Value)
        
        
        
        
        
        
        '4º comentario
        
        '
        '
        'AÇÃO: COMENTARIO
        '
        Clear elementClean
        '
        'root-> AppObj   |   Painel Acao-> AppObjGrade   |   Grade Filtro-> AppObjPainelAcao
        '
        Call Search(AppObjPainelAcao, eUIA_AutomationIdPropertyId, "ComentariosRichText", elementClean)
        
        'show " >>TEXTO Comentarios: ", elementClean.CurrentClassName, elementClean.CurrentAutomationId
        '
        '
        'TEXTO COMENTARIO
        '
        Set o_LegacyAccessiblePattern = elementClean.GetCurrentPattern(UIA_LegacyIAccessiblePatternId)
        '
        o_LegacyAccessiblePattern.SetValue (StoragePlan.Range("C3").Value)
        
        
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


Public Function Search(ByVal obj As UIAutomationClient.IUIAutomationElement, _
    typed As Condition, strFinalElemSearch As String, _
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
        
        'aguarda execução para q o pc possa fazer outras tarefas
        DoEvents
        If element1.Length <> 0 Then
                Set element2 = obj.FindFirst(TreeScope_Children, condition1)
        End If
        Do While Not element2 Is Nothing
            'verificar como tratar elemento atual
            Dim oPattern As UIAutomationClient.IUIAutomationLegacyIAccessiblePattern
            Set oPattern = element2.GetCurrentPattern(UIA_LegacyIAccessiblePatternId)
            'show element2.CurrentLocalizedControlType
            'se pattern é o certo, temos mais um elemento pra adicionar aqui
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
                        'show "result CurrentLocalizedControlType $$ ", strFinalElemSearch
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
            
            'não é o elemento que pedi pq saiu do SELECT SEM achar o     elem     ENTÃO
            'andando pra frente busca elemento interno
            Search element2, typed, strFinalElemSearch, elem

            'show "não é o elemento que pedi : ", element2.CurrentControlType, elem.CurrentAutomationId
            
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
        
        'aguarda execução para q o pc possa fazer outras tarefas
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
        
            'se pattern é o certo, temos mais um elemento pra adicionar aqui
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
             
            'não é o elemento que pedi pq saiu do SELECT SEM achar o     elem     ENTÃO
            'andando pra frente busca elemento interno
            SearchPattern element2, typed, strFinalElemSearch, elem

       '     show "não é o elemento que pedi : ", element2.CurrentControlType, elem.CurrentName
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
            'StrComp return 0 – if strings are equal;
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
 
