#If VBA7 Then
    Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
#Else
    Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If
Public oAutomation As New CUIAutomation ' the UI Automation API\
Public AppObj As UIAutomationClient.IUIAutomationElement

Public AppObjGrade As UIAutomationClient.IUIAutomationElement
Public AppObjPainelAcao As UIAutomationClient.IUIAutomationElement

Public elementClean As UIAutomationClient.IUIAutomationElement
Public MyElement1 As UIAutomationClient.IUIAutomationElement
Public MyElement2 As UIAutomationClient.IUIAutomationElement
Public MyElement3 As UIAutomationClient.IUIAutomationElement

Public o_InvokePattern As UIAutomationClient.IUIAutomationInvokePattern
Public o_LegacyAccessiblePattern As UIAutomationClient.IUIAutomationLegacyIAccessiblePattern

Public protocolo As String
Public emailcliente As String
Public cnpjcliente As String
Public comentario As String

Public Enum Condition
   eUIA_NamePropertyId
   eUIA_AutomationIdPropertyId
   eUIA_ClassNamePropertyId
   eUIA_LocalizedControlTypePropertyId
End Enum

Public Function Clear(c As UIAutomationClient.IUIAutomationElement)
     CNPJ = Sheets("Storage").Range("C12").Value
     UF = "A" + CStr(ActiveCell.row)
     Movel = "B" + CStr(ActiveCell.row)
     OS = "C" + CStr(ActiveCell.row)
     show Range(UF).Value, Range(Movel).Value, Range(OS).Value
     
     
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
        'If elementClean Is Nothing Then
            'show " elementClean Is Nothing"
        'End If
        'If Not elementClean Is Nothing Then
        '    Set elementClean = Nothing
        '    If elementClean Is Nothing Then
            'show " INCEPTION Is Nothing"
        '    End If
        'End If
       
        'root-> AppObj   |   Painel Acao-> AppObjGrade   |   Grade Filtro-> AppObjPainelAcao
        '
        Call Search(AppObj, eUIA_AutomationIdPropertyId, "NovaSolicitacaoButton", elementClean)
       
        'show ">>>>>>BOTAO Nova Solicitacao:   ", elementClean.CurrentAutomationId
        '
        '
        Set o_InvokePattern = elementClean.GetCurrentPattern(UIAutomationClient.UIA_InvokePatternId)
        '
        o_InvokePattern.Invoke
       
       
       
       
       
       
       
       
       
       
       
        '
        'AÇÃO COMBO BOX  Dados - Serviço
        '
        Clear elementClean
        '
        'root-> AppObj   |   Painel Acao-> AppObjGrade   |   Grade Filtro-> AppObjPainelAcao
        '
        Call Search(AppObj, eUIA_AutomationIdPropertyId, "EquipePersonalizadoComboBox", elementClean)
       
        'show ">> element FOUND: " & elementClean.CurrentAutomationId
        '
        '
        '
        '
        Set o_LegacyAccessiblePattern = elementClean.GetCurrentPattern(UIA_LegacyIAccessiblePatternId)
        '
        o_LegacyAccessiblePattern.SetValue ("Dados - Serviço")
        '
        '
        'AÇÃO COMBO BOX  TT
        '
        Clear elementClean
        '
        'root-> AppObj   |   Painel Acao-> AppObjGrade   |   Grade Filtro-> AppObjPainelAcao
        '
        Call Search(AppObj, eUIA_AutomationIdPropertyId, "FonteDaPerfilacaoComboBox", elementClean)
       
        'show ">> element FOUND: " & elementClean.CurrentAutomationId
        '
        '
        '
        '
        Set o_LegacyAccessiblePattern = elementClean.GetCurrentPattern(UIA_LegacyIAccessiblePatternId)
        '
        o_LegacyAccessiblePattern.SetValue ("TT")
       
       
       
        'AÇÃO COMBO BOX Conclído
        '
        Clear elementClean
        '
        'root-> AppObj   |   Painel Acao-> AppObjGrade   |   Grade Filtro-> AppObjPainelAcao
        '
        Call Search(AppObj, eUIA_AutomationIdPropertyId, "StatusComboBox", elementClean)
       
        'show ">> element FOUND: " & elementClean.CurrentAutomationId
        '
        '
        '
        '
        Set o_LegacyAccessiblePattern = elementClean.GetCurrentPattern(UIA_LegacyIAccessiblePatternId)
        '
        o_LegacyAccessiblePattern.SetValue ("Concluído")
       
        'AÇÃO COMBO BOX Solicitação
        '
        Clear elementClean
        '
        'root-> AppObj   |   Painel Acao-> AppObjGrade   |   Grade Filtro-> AppObjPainelAcao
        '
        Call Search(AppObj, eUIA_AutomationIdPropertyId, "MotivoComboBox", elementClean)
       
        'show ">> element FOUND: " & elementClean.CurrentAutomationId
        '
        '
        '
        '
        Set o_LegacyAccessiblePattern = elementClean.GetCurrentPattern(UIA_LegacyIAccessiblePatternId)
        '
        o_LegacyAccessiblePattern.SetValue ("Solicitação")
       
       
       
       
       
       
       
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
        o_LegacyAccessiblePattern.SetValue (CNPJ)
       
       
       
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
        o_LegacyAccessiblePattern.SetValue (Sheets("Storage").Range("C13").Value)
       
       
       
       
       
       
        'AÇÃO COMBO BOX TIPO SOLICITAÇÃO GRADE
       
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
        o_LegacyAccessiblePattern.SetValue (Sheets("Storage").Range("C14").Value)
       
       
       
       
        'AÇÃO COMBO BOX MOTIVO GRADE

       
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
        o_LegacyAccessiblePattern.SetValue (Sheets("Storage").Range("C15").Value)
       
       
       
       
       
       
        'AÇÃO COMBO BOX OS GERADA / TT GRADE
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
       
       
       
        'AÇÃO COMBO BOX QTD. GRADE
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
       
       
       
       
       
        'AÇÃO COMBO BOX RESPONSAVEL GRADE
       
                'AÇÃO COMBO BOX QTD. GRADE
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
        o_LegacyAccessiblePattern.SetValue (Sheets("Storage").Range("B5").Value)
       
       
       
       
       
       
       
       
       
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
       
        o_LegacyAccessiblePattern.SetValue (protocolo)
       
       
       
       
       
       
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
        o_LegacyAccessiblePattern.SetValue (emailcliente)
       
       
       
       
       
       
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
        o_LegacyAccessiblePattern.SetValue (comentario)
       
       
       
       
End Sub


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
            'se pattern é o certo, temos mais um elemento pra adicionar aqui
            'try catch bloc to
            Select Case typed
                Case eUIA_NamePropertyId
                    If StrComp(element2.CurrentName, strFinalElemSearch) = 0 Then
                        'show "result CurrentLocalizedControlType $$ ", strFinalElemSearch
                        ended = True
                        Set elem = element2
                        Set Search = elem
                        Exit Do
                    End If
                   
                Case eUIA_AutomationIdPropertyId
                    If StrComp(element2.CurrentAutomationId, strFinalElemSearch) = 0 Then
                        'show "result CurrentLocalizedControlType $$ ", strFinalElemSearch
                        ended = True
                        Set elem = element2
                        Set Search = elem
                        Exit Do
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

            'show "não é o elemento que pedi : ", element2.CurrentControlType
           
            If Not elem Is Nothing Then Exit Function
           
            Set element2 = walker.GetNextSiblingElement(element2)
            'andando pra traz
 
        Loop
        If ended Then Exit Function
End Function
