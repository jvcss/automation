Attribute VB_Name = "Atividade"
#If VBA7 Then
    Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
#Else
    Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If
Public oAutomation As New CUIAutomation ' the UI Automation API\
Public AppObj As UIAutomationClient.IUIAutomationElement
Public elementClean As UIAutomationClient.IUIAutomationElement
Public MyElement1 As UIAutomationClient.IUIAutomationElement
Public MyElement2 As UIAutomationClient.IUIAutomationElement
Public MyElement3 As UIAutomationClient.IUIAutomationElement
Public MyElement4 As UIAutomationClient.IUIAutomationElement
Public MyElement5 As UIAutomationClient.IUIAutomationElement
Public MyElement6 As UIAutomationClient.IUIAutomationElement
Public MyElement7 As UIAutomationClient.IUIAutomationElement
Public MyElement8 As UIAutomationClient.IUIAutomationElement
Public MyElement9 As UIAutomationClient.IUIAutomationElement
Public MyElement10 As UIAutomationClient.IUIAutomationElement
Public MyElement11 As UIAutomationClient.IUIAutomationElement
Public MyElement12 As UIAutomationClient.IUIAutomationElement
Public MyElement13 As UIAutomationClient.IUIAutomationElement
Public MyElement14 As UIAutomationClient.IUIAutomationElement
Public MyElement15 As UIAutomationClient.IUIAutomationElement
Public MyElement16 As UIAutomationClient.IUIAutomationElement
Public MyElement17 As UIAutomationClient.IUIAutomationElement
Public MyElement18 As UIAutomationClient.IUIAutomationElement
Public MyElement19 As UIAutomationClient.IUIAutomationElement
Public MyElement20 As UIAutomationClient.IUIAutomationElement
Public MyElement21 As UIAutomationClient.IUIAutomationElement
Public MyElement22 As UIAutomationClient.IUIAutomationElement
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
Public Function Clear(ByVal c As UIAutomationClient.IUIAutomationElement)
   
    If Not c Is Nothing Then
        Set c = MyElement1
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



        'Set o_InvokePattern = elementClean.GetCurrentPattern(UIAutomationClient.UIA_InvokePatternId)
        'Set o_LegacyAccessiblePattern = elementClean.GetCurrentPattern(UIA_LegacyIAccessiblePatternId)
        'Set o_InvokePattern = elementClean.GetCurrentPropertyValue(50000)
Public Sub Test()

        protocolo = "202100016611927"
        emailcliente = "não tem"
        cnpjcliente = "06649157000129"
        comentario = "retirar: Envio de Oi Torpedo Recepção de Sons & Imagens e Bate Papo Recepção de Mensagens da Oi"







        '1^ botao nova sol
        Set AppObj = oAutomation.GetRootElement.FindFirst( _
        TreeScope_Children, _
        PropCondition(oAutomation, _
        eUIA_AutomationIdPropertyId, "Form_Perfilacao_Outros"))
        
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
        
        Set AppObj = AppObj.FindFirst( _
        TreeScope_Children, _
        PropCondition(oAutomation, _
        eUIA_AutomationIdPropertyId, "TableLayoutPanel5"))
        '----------------------------------------------------------------------root
        
        'PainelAcaoBO
        Set elementClean = AppObj.FindFirst( _
        TreeScope_Children, _
        PropCondition(oAutomation, _
        eUIA_AutomationIdPropertyId, "TableLayoutPanel9"))
        Set elementClean = elementClean.FindFirst( _
        TreeScope_Children, _
        PropCondition(oAutomation, _
        eUIA_AutomationIdPropertyId, "PainelAcaoBO"))
        '----------------------------------------------------------------------PainelAcaoBO
        'NovaSolicitacaoButton
        Call Search(elementClean, eUIA_AutomationIdPropertyId, _
        "NovaSolicitacaoButton", MyElement1)
        
        Set o_InvokePattern = MyElement1.GetCurrentPattern(UIAutomationClient.UIA_InvokePatternId)
        o_InvokePattern.Invoke
       
       
       
       
       
       
       
       
       
       
       
       
       
       
       
       
        '2º tt origem
        Set AppObj = oAutomation.GetRootElement.FindFirst( _
        TreeScope_Children, _
        PropCondition(oAutomation, _
        eUIA_AutomationIdPropertyId, "Form_Perfilacao_Outros"))
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
        Set AppObj = AppObj.FindFirst( _
        TreeScope_Children, _
        PropCondition(oAutomation, _
        eUIA_AutomationIdPropertyId, "TableLayoutPanel5"))
       ''----------------------------------------------------------------------root
       
          'PainelAcaoBO
        Set MyElement2 = AppObj.FindFirst( _
        TreeScope_Children, _
        PropCondition(oAutomation, _
        eUIA_AutomationIdPropertyId, "TableLayoutPanel9"))
        Set MyElement2 = AppObj.FindFirst( _
        TreeScope_Children, _
        PropCondition(oAutomation, _
        eUIA_AutomationIdPropertyId, "PainelAcaoBO"))
       
        'Protocolo Text Box > PainelAcaoBO
        Call Search(AppObj, eUIA_AutomationIdPropertyId, _
        "ProtocoloTextBox", MyElement4)
        Set o_LegacyAccessiblePattern = elementClean.GetCurrentPattern(UIA_LegacyIAccessiblePatternId)
        o_LegacyAccessiblePattern.SetValue (protocolo)
        show " >>ProtocoloTextBox<< ", MyElement3.CurrentAutomationId

        'Call Search(AppObj, eUIA_AutomationIdPropertyId, "ProtocoloTextBox")
        'Debug.Print " >>autoID<< " & elementClean.CurrentAutomationId
        'Set o_LegacyAccessiblePattern = elementClean.GetCurrentPattern(UIA_LegacyIAccessiblePatternId)
        'o_LegacyAccessiblePattern.SetValue = protocolo
       
       
       
       
       
       
       
       
       
       
       '3^ email cliente
       Set AppObj = oAutomation.GetRootElement.FindFirst( _
        TreeScope_Children, _
        PropCondition(oAutomation, _
        eUIA_AutomationIdPropertyId, "Form_Perfilacao_Outros"))
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
        Set AppObj = AppObj.FindFirst( _
        TreeScope_Children, _
        PropCondition(oAutomation, _
        eUIA_AutomationIdPropertyId, "TableLayoutPanel5"))
       ''root
       
          'PainelAcaoBO
        Set elementClean = AppObj.FindFirst( _
        TreeScope_Children, _
        PropCondition(oAutomation, _
        eUIA_AutomationIdPropertyId, "TableLayoutPanel9"))
        
        Set elementClean = AppObj.FindFirst( _
        TreeScope_Children, _
        PropCondition(oAutomation, _
        eUIA_AutomationIdPropertyId, "PainelAcaoBO"))
       
        'Protocolo Text Box > PainelAcaoBO
        Call Search(elementClean, eUIA_AutomationIdPropertyId, _
        "EmailTextBox", MyElement3)
        Set o_LegacyAccessiblePattern = elementClean.GetCurrentPattern(UIA_LegacyIAccessiblePatternId)
        o_LegacyAccessiblePattern.SetValue (protocolo)
        show " >>ProtocoloTextBox<< ", MyElement4.CurrentAutomationId

        'Call Search(AppObj, eUIA_AutomationIdPropertyId, "ProtocoloTextBox")
        'Debug.Print " >>autoID<< " & elementClean.CurrentAutomationId
        'Set o_LegacyAccessiblePattern = elementClean.GetCurrentPattern(UIA_LegacyIAccessiblePatternId)
        'o_LegacyAccessiblePattern.SetValue = protocolo
       
       
       
       
       
       
       
       
       
       
       
       
       
       
       
       
       
       
       
       
       
       
       
       
       
       
       
       
       
       
       
       
       
       
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

Sub ok()
    Set AppObj = oAutomation.GetRootElement.FindFirst( _
        TreeScope_Children, _
        PropCondition(oAutomation, eUIA_NamePropertyId, "Calculadora"))
        
        show AppObj.CurrentClassName
        Set AppObj = AppObj.FindFirst( _
        TreeScope_Children, _
        PropCondition(oAutomation, eUIA_NamePropertyId, "Calculadora"))
        
        show AppObj.CurrentClassName
        
        Search AppObj, eUIA_NamePropertyId, "Oito", elementClean
        
        show "e: ", elementClean.CurrentClassName
End Sub


Public Function Search( _
ByVal obj As UIAutomationClient.IUIAutomationElement, _
typed As Condition, _
strFinalElemSearch As String, _
ByRef elem As UIAutomationClient.IUIAutomationElement) _
As UIAutomationClient.IUIAutomationElement

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

            Select Case typed
                Case eUIA_NamePropertyId
                '   InStr(1, element2.CurrentName, strFinalElemSearch) > 0
                    If StrComp(element2.CurrentName, strFinalElemSearch) = 0 Then
                    show "result CurrentName>", element2.CurrentName
                        ended = True
                        Set elem = element2
                        Set Search = elem
                         Exit Function
                    End If
                   
                Case eUIA_AutomationIdPropertyId
                    If StrComp(element2.CurrentAutomationId, strFinalElemSearch) = 0 Then
                        show "result CurrentAutomationId>", strFinalElemSearch
                        ended = True
                        Set elem = element2
                        Set Search = elem
                        Exit Function
                    End If
                   
                Case eUIA_ClassNamePropertyId
                    If StrComp(element2.CurrentClassName, strFinalElemSearch) = 0 Then
                        show "result CurrentClassName>", elem.CurrentClassName
                        ended = True
                        Set elem = element2
                        Set Search = elem
                        Exit Function
                    End If
                   
                Case eUIA_LocalizedControlTypePropertyId
                    If StrComp(element2.CurrentLocalizedControlType, strFinalElemSearch) = 0 Then
                        show "result CurrentLocalizedControlType>", elem.CurrentLocalizedControlType
                        ended = True
                        Set elem = element2
                        Set Search = elem
                        Exit Function
                    End If
            End Select
             
           
            'andando rumo aos filhos
            Search element2, typed, strFinalElemSearch, elem
            If ended Then
                Exit Function
            End If
            show "elemento atual>", element2.CurrentAutomationId, "::", element2.CurrentClassName
            If Not elem Is Nothing Then Exit Function
           

            Set element2 = walker.GetNextSiblingElement(element2)
            'andando de lado
 
        Loop
        If ended Then Exit Function
End Function
