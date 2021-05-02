Attribute VB_Name = "Main"
Option Explicit

Public STR_Element As String
Public UI_Layer_Number_Element As Integer
Public UI_Automation As New CUIAutomation
Public UI_ElementArray As IUIAutomationElementArray
Public UI_Element As IUIAutomationElement
Public UI_ElementEmpty As IUIAutomationElement

Public UI_TreeWalker As IUIAutomationTreeWalker
'Public UI_AutoCacheRequest As IUIAutomationCacheRequest

Public UI_LegacyAccPattern As IUIAutomationLegacyIAccessiblePattern

Private UI_TrueCondition As IUIAutomationCondition
Private UI_PropertiesDictionary As New Scripting.Dictionary
Private UI_DictionaryControl As New Scripting.Dictionary

Sub PlayLoop()
    Set UI_Element = UI_Automation.GetRootElement
    Busca UI_Element, eUIA_AutomationIdPropertyId, "NovaSolicitacaoButton", UI_ElementEmpty
    show UI_Element.CurrentClassName
End Sub

Public Function Busca(ByVal obj As UIAutomationClient.IUIAutomationElement, typed As Condition, strFinalElemSearch As String, ByRef elem As UIAutomationClient.IUIAutomationElement, Optional Layer_Number As Integer = 1) As UIAutomationClient.IUIAutomationElement
        On Error Resume Next
        Dim ended As Boolean
        ended = False
        Dim walker As UIAutomationClient.IUIAutomationTreeWalker
        Dim element1 As UIAutomationClient.IUIAutomationElementArray
        Dim element2 As UIAutomationClient.IUIAutomationElement

        Set walker = oAutomation.ControlViewWalker
        Dim condition1 As UIAutomationClient.IUIAutomationCondition

        Set condition1 = oAutomation.CreateTrueCondition

        Set element1 = obj.FindAll(TreeScope_Descendants, condition1)
        
        DoEvents
        If element1.Length <> 0 Then

                Set element2 = obj.FindFirst(TreeScope_Children, condition1)
                'colocar na celula a CLASS do elemento
        Else
            'caso nao ache valor, colocar NULL na celula e voltar para LAYER acima.
            UI_Layer_Number_Element = UI_Layer_Number_Element + 1
            UI_Matrix.Cells(UI_Layer_Number_Element, Layer_Number).Value = "NOTHING"
            'Layer_Number = Layer_Number - 1
            show "-1"
        End If
        
        UI_Layer_Number_Element = UI_Layer_Number_Element + 1
        
        Do While Not element2 Is Nothing
            Select Case typed
                Case eUIA_AutomationIdPropertyId
                    If StrComp(element2.CurrentAutomationId, strFinalElemSearch) = 0 Then
                        show "result Automation Id Property Id $$ ", strFinalElemSearch
                        ended = True
                        Set elem = element2
                        Set Busca = elem
                        Exit Function
                    End If
            End Select
            
            show "+1"
            If element2.CurrentClassName = "" And Not element2.CurrentAutomationId = "" Then
                UI_Matrix.Cells(UI_Layer_Number_Element, Layer_Number).Value = element2.CurrentAutomationId
                
                
                show "> ", element2.CurrentAutomationId
            ElseIf Not element2.CurrentClassName = "" Then
                UI_Matrix.Cells(UI_Layer_Number_Element, Layer_Number).Value = element2.CurrentClassName
                
                show "> ", element2.CurrentClassName
            ElseIf Not element2.CurrentName = "" Then
                UI_Matrix.Cells(UI_Layer_Number_Element, Layer_Number).Value = element2.CurrentName
                
                show "> ", element2.CurrentName
            End If
            
            Set UI_LegacyAccPattern = element2.GetCurrentPattern(UIA_LegacyIAccessiblePatternId)
            
            If UI_LegacyAccPattern.CurrentName = "" And Not UI_LegacyAccPattern.CurrentDescription = "" Then
                show "Legacy Description>:", UI_LegacyAccPattern.CurrentDescription
            ElseIf Not UI_LegacyAccPattern.CurrentName = "" Then
                show "Legacy Name>:", UI_LegacyAccPattern.CurrentName
            End If
            
            
            
            
            
            Layer_Number = Layer_Number + 1
            'caso não tenha filho vai pular o while
            Busca element2, typed, strFinalElemSearch, elem, Layer_Number
            
            'encerra loop que vai para filho
            show "não é o elemento que pedi : ", element2.CurrentControlType, elem.CurrentClassName, element2.CurrentClassName
            
            If Not elem Is Nothing Then Exit Function
            Set element2 = walker.GetNextSiblingElement(element2)
            Layer_Number = Layer_Number - 1
            
            show "--1 "
            Loop
            
        If ended Then Exit Function

End Function
























Function Execute(ByVal obj As UIAutomationClient.IUIAutomationElement, Optional STR_UI_Elemnt As String = "")
    On Error Resume Next
    Dim count As Integer
    Set UI_TreeWalker = UI_Automation.ControlViewWalker
    Set UI_TrueCondition = UI_Automation.CreateTrueCondition
    Set UI_ElementArray = UI_Automation.GetRootElement.FindAll(TreeScope_Descendants, UI_TrueCondition)
    DoEvents
    'If UI_ElementArray.Length <> 0 Then
    '    Set UI_Element = obj.FindFirst(TreeScope_Children, UI_TrueCondition)
    'End If
    For count = 0 To UI_ElementArray.Length

        Set UI_Element = UI_ElementArray.GetElement(count)
        show UI_Element.CurrentName
    Next count
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
            
            Search element2, typed, strFinalElemSearch, elem

            show "não é o elemento que pedi : ", element2.CurrentControlType, elem.CurrentAutomationId
            
            If Not elem Is Nothing Then Exit Function
            
            Set element2 = walker.GetNextSiblingElement(element2)
        Loop
        If ended Then Exit Function
End Function



