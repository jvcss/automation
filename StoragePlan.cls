VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StoragePlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Dim element As String
'    oAutomation    UIAutomationClient . CUIAutomation
' Element   AppObj[AppObjGrade,AppObjPainelAcao, elementClean, MyElement1.2.3]
' Pattern   o_ValuePattern, o_ValuePattern, o_LegacyAccessiblePattern  '
' Enum      ConditionLegacyPattern, Condition, walker, count



Private Sub AddAction_Click()

End Sub

    'SearchActiveX ( Root,iN Action ,iN ElemAttr , ElemInfo , ElemResult )
    '   .walk trought tree with true condition with array *can reuse this to activate combobox-listbox
    '   .while get element get pattern?
    '       'case element select elemenAttr
    '           result = element
    '       'case action
    '           element.action
    '   ?
    '
    'click
    'set
    'get
    'check
    'uncheck
    '
    ' o_LegacyAccessiblePattern.SetValue ("Ctrl + Shift")
    'show AppObj.CurrentName, o_LegacyAccessiblePattern.CurrentName, o_LegacyAccessiblePattern.CurrentDefaultAction


Sub ComboBox2_Change()
    show "Change event", StoragePlan.OLEObjects("ComboBox2").Object.Value
    
  '  ComboBox2.Value = StoragePlan.OLEObjects("ComboBox2").Object.Value
End Sub

Sub ComboBox2_LostFocus()
  '  show "LostFocus event", StoragePlan.OLEObjects("ComboBox2").Object.Value
   Set element = bodyasync
    show element.CurrentClassName
 '   ADM_Action_Element (element)
    
    
    
    
End Sub




Public Function bodyasync() As UIAutomationClient.IUIAutomationElementArray

    Dim i As Integer
    i = 0
    show i, "conta"
    Clear AppObj
    Clear elementClean
    Clean o_LegacyAccessiblePattern
    
    Set walker = oAutomation.ControlViewWalker
    Set elementClean = walker.GetFirstChildElement(oAutomation.GetRootElement)
    'Set element = elementClean
    'With ComboBox2
   '         .Clear
   ' End With
   
    show elementClean.CurrentClassName
 '   MyElementArray(1) = elementClean
 '   show MyElementArray.GetElement(1)
    Do While Not elementClean Is Nothing
        Set o_LegacyAccessiblePattern = elementClean.GetCurrentPattern(UIA_LegacyIAccessiblePatternId)
        Set elementClean = walker.GetNextSiblingElement(elementClean)
    '    Set MyElementArray(i) = elementClean
     '   show elementClean.CurrentClassName
        i = i + 1
        'ELEMENTO NA MAO elementClean / o_LegacyAccessiblePattern
        
        If Not elementClean Is Nothing Then
            With ComboBox2
               ' .Clear
                .AddItem o_LegacyAccessiblePattern.CurrentName
            End With
        End If
    Loop
    bodyasync ' = MyElementArray
End Function

Private Sub ComboBox2_DropButtonClick()
    With ComboBox2
           .Clear
    End With
    bodyasync
End Sub

Private Sub ComboBoxElement_DropButtonClick()
    'bodyasync
    'elementClean
    With ComboBoxElement
           .Clear
    End With
    bodyasync
    'element = bodyasync.GetElement(1)
    'ADM_Action_Element (element)
   ' show "Clean", element.CurrentClassName
    
End Sub


Private Sub Worksheet_SelectionChange(ByVal Target As Range)

End Sub


Function ADM_Action_Element(element As UIAutomationClient.IUIAutomationElement) 'ByVal element As UIAutomationClient.IUIAutomationElement)
    On Error Resume Next
    Dim walker As UIAutomationClient.IUIAutomationTreeWalker
    Dim element1 As UIAutomationClient.IUIAutomationElementArray
    Dim element2 As UIAutomationClient.IUIAutomationElement
    Dim condition As UIAutomationClient.IUIAutomationCondition
    Set walker = oAutomation.ControlViewWalker
    Set condition = oAutomation.CreateTrueCondition
    
    
    Set element1 = element.FindAll(TreeScope_Children, condition)
    DoEvents
    If element1.Length <> 0 Then
        Set element2 = element.FindFirst(TreeScope_Children, condition)
    End If
    
    Do While Not element2 Is Nothing
        Set o_LegacyAccessiblePattern = element2.GetCurrentPattern(UIA_LegacyIAccessiblePatternId)
        show "current Pattern.", o_LegacyAccessiblePattern.CurrentName
        
        If Not element2 Is Nothing Then
            With ComboBoxElement
               ' .Clear
                .AddItem o_LegacyAccessiblePattern.CurrentName
            End With
        End If
        
        ADM_Action_Element element2
        If Not element Is Nothing Then Exit Function
        Set element2 = walker.GetNextSiblingElement(element2)
    Loop
End Function
