VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StoragePlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'    oAutomation    UIAutomationClient . CUIAutomation
' Element   AppObj[AppObjGrade,AppObjPainelAcao, elementClean, MyElement1.2.3]
' Pattern   o_ValuePattern, o_ValuePattern, o_LegacyAccessiblePattern
' Enum      ConditionLegacyPattern, Condition, walker, count
Dim i As Integer
Private Sub ComboBox2_DropButtonClick()
    'ComboBox2_DropDown
'End Sub
'Private Sub ComboBox2_initialize()
   
'End Sub
'Private Sub ComboBox2_DropDown()
    i = i + 1
    show i
    Clear AppObj
    Clear elementClean
    Set walker = oAutomation.ControlViewWalker
    Set elementClean = walker.GetFirstChildElement(oAutomation.GetRootElement)
    With ComboBox2
        .Clear
        .Value = "LISTA DE APP"
        
        
    End With
    
    Do While Not elementClean Is Nothing
        Set elementClean = walker.GetNextSiblingElement(elementClean)
        If Not elementClean Is Nothing Then
            With ComboBox2
                .AddItem elementClean.CurrentClassName
            End With
             'show elementClean.CurrentClassName
        End If
       
    Loop
    'Clear AppObj
End Sub
  
  
  
  
    '
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





Private Sub worksheet_activate()
    show "foi a"

    Application.SendKeys "^g^a{DEL}"
    
    With ComboBox2
        .Clear
        .Value = "LISTA DE APP"
        
        
    End With
End Sub
