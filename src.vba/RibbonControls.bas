Dim ribbonUI As IRibbonUI


Public Sub AddInLoaded(ribbon As IRibbonUI)
    Set ribbonUI = ribbon
End Sub


' Get status for specific controls

Function isAutoUpdateSectionTimeline() As Boolean
    isAutoUpdateSectionTimeline = CBool(ActivePresentation.Tags("isAutoUpdateSectionTimeline") = "True")
End Function

Function isBoldActiveSection() As Boolean
    isBoldActiveSection = CBool(ActivePresentation.Tags("IsBoldActiveSection") = "True")
End Function


' Checkboxes

Sub GetControlState(control As IRibbonControl, ByRef returnedVal)
    Dim tagVal As String
    Dim tagId As String
    tagId = control.Tag
    tagVal = CBool(ActivePresentation.Tags(tagId) = "True")

    If tagVal <> "" Then
        returnedVal = tagVal
    Else
        ActivePresentation.Tags.Add tagId, "True"
        returnedVal = True
    End If
End Sub


Sub SetControlState(control As IRibbonControl, pressed As Boolean)
    Dim tagId As String
    tagId = control.Tag
    ActivePresentation.Tags.Add tagId, CStr(pressed)
    
    If tagId = "IsBoldActiveSection" Then
        UpdateFooterTimeline showMsgBox:=False
    End If
End Sub


' Editbox

Sub GetSectionSeparator(control As IRibbonControl, ByRef returnedVal)
    Dim tagVal As String
    tagVal = ActivePresentation.Tags("SectionSeparator")
    If tagVal <> "" Then
        returnedVal = tagVal
    Else
        ActivePresentation.Tags.Add "SectionSeparator", "  >  "
        returnedVal = "  >  "
    End If
End Sub


Sub SetSectionSeparator(control As IRibbonControl, newValue As String)
    If newValue <> ActivePresentation.Tags("SectionSeparator") Then
        ActivePresentation.Tags.Add "SectionSeparator", newValue
        UpdateFooterTimeline showMsgBox:=False
    End If
End Sub