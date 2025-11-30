Public WithEvents oApp As Application

Private Sub oApp_PresentationNewSlide(ByVal oSlide As slide)
    If HasMinSectionCount(ActivePresentation.SectionProperties) Then
        UpdateFooterOnSlide oSlide, GetFooterText(), GetFirstSectionIndex()
    End If
End Sub


Private Sub oApp_SlideSelectionChanged(ByVal SldRange As SlideRange)
    If isAutoUpdateSectionTimeline Then
        SlideOrderChanged
    End If
End Sub