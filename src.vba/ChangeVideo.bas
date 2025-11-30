Option Explicit


Sub ReplaceVideo()
    Dim oldVideo As Shape
    Dim oShape As Shape
    Dim fileDialog As fileDialog
    Dim selectedFile As String
    Dim animationSettings As Effect
    Dim i As Integer

    ' Check if a shape is selected and if it is a media object
    If ActiveWindow.Selection.Type <> ppSelectionShapes Then
        MsgBox "No video selected.", vbExclamation
        Exit Sub
    End If
    
    If ActiveWindow.Selection.ShapeRange.Count <> 1 Then
        MsgBox "Must select only one video.", vbExclamation
        Exit Sub
    End If

    ' Get selected item in group (if applicable)
    If ActiveWindow.Selection.HasChildShapeRange Then
        For Each oShape In ActiveWindow.Selection.ChildShapeRange
            If oShape.Type = msoMedia Then
                If oldVideo Is Nothing Then
                    Set oldVideo = oShape
                Else
                    MsgBox "Only one video may be selected", vbExclamation
                    Exit Sub
                End If
            End If
        Next oShape
    ' Get normal selection
    Else
        For Each oShape In ActiveWindow.Selection.ShapeRange
            If oShape.Type = msoMedia Then
                If oldVideo Is Nothing Then
                    Set oldVideo = oShape
                Else
                    MsgBox "Only one video may be selected", vbExclamation
                    Exit Sub
                End If
            End If
        Next oShape
    End If
    
    If oldVideo Is Nothing Then
        MsgBox "No video selected", vbExclamation
        Exit Sub
    End If

    ' Open file dialog to select a new video file
    Set fileDialog = Application.fileDialog(msoFileDialogFilePicker)
    With fileDialog
        .Filters.Add "Video Files", "*.mp4;*.avi;*.wmv;*.mov", 1
        .Title = "Select a New Video File"
        If .Show = -1 Then
            selectedFile = .SelectedItems(1)
        Else
            ' Silence this because it gets annoying otherwise
            ' MsgBox "No file selected.", vbExclamation
            Exit Sub
        End If
    End With
    
    ReplaceVideoItem oldVideo, selectedFile
End Sub


Sub ReplaceVideoItem(oldVideo As Shape, newVideoPath As String)
    Dim newVideo As Shape
    Dim oSlide As slide
    
    Dim Effect As Effect
    Dim i As Integer, j As Integer
    
    ' Get the slide containing the selected video
    Set oSlide = oldVideo.Parent

    ' Insert the new video
    Set newVideo = oSlide.Shapes.AddMediaObject2( _
        newVideoPath, _
        LinkToFile:=msoFalse, _
        SaveWithDocument:=msoTrue, _
        Left:=oldVideo.Left, _
        Top:=oldVideo.Top)
        
    newVideo.name = oldVideo.name

    ' Resize to same width
    newVideo.LockAspectRatio = True
    newVideo.Width = oldVideo.Width
    newVideo.Top = oldVideo.Top + (oldVideo.Height - newVideo.Height) / 2
    newVideo.LockAspectRatio = oldVideo.LockAspectRatio

    ' Copy style from old video to new video
    newVideo.Line.Visible = oldVideo.Line.Visible
    If newVideo.Line.Visible = msoTrue Then
        newVideo.Line.ForeColor.rgb = oldVideo.Line.ForeColor.rgb
        newVideo.Line.Weight = oldVideo.Line.Weight
        newVideo.Line.Style = oldVideo.Line.Style
    End If
    newVideo.Fill.Visible = oldVideo.Fill.Visible
    newVideo.Fill.ForeColor.rgb = oldVideo.Fill.ForeColor.rgb
    ' Set Z-Order
    SetZIndex newVideo, oldVideo.ZOrderPosition
    
    ' Remove any default animations applied to the object
    DeleteAnimations newVideo
    If oldVideo.Child Then
        AddToGroup.AddToGroup oldVideo.ParentGroup, newVideo, oldVideo
    Else:
        TransferAnimations oldVideo, newVideo
    End If
    
    ' Remove the old video
    oldVideo.Delete
End Sub