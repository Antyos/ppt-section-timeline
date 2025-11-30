' Export to PowerPoint Add-In
Sub ExportToPPAM()
    Dim pptPresentation As Presentation
    Dim outDir As String
    Dim baseName As String
    Dim ppamPath As String
    Dim oAddIn As AddIn
    Dim isAddInEnabled As Boolean
    Dim addInPath As String
    
    ' Get the active presentation
    Set pptPresentation = ActivePresentation

    ' Get the current full path of the active presentation
    outDir = pptPresentation.path
    baseName = Stem(pptPresentation.name)
    ' Construct the new file name with .ppam extension
    ppamPath = outDir & "\" & baseName & ".ppam"
    
    ' Disable addin so we can update the compiled file addin without powerpoint reserving the file
    isAddInEnabled = False
    On Error GoTo AddInError
    Set oAddIn = Application.AddIns(baseName)
    If oAddIn.Loaded = True Then
        addInPath = oAddIn.FullName
        Application.AddIns.Remove baseName
        isAddInEnabled = True
    End If
    On Error GoTo 0
AddInError:
    
    ' Delete existing file
    If Dir(ppamPath) <> "" Then Kill ppamPath

    ' Save the active presentation as a PPAM file in the same location.
    ' Specifying 'ppSaveAsAddIn' saves it as a .ppa file instead of '.ppam',
    ' so we give it the file extension above and let it infer the file type.
    pptPresentation.SaveAs ppamPath
    
    ' Re-enable addin. if isAddInEnabled is True, we already know the addin exists
    If isAddInEnabled = True Then
        Application.AddIns.Add addInPath
        Application.AddIns(baseName).Loaded = True
    End If

    MsgBox "Presentation has been saved as a PPAM file at " & ppamPath, vbInformation
End Sub


Function Stem(path As String)
    Stem = Left(path, InStrRev(path, ".") - 1)
End Function