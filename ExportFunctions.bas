Attribute VB_Name = "ExportFunctions"
Sub ExportSlidesToHTML()
    Dim ppt As Presentation
    Dim sld As Slide
    Dim filePath As String
    Dim imagesFolderPath As String
    Dim FSO As Object
    
    ' Create FileSystemObject for file operations
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    ' Get current presentation
    Set ppt = Application.ActivePresentation
    
    ' Create output folders
    filePath = ppt.Path & "\HTMLSlides\"
    imagesFolderPath = filePath & "images\"
    
    ' Use new folder creation function
    If Not EnsureFolderExists(FSO, filePath) Or _
       Not EnsureFolderExists(FSO, imagesFolderPath) Then
        MsgBox "Error creating necessary folders", vbCritical
        Exit Sub
    End If
    
    ' Process each slide
    For Each sld In ppt.Slides
        ProcessSingleSlide sld, filePath, imagesFolderPath
    Next sld
    
    MsgBox "HTML files have been created in: " & filePath
End Sub
' Helper function to process current slide
Sub ExportCurrentSlideToHTML()
    Dim ppt As Presentation
    Dim currentSlide As Slide
    Dim filePath As String
    Dim imagesFolderPath As String
    Dim FSO As Object
    
    ' Create FileSystemObject for file operations
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    ' Get current presentation and slide
    Set ppt = Application.ActivePresentation
    Set currentSlide = Application.ActiveWindow.View.Slide
    
    ' Create output folders
    filePath = ppt.Path & "\HTMLSlides\"
    imagesFolderPath = filePath & "images\"
    
    ' Create folders if they don't exist
    If Not FSO.FolderExists(filePath) Then
        FSO.CreateFolder filePath
    End If
    If Not FSO.FolderExists(imagesFolderPath) Then
        FSO.CreateFolder imagesFolderPath
    End If
    
    ' Process current slide
    ProcessSingleSlide currentSlide, filePath, imagesFolderPath
    Debug.Print filePath
    MsgBox "HTML file has been created for slide " & currentSlide.slideNumber & " in: " & filePath
End Sub

' Helper function to process a specific slide number
Sub ExportSlideNumberToHTML(slideNumber As Integer)
    Dim ppt As Presentation
    Dim targetSlide As Slide
    Dim filePath As String
    Dim imagesFolderPath As String
    Dim FSO As Object
    
    ' Create FileSystemObject for file operations
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    ' Get current presentation
    Set ppt = Application.ActivePresentation
    
    ' Check if slide number exists
    If slideNumber <= 0 Or slideNumber > ppt.Slides.Count Then
        MsgBox "Invalid slide number. Please enter a number between 1 and " & ppt.Slides.Count
        Exit Sub
    End If
    
    Set targetSlide = ppt.Slides(slideNumber)
    
    ' Create output folders
    filePath = ppt.Path & "\HTMLSlides\"
    imagesFolderPath = filePath & "images\"
    
    ' Create folders if they don't exist
    If Not FSO.FolderExists(filePath) Then
        FSO.CreateFolder filePath
    End If
    If Not FSO.FolderExists(imagesFolderPath) Then
        FSO.CreateFolder imagesFolderPath
    End If
    
    ' Process target slide
    ProcessSingleSlide targetSlide, filePath, imagesFolderPath
    
    MsgBox "HTML file has been created for slide " & slideNumber & " in: " & filePath
End Sub

' Helper function to process a range of slides
Sub ExportSlideRangeToHTML(startSlide As Integer, endSlide As Integer)
    Dim ppt As Presentation
    Dim i As Integer
    Dim filePath As String
    Dim imagesFolderPath As String
    Dim FSO As Object
    
    ' Create FileSystemObject for file operations
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    ' Get current presentation
    Set ppt = Application.ActivePresentation
    
    ' Validate slide range
    If startSlide <= 0 Or endSlide > ppt.Slides.Count Or startSlide > endSlide Then
        MsgBox "Invalid slide range. Please enter numbers between 1 and " & ppt.Slides.Count
        Exit Sub
    End If
    
    ' Create output folders
    filePath = ppt.Path & "\HTMLSlides\"
    imagesFolderPath = filePath & "images\"
    
    ' Create folders if they don't exist
    If Not FSO.FolderExists(filePath) Then
        FSO.CreateFolder filePath
    End If
    If Not FSO.FolderExists(imagesFolderPath) Then
        FSO.CreateFolder imagesFolderPath
    End If
    
    ' Process each slide in range
    For i = startSlide To endSlide
        ProcessSingleSlide ppt.Slides(i), filePath, imagesFolderPath
    Next i
    
    Debug.Print filePath
    MsgBox "HTML files have been created for slides " & startSlide & " to " & endSlide & " in: " & filePath
End Sub
