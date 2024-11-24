Attribute VB_Name = "MainSub"
'Function to export all slides and make navigation page
Sub ExportWithNavigation()
    ExportSlidesToHTML
    CreateNavigationPage
End Sub

' Main processing function that handles a single slide
Function ProcessSingleSlide(sld As Slide, filePath As String, imagesFolderPath As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim htmlContent As String
    Dim slideNum As Integer
    Dim adoStream As Object
    Dim FSO As Object
    Dim category As String
    Dim categoryPath As String
    Dim categoryImagesPath As String
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set adoStream = CreateObject("ADODB.Stream")
    slideNum = sld.slideNumber
    
    ' Get the category for this slide
    category = GetSlideCategory(sld)
    
    ' Create category-specific paths
    categoryPath = EnsureCategoryFolder(filePath, category, FSO)
    categoryImagesPath = EnsureCategoryFolder(imagesFolderPath, category, FSO)
    
    ' Ensure base folders exist without deleting content
    If Not EnsureFolderExists(FSO, filePath) Or _
       Not EnsureFolderExists(FSO, imagesFolderPath) Then
        Debug.Print "Error creating necessary folders"
        ProcessSingleSlide = False
        Exit Function
    End If
    
    ' Start HTML document
    htmlContent = "<!DOCTYPE html>" & vbNewLine & _
                 "<html dir='rtl' lang='he'><head>" & vbNewLine & _
                 "<meta charset='UTF-8'>" & vbNewLine & _
                 "<title>Slide " & slideNum & " - " & category & "</title>" & vbNewLine & _
                 "<style>" & vbNewLine & _
                 "@font-face {" & vbNewLine & _
                 "  font-family: 'Hebrew';" & vbNewLine & _
                 "  src: local('Arial');" & vbNewLine & _
                 "}" & vbNewLine & _
                 "body { font-family: 'Hebrew', Arial, sans-serif; margin: 1vw 1vw; direction: rtl; }" & vbNewLine & _
                 "p { margin: 10px 0; }" & vbNewLine & _
                 "table { border-collapse: collapse; width: 100vw; }" & vbNewLine & _
                 "td, th { border: 1px solid #ddd; padding: 8px; }" & vbNewLine & _
                 "ul { margin: 2% 4%; }" & vbNewLine & _
                 ".image-container { margin: 20px 0; text-align: center; }" & vbNewLine & _
                 ".slide-image { max-width: 100%; height: auto; }" & vbNewLine & _
                 ".category { font-size: 1.2em; color: #666; margin-bottom: 20px; }" & vbNewLine & _
                 "</style></head>" & vbNewLine & _
                 "<body>"
                 
                 
                 'maybe add it back later?
                 '&"<div class='category'>Category: " & category & "</div>"
                               
                 '& vbNewLine & _
                 '"<h1>Slide " & slideNum & "</h1>" & vbNewLine

    ' Process shapes hierarchically
    htmlContent = htmlContent & ProcessShapesHierarchically(sld, categoryImagesPath)
    
    ' Close HTML document
    htmlContent = htmlContent & "</body></html>"
    
    ' Save using ADODB.Stream to handle Unicode correctly
    adoStream.Open
    adoStream.Type = 2 'Text
    adoStream.Charset = "UTF-8"
    adoStream.WriteText htmlContent
    adoStream.SaveToFile categoryPath & "slide" & slideNum & ".html", 2 'Create/Overwrite
    adoStream.Close
    
    ProcessSingleSlide = True
    
    
    Exit Function

ErrorHandler:
    Debug.Print "Error processing slide " & slideNum & ": " & Err.Description
    ProcessSingleSlide = False
End Function
