Attribute VB_Name = "CategoriesManager"
Function GetCategoryFromPath(categoryPath As String) As String
    Dim parts() As String
    Dim lastPart As String
    
    ' Split path by backslash
    parts = Split(categoryPath, "\")
    
    ' Get last non-empty part
    For i = UBound(parts) To 0 Step -1
        If Trim(parts(i)) <> "" Then
            lastPart = parts(i)
            Exit For
        End If
    Next i
    
    GetCategoryFromPath = lastPart
End Function

Function GetSlideCategory(sld As Slide) As String
    Dim shp As Shape
    
    Debug.Print "Total shapes on slide: " & sld.shapes.Count
    
    ' Look for shapes on the right side of the slide
    For Each shp In sld.shapes
        Debug.Print "----------------------------------------"
        Debug.Print "Shape name: " & shp.Name
        Debug.Print "Shape type: " & shp.Type
        Debug.Print "Position: Left=" & shp.Left & ", Top=" & shp.Top
        
        If shp.HasTextFrame Then
            Debug.Print "Has text frame: Yes"
            Debug.Print "Text content: " & shp.TextFrame.textRange.text
            Debug.Print "Text orientation: " & shp.TextFrame.Orientation
            
            ' Check if the shape is on the right side of the slide
            If shp.Left > (ActivePresentation.PageSetup.SlideWidth * 0.8) Then
                Debug.Print "Is on right side: Yes"
                
                ' Check for vertical text orientation
                If shp.TextFrame.Orientation = msoTextOrientationVerticalFarEast Or _
                   shp.TextFrame.Orientation = msoTextOrientationUpward Or _
                   shp.TextFrame.Orientation = msoTextOrientationDownward Then
                    Debug.Print "Has vertical text: Yes"
                    Debug.Print "FOUND CATEGORY: " & shp.TextFrame.textRange.text
                    
                    GetSlideCategory = CleanFolderName(Trim(shp.TextFrame.textRange.text))
                    Exit Function
                Else
                    Debug.Print "Has vertical text: No"
                End If
            Else
                Debug.Print "Is on right side: No"
            End If
        Else
            Debug.Print "Has text frame: No"
        End If
    Next shp
    
    Debug.Print "No category found, returning 'Uncategorized'"
    GetSlideCategory = "Uncategorized"
End Function
Sub TestGetSlideCategory()
    Dim currentSlide As Slide
    Dim category As String
    
    ' Get the current slide
    Set currentSlide = Application.ActiveWindow.View.Slide
    
    Debug.Print "--------------------"
    Debug.Print "Testing slide #" & currentSlide.slideNumber
    
    ' Call the function and get result
    category = GetSlideCategory(currentSlide)
    
    ' Display result in immediate window
    Debug.Print "Category found: " & category
    
    ' Also show in a message box for easier viewing
    MsgBox "Category found: " & category, vbInformation, "Category Detection Result"
End Sub

Function IsCategoryTextBox(shp As Shape) As Boolean
    ' Criteria for identifying category text boxes
    ' Adjust these based on your specific presentation format
    
    ' Must have text
    If Not shp.HasTextFrame Then Exit Function
    If Trim(shp.TextFrame.textRange.text) = "" Then Exit Function
    
    ' Check text length (adjust max length as needed)
    If Len(shp.TextFrame.textRange.text) > 50 Then Exit Function
    
    ' Check if it's positioned near the top of the slide
    If shp.Top > 150 Then Exit Function  ' Adjust threshold as needed
    
    ' Additional criteria could include:
    ' - Font size (typically larger for categories)
    ' - Text formatting (bold, specific color, etc.)
    ' - Position on slide
    ' - Specific naming convention in the shape name
    
    IsCategoryTextBox = True
End Function

Function CleanFolderName(folderName As String) As String
    Dim cleaned As String
    cleaned = folderName
    
    ' Remove invalid characters
    cleaned = Replace(cleaned, "\", "-")
    cleaned = Replace(cleaned, "/", "-")
    cleaned = Replace(cleaned, ":", "-")
    cleaned = Replace(cleaned, "*", "-")
    cleaned = Replace(cleaned, "?", "-")
    cleaned = Replace(cleaned, """", "-")
    cleaned = Replace(cleaned, "<", "-")
    cleaned = Replace(cleaned, ">", "-")
    cleaned = Replace(cleaned, "|", "-")
    
    ' Remove multiple dashes
    Do While InStr(cleaned, "--") > 0
        cleaned = Replace(cleaned, "--", "-")
    Loop
    
    ' Trim spaces and dashes from ends
    cleaned = Trim(cleaned)
    If Left(cleaned, 1) = "-" Then cleaned = Mid(cleaned, 2)
    If Right(cleaned, 1) = "-" Then cleaned = Left(cleaned, Len(cleaned) - 1)
    
    CleanFolderName = cleaned
End Function

Function EnsureCategoryFolder(baseFolder As String, category As String, FSO As Object) As String
    Dim categoryPath As String
    
    ' Remove any trailing backslash from baseFolder
    If Right(baseFolder, 1) = "\" Then
        baseFolder = Left(baseFolder, Len(baseFolder) - 1)
    End If
    
    ' Clean category name to ensure it's valid for folder creation
    category = CleanFolderName(category)
    
    ' Create the proper path
    categoryPath = baseFolder & "\" & category
    
    ' Create category folder if it doesn't exist
    If Not FSO.FolderExists(categoryPath) Then
        FSO.CreateFolder categoryPath
    End If
    
    ' Return path with trailing backslash
    EnsureCategoryFolder = categoryPath & "\"
End Function
