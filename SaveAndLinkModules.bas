Attribute VB_Name = "SaveAndLinkModules"
Function SaveAndLinkImage(shp As Shape, slideNum As Integer, categoryImagesPath As String) As String
    Dim imageName As String
    Dim imageFormat As String
    Dim htmlImage As String
    Dim categoryName As String
    
    ' Determine image format (default to PNG)
    imageFormat = ".png"
    
    ' Get category name from path
    categoryName = GetCategoryFromPath(categoryImagesPath)
    
    ' Generate unique image name
    imageName = "slide" & slideNum & "_image" & shp.Id & imageFormat
    
    ' Export the image to category subfolder
    On Error Resume Next
    shp.Export categoryImagesPath & imageName, ppShapeFormatPNG
    On Error GoTo 0
    
    ' Create HTML for the image with path that includes category
    htmlImage = "<div class='image-container'>" & vbNewLine & _
                "<img src='../images/" & categoryName & "/" & imageName & "' " & _
                "alt='Slide " & slideNum & " Image' " & _
                "class='slide-image'>" & vbNewLine & _
                "</div>" & vbNewLine
    
    SaveAndLinkImage = htmlImage
End Function

Function SaveAndLinkChart(shp As Shape, slideNum As Integer, categoryImagesPath As String) As String
    Dim chartName As String
    Dim htmlChart As String
    Dim categoryName As String
    
    ' Get category name from path
    categoryName = GetCategoryFromPath(categoryImagesPath)
    
    ' Generate unique chart name
    chartName = "slide" & slideNum & "_chart" & shp.Id & ".png"
    
    ' Export the chart as image to category subfolder
    On Error Resume Next
    shp.Export categoryImagesPath & chartName, ppShapeFormatPNG
    On Error GoTo 0
    
    ' Create HTML for the chart with path that includes category
    htmlChart = "<div class='image-container'>" & vbNewLine & _
                "<img src='../images/" & categoryName & "/" & chartName & "' " & _
                "alt='Slide " & slideNum & " Chart' " & _
                "class='slide-image'>" & vbNewLine & _
                "</div>" & vbNewLine
    
    SaveAndLinkChart = htmlChart
End Function

Function SaveAndLinkGraphic(shp As Shape, slideNum As Integer, categoryImagesPath As String) As String
    Dim graphicName As String
    Dim htmlGraphic As String
    Dim categoryName As String
    
    ' Get category name from path
    categoryName = GetCategoryFromPath(categoryImagesPath)
    
    ' Generate unique graphic name
    graphicName = "slide" & slideNum & "_graphic" & shp.Id & ".png"
    
    ' Export the graphic as image to category subfolder
    On Error Resume Next
    shp.Export categoryImagesPath & graphicName, ppShapeFormatPNG
    On Error GoTo 0
    
    ' Create HTML for the graphic with path that includes category
    htmlGraphic = "<div class='image-container'>" & vbNewLine & _
                  "<img src='../images/" & categoryName & "/" & graphicName & "' " & _
                  "alt='Slide " & slideNum & " Graphic' " & _
                  "class='slide-image'>" & vbNewLine & _
                  "</div>" & vbNewLine
    
    SaveAndLinkGraphic = htmlGraphic
End Function

