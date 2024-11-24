Attribute VB_Name = "VisualProcessing"
Function GetVisuallyOrderedShapes(sld As Slide) As Collection
    Dim orderedShapes As Collection
    Dim shp As Shape
    Dim shapeInfo As Dictionary
    Dim shapeInfos() As Dictionary
    Dim i As Long
    
    ' Initialize collection
    Set orderedShapes = New Collection
    
    ' Create array to hold shape information
    ReDim shapeInfos(1 To sld.shapes.Count)
    i = 1
    
    ' Collect shape information with position data
    For Each shp In sld.shapes
        shp.ZOrder msoSendToBack
        ' Skip only slide number and footer placeholders
        If shp.Type = msoPlaceholder Then
            If shp.PlaceholderFormat.Type = ppPlaceholderSlideNumber Or _
               shp.PlaceholderFormat.Type = ppPlaceholderFooter Or _
               shp.PlaceholderFormat.Type = ppPlaceholderHeader Then
                GoTo nextShape
            End If
        End If
        
        ' Modified condition: Only skip shapes that are explicitly marked as category textboxes
        If shp.Type = msoTextBox Or shp.Type = msoPlaceholder Or shp.Type = msoAutoShape Then
            If shp.HasTextFrame Then
                If IsCategoryTextBox(shp) And _
                   shp.Left > (ActivePresentation.PageSetup.SlideWidth * 0.8) Then
                    GoTo nextShape
                End If
            End If
        End If
        
        ' Include all shapes with text, regardless of type
        ' Only skip pure rectangles without text that are likely background shapes
        If shp.Type = msoAutoShape And Not shp.HasTextFrame Then
            ' Check if it's a regular rectangle (not rounded) without text
            If shp.AutoShapeType = msoShapeRectangle And _
               Not shp.AutoShapeType = msoShapeRoundedRectangle Then
                GoTo nextShape
            End If
        End If
        
        ' Always include rounded rectangles, whether they have text or not
        If shp.AutoShapeType = msoShapeRoundedRectangle Then
            ' Include it in our collection
        End If
        
        Set shapeInfo = New Dictionary
        shapeInfo.Add "Shape", shp
        shapeInfo.Add "Top", shp.Top
        shapeInfo.Add "Left", shp.Left
        Set shapeInfos(i) = shapeInfo
        i = i + 1
nextShape:
    Next shp
    
    ' Adjust array size to account for skipped shapes
    If i <= UBound(shapeInfos) Then
        ReDim Preserve shapeInfos(1 To i - 1)
    End If
    
    ' Sort shapes by Top position (and Left for same vertical position)
    BubbleSortShapes shapeInfos
    
    ' Add sorted shapes to collection
    For i = LBound(shapeInfos) To UBound(shapeInfos)
        orderedShapes.Add shapeInfos(i)("Shape")
    Next i
    
    Set GetVisuallyOrderedShapes = orderedShapes
End Function
Function ProcessShapesVisually(sld As Slide) As String
    Dim orderedShapes As Collection
    Dim shp As Shape
    Dim htmlContent As String
    
    Set orderedShapes = GetVisuallyOrderedShapes(sld)
    htmlContent = ""
    
    For Each shp In orderedShapes
        Select Case shp.Type
            Case msoTable
                htmlContent = htmlContent & ConvertTableToHTML(shp.Table)
            Case msoTextBox, msoPlaceholder
                If shp.HasTextFrame Then
                    htmlContent = htmlContent & ConvertTextFrameToHTML(shp.TextFrame)
                End If
            Case msoPicture, msoLinkedPicture
                htmlContent = htmlContent & SaveAndLinkImage(shp, sld.slideNumber, imagesFolderPath)
            Case msoChart
                htmlContent = htmlContent & SaveAndLinkChart(shp, sld.slideNumber, imagesFolderPath)
            Case msoGraphic
                htmlContent = htmlContent & SaveAndLinkGraphic(shp, sld.slideNumber, imagesFolderPath)
        End Select
    Next shp
    
    ProcessShapesVisually = htmlContent
End Function
