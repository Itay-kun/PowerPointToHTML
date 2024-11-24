Attribute VB_Name = "HierarchicalProcessing"
Function GetHierarchicalShapes(sld As Slide) As Collection
    Dim hierarchicalShapes As Collection
    Dim shp As Shape
    Dim shapeGroups As Collection
    Dim currentGroup As Collection
    Dim lastTop As Double
    Dim VERTICAL_THRESHOLD As Double
    
    ' Initialize collections
    Set hierarchicalShapes = New Collection
    Set shapeGroups = New Collection
    Set currentGroup = New Collection
    
    ' Set threshold for considering shapes as part of the same group (in points)
    VERTICAL_THRESHOLD = 20
    
    ' First get shapes in vertical order
    Set orderedShapes = GetVisuallyOrderedShapes(sld)
    
    ' Group shapes that are visually related
    lastTop = -1
    For Each shp In orderedShapes
        If lastTop = -1 Then
            ' First shape
            currentGroup.Add shp
        ElseIf Abs(shp.Top - lastTop) <= VERTICAL_THRESHOLD Then
            ' Shape is close enough vertically to be part of current group
            currentGroup.Add shp
        Else
            ' Start new group
            If currentGroup.Count > 0 Then
                shapeGroups.Add currentGroup
            End If
            Set currentGroup = New Collection
            currentGroup.Add shp
        End If
        
        lastTop = shp.Top
    Next shp
    
    ' Add last group if not empty
    If currentGroup.Count > 0 Then
        shapeGroups.Add currentGroup
    End If
    
    ' Process each group
    Dim grp As Collection
    For Each grp In shapeGroups
        ' Sort shapes within group by left position
        ProcessGroupHierarchy grp, hierarchicalShapes
    Next grp
    
    Set GetHierarchicalShapes = hierarchicalShapes
End Function
Sub ProcessGroupHierarchy(grp As Collection, hierarchicalShapes As Collection)
    Dim shp As Shape
    Dim sortedShapes() As Shape
    Dim i As Long
    
    ' Convert group to array for sorting
    ReDim sortedShapes(1 To grp.Count)
    i = 1
    For Each shp In grp
        Set sortedShapes(i) = shp
        i = i + 1
    Next shp
    
    ' Sort by left position
    Dim j As Long, tempShape As Shape
    For i = LBound(sortedShapes) To UBound(sortedShapes) - 1
        For j = i + 1 To UBound(sortedShapes)
            If sortedShapes(i).Left > sortedShapes(j).Left Then
                Set tempShape = sortedShapes(i)
                Set sortedShapes(i) = sortedShapes(j)
                Set sortedShapes(j) = tempShape
            End If
        Next j
    Next i
    
    ' Add sorted shapes to final collection
    For i = LBound(sortedShapes) To UBound(sortedShapes)
        hierarchicalShapes.Add sortedShapes(i)
    Next i
End Sub
Function ProcessShapesHierarchically(sld As Slide, imagesFolderPath As String) As String
    Dim hierarchicalShapes As Collection
    Dim shp As Shape
    Dim htmlContent As String
    Dim isOverlapping As Boolean
    Dim overlapContainer As String
    
    Set hierarchicalShapes = GetHierarchicalShapes(sld)
    htmlContent = ""
    isOverlapping = False
    
    For Each shp In hierarchicalShapes
        ' Check if this shape overlaps with the next one
        isOverlapping = IsShapeOverlapping(shp, hierarchicalShapes)
        
        If isOverlapping And Not isOverlapContainer Then
            ' Start a new container for overlapping shapes
            htmlContent = htmlContent & "<div class='shape-container' style='position: relative;'>" & vbNewLine
            isOverlapContainer = True
        End If
        
        ' Process the shape based on its type
        Select Case shp.Type
            Case msoTable
                htmlContent = htmlContent & WrapShapeContent(ConvertTableToHTML(shp.Table), shp, isOverlapping)
            Case msoTextBox, msoPlaceholder
                If shp.HasTextFrame Then
                    htmlContent = htmlContent & WrapShapeContent(ConvertTextFrameToHTML(shp.TextFrame), shp, isOverlapping)
                End If
            Case msoAutoShape
                If shp.AutoShapeType = msoShapeRectangle Then
                    ' Handle rectangle shapes
                    If shp.HasTextFrame Then
                        ' Rectangle with text
                        htmlContent = htmlContent & WrapShapeContent(ConvertTextFrameToHTML(shp.TextFrame), shp, isOverlapping)
                    Else
                        ' Rectangle without text (background shape)
                        htmlContent = htmlContent & CreateBackgroundShape(shp)
                    End If
                End If
            Case msoPicture, msoLinkedPicture
                htmlContent = htmlContent & WrapShapeContent(SaveAndLinkImage(shp, sld.slideNumber, imagesFolderPath), shp, isOverlapping)
            Case msoChart
                htmlContent = htmlContent & WrapShapeContent(SaveAndLinkChart(shp, sld.slideNumber, imagesFolderPath), shp, isOverlapping)
            Case msoGraphic
                htmlContent = htmlContent & WrapShapeContent(SaveAndLinkGraphic(shp, sld.slideNumber, imagesFolderPath), shp, isOverlapping)
        End Select
        
        ' Check if we need to close the overlap container
        If isOverlapContainer And Not IsNextShapeOverlapping(shp, hierarchicalShapes) Then
            htmlContent = htmlContent & "</div>" & vbNewLine
            isOverlapContainer = False
        End If
    Next shp
    
    ' Ensure we close any remaining container
    If isOverlapContainer Then
        htmlContent = htmlContent & "</div>" & vbNewLine
    End If
    
    ProcessShapesHierarchically = htmlContent
End Function
