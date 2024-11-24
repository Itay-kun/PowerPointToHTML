Attribute VB_Name = "ShapeOverlapFunctions"
Function IsShapeOverlapping(currentShape As Shape, shapes As Collection) As Boolean
    Dim nextShape As Shape
    Dim currentIndex As Long
    
    ' Find current shape index
    For i = 1 To shapes.Count
        If shapes(i).Id = currentShape.Id Then
            currentIndex = i
            Exit For
        End If
    Next i
    
    ' Check if there's a next shape to compare
    If currentIndex < shapes.Count Then
        Set nextShape = shapes(currentIndex + 1)
        
        ' Check for overlap by comparing bounds
        If ShapesOverlap(currentShape, nextShape) Then
            IsShapeOverlapping = True
            Exit Function
        End If
    End If
    
    IsShapeOverlapping = False
End Function
Function IsNextShapeOverlapping(currentShape As Shape, shapes As Collection) As Boolean
    Dim nextShape As Shape
    Dim currentIndex As Long
    
    ' Find current shape index
    For i = 1 To shapes.Count
        If shapes(i).Id = currentShape.Id Then
            currentIndex = i
            Exit For
        End If
    Next i
    
    ' Check if there's a next shape to compare
    If currentIndex < shapes.Count Then
        Set nextShape = shapes(currentIndex + 1)
        
        ' Check for overlap between current shape and next shape
        IsNextShapeOverlapping = ShapesOverlap(currentShape, nextShape)
    Else
        ' No next shape, so no overlap
        IsNextShapeOverlapping = False
    End If
End Function

Function ShapesOverlap(shape1 As Shape, shape2 As Shape) As Boolean
    ' Check if two shapes overlap by comparing their boundaries
    Dim rect1Left As Single
    Dim rect1Right As Single
    Dim rect1Top As Single
    Dim rect1Bottom As Single
    Dim rect2Left As Single
    Dim rect2Right As Single
    Dim rect2Top As Single
    Dim rect2Bottom As Single
    
    ' Get boundaries for first shape
    rect1Left = shape1.Left
    rect1Right = shape1.Left + shape1.width
    rect1Top = shape1.Top
    rect1Bottom = shape1.Top + shape1.height
    
    ' Get boundaries for second shape
    rect2Left = shape2.Left
    rect2Right = shape2.Left + shape2.width
    rect2Top = shape2.Top
    rect2Bottom = shape2.Top + shape2.height
    
    ' Check for overlap
    If rect1Left < rect2Right And _
       rect1Right > rect2Left And _
       rect1Top < rect2Bottom And _
       rect1Bottom > rect2Top Then
        ShapesOverlap = True
    Else
        ShapesOverlap = False
    End If
End Function

Function WrapShapeContent(content As String, shp As Shape, isOverlapping As Boolean) As String
    Dim wrappedContent As String
    
    If isOverlapping Then
        ' Create a positioned wrapper for overlapping content
        wrappedContent = "<div style='position: absolute; " & _
                        "top: " & shp.Top & "px; " & _
                        "left: " & shp.Left & "px; " & _
                        "width: " & shp.width & "px; " & _
                        "height: " & shp.height & "px; " & _
                        "z-index: " & shp.ZOrderPosition & ";'>" & _
                        vbNewLine & content & vbNewLine & "</div>"
    Else
        ' Return content as-is if not overlapping
        wrappedContent = content
    End If
    
    WrapShapeContent = wrappedContent
End Function

Function CreateBackgroundShape(shp As Shape) As String
    Dim bgColor As String
    Dim htmlShape As String
    
    ' Get background color if it exists
    If shp.Fill.Visible Then
        bgColor = RGBToHex(shp.Fill.ForeColor.rgb)
    Else
        bgColor = "#FFFFFF" ' Default to white if no fill
    End If
    
    ' Create HTML for background shape
    htmlShape = "<div style='position: absolute; " & _
                "top: " & shp.Top & "px; " & _
                "left: " & shp.Left & "px; " & _
                "width: " & shp.width & "px; " & _
                "height: " & shp.height & "px; " & _
                "background-color: " & bgColor & "; " & _
                "z-index: " & shp.ZOrderPosition & ";'></div>" & vbNewLine
    
    CreateBackgroundShape = htmlShape
End Function

