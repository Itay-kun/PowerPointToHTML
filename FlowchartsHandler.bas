Attribute VB_Name = "FlowchartsHandler"
Function ConvertFlowchartToHTML(shp As Shape) As String
    Dim htmlContent As String
    Dim connectors As Collection
    Dim shapes As Collection
    Dim flowShape As Shape
    Dim connector As Shape
    
    ' Initialize collections for shapes and connectors
    Set shapes = New Collection
    Set connectors = New Collection
    
    ' First pass: Categorize shapes and connectors
    For Each flowShape In shp.GroupItems
        If flowShape.Type = MsoConnector Then
            connectors.Add flowShape
        Else
            shapes.Add flowShape
        End If
    Next flowShape
    
    ' Start flowchart container
    htmlContent = "<div class='flowchart-container'" & GetShapeStyles(shp) & ">" & vbNewLine
    
    ' Process shapes (nodes)
    For Each flowShape In shapes
        htmlContent = htmlContent & CreateFlowchartNode(flowShape)
    Next flowShape
    
    ' Process connectors (arrows/lines)
    For Each connector In connectors
        htmlContent = htmlContent & CreateFlowchartConnector(connector)
    Next connector
    
    ' Close flowchart container
    htmlContent = htmlContent & "</div>" & vbNewLine
    
    ConvertFlowchartToHTML = htmlContent
End Function

Function CreateFlowchartNode(shp As Shape) As String
    Dim nodeHtml As String
    Dim nodeStyle As String
    Dim nodeContent As String
    
    ' Get shape styles (background, border, etc.)
    nodeStyle = GetShapeStyles(shp)
    
    ' Get text content if available
    If shp.HasTextFrame Then
        nodeContent = ConvertTextFrameToHTML(shp.TextFrame)
    Else
        nodeContent = ""
    End If
    
    ' Create node HTML with positioning
    nodeHtml = "<div class='flowchart-node' style='" & _
              "position: absolute; " & _
              "left: " & shp.Left & "px; " & _
              "top: " & shp.Top & "px; " & _
              "width: " & shp.width & "px; " & _
              "height: " & shp.height & "px;" & _
              nodeStyle & "'>" & _
              nodeContent & _
              "</div>" & vbNewLine
    
    CreateFlowchartNode = nodeHtml
End Function

Function CreateFlowchartConnector(connector As Shape) As String
    Dim connectorHtml As String
    Dim connectorStyle As String
    Dim startPoint As String
    Dim endPoint As String
    
    ' Get connector style (line color, thickness, etc.)
    connectorStyle = GetConnectorStyles(connector)
    
    ' Get start and end points
    startPoint = connector.ConnectorFormat.BeginConnected.Left & "," & _
                connector.ConnectorFormat.BeginConnected.Top
    endPoint = connector.ConnectorFormat.EndConnected.Left & "," & _
              connector.ConnectorFormat.EndConnected.Top
    
    ' Create SVG path for connector
    connectorHtml = "<svg class='flowchart-connector' " & _
                   "style='position: absolute; left: 0; top: 0; " & _
                   "width: 100%; height: 100%; pointer-events: none;'>" & _
                   "<path d='M " & startPoint & " L " & endPoint & "' " & _
                   connectorStyle & "/></svg>" & vbNewLine
    
    CreateFlowchartConnector = connectorHtml
End Function

Function GetConnectorStyles(connector As Shape) As String
    Dim styleString As String
    
    styleString = " style='"
    
    ' Line color
    If connector.Line.Visible Then
        styleString = styleString & "stroke: " & _
                     RGBToHex(connector.Line.ForeColor.rgb) & ";"
    End If
    
    ' Line width
    styleString = styleString & "stroke-width: " & _
                 connector.Line.Weight & "px;"
    
    ' Line style (dashed, dotted, etc.)
    Select Case connector.Line.DashStyle
        Case msoLineDash
            styleString = styleString & "stroke-dasharray: 5,5;"
        Case msoLineDashDot
            styleString = styleString & "stroke-dasharray: 5,3,2,3;"
        Case msoLineDashDotDot
            styleString = styleString & "stroke-dasharray: 5,3,2,3,2,3;"
    End Select
    
    ' Add arrow markers if needed
    If connector.Line.BeginArrowheadStyle <> msoArrowheadNone Or _
       connector.Line.EndArrowheadStyle <> msoArrowheadNone Then
        styleString = styleString & "marker-end: url(#arrow);"
    End If
    
    styleString = styleString & "'"
    
    GetConnectorStyles = styleString
End Function

Function IsFlowchartGroup(shp As Shape) As Boolean
    Dim groupItem As Shape
    Dim hasConnector As Boolean
    Dim hasShape As Boolean
    Dim connectorCount As Integer
    Dim shapeCount As Integer
    
    ' First check if it's a group
    If shp.Type <> msoGroup Then
        IsFlowchartGroup = False
        Exit Function
    End If
    
    ' Analyze group contents
    For Each groupItem In shp.GroupItems
        If groupItem.Type = MsoConnector Then
            hasConnector = True
            connectorCount = connectorCount + 1
        ElseIf groupItem.Type = msoAutoShape Or groupItem.Type = msoTextBox Then
            hasShape = True
            shapeCount = shapeCount + 1
        End If
    Next groupItem
    
    ' Debug output
    Debug.Print "Group analysis:"
    Debug.Print "- Number of connectors: " & connectorCount
    Debug.Print "- Number of shapes: " & shapeCount
    
    ' Criteria for a flowchart:
    ' 1. Must have at least one connector
    ' 2. Must have at least two shapes
    ' 3. Must have connected shapes (check connector connections)
    If hasConnector And hasShape And shapeCount >= 2 Then
        ' Additional check for connected shapes
        Dim isConnected As Boolean
        isConnected = False
        
        For Each groupItem In shp.GroupItems
            If groupItem.Type = MsoConnector Then
                ' Check if connector is actually connected
                If groupItem.ConnectorFormat.BeginConnected And _
                   groupItem.ConnectorFormat.EndConnected Then
                    isConnected = True
                    Exit For
                End If
            End If
        Next groupItem
        
        IsFlowchartGroup = isConnected
        
        Debug.Print "- Has connected shapes: " & isConnected
        Debug.Print "=> Is Flowchart: " & isConnected
    Else
        IsFlowchartGroup = False
        Debug.Print "=> Is Flowchart: False (insufficient elements)"
    End If
End Function

Function GetFlowchartShapeType(shp As Shape) As String
    ' This function identifies common flowchart shapes
    Select Case shp.AutoShapeType
        ' Process/Action shapes
        Case msoShapeRectangle
            GetFlowchartShapeType = "process"
        Case msoShapeFlowchartProcess
            GetFlowchartShapeType = "process"
            
        ' Decision shapes
        Case msoShapeFlowchartDecision
            GetFlowchartShapeType = "decision"
        Case msoShapeDiamond
            GetFlowchartShapeType = "decision"
            
        ' Terminal shapes (Start/End)
        Case msoShapeFlowchartTerminator
            GetFlowchartShapeType = "terminal"
        Case msoShapeOval
            GetFlowchartShapeType = "terminal"
            
        ' Input/Output shapes
        Case msoShapeFlowchartData
            GetFlowchartShapeType = "data"
        Case msoShapeParallelogram
            GetFlowchartShapeType = "data"
            
        ' Document shapes
        Case msoShapeFlowchartDocument
            GetFlowchartShapeType = "document"
            
        ' Preparation shapes
        Case msoShapeFlowchartPreparation
            GetFlowchartShapeType = "preparation"
            
        ' Default case
        Case Else
            GetFlowchartShapeType = "generic"
    End Select
End Function

Function GetShapeConnectionPoints(shp As Shape) As Collection
    Dim points As New Collection
    Dim point As Dictionary
    
    ' Top connection point
    Set point = New Dictionary
    point.Add "x", shp.Left + (shp.width / 2)
    point.Add "y", shp.Top
    point.Add "position", "top"
    points.Add point
    
    ' Right connection point
    Set point = New Dictionary
    point.Add "x", shp.Left + shp.width
    point.Add "y", shp.Top + (shp.height / 2)
    point.Add "position", "right"
    points.Add point
    
    ' Bottom connection point
    Set point = New Dictionary
    point.Add "x", shp.Left + (shp.width / 2)
    point.Add "y", shp.Top + shp.height
    point.Add "position", "bottom"
    points.Add point
    
    ' Left connection point
    Set point = New Dictionary
    point.Add "x", shp.Left
    point.Add "y", shp.Top + (shp.height / 2)
    point.Add "position", "left"
    points.Add point
    
    Set GetShapeConnectionPoints = points
End Function

