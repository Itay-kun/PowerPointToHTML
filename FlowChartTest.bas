Attribute VB_Name = "FlowChartTest"
Sub DiagnoseFlowchartOnCurrentSlide()
    Dim sld As Slide
    Dim shp As Shape
    Dim FSO As Object
    Dim filePath As String
    Dim connectorCount As Integer
    Dim shapeCount As Integer
    Dim flowchartShapes As Collection
    Dim flowchartConnectors As Collection
    
    ' Initialize
    Set sld = Application.ActiveWindow.View.Slide
    Set flowchartShapes = New Collection
    Set flowchartConnectors = New Collection
    connectorCount = 0
    shapeCount = 0
    
    Debug.Print "====== Flowchart Diagnosis ======"
    Debug.Print "Analyzing slide " & sld.slideNumber
    Debug.Print "Total objects on slide: " & sld.shapes.Count
    
    ' First pass: Process connectors
    Debug.Print vbNewLine & "Processing connectors..."
    For Each shp In sld.shapes
        If shp.Type = -2 Then  ' This is a connector
            connectorCount = connectorCount + 1
            flowchartConnectors.Add shp
            Debug.Print vbNewLine & "Found connector #" & connectorCount & ":"
            Debug.Print "- Name: " & shp.Name
            Debug.Print "- Position: Left=" & shp.Left & ", Top=" & shp.Top
            Debug.Print "- Dimensions: Width=" & shp.width & ", Height=" & shp.height
        End If
    Next shp
    
    ' Second pass: Process shapes
    Debug.Print vbNewLine & "Processing shapes..."
    For Each shp In sld.shapes
        If shp.Type <> -2 Then  ' Skip connectors this time
            Select Case shp.Type
                Case msoAutoShape
                    shapeCount = shapeCount + 1
                    flowchartShapes.Add shp
                    Debug.Print vbNewLine & "Found shape #" & shapeCount & ":"
                    Debug.Print "- Type: " & shp.AutoShapeType
                    Debug.Print "- Name: " & shp.Name
                    Debug.Print "- Position: Left=" & shp.Left & ", Top=" & shp.Top
                    If shp.HasTextFrame Then
                        Debug.Print "- Text: " & CleanDebugText(shp.TextFrame.textRange.text)
                    End If
                    
                Case msoTextBox
                    shapeCount = shapeCount + 1
                    flowchartShapes.Add shp
                    Debug.Print vbNewLine & "Found textbox #" & shapeCount & ":"
                    Debug.Print "- Name: " & shp.Name
                    Debug.Print "- Position: Left=" & shp.Left & ", Top=" & shp.Top
                    If shp.HasTextFrame Then
                        Debug.Print "- Text: " & CleanDebugText(shp.TextFrame.textRange.text)
                    End If
            End Select
        End If
    Next shp
    
    Debug.Print vbNewLine & "Summary:"
    Debug.Print "- Total connectors found: " & connectorCount
    Debug.Print "- Total shapes found: " & shapeCount
    
    ' Determine if this looks like a flowchart
    If connectorCount > 0 And shapeCount >= 2 Then
        Debug.Print vbNewLine & "This slide appears to contain a flowchart structure!"
        Debug.Print "(Found " & connectorCount & " connectors and " & shapeCount & " shapes)"
        
        ' Create test HTML
        filePath = CreateTestHTML(sld, flowchartShapes, flowchartConnectors)
        MsgBox "Found flowchart structure! HTML file has been saved to:" & vbNewLine & vbNewLine & filePath, _
               vbInformation, "Flowchart Analysis Result"
    Else
        Debug.Print vbNewLine & "This slide does not appear to contain a flowchart structure."
        Debug.Print "(Found " & connectorCount & " connectors and " & shapeCount & " shapes)"
    End If
    
    Debug.Print "====== Diagnosis Complete ======"
End Sub

Function CreateTestHTML(sld As Slide, shapes As Collection, connectors As Collection) As String
    Dim FSO As Object
    Dim htmlFile As Object
    Dim filePath As String
    Dim htmlContent As String
    Dim shp As Shape
    
    Debug.Print "====== HTML File Creation ======"
    Debug.Print "PowerPoint location: " & Application.ActivePresentation.Path
    
    ' Create FileSystemObject
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    ' Create path for test file
    filePath = Application.ActivePresentation.Path & "\flowchart_test.html"
    
    ' Create HTML content
    htmlContent = "<!DOCTYPE html>" & vbNewLine & _
                 "<html dir='rtl'><head>" & vbNewLine & _
                 "<meta charset='UTF-8'>" & vbNewLine & _
                 "<title>Flowchart Test</title>" & vbNewLine & _
                 "<style>" & vbNewLine & _
                 "body { font-family: Arial, sans-serif; margin: 20px; direction: rtl; }" & vbNewLine & _
                 ".flowchart-container { position: relative; width: 1000px; height: 800px; border: 1px solid #ccc; margin: 20px; }" & vbNewLine & _
                 ".shape { position: absolute; border: 1px solid #000; padding: 10px; text-align: center; box-sizing: border-box; }" & vbNewLine & _
                 ".connector { position: absolute; height: 2px; background: black; transform-origin: 0 0; }" & vbNewLine & _
                 "</style></head>" & vbNewLine & _
                 "<body>" & vbNewLine & _
                 "<h2>Flowchart from Slide " & sld.slideNumber & "</h2>" & vbNewLine & _
                 "<div class='flowchart-container'>" & vbNewLine

    ' First add all shapes
    For Each shp In shapes
        ' Get background color if it exists
        Dim bgColor As String
        bgColor = ""
        If shp.Fill.Visible Then
            bgColor = "background-color: " & RGBToHex(shp.Fill.ForeColor.rgb) & ";"
        End If
        
        htmlContent = htmlContent & _
            "<div class='shape' id='" & CleanID(shp.Name) & "' style='left: " & shp.Left & "px; top: " & shp.Top & "px; " & _
            "width: " & shp.width & "px; height: " & shp.height & "px; " & bgColor & "'>" & _
            IIf(shp.HasTextFrame, shp.TextFrame.textRange.text, "") & "</div>" & vbNewLine
    Next shp

    ' Then add all connectors
    For Each shp In connectors
        Dim x1 As Long, y1 As Long, x2 As Long, y2 As Long
        Dim length As Double, angle As Double
        
        ' Get start and end points
        x1 = shp.Left
        y1 = shp.Top
        x2 = shp.Left + shp.width
        y2 = shp.Top + shp.height
        
        ' Calculate length and angle
        length = Sqr((x2 - x1) ^ 2 + (y2 - y1) ^ 2)
        If x2 - x1 <> 0 Then
            angle = Atn((y2 - y1) / (x2 - x1)) * 180 / 3.14159
        Else
            angle = IIf(y2 > y1, 90, -90)
        End If
        
        htmlContent = htmlContent & _
            "<div class='connector' style='left: " & x1 & "px; top: " & y1 & "px; " & _
            "width: " & length & "px; " & _
            "transform: rotate(" & angle & "deg);'></div>" & vbNewLine
    Next shp

    ' Close containers
    htmlContent = htmlContent & "</div></body></html>"
    
    ' Write to file
    Set htmlFile = FSO.CreateTextFile(filePath, True, True) ' True for overwrite, True for Unicode
    htmlFile.Write htmlContent
    htmlFile.Close
    
    Debug.Print "HTML file saved to: " & filePath
    Debug.Print "====== HTML File Creation Complete ======"
    
    CreateTestHTML = filePath
End Function

Function CleanDebugText(txt As String) As String
    ' This helps make the debug output more readable
    CleanDebugText = Replace(Replace(txt, vbCr, " "), vbLf, " ")
End Function

Function CleanID(txt As String) As String
    ' Clean shape name for use as HTML id
    CleanID = Replace(Replace(txt, " ", "_"), ".", "_")
End Function

Function IsConnector(shp As Shape) As Boolean
    ' Returns true if the shape is a connector based on both type and name
    IsConnector = (shp.Type = -2) Or (InStr(1, shp.Name, "Connector", vbTextCompare) > 0)
End Function
