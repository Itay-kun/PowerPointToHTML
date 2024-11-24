Attribute VB_Name = "debugModule"
Sub DebugTableOnCurrentSlide()
    Dim sld As Slide
    Dim shp As Shape
    
    ' Get current slide
    Set sld = Application.ActiveWindow.View.Slide
    
    Debug.Print "Analyzing slide " & sld.slideNumber
    Debug.Print "Total shapes on slide: " & sld.shapes.Count
    
    ' Look for tables
    For Each shp In sld.shapes
        Debug.Print "----------------------------------------"
        Debug.Print "Shape name: " & shp.Name
        Debug.Print "Shape type: " & shp.Type
        
        If shp.Type = msoTable Then
            Debug.Print "TABLE FOUND!"
            Debug.Print "Table dimensions: " & shp.Table.Rows.Count & " rows x " & shp.Table.Columns.Count & " columns"
            Debug.Print "Table position: Left=" & shp.Left & ", Top=" & shp.Top
            
            ' Print first cell content
            Debug.Print "First cell content: " & shp.Table.cell(1, 1).Shape.TextFrame.textRange.text
        End If
    Next shp
    
    Debug.Print "Analysis complete"
End Sub

Sub DiagnoseTableOnCurrentSlide()
    Dim sld As Slide
    Dim shp As Shape
    
    ' Get current slide
    Set sld = Application.ActiveWindow.View.Slide
    
    Debug.Print "====== Table Diagnosis ======"
    Debug.Print "Analyzing slide " & sld.slideNumber
    Debug.Print "Total shapes on slide: " & sld.shapes.Count
    
    ' Look for tables
    For Each shp In sld.shapes
        If shp.Type = msoTable Then
            Debug.Print vbNewLine & "TABLE FOUND!"
            Debug.Print "Table position: Left=" & shp.Left & ", Top=" & shp.Top
            Debug.Print "Table dimensions: " & shp.Table.Rows.Count & " rows x " & shp.Table.Columns.Count & " columns"
            
            ' Try to access each cell
            Dim row As Long, col As Long
            For row = 1 To shp.Table.Rows.Count
                For col = 1 To shp.Table.Columns.Count
                    On Error Resume Next
                    Dim cellText As String
                    cellText = shp.Table.cell(row, col).Shape.TextFrame.textRange.text
                    If Err.Number <> 0 Then
                        Debug.Print "ERROR accessing cell [" & row & "," & col & "]: " & Err.Description
                        Err.Clear
                    Else
                        Debug.Print "Cell [" & row & "," & col & "] = " & cellText
                    End If
                    On Error GoTo 0
                Next col
            Next row
        End If
    Next shp
    
    Debug.Print "====== Diagnosis Complete ======"
End Sub
