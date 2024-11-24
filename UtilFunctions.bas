Attribute VB_Name = "UtilFunctions"
Function IsHebrewText(txt As String) As Boolean
    Dim i As Long
    Dim ch As Long
    
    For i = 1 To Len(txt)
        ch = AscW(Mid(txt, i, 1))
        ' Check for Hebrew character range (1424-1535)
        If ch >= 1424 And ch <= 1535 Then
            IsHebrewText = True
            Exit Function
        End If
    Next i
    
    IsHebrewText = False
End Function
Function GetTextAlignment(alignment As PpParagraphAlignment) As String
    Select Case alignment
        Case ppAlignLeft
            GetTextAlignment = "left"
        Case ppAlignCenter
            GetTextAlignment = "center"
        Case ppAlignRight
            GetTextAlignment = "right"
        Case ppAlignJustify
            GetTextAlignment = "justify"
        Case Else
            GetTextAlignment = "left"
    End Select
End Function
Function GetTextStyle(textRange As textRange) As String
    Dim styleString As String
    Dim font As font
    
    Set font = textRange.font
    styleString = " style='"
    
    ' Font size (convert points to pixels - multiply by 1.33)
    If font.Size > 0 Then
        styleString = styleString & "font-size: " & Format(font.Size * 1.33, "0.00") & "px; "
    End If
    
    ' Font color
    If font.Color.rgb <> 0 Then
        styleString = styleString & "color: " & RGBToHex(font.Color.rgb) & "; "
    End If
    
    ' Font family
    If font.Name <> "" Then
        styleString = styleString & "font-family: '" & font.Name & "', Arial, sans-serif; "
    End If
    
    ' Bold
    If font.Bold Then
        styleString = styleString & "font-weight: bold; "
    End If
    
    ' Italic
    If font.Italic Then
        styleString = styleString & "font-style: italic; "
    End If
    
    ' Underline
    If font.Underline Then
        styleString = styleString & "text-decoration: underline; "
    End If
    
    ' Only return style attribute if there are actual styles
    If styleString = " style='" Then
        GetTextStyle = ""
    Else
        GetTextStyle = styleString & "'"
    End If
End Function
Function RGBToHex(rgbColor As Long) As String
    Dim rValue As Integer
    Dim gValue As Integer
    Dim bValue As Integer
    
    rValue = rgbColor Mod 256
    gValue = (rgbColor \ 256) Mod 256
    bValue = (rgbColor \ 65536) Mod 256
    
    RGBToHex = "#" & Right("0" & Hex(rValue), 2) & _
                     Right("0" & Hex(gValue), 2) & _
                     Right("0" & Hex(bValue), 2)
End Function

Function FormatTextWithStyles(textRange As textRange) As String
    Dim result As String
    Dim char As textRange
    Dim currentText As String
    Dim i As Long
    
    result = ""
    
    ' Process each character to capture format changes
    For i = 1 To textRange.length
        Set char = textRange.Characters(i, 1)
        currentText = CleanHTML(char.text)
        
        ' Skip if empty
        If Len(currentText) = 0 Then
            GoTo NextChar
        End If
        
        ' Apply formatting
        currentText = ApplyTextFormatting(char, currentText)
        
        result = result & currentText
NextChar:
    Next i
    
    FormatTextWithStyles = result
End Function
Function ApplyTextFormatting(textRange As textRange, text As String) As String
    Dim result As String
    Dim styleString As String
    Dim classes As String
    Dim isHebrew As Boolean
    
    result = text
    styleString = ""
    classes = ""
    isHebrew = IsHebrewText(text)
    
    ' Add direction class
    classes = classes & IIf(isHebrew, " rtl", " ltr")
    
    ' Font family
    If textRange.font.Name <> "" Then
        styleString = styleString & "font-family: '" & textRange.font.Name & "';"
        classes = classes & " custom-font"
    End If
    
    ' Font size
    If textRange.font.Size > 0 Then
        styleString = styleString & "font-size: " & CStr(textRange.font.Size * 1.33) & "px;"
    End If
    
    ' Font color
    If textRange.font.Color.rgb <> 0 Then
        styleString = styleString & "color: " & RGBToHex(textRange.font.Color.rgb) & ";"
    End If
    
    ' Bold
    If textRange.font.Bold Then
        classes = classes & " bold"
    End If
    
    ' Italic
    If textRange.font.Italic Then
        classes = classes & " italic"
    End If
    
    ' Underline
    If textRange.font.Underline Then
        classes = classes & " underline"
    End If
    
    ' Build the span tag with styles and classes
    If styleString <> "" Or classes <> "" Then
        result = "<span" & _
                IIf(classes <> "", " class='mixed" & classes & "'", "") & _
                IIf(styleString <> "", " style='" & styleString & "'", "") & _
                ">" & result & "</span>"
    End If
    
    ApplyTextFormatting = result
End Function
Function CleanHTML(txt As String) As String
    ' Clean up text for HTML without changing character encoding
    Dim cleanText As String
    cleanText = txt
    
    ' Handle newlines properly - do this first
    cleanText = Replace(cleanText, vbCr, vbCrLf)
    cleanText = Replace(cleanText, vbLf, "")
    
    ' Replace special HTML characters, but preserve <br> tags
    ' First, temporarily replace <br> with a unique marker
    cleanText = Replace(cleanText, "<br>", "[[BR_MARKER]]")
    
    ' Replace special HTML characters
    cleanText = Replace(cleanText, "&", "&amp;")
    cleanText = Replace(cleanText, "<", "&lt;")
    cleanText = Replace(cleanText, ">", "&gt;")
    cleanText = Replace(cleanText, """", "&quot;")
    
    cleanText = Replace(cleanText, ChrW(8593), "&#8593;")        ' Direct up arrow if present
    cleanText = Replace(cleanText, ChrW(8595), "&#8595;")        ' Direct down arrow if present
    
    ' Restore <br> tags
    cleanText = Replace(cleanText, "[[BR_MARKER]]", "<br>")
    
    ' Handle PowerPoint's line breaks
    cleanText = Replace(cleanText, vbCrLf, "<br>")
    
    CleanHTML = cleanText
End Function
Function GetCellBackgroundColor(cell As cell) As String
    Dim fillColor As Long
    Dim styleAttr As String
    
    ' Check if cell has fill
    If cell.Shape.Fill.Visible Then
        ' Get the RGB color
        fillColor = cell.Shape.Fill.ForeColor.rgb
        
        ' Convert RGB to hex and create style attribute
        styleAttr = " style='background-color: " & RGBToHex(fillColor) & ";'"
        GetCellBackgroundColor = styleAttr
    Else
        ' No background color
        GetCellBackgroundColor = ""
    End If
End Function
Function ProcessFormattedText(textRange As textRange) As String
    Dim result As String
    Dim runText As String
    Dim currentRun As textRange
    Dim lastFormatting As String
    Dim currentFormatting As String
    
    result = ""
    lastFormatting = ""
    
    ' Process each run of text (a run is a segment with consistent formatting)
    For Each currentRun In textRange.Runs
        runText = CleanHTML(currentRun.text)
        
        ' Skip empty runs
        If Len(runText) = 0 Then
            GoTo NextRun
        End If
        
        ' Get formatting for current run
        currentFormatting = GetRunFormatting(currentRun)
        
        ' If formatting changed, close previous span and start new one
        If currentFormatting <> lastFormatting Then
            ' Close previous formatting if exists
            If lastFormatting <> "" Then
                result = result & "</span>"
            End If
            
            ' Start new formatting if needed
            If currentFormatting <> "" Then
                result = result & "<span" & currentFormatting & ">"
            End If
        End If
        
        ' Add the text
        result = result & runText
        
        ' Update last formatting
        lastFormatting = currentFormatting
        
NextRun:
    Next currentRun
    
    ' Close final span if needed
    If lastFormatting <> "" Then
        result = result & "</span>"
    End If
    
    Debug.Print result
    ProcessFormattedText = result
End Function

Sub BubbleSortShapes(shapeInfos() As Dictionary)
    Dim i As Long, j As Long
    Dim tempDict As Dictionary
    
    For i = LBound(shapeInfos) To UBound(shapeInfos) - 1
        For j = i + 1 To UBound(shapeInfos)
            ' Compare Top positions first
            If shapeInfos(i)("Top") > shapeInfos(j)("Top") Then
                ' Swap
                Set tempDict = shapeInfos(i)
                Set shapeInfos(i) = shapeInfos(j)
                Set shapeInfos(j) = tempDict
            ElseIf shapeInfos(i)("Top") = shapeInfos(j)("Top") Then
                ' If same vertical position, order by Left position
                If shapeInfos(i)("Left") > shapeInfos(j)("Left") Then
                    ' Swap
                    Set tempDict = shapeInfos(i)
                    Set shapeInfos(i) = shapeInfos(j)
                    Set shapeInfos(j) = tempDict
                End If
            End If
        Next j
    Next i
End Sub
Function GetTextDirectionAttribute(txt As String) As String
    ' Return appropriate dir attribute based on text content
    If IsHebrewText(txt) Then
        GetTextDirectionAttribute = " dir='rtl'"
    Else
        ' For English or other LTR text
        GetTextDirectionAttribute = " dir='ltr'"
    End If
End Function

Function GetShapeStyles(shp As Shape) As String
    Dim styleString As String
    styleString = " style='"
    
    ' Add background color if shape has fill
    If shp.Fill.Visible Then
        If shp.Fill.ForeColor.rgb <> 0 Then
            styleString = styleString & "background-color: " & RGBToHex(shp.Fill.ForeColor.rgb) & "; "
        End If
        
        ' Handle transparency
        If shp.Fill.Transparency > 0 Then
            styleString = styleString & "opacity: " & Format(1 - shp.Fill.Transparency, "0.00") & "; "
        End If
    End If
    
    ' Add border if shape has line
    If shp.Line.Visible Then
        ' Border color
        If shp.Line.ForeColor.rgb <> 0 Then
            styleString = styleString & "border-color: " & RGBToHex(shp.Line.ForeColor.rgb) & "; "
        End If
        
        ' Border width
        If shp.Line.Weight > 0 Then
            styleString = styleString & "border-width: " & shp.Line.Weight & "px; "
            styleString = styleString & "border-style: solid; "
        End If
        
        ' Border transparency
        If shp.Line.Transparency > 0 Then
            styleString = styleString & "border-opacity: " & Format(1 - shp.Line.Transparency, "0.00") & "; "
        End If
        
        ' Handle different line styles
        Select Case shp.Line.DashStyle
            Case msoLineDash
                styleString = styleString & "border-style: dashed; "
            Case msoLineDashDot
                styleString = styleString & "border-style: dashed; "
            Case msoLineDashDotDot
                styleString = styleString & "border-style: dotted; "
            Case msoLineRoundDot
                styleString = styleString & "border-style: dotted; "
            Case msoLineSolid
                styleString = styleString & "border-style: solid; "
        End Select
    End If
    
    ' Add border radius for rounded rectangle shapes
    If shp.Type = msoAutoShape Then
        If shp.AutoShapeType = msoShapeRoundedRectangle Then
            styleString = styleString & "border-radius: 15px; "  ' Increased from 10px to 15px for better visibility
            styleString = styleString & "padding: 10px; "        ' Added padding for better text spacing
        End If
    End If
    
    ' Only return style attribute if there are actual styles
    If styleString = " style='" Then
        GetShapeStyles = ""
    Else
        GetShapeStyles = styleString & "'"
    End If
End Function
Function EnsureFolderExists(FSO As Object, folderPath As String) As Boolean
    On Error Resume Next
    
    ' Check if folder exists
    If Not FSO.FolderExists(folderPath) Then
        ' Create folder if it doesn't exist
        FSO.CreateFolder folderPath
    End If
    
    ' Return True if folder exists/was created successfully
    EnsureFolderExists = FSO.FolderExists(folderPath)
    
    On Error GoTo 0
End Function

Function GetSpanAttributes(cell As cell, row As Integer, col As Integer, mergeTracker As Collection) As String
    Dim spanAttr As String
    Dim mergeInfo As Variant
    
    spanAttr = ""
    
    ' Only proceed if this is a merged cell
    If cell.Merged Then
        On Error Resume Next
        mergeInfo = mergeTracker.item(CStr(row) & "_" & CStr(col))
        On Error GoTo 0
        
        If Not IsEmpty(mergeInfo) Then
            ' Split merge info into rowspan,colspan
            Dim spans() As String
            spans = Split(mergeInfo, ",")
            
            ' Add rowspan if > 1
            If CLng(spans(0)) > 1 Then
                spanAttr = spanAttr & " rowspan='" & spans(0) & "'"
            End If
            
            ' Add colspan if > 1
            If CLng(spans(1)) > 1 Then
                spanAttr = spanAttr & " colspan='" & spans(1) & "'"
            End If
        End If
    End If
    
    GetSpanAttributes = spanAttr
End Function
Function GetCellText(cell As cell) As String
    ' Safely extract cell text
    On Error Resume Next
    If Not cell.Shape Is Nothing Then
        If cell.Shape.HasTextFrame Then
            GetCellText = cell.Shape.TextFrame.textRange.text
        End If
    End If
    On Error GoTo 0
    
    ' Clean up the text
    GetCellText = Trim(GetCellText)
End Function
Function GetCellStyles(cell As cell) As String
    Dim styleStr As String
    styleStr = "border: 1px solid #ddd; padding: 8px;"
    
    ' Add background color if present
    If cell.Shape.Fill.Visible Then
        styleStr = styleStr & " background-color: " & RGBToHex(cell.Shape.Fill.ForeColor.rgb) & ";"
    End If
    
    ' Add text alignment based on content direction
    ' Note: We'll handle this via dir attribute instead
    
    GetCellStyles = styleStr
End Function
Function CreateTableCell(cellText As String, cellStyles As String, textDirection As String, isHeader As Boolean) As String
    Dim cellHtml As String
    
    ' Determine cell tag type
    Dim tagName As String
    tagName = IIf(isHeader, "th", "td")
    
    ' Create the cell with all attributes
    cellHtml = "<" & tagName & _
               " style='" & cellStyles & "'" & _
               textDirection & ">" & _
               CleanHTML(cellText) & _
               "</" & tagName & ">" & vbNewLine
    
    CreateTableCell = cellHtml
End Function
Function GetRunFormatting(textRun As textRange) As String
    Dim formatting As String
    Dim styleClasses As String
    Dim inlineStyles As String
    Dim font As font
    
    Set font = textRun.font
    formatting = ""
    styleClasses = ""
    inlineStyles = ""
    
    ' Handle text direction
    If IsHebrewText(textRun.text) Then
        styleClasses = styleClasses & " rtl"
    Else
        styleClasses = styleClasses & " ltr"
    End If
    
    ' Font family
    If font.Name <> "" Then
        inlineStyles = inlineStyles & "font-family: '" & font.Name & "', Arial, sans-serif; "
    End If
    
    ' Font size (convert points to pixels - multiply by 1.33)
    If font.Size > 0 Then
        inlineStyles = inlineStyles & "font-size: " & Format(font.Size * 1.33, "0.00") & "px; "
    End If
    
    ' Font color
    If font.Color.rgb <> 0 Then
        inlineStyles = inlineStyles & "color: " & RGBToHex(font.Color.rgb) & "; "
    End If
    
    ' Bold
    If font.Bold Then
        styleClasses = styleClasses & " bold"
        inlineStyles = inlineStyles & "font-weight: bold; "
    End If
    
    ' Italic
    If font.Italic Then
        styleClasses = styleClasses & " italic"
        inlineStyles = inlineStyles & "font-style: italic; "
    End If
    
    ' Underline
    If font.Underline Then
        styleClasses = styleClasses & " underline"
        inlineStyles = inlineStyles & "text-decoration: underline; "
    End If
    
    ' Combine classes and styles
    If styleClasses <> "" Or inlineStyles <> "" Then
        formatting = IIf(styleClasses <> "", " class='mixed" & styleClasses & "'", "") & _
                    IIf(inlineStyles <> "", " style='" & inlineStyles & "'", "")
    End If
    
    GetRunFormatting = formatting
End Function


