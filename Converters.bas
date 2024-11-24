Attribute VB_Name = "Converters"
Function ConvertTableToHTML(tbl As Table) As String
    Dim row As Long
    Dim col As Long
    Dim htmlTable As String
    Dim cell As cell
    Dim cellText As String
    Dim IsHeaderRow As Boolean
    
    On Error GoTo ErrorHandler
    
    ' Initialize row and col
    row = 1
    col = 1
    
    ' Debug logging
    Debug.Print "Starting table conversion"
    Debug.Print "Table dimensions: " & tbl.Rows.Count & " rows x " & tbl.Columns.Count & " columns"
    
    ' Add table with basic styling
    htmlTable = "<table style='border-collapse: collapse; width: 100%; margin: 20px 0;'>" & vbNewLine
    
    ' Process each row
    For row = 1 To tbl.Rows.Count
        htmlTable = htmlTable & "<tr>" & vbNewLine
        
        ' Determine if this is a header row based on formatting
        IsHeaderRow = False 'IsHeaderRow(tbl, row)
        
        ' Process each column
        For col = 1 To tbl.Columns.Count
            Set cell = tbl.cell(row, col)
            
            ' Get cell text
            cellText = GetCellText(cell)
            
            ' Get cell styles including background color and text alignment
            Dim cellStyles As String
            cellStyles = GetCellStyles(cell)
            
            ' Determine text direction
            Dim textDirection As String
            textDirection = GetTextDirectionAttribute(cellText)
            
            ' Create the cell with appropriate styling and direction
            htmlTable = htmlTable & CreateTableCell(cellText, cellStyles, textDirection, IsHeaderRow)
        Next col
        
        htmlTable = htmlTable & "</tr>" & vbNewLine
    Next row
    
    htmlTable = htmlTable & "</table>" & vbNewLine
    
    Debug.Print "Table conversion completed successfully"
    ConvertTableToHTML = htmlTable
    Exit Function

ErrorHandler:
    Debug.Print "Error in ConvertTableToHTML at row " & row & ", col " & col & ": " & Err.Description
    Debug.Print "Error number: " & Err.Number
    Debug.Print "Error source: " & Err.Source
    ConvertTableToHTML = "<table style='border-collapse: collapse; width: 100%; margin: 20px 0;'>" & _
                        "<tr><td style='border: 1px solid #ddd; padding: 8px;'>" & _
                        "Error converting table: " & Err.Description & _
                        "</td></tr></table>"
End Function

Function IsHeaderRow(tbl As Table, rowIndex As Long) As Boolean
    ' Check if this row has header-like formatting
    Dim cell As cell
    Set cell = tbl.cell(rowIndex, 1)
    
    ' Check for bold text or different background color
    If cell.Shape.TextFrame.textRange.font.Bold Then
        IsHeaderRow = True
        Exit Function
    End If
    
    ' You could add more header detection logic here
    ' For example, check background color, font size, etc.
    IsHeaderRow = (rowIndex = 1)  ' Default to treating first row as header
End Function
Function GetMergedRowSpan(tbl As Table, startRow As Integer, startCol As Integer) As Integer
    Dim rowSpan As Integer
    Dim currentRow As Integer
    Dim cell As cell
    
    rowSpan = 1
    currentRow = startRow + 1
    
    While currentRow <= tbl.Rows.Count
        Set cell = tbl.cell(currentRow, startCol)
        If cell.Merged Then
            rowSpan = rowSpan + 1
            currentRow = currentRow + 1
        Else
            Exit Do
        End If
    Wend
    
    GetMergedRowSpan = rowSpan
End Function
Function GetMergedColSpan(tbl As Table, startRow As Integer, startCol As Integer) As Integer
    Dim colSpan As Integer
    Dim currentCol As Integer
    Dim cell As cell
    
    colSpan = 1
    currentCol = startCol + 1
    
    While currentCol <= tbl.Columns.Count
        Set cell = tbl.cell(startRow, currentCol)
        If cell.Merged Then
            colSpan = colSpan + 1
            currentCol = currentCol + 1
        Else
            Exit Do
        End If
    Wend
    
    GetMergedColSpan = colSpan
End Function
Function IsMergedButNotMain(cell As cell, row As Integer, col As Integer, mergeTracker As Collection) As Boolean
    Dim mainCell As cell
    Dim mergeInfo As Variant
    
    If Not cell.Merged Then
        IsMergedButNotMain = False
        Exit Function
    End If
    
    ' Check if this cell is in the merge tracker as a main cell
    On Error Resume Next
    mergeInfo = mergeTracker.item(CStr(row) & "_" & CStr(col))
    On Error GoTo 0
    
    ' If we found it in the tracker, it's a main cell
    If Not IsEmpty(mergeInfo) Then
        IsMergedButNotMain = False
    Else
        IsMergedButNotMain = True
    End If
End Function
Function ConvertTextFrameToHTML(txtFrame As TextFrame) As String
    Dim txt As textRange
    Dim para As textRange
    Dim htmlText As String
    Dim isBulletList As Boolean
    Dim textStyle As String
    Dim shapeStyle As String
    
    Set txt = txtFrame.textRange
    isBulletList = False
    htmlText = ""
    
    ' Add CSS styles for text formatting
    htmlText = "<style>" & vbNewLine & _
              ".bold { font-weight: bold; }" & vbNewLine & _
              ".italic { font-style: italic; }" & vbNewLine & _
              ".underline { text-decoration: underline; }" & vbNewLine & _
              ".rtl { direction: rtl; }" & vbNewLine & _
              ".ltr { direction: ltr; }" & vbNewLine & _
              "</style>" & vbNewLine
    
    ' Get shape styles
    shapeStyle = GetShapeStyles(txtFrame.Parent)
    
    ' Start a div with shape styles if any exist
    If shapeStyle <> "" Then
        htmlText = htmlText & "<div" & shapeStyle & ">" & vbNewLine
    End If
    
    ' Process each paragraph
    For Each para In txt.paragraphs
        If para.ParagraphFormat.Bullet.Visible Then
            ' Start bullet list if not already started
            If Not isBulletList Then
                htmlText = htmlText & "<ul>" & vbNewLine
                isBulletList = True
            End If
            htmlText = htmlText & "<li>" & ProcessFormattedText(para) & "</li>" & vbNewLine
        Else
            ' End bullet list if it was active
            If isBulletList Then
                htmlText = htmlText & "</ul>" & vbNewLine
                isBulletList = False
            End If
            
            ' Add paragraph
            htmlText = htmlText & "<p>" & ProcessFormattedText(para) & "</p>" & vbNewLine
        End If
    Next para
    
    ' Close bullet list if still open
    If isBulletList Then
        htmlText = htmlText & "</ul>" & vbNewLine
    End If
    
    ' Close shape style div if it was opened
    If shapeStyle <> "" Then
        htmlText = htmlText & "</div>" & vbNewLine
    End If
    
    Debug.Print htmlText
    ConvertTextFrameToHTML = htmlText
End Function

