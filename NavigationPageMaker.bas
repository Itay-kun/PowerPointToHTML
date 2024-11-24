Attribute VB_Name = "NavigationPageMaker"
Sub CreateNavigationPage()
    Dim FSO As Object
    Dim htmlFile As Object
    Dim baseFolder As String
    Dim htmlContent As String
    
    ' Create FileSystemObject
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    ' Get base folder path
    baseFolder = Application.ActivePresentation.Path & "\HTMLSlides\"
    
    ' Create navigation HTML content
htmlContent = "<!DOCTYPE html>" & vbNewLine
htmlContent = htmlContent + "<html dir='rtl' lang='he'>" & vbNewLine
htmlContent = htmlContent + "<head>" & vbNewLine
htmlContent = htmlContent + "    <meta charset='UTF-8'>" & vbNewLine
htmlContent = htmlContent + "    <title>Slides Navigation</title>" & vbNewLine
htmlContent = htmlContent + "    <style>" & vbNewLine
htmlContent = htmlContent + "        body { font-family: Arial, sans-serif; margin: 40px; direction: rtl; }" & vbNewLine
htmlContent = htmlContent + "        .tree { margin: 20px; }" & vbNewLine
htmlContent = htmlContent + "        .tree-folder { margin-bottom: 10px; }" & vbNewLine
htmlContent = htmlContent + "        .tree-folder-header { "
htmlContent = htmlContent + "            background-color: #f0f0f0; "
htmlContent = htmlContent + "            padding: 8px; "
htmlContent = htmlContent + "            margin: 2px 0; "
htmlContent = htmlContent + "            border-radius: 4px; "
htmlContent = htmlContent + "            font-weight: bold; "
htmlContent = htmlContent + "        }" & vbNewLine
htmlContent = htmlContent + "        .tree-folder-content { margin-right: 20px; }" & vbNewLine
htmlContent = htmlContent + "        .tree-item { "
htmlContent = htmlContent + "            padding: 5px 8px; "
htmlContent = htmlContent + "            margin: 2px 0; "
htmlContent = htmlContent + "            border-radius: 4px; "
htmlContent = htmlContent + "        }" & vbNewLine
htmlContent = htmlContent + "        .tree-item:hover { background-color: #f8f8f8; }" & vbNewLine
htmlContent = htmlContent + "        a { "
htmlContent = htmlContent + "            color: #0066cc; "
htmlContent = htmlContent + "            text-decoration: none; "
htmlContent = htmlContent + "            display: block; "
htmlContent = htmlContent + "        }" & vbNewLine
htmlContent = htmlContent + "        a:hover { text-decoration: underline; }" & vbNewLine
htmlContent = htmlContent + "        h1 { color: #333; margin-bottom: 30px; }" & vbNewLine
htmlContent = htmlContent + "    </style>" & vbNewLine
htmlContent = htmlContent + "</head>" & vbNewLine
htmlContent = htmlContent + "<body>" & vbNewLine
htmlContent = htmlContent + "    <h1>Slides Navigation</h1>" & vbNewLine
htmlContent = htmlContent + "    <div class='tree'>" & vbNewLine
htmlContent = htmlContent + GenerateTreeHTML(FSO, baseFolder, "")
htmlContent = htmlContent + "    </div>" & vbNewLine
htmlContent = htmlContent + "</body>" & vbNewLine
htmlContent = htmlContent + "</html>"
    
    ' Save the navigation file
    Set htmlFile = FSO.CreateTextFile(baseFolder & "index.html", True, True)
    htmlFile.Write htmlContent
    htmlFile.Close
    
    MsgBox "Navigation page has been created at: " & baseFolder & "index.html"
End Sub

Function GenerateFileStructureJSON(FSO As Object, folderPath As String) As String
    Dim folder As Object
    Dim subFolder As Object
    Dim file As Object
    Dim jsonStr As String
    
    Set folder = FSO.GetFolder(folderPath)
    
    jsonStr = "{"
    
    ' Add subfolders
    For Each subFolder In folder.SubFolders
        If subFolder.Name <> "images" Then  ' Skip the images folder
            jsonStr = jsonStr & """" & subFolder.Name & """: {"
            
            ' Add files in subfolder
            For Each file In subFolder.Files
                If LCase(FSO.GetExtensionName(file.Name)) = "html" Then
                    jsonStr = jsonStr & """" & file.Name & """: {},"
                End If
            Next file
            
            ' Remove trailing comma if exists
            If Right(jsonStr, 1) = "," Then
                jsonStr = Left(jsonStr, Len(jsonStr) - 1)
            End If
            
            jsonStr = jsonStr & "},"
        End If
    Next subFolder
    
    ' Remove trailing comma if exists
    If Right(jsonStr, 1) = "," Then
        jsonStr = Left(jsonStr, Len(jsonStr) - 1)
    End If
    
    jsonStr = jsonStr & "}"
    
    GenerateFileStructureJSON = jsonStr
End Function

Function GenerateTreeHTML(FSO As Object, folderPath As String, relativePath As String) As String
    Dim folder As Object
    Dim subFolder As Object
    Dim file As Object
    Dim htmlStr As String
    Dim fileCount As Long
    
    Set folder = FSO.GetFolder(folderPath)
    htmlStr = ""
    fileCount = 0
    
    ' Process subfolders first
    For Each subFolder In folder.SubFolders
        If subFolder.Name <> "images" Then  ' Skip the images folder
            Dim subFolderContent As String
            Dim newRelativePath As String
            newRelativePath = IIf(relativePath = "", "", relativePath & "/") & subFolder.Name
            
            ' Get subfolder content first
            subFolderContent = GenerateTreeHTML(FSO, subFolder.Path, newRelativePath)
            
            ' Only add folder if it has content
            If subFolderContent <> "" Then
                htmlStr = htmlStr & "        <div class='tree-folder'>" & vbNewLine & _
                         "            <div class='tree-folder-header'>" & subFolder.Name & "</div>" & vbNewLine & _
                         "            <div class='tree-folder-content'>" & vbNewLine & _
                         subFolderContent & _
                         "            </div>" & vbNewLine & _
                         "        </div>" & vbNewLine
                fileCount = fileCount + 1
            End If
        End If
    Next subFolder
    
    ' Then process HTML files in current folder
    For Each file In folder.Files
        If LCase(FSO.GetExtensionName(file.Name)) = "html" And file.Name <> "index.html" Then
            Dim fileRelativePath As String
            fileRelativePath = IIf(relativePath = "", "", relativePath & "/") & file.Name
            
            htmlStr = htmlStr & "            <div class='tree-item'>" & vbNewLine & _
                     "                <a href='" & fileRelativePath & "'>" & _
                     file.Name & "</a>" & vbNewLine & _
                     "            </div>" & vbNewLine
            fileCount = fileCount + 1
        End If
    Next file
    
    ' Only return content if we found files
    If fileCount > 0 Then
        GenerateTreeHTML = htmlStr
    Else
        GenerateTreeHTML = ""
    End If
End Function

