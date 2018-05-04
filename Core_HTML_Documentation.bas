Attribute VB_Name = "Core_HTML_Documentation"
'HTML Documentation'
Dim HTMLFile As String

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
     (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal _
     lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Function OpenHTML()
ShellExecute 0&, vbNullString, HTMLFile, vbNullString, _
      vbNullString, SW_SHOWNORMAL
      
End Function

Sub HTML()
    HTMLFile = ActiveWorkbook.Path & "\ExcelentialsDocumentation.html"
    Call ClearDoc
    
    Close
    Open HTMLFile For Append As #1
        Print #1, "<html>"
        Print #1, "<head>"
        Print #1, "<style type=""text/css"">"
        Print #1, "  body { font-size:12px;font-family:tahoma } "
        Print #1, "</style>"
        Print #1, "</head>"
        Print #1, "<body>"

        Print #1, "<h2> Excelentials! </h2>"
        Print #1, "<p>Author:Devon Fuller</p>"
        Print #1, "<p>Version: 0.5B</p>"

        Print #1, "</body>"
        Print #1, "</html>"
    Close
End Sub

Sub WriteHeader(dMessage As String)
 Open HTMLFile For Append As #1
        Print #1, "<html>"
        Print #1, "<body>"
        Print #1, "<h2>" & dMessage & "</h2>"
        Print #1, "</body>"
        Print #1, "</html>"
    Close
End Sub
 
Sub WriteSub(dMessage As String)
 Open HTMLFile For Append As #1
 Print #1, "<html>"
        Print #1, "<head>"
        Print #1, "<style type=""text/css"">"
        Print #1, "  body { font-size:12px;font-family:tahoma } "
        Print #1, "</style>"
        Print #1, "</head>"
        Print #1, "<body>"

        Print #1, "<p>" & dMessage & "</p>"

        Print #1, "</body>"
        Print #1, "</html>"
    Close
End Sub

Sub ClearDoc()
    Open HTMLFile For Output As #1
    
        Print #1, ""
    Close
    
End Sub

