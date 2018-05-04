Attribute VB_Name = "Core_Utils_Workbook"

Public Function WorkbookStart(ApplicationName As String)
Dim Wb As Workbook
Core.Log (Workbooks.Count)
Set Wb = Workbooks.Open(ApplicationName)
Core.Log (Workbooks.Count)
Wb.Activate
End Function

Public Function WorkbookSelection(ApplicationName As String)
Dim Wb As Workbook
Wb(ApplicationName).Selected
End Function

