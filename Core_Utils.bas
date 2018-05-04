Attribute VB_Name = "Core_Utils"
Function Converter(Letter As String) As Integer
For I = 1 To 26
If (ColConvert(I) = LCase(Letter)) Then
Converter = I
Exit Function
End If
Next I
Converter = 0
End Function

Function CharAt(Message As String, Index As Integer) As String
CharAt = Left(Mid(Message, Index), 1)
End Function

Public Function CellReaderL(X As Integer, Y As String) As String
CellReaderL = Cells(X, Converter(Y).Value)
End Function

Public Function CellReaderI(X As Integer, Y As Integer) As String
CellReaderI = Cells(X, Y).Value
End Function

Public Function Copy(SelRange As Range)
Range(SelRange).Select
Selection.Copy
End Function

Public Function Paste(SelRange As Range)
SelRange.Select
ActiveSheet.Paste
End Function

