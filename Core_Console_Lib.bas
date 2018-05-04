Attribute VB_Name = "Core_Console_Lib"

'Cell Writer'
Function CW(Message As String, Row As Integer, Col As String)

On Error GoTo ErrHandler

Dim ColNumber As Integer
Dim I As Integer

For I = 1 To Len(Col)
ColNumber = ColNumber + Converter(CharAt(Col, I))
Next I

Cells(Row, ColNumber).Value = Message
Log ("Cell(" & Row & " , " & ColNumber & ") Has been modified to " & Message & "!")

ErrHandler:
If (Err.Number = 13) Then
Log ("[X] Error Incorrect Syntax! /WriteCell {RowInt} {ColString} {TextString}")
End If

Exit Function
End Function

Function ChangeConsoleBackColor()
UCI.ConsoleOutput.BackColor = RGB(CArg(1), CArg(2), CArg(3))
End Function

Function ChangeConsoleTextColor()
UCI.ConsoleOutput.ForeColor = RGB(CArg(1), CArg(2), CArg(3))
End Function

Function ConsoleClear()
UCI.ConsoleOutput.Clear
End Function

Function ListCommands()
Dim I As Integer
Dim J As Integer

Log ("====Help====")
Log ("Just click and send to get command details")
For I = 0 To 20
For J = 1 To 20
If (Core_Util.GetPluginCommand(I, J) <> "") Then
Log ("/" + Core_Util.GetPluginCommand(I, J))
Else
I = I + 1
J = 1
End If
Next J
Next I


End Function
