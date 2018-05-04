Attribute VB_Name = "Core_Commands"
'TODO: Add more top level commands'
Sub Restart()

End Sub

Sub CommandRegistering()

Call Core.RegisterCommands(0, "ChangeScale")
Call Core.RegisterCommands(0, "BackColorChange")
Call Core.RegisterCommands(0, "ClearConsole")
Call Core.RegisterCommands(0, "Plugins")
Call Core.RegisterCommands(0, "ViewCell")
Call Core.RegisterCommands(0, "WriteCell")

End Sub

Sub ChangeScale()
ResolutionScaler.Show
End Sub

Sub BackColorChange()
Call Core_Console_Lib.ChangeConsoleBackColor
End Sub

Sub ClearConsole()
UCI.ConsoleOutput.Clear
End Sub

Sub Plugins()
Dim I As Integer
Core.Log ("Excelentials")
For I = 1 To 20
If (Core.GetmNamePos(I) <> "") Then
Core.Log ("|")
Core.Log ("|-" & Core.GetmNamePos(I))
End If

Next I
End Sub

Sub ViewCell()
Call Core_Console_Lib.CV
End Sub
Sub WriteCell()
Call Core_Console_Lib.CW
End Sub

Sub SendEmail()
Core_Email.iniEmail
Core.SetEmailTime (1)
End Sub

