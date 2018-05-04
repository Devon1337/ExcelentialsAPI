Attribute VB_Name = "Core"
'Excellentials'
'Author: Devon Fuller'
'Version: 0.5D'

'Things needed for 0.6A'
'Add WorkbookAccess'

'TODO: Clean File'
'TODO: Remove unneeded If Statements'
'TODO: Remove unneeded functions/subroutines'
'TODO: Update Documentation'

Dim PluginName(20) As String
Dim PluginCommands(20, 20) As String
Dim ColConvert(26) As String
Dim CArg(25) As String
Dim ModId As New Scripting.Dictionary
Dim TempMid As Integer
Dim IndexPointer As Integer
Dim TempMName As String
Dim Enabler As Boolean
Dim Confirmer As Boolean
Dim FileWriteEn As Boolean
Dim EmailStage As Integer
Dim TempMid2 As Integer

Function Ini()
EmailStage = 0
Enabler = False
Confirmer = False
IndexPoint = 1

Log ("Starting...")
Log ("Calling Core_Lang.ENG_KEY")
Call Core_Lang.ENG_KEY

Call Core_Commands.CommandRegistering

Log ("Documenting Core...")
Call Core_HTML_Documentation.HTML
Call Core_Lang_HTML.HTMLDOC
Log ("Written HTML Documentation")

Log ("Done!")

End Function

Function Log(mStr As String)
Call UCI.ConsoleOutput_Change
UCI.ConsoleOutput.AddItem (mStr)
End Function

Function RegisterCommands(Mid As Integer, Command As String)
If (TempMid2 <> Mid) Then
TempMid2 = Mid
IndexPointer = 1
End If
Call Core_Mod_Data.SetPluginCommand(Mid, IndexPointer, Command)
PluginCommands(Mid, IndexPointer) = Command
IndexPointer = IndexPointer + 1
End Function

'Set and Getters'
Function SetLetters(Index As Integer, Letter As String)
ColConvert(Index) = Letter
End Function
Function SetMId(Mid As Integer)
TempMid = Mid
End Function
Function SetMName(mName As String)
TempMName = mName
End Function
Function GetmId() As Integer
GetmId = TempMid
End Function
Function GetmName() As String
GetmName = TempMName
End Function
Function GetmNamePos(Index As Integer) As String
If (PluginName(Index) <> "") Then
GetmNamePos = PluginName(Index)
End If
End Function
Function SetArg(ArgNumber As Integer, ArgReturn As Variant)
CArg(ArgNumber) = ArgReturn
End Function
Function SetEmailTime(Index As Integer)
EmailStage = Index
End Function

'Cell Viewer'
Function CV() As String
Dim ColNumber As Integer
Dim I As Integer
For I = 1 To Len(CArg(2))
ColNumber = ColNumber + Converter(CharAt(CArg(2), I))
Next I
Log (Cells(CArg(1), CArg(2)).Value)
CV = Cells(CArg(1), CArg(2)).Value
End Function

'Gathers Mod Credentials'
Function EnabledPlugin(Mid As Integer, mName As String)

If (UCI.Visible = False) Then
SetMId (Mid)
SetMName (mName)
UCI.Show
End If

If (UCI.Visible = True) Then
If (ModId.Exists(GetmId)) Then
Call Log("modID Already exists! {" & GetmName & "}")
Else
PluginName(Mid) = mName
Call ModId.Add(GetmId, GetmName)
Call Log("[$]" & GetmName & " has been enabled")
End If
End If

End Function

Function NewCommandProcessor(Message As String)

On Error GoTo ErrHandler

'Declaration'
Dim Index As Integer
Dim TempString As String
Dim LengIdentifier As Integer
Dim ArgConcate As String
Dim I As Integer
Dim J As Integer


Log (Message)

'Variable Definition'
Index = 1

If (Core_Utils.CharAt(Message, 1) = "/") Then GoTo CommandMarker

CommandMarker:
For I = 2 To Len(Message)
ArgConcate = ArgConcate & Core_Utils.CharAt(Message, I)
Next I
    Application.Run "Application.Run " & ArgConcate
HelpMarker:


If (Message = "?") Then
Core_Console_Lib.ListCommands
End If

ErrHandler:
If (Err.Number = 1004) Then
Log ("[X] Error: unknown command! Type '?' for help")
End If
Exit Function

End Function

