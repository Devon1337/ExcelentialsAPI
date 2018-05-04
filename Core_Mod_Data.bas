Attribute VB_Name = "Core_Mod_Data"
Dim PluginCommands(20, 20) As String

Function GetPluginCommand(X As Integer, Y As Integer) As String
GetPluginCommand = PluginCommands(X, Y)
End Function

Function SetPluginCommand(PluginPosition As Integer, PluginCommandPosition As Integer, PluginCommand As String)
PluginCommands(PluginPosition, PluginCommandPosition) = PluginCommand
End Function
