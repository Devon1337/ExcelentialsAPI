VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UCI 
   Caption         =   "User Console Input"
   ClientHeight    =   3876
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   6144
   OleObjectBlob   =   "UCI.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UCI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CurrentCommand As String

Sub ConsoleOutput_Change()
If (ConsoleOutput.ListCount > 14) Then
ConsoleOutput.RemoveItem (0)
End If
End Sub

Private Sub ConsoleOutput_Click()
UserInput.Text = ConsoleOutput.Text
End Sub

Private Sub SCMD_Click()

Call Core.NewCommandProcessor(UserInput.Text)

End Sub

Private Sub UserForm_Initialize()


End Sub

Private Sub UserForm_Activate()

ActivatedBool = True
Dim Mid As Integer
Dim mName As String
Mid = Core.GetmId
mName = Core.GetmName
Call Core.Ini
If (Mid <> 0) Then
Call Core.EnabledPlugin(Mid, mName)
End If

End Sub

Private Sub UserInput_Change()

End Sub
