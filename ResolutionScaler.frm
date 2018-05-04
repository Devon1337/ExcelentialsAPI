VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ResolutionScaler 
   Caption         =   "UserForm1"
   ClientHeight    =   79464
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   79788
   OleObjectBlob   =   "ResolutionScaler.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ResolutionScaler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Speed As Integer
Dim x1 As Integer
Dim x2 As Integer
Dim y1 As Integer
Dim y2 As Integer
Dim Modx As Double
Dim Mody As Double

Private Sub CommandButton1_Click()
x1 = CommandButton1.Left
y1 = CommandButton1.Top
End Sub

Private Sub CommandButton12_Click()
Call ScaleRes
ResolutionScaler.Hide
End Sub

Sub ScaleRes()
Modx = x2 / UCI.Width
Mody = y2 / UCI.Height

UCI.Height = y2
UCI.Width = x2
UCI.ConsoleOutput.Width = UCI.ConsoleOutput.Width * Modx
UCI.ConsoleOutput.Height = UCI.ConsoleOutput.Height * Mody
UCI.UserInput.Width = UCI.UserInput.Width * Modx
UCI.UserInput.Height = UCI.UserInput.Height * Mody
UCI.SCMD.Width = UCI.SCMD.Width * Modx
UCI.SCMD.Height = UCI.SCMD.Height * Mody

End Sub

Private Sub CommandButton5_Click()
x2 = CommandButton5.Left
y2 = CommandButton5.Top

End Sub

'Right'
Private Sub CommandButton10_Click()
ResolutionScaler.CommandButton5.Left = ResolutionScaler.CommandButton5.Left + Speed
End Sub
'Left'
Private Sub CommandButton11_Click()
ResolutionScaler.CommandButton5.Left = ResolutionScaler.CommandButton5.Left - Speed
End Sub

Private Sub CommandButton6_Click()
ResolutionScaler.CommandButton5.Left = ResolutionScaler.CommandButton5.Left / 2
ResolutionScaler.CommandButton5.Top = ResolutionScaler.CommandButton5.Top / 2
End Sub
'Down'
Private Sub CommandButton8_Click()
ResolutionScaler.CommandButton5.Top = ResolutionScaler.CommandButton5.Top + Speed
End Sub

'Up'
Private Sub CommandButton9_Click()
ResolutionScaler.CommandButton5.Top = ResolutionScaler.CommandButton5.Top - Speed
End Sub

Private Sub UserForm_Click()
Speed = 50
End Sub
