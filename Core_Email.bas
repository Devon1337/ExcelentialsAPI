Attribute VB_Name = "Core_Email"
'TODO: Add a form of drafting/editting/deleting'
Dim Recip As String
Dim Title As String
Dim strbody As String

Function iniEmail()
Call Core.ConsoleClear
Call Core.Log("Send To?")
End Function

Function EmailSend()
Core.SetEmailTime (0)

Dim OutApp As Object
Dim OutMail As Object

Set OutApp = CreateObject("Outlook.Application")
Set OutMail = OutApp.CreateItem(0)

On Error Resume Next
With OutMail
    .To = Recip
    .Subject = Title
    .Body = strbody
    .Send
End With
On Error GoTo 0

Set OutMail = Nothing
Set OutApp = Nothing

End Function

Function SetEmail(Recipiant As String)
Recip = Recipiant
Core.SetEmailTime (2)
Core.ConsoleClear
Core.Log ("What is the subject?")
End Function

Function SetSubject(Subject As String)
Title = Subject
Core.SetEmailTime (3)
Core.ConsoleClear
Core.Log ("What is the Message?")
End Function

Function SetMessage(Message As String)
strbody = Message
Core.SetEmailTime (4)
Core.ConsoleClear
Core.Log ("Is this correct?")
Core.Log ("Send To: " & Recip)
Core.Log ("Subject: " & Title)
Core.Log ("")
Core.Log (strbody)

End Function
