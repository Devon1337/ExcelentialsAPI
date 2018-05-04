Attribute VB_Name = "Core_Lang_HTML"
'Documentation of the Core*'
Sub HTMLDOC()
Core_HTML_Documentation.WriteHeader ("Resp: Core")
Core_HTML_Documentation.WriteHeader ("Log(mStr)")
Core_HTML_Documentation.WriteSub ("Log(mStr): Used to output to UCI.ConsoleOutput.")
Core_HTML_Documentation.WriteSub ("mStr: String parameter that is used for the message.")

Core_HTML_Documentation.WriteHeader ("CharAt(Index, Message)")
Core_HTML_Documentation.WriteHeader ("ConsoleClear()")
Core_HTML_Documentation.WriteHeader ("Converter()")
Core_HTML_Documentation.WriteHeader ("CV(Row, Col)")
Core_HTML_Documentation.WriteHeader ("GetLetter")
Core_HTML_Documentation.WriteHeader ("SetLetters")
Core_HTML_Documentation.WriteHeader ("GetmId")
Core_HTML_Documentation.WriteHeader ("SetMId")
Core_HTML_Documentation.WriteHeader ("GetmName")
Core_HTML_Documentation.WriteHeader ("SetMName")
Core_HTML_Documentation.WriteHeader ("GetmNamePos")
Core_HTML_Documentation.WriteHeader ("SetEmailTime")
Core_HTML_Documentation.WriteHeader ("Ini")
Core_HTML_Documentation.WriteHeader ("NewCommandProcessor")

Core_HTML_Documentation.WriteHeader ("EnabledPlugin(mId, mName)")
Core_HTML_Documentation.WriteSub ("EnabledPlugin(mId,mName): Used to enable bridging from the core and the plugin.")
Core_HTML_Documentation.WriteSub ("mId: Integer param to link mName through dictionary.")
Core_HTML_Documentation.WriteSub ("mName: String param used for function identification.")

Core_HTML_Documentation.WriteHeader ("ChangeConsoleBackColor(Red, Green, Blue)")
Core_HTML_Documentation.WriteSub ("ChangeConsoleBackColor(Red, Green, Blue): Used to manipulate the background color of UCI.ConsoleOutput.")
Core_HTML_Documentation.WriteSub ("Red: 0 -> 255 long param Color intensity of red.")
Core_HTML_Documentation.WriteSub ("green: 0 -> 255 long param Color intensity of green.")
Core_HTML_Documentation.WriteSub ("blue: 0 -> 255 long param Color intensity of blue.")

Core_HTML_Documentation.WriteHeader ("ChangeConsoleTextColor(Red, Green, Blue)")
Core_HTML_Documentation.WriteSub ("ChangeConsoleTextColor(Red, Green, Blue): Used to manipulate the foreground color of UCI.ConsoleOutput.")
Core_HTML_Documentation.WriteSub ("Red: 0 -> 255 long param Color intensity of red.")
Core_HTML_Documentation.WriteSub ("green: 0 -> 255 long param Color intensity of green.")
Core_HTML_Documentation.WriteSub ("blue: 0 -> 255 long param Color intensity of blue.")

Core_HTML_Documentation.WriteHeader ("FLCommand()")
Core_HTML_Documentation.WriteSub ("FLCommand(): Used for transfering commands to functions.")

Core_HTML_Documentation.WriteHeader ("Resp: Core_Commands")
Core_HTML_Documentation.WriteHeader ("Resp: Core_Debug")
Core_HTML_Documentation.WriteHeader ("Resp: Core_Email")
Core_HTML_Documentation.WriteHeader ("Resp: Core_HTML_Documentation")
End Sub
