Set WshShell = WScript.CreateObject("WScript.Shell")
WshShell.Run "winword.exe", 9

WScript.Sleep 500 


Dim index
index = 0

Do While index <= 1000
	index += 1

	WshShell.SendKeys "Hello World!"
	WshShell.SendKeys "{ENTER}"
	WScript.Sleep 500 

Loop