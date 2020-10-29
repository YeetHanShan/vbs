' -s = shutdown, -t 0 = no timeout, -f = force programs to close 
strShutdown = "shutdown.exe -s -t 0 -f -m \\" & strComputer 
set objShell = CreateObject("WScript.Shell") 
objShell.Run strShutdown, 0, false
