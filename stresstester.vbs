msg = InputBox("How much do you want to stress test your PC? (input a number)")
    for i = 1 to msg
        Set objShell = CreateObject("WScript.Shell")
        objShell.Run "notepad.exe"
        WScript.Sleep 1000
    Next
