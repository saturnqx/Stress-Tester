msg = MsgBox("Do you want to choose the stress amount or do you want to go infinite? (yes is infinite, no is custom)",4)
if msg = 6 Then
    Do
        Set objShell = CreateObject("WScript.Shell")
        objShell.Run "notepad.exe"
        WScript.Sleep 3000
        
    Loop
ElseIf msg = 7 Then
    msg = InputBox("How much do you want to stress test your PC? (input a number)")
    for i = 1 to msg
        Set objShell = CreateObject("WScript.Shell")
        objShell.Run "notepad.exe"
        WScript.Sleep 1000
    Next
end If
