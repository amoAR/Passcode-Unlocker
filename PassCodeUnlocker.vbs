main

Sub main
    
    Dim WshShell: Set WshShell = WScript.CreateObject("WScript.Shell")

    Const MIN_LOOP = 0
    Const MAX_LOOP = 9999
    Execute "Const LENGTH = " & Len(CStr(MAX_LOOP))
    
    Dim PADDING
    For i = 1 To Len(CStr(MAX_LOOP)) - 1
        PADDING = PADDING & "0"
    Next

    Set Rtn = WshShell.Exec("%windir%\system32\notepad.exe")
    WScript.Sleep 2000
    Dim result: result = WshShell.AppActivate(Rtn.ProcessID, 2000)
    If result <> True Then
        Msgbox "Failed to lunch program!", vbOKOnly, "Error"
        Exit Sub
    End If

    Dim ones, tens, hundreds, thousands
    For i = MIN_LOOP To MAX_LOOP
        number = Right(PADDING & i, LENGTH)
        thousands = Mid(number, 1, 1)
        hundreds = Mid(number, 2, 1)
        tens = Mid(number, 3, 1)
        ones = Mid(number, 4, 1)
        WshShell.SendKeys thousands
        WScript.Sleep 2000
        WshShell.SendKeys hundreds
        WScript.Sleep 2000
        WshShell.SendKeys tens
        WScript.Sleep 2000
        WshShell.SendKeys ones
        WshShell.SendKeys "{ENTER}"
        WScript.Sleep 5000
    Next

    Set WshShell = Nothing
    WScript.Quit 1

End Sub