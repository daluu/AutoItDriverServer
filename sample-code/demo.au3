Local $myArgs = ""
If $CmdLine[0] > 0 Then
	For $i = 1 To $CmdLine[0] Step 1
		$myArgs &= " " & $CmdLine[$i]
	Next
EndIf
MsgBox(64, "Demo", "An AutoIt au3 script was ran, with arguments (if any) of: " & $myArgs)