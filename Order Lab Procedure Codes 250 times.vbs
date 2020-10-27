LoopQuestionForUser = "Which lab procedure code?" 'how many loops?
LoopQuestionForUser2 = "How many batches of 250?" 'how many loops?

LabProcedure = InputBox(LoopQuestionForUser, "Order Lab Procedure Codes 250 times", "SNSWEDU")

For Attempt = 0 To 5
    NumberOfLoops = InputBox(LoopQuestionForUser2, "Order Lab Procedure Codes 250 times", NumberOfLoops)
    If IsNumeric(NumberOfLoops) Then Exit For
Next

Dim Strings(7)

Strings(0) = "~"    	'opens order
Strings(1) = LabProcedure
Strings(2) = "~"        'confirm with Enter
Strings(3) = "{end}"    'go to quantity
Strings(4) = "250"      'BATCH 250
Strings(5) = "~"        'confirm
Strings(6) = "{home}"   'go back to Code field
Strings(7) = "{down}"   'New line

Set Shell = CreateObject("Wscript.Shell")

If Not Shell.AppActivate("New Orders") Then
    MsgBox "PowerPath is not open"
    Wscript.Quit
End If

For LoopIndex = 1 To NumberOfLoops
    For StringIndex = 0 To UBound(Strings)
        Wscript.Sleep 100
        Shell.SendKeys Strings(StringIndex), True
    Next
    If NumberOfLoops > 1000 Then Exit For
Next

Set Shell = Nothing
