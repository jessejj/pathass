LoopQuestionForUser = "Which lab procedure code?" 'how many loops?
LoopQuestionForUser2 = "How many batches of 250?" 'how many loops?

LabProcedure = InputBox(LoopQuestionForUser, "Order Lab Procedure Codes 250 times", "SNSWEDU")

For Attempt = 0 To 5
    NumberOfLoops = InputBox(LoopQuestionForUser2, "Order Lab Procedure Codes 250 times", NumberOfLoops)
    If IsNumeric(NumberOfLoops) Then Exit For
Next

Dim Strings(6)

Strings(0) = "{F2}250"    	'opens order
Strings(1) = "~"        'confirm
Strings(2) = "{home}"   'go back to Code field
Strings(3) = LabProcedure
Strings(4) = "~"        'confirm
Strings(5) = "{tab 3}"        'confirm with Enter
Strings(6) = "^i"   'New line

Set Shell = CreateObject("Wscript.Shell")


If Not Shell.AppActivate("New Orders") Then
    MsgBox "PowerPath is not open"
    Wscript.Quit
End If

Shell.SendKeys "{end}", True

For LoopIndex = 1 To NumberOfLoops

If Not Shell.AppActivate("New Orders") Then
    MsgBox "PowerPath is not open"
    Wscript.Quit
End If

    For StringIndex = 0 To UBound(Strings)
        Wscript.Sleep 50
        Shell.SendKeys Strings(StringIndex), True
    Next
    If NumberOfLoops > 1000 Then Exit For
Next

Set Shell = Nothing
