'this simple script will loop a set of keystrokes with a pause in between each loop

LoopQuestionForUser = "How many orders are there that need to be toggled for the current specimen?" 'how many loops?

For Attempt = 0 to 5
    NumberOfLoops = InputBox(LoopQuestionForUser, "Rush toggle for PowerPath Orders", NumberOfLoops)
    If IsNumeric(NumberOfLoops) Then Exit For
Next

If Not IsNumeric(NumberOfLoops) Or NumberOfLoops < 1 Then Wscript.Quit

Dim Strings(3)

Strings(0) = "~"    'opens order
Strings(1) = "%h"   'toggles Rush checkbox
Strings(2) = "~"    'confirm with Enter
Strings(3) = "{down}"   'select next order


Set Shell = CreateObject("Wscript.Shell")

If Not Shell.AppActivate("PowerPath") Then
    msgbox "PowerPath is not open"
    Wscript.Quit
End If

For LoopIndex = 1 To NumberOfLoops 
    For StringIndex = 0 to UBound(Strings)
        Wscript.Sleep 100
        Shell.Sendkeys Strings(StringIndex), True
    Next
    If NumberOfLoops > 30 Then Exit For
Next

Set Shell = Nothing
