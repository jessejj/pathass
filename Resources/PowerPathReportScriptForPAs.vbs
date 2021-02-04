Dim wordObj

Dim lastMessage
pauseTime = 250
AutoNextItem = True 'user setting, future command flag

Dim grossHeader, microHeader, synoHeader
grossHeader = "GROSS DESCRIPTION:"
microHeader = "MICROSCOPIC DESCRIPTION:"
synoHeader = "SYNOPTIC REPORT:"

Say "Script starting..."
Wait 1

Do
WScript.Sleep 1000
If WordIsRunning Then
    If ReportIsOpen Then
        'do things, set user settings, etc
        Say "If you can read this, good things are happening"
        If AutoNextItem Then
            NumberOfFields = CountFields()
            If NumberOfFields = KnownNumberOfFields - 1 Then
                'user cleared one field
                nextItem
                KnownNumberOfFields = NumberOfFields
            ElseIf NumberOfFields > KnownNumberOfFields Then
                'user loaded a template or did something else
                Say "The number of fields was " & KnownNumberOfFields
                KnownNumberOfFields = NumberOfFields
                Say "Now there are " & KnownNumberOfFields
            ElseIf NumberOfFields = 0 Then
                Say "There are no fields"
                Wait 10
            ElseIf NumberOfFields < KnownNumberOfFields - 1 Then
                'user deleted a bunch of fields
                KnownNumberOfFields = NumberOfFields
            ElseIf NumberOfFields = KnownNumberOfFields Then
                Wait 1
            End If
        End If
    End If
End If
Loop 'Can run indenfinitely because user has access to cmd window

Function Say(message)
    'If Not message = lastMessage Then
        lastMessage = message
        On Error Resume Next
        WScript.StdOut.Write message
        If Err.Number <> 0 Then
            Msgbox "This script is only supported when launched from the App"
            WScript.Quit
        End If
        WScript.StdOut.WriteLine
        'If WordIsRunning Then wordObj.StatusBar = message
    'End If
End Function

Function Wait(seconds)
    WScript.Sleep seconds * 1000
End Function

Function WordIsRunning
    On Error Resume Next
    Set wordObj = GetObject(,"Word.Application")
    If Err.Number = 0 Then 
        WordIsRunning = True 
        Say "Word is running"
    Else
        Say "Waiting for Microsoft Word."
        WordIsRunning = False
        Wait 10
    End If
End Function

Function ReportIsOpen
    If wordObj.Documents.Count = 0 Then Say "There are no documents open."
    If Not wordObj.Documents.Count = DocumentCount Or DocumentCount = 1 Then
        If wordObj.Documents.Count = 0 Then
            Say "There are no documents open."
        Else
            For Each prop In wordObj.ActiveDocument.CustomDocumentProperties
                If prop.Name = "CaseID" Then
                    checkCaseID = prop.Value
                    Exit For
                End If
            Next

            If Not Len(checkCaseID) = 7 Then 
                Say "Not a recognized PowerPath report."
                ReportIsOpen = False
                Wait 5
            Else
                Say "Report is ready."
                ReportIsOpen = True
            End If
        End If
        DocumentCount = wordObj.Documents.Count
    End If
End Function

Function CountFields
    docText = wordObj.ActiveDocument.Range.Text
    If InStr(docText, microHeader) > InStr(docText, grossHeader) Then 'this is a Montfort report
        docText = Mid(docText, InStr(docText, grossHeader), InStr(docText, microHeader)-InStr(docText, grossHeader)) 'start at gross, end at micro
    Else
        docText = Mid(docText, InStr(docText, grossHeader))
    End If
    CountFields = len(docText) - len(replace(docText, "]", ""))
End Function

Sub nextItem
    If InStr(wordObj.Selection.text, "]") > 0 Then Exit Sub
    With wordObj.Selection.Find
        .Text = "(\[)*(\])"
        .Forward = True
        .Wrap = 0 'wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    wordObj.Selection.Find.Execute
    If wordObj.Selection.Find.Found = True Then 
        WScript.StdOut.Write Chr(7)     'beep
    End If
End Sub