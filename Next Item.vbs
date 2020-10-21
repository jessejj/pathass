Option Explicit
Dim scriptVersion
scriptVersion = "Next Item Script 2020.07.21" 'by Jesse Pelley

'This script was written to help Pathologist Assistants describe specimens more effeciently while they dissect

Dim pauseValue, greetingsMessage, continuingText, continuingTextArr, settingsPath, userInitials, cassetteKey 'for checkSettings
Dim numberofBlanks, previousnumberofBlanks, advancedFields, timerStart, CaseID

Dim grossHeader, microHeader, synoHeader
grossHeader = "GROSS DESCRIPTION:"
microHeader = "MICROSCOPIC DESCRIPTION:"
synoHeader = "SYNOPTIC REPORT:"



Dim wordObj, docText
Set wordObj = GetObject(,"word.application")

If wordObj.Documents.Count = 0 Then
Say "Please open a report to dictate"
    WScript.sleep 3000
    Set wordObj = Nothing
    WScript.Quit
End If

checkSettings
greeting

If userInitials = "" Then 
    userInitials = InputBox("Please type your initials that you use in your gross descriptions",scriptVersion,"JJP")
    userInitials = Ucase(userInitials)
    WriteSettings "userInitials=", userInitials
End If

If Not Len(userInitials) = 3 Then
    Say "Please use 3 letters for your initials. Say 'open settings' to fix"
    WScript.sleep 3000
    Set wordObj = Nothing
    WScript.Quit
End If

CaseID = checkCaseID


Dim shell
        Set shell = createobject("wscript.shell")
            shell.AppActivate("Case")
        Set shell = Nothing



nextItem




Do
    If reportSafe = False Then 
        WScript.Sleep 3000 'to display function's message
        Exit Do
    End If

    If Not checkCaseID = CaseID Then Exit Do

    timerStart = Timer()

    previousnumberofBlanks = numberofBlanks
    numberofBlanks = countBlanks
    If InStr(docText, userInitials) > 0 Then exit do

    'Say "Advancing to next field in " & pauseValue & " milliseconds"
    WScript.Sleep pauseValue

    If incompleteParagraph Then previousnumberofBlanks = previousnumberofBlanks - 1

    If numberofBlanks < previousnumberofBlanks Then
        nextItem
    Else
        'Say "Waiting..."
        waitLoop
    End If

    If cassetteKey = "Auto" then blockAutofill
    
    Say "There are " & numberofBlanks & " blank fields left in the gross description"

Loop Until numberofBlanks = 0

Say scriptVersion & " Finished."
Set wordObj = Nothing

Function waitLoop

    Do 
        WScript.Sleep pauseValue 
        If InStr(wordObj.Selection.Text, "]") = 0 Then 'Fill in not present or filled out, so stop waiting
            Exit Do
        End If
    Loop

End Function

Function Say(message)
    wordObj.StatusBar = message
WScript.StdOut.Write message
WScript.StdOut.WriteLine
End Function

Function countBlanks

    docText = wordObj.ActiveDocument.Range.Text

    If InStr(docText, microHeader) > InStr(docText, grossHeader) Then 'this is a Montfort report
        docText = Mid(docText, InStr(docText, grossHeader), InStr(docText, microHeader)-InStr(docText, grossHeader)) 'start at gross, end at micro
    Else
        docText = Mid(docText, InStr(docText, grossHeader))
    End If

    countBlanks = len(docText) - len(replace(docText, "]", ""))

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
        advancedFields = advancedFields + 1

WScript.StdOut.Write Chr(7)
    End If



End Sub

Function incompleteParagraph
    dim lastCharacters, Index
    Set lastCharacters = wordObj.Selection.Range
    lastCharacters.Start = lastCharacters.Start - 2

    For Index = 0 to Ubound(continuingTextArr)
        If InStr(lastCharacters, continuingTextArr(Index)) > 0 Then
            incompleteParagraph = True
            exit for
        end If
    Next

End Function

Function reportSafe
    reportSafe = True
    docText = wordObj.ActiveDocument.Range.Text

    If InStr(docText, grossHeader) = 0 Then
        reportSafe = False
        Say grossHeader & " is missing."
    End If

    If InStr(docText, microHeader) = 0 And InStr(docText, synoHeader) = 0 Then
        reportSafe = False
        Say microHeader & " or " & synoHeader & " is missing."
    End If

End Function

Sub checkSettings

    Dim FSO
    settingsPath = "Next Item Settings.txt"

    Set FSO = CreateObject("Scripting.FileSystemObject")

    If Not FSO.FileExists(settingsPath) Then
        Say "Settings file not found."
        WScript.sleep 3000
        Set FSO = Nothing
        Set wordObj = Nothing
        WScript.Quit
    End If

    Dim SettingsFile, currentLine
    Set SettingsFile = FSO.OpenTextFile(settingsPath, 1)
        Do While SettingsFile.AtEndOfStream = False

            currentLine = SettingsFile.ReadLine

            If InStr(currentLine, "userInitials=") > 0 Then
                userInitials = Trim(Mid(currentLine, InStr(currentLine,"=")+1))
            End If

            If InStr(currentLine, "pauseValue=") > 0 Then
                pauseValue = Trim(Mid(currentLine, Instr(currentLine,"=")+1))
                If pauseValue = 0 Or pauseValue > 5000 Then
                    Say "Please choose a pauseValue between 0 and 5001 in Settings"
                        WScript.sleep 3000
                        Set FSO = Nothing
                        Set wordObj = Nothing
                        WScript.Quit
                End If
            End If

            If InStr(currentLine, "greetingsMessage=") > 0 Then
                greetingsMessage = Trim(Mid(currentLine, InStr(currentLine,"=")+1))
            End If
            
            If InStr(currentLine, "continuingText=") > 0 Then
                continuingText = Trim(Mid(currentLine, Instr(currentLine,"=")+1))
                continuingTextArr = Split(continuingText, "_")
            End If

            If InStr(currentLine, "cassetteKey=") > 0 Then
                cassetteKey = Trim(Mid(currentLine, Instr(currentLine,"=")+1))
            End If

        Loop
        SettingsFile.Close

    Set FSO = Nothing
End Sub

Sub greeting

    If greetingsMessage = scriptVersion Then exit Sub

    Dim messageText, Answer
    messageText = "Please read carefully. This script will run until all of the blanks in your gross are completed or when your initials are detected. You can say 'Open settings' to change how this works." & vbcrlf & "This is entirely optional. DO NOT USE IF YOU ARE HAVING POOR VOICE RECOGNITION."

    Answer = Msgbox (messageText & vbcrlf & vbcrlf & "Do you want to stop seeing this message?",vbsystemmodal+vbyesno,scriptVersion)

    If Answer = vbYes Then
        WriteSettings "greetingsMessage=", scriptVersion
    End If

End Sub

Function WriteSettings(key,value)
    Dim SettingsFile, settingsText, settingsArr, FSO, Index
        
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set SettingsFile = FSO.OpenTextFile(settingsPath, 1)
    settingsText = SettingsFile.ReadAll
    SettingsFile.Close

    settingsArr = Split(settingsText, vbCrLf)

    For Index = 0 to Ubound(settingsArr)
        If InStr(settingsArr(Index), key) > 0 Then
            settingsArr(Index)=key & value
        End If
    Next

    Set SettingsFile = FSO.OpenTextFile(settingsPath, 2)

    For Index = 0 to Ubound(settingsArr)
        SettingsFile.WriteLine settingsArr(Index)
    Next

    SettingsFile.Close

    Set FSO = Nothing

End Function

Function checkCaseID
    Dim prop
	For Each prop In wordObj.ActiveDocument.CustomDocumentProperties
		If prop.Name = "CaseID" Then
			checkCaseID = prop.Value
			Exit For
		End If
	Next
exit function
    If Not Len(checkCaseID) = 7 Then 
        Say "Bad CaseID. Is this a PowerPath report?"
        WScript.sleep 3000
        Set wordObj = Nothing
        WScript.Quit
    End If
End Function

Sub AccessDB
    Dim objConnection, myRecordSet, logTime, dbTable, DatabasePath, adStateClosed, SQL
    Set objConnection = CreateObject("ADODB.Connection")
    Set myRecordSet = CreateObject("ADODB.Recordset")

    logTime = Now()

    If timer() - timerStart < 10 Then 
        dbTable = "NextItemDebug"
    Else
        dbTable = "NextItem"
    End If

    DatabasePath = "\\ohpathdragon01\voicebrook\Admin\Commands\STATS\Helping PAs.accdb"


    If objConnection.State = adStateClosed Then
        objConnection.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & DatabasePath & ";"
        objConnection.Open
    End If

    myRecordSet.ActiveConnection = objConnection
    SQL = "INSERT INTO " & dbTable & " ("


        SQL = SQL & "userInitials"
        SQL = SQL & "CaseID"
        SQL = SQL & ",advancedFields"        
        SQL = SQL & ",elapsedTime"
        SQL = SQL & ",logTime"

        SQL = SQL & ") VALUES ("

        SQL = SQL & "'" & userInitials & "',"
        SQL = SQL & "'" & CaseID & "',"
        SQL = SQL & "'" & advancedFields & "',"
        SQL = SQL & "'" & timer()-timerStart & "',"
        SQL = SQL & "'" & logTime & "'"

    'MsgBox SQL

    myRecordSet.Open SQL & ")"

    'myRecordSet.Close

    Set objConnection = Nothing
    Set myRecordSet = Nothing
End Sub

Sub blockAutofill
    dim quantityText, quantityTextArr
    quantityTextArr = Split(wordObj.ActiveDocument.Range.Text, "CLINICAL HISTORY:", 2)
    quantityText = quantityTextArr(0)

    dim blockText, specimenNumber, specimenQuantity, singleCase, currentSpecimenGross, numberOfPieces, extraBlocks, Index

        specimenQuantity = 1
        For Index = 1 To 50
            If InStr(quantityText, Index & ". ") > 0 Then specimenQuantity = Index
        Next

    blockText = wordObj.ActiveDocument.Range.Text
    Set blockText = wordObj.Selection.Range
    blockText.StartOf 6,1

    blockText = Mid(blockText,InStr(blockText, "GROSS DESCRIPTION:"))            'i forget the significance of this, but it causes intermittent errors so it's removed
    'find current specimen
    specimenNumber = 0
    For Index = 1 To specimenQuantity
        If InStr(blockText, vbCr & Index & ". ") = 0 Then Exit For 'changed from > 0
        specimenNumber = Index
    Next

    If specimenNumber = 0 then
        singleCase = True
        specimenNumber = 1
    End If
    
    If singleCase Then
        currentSpecimenGross = Split(blockText, vbCr)
    Else
        'msgbox specimenNumber
        currentSpecimenGross = Split(blockText, specimenNumber & ". The specimen")
dim currentSpecimenGross2
        currentSpecimenGross2 = Split(currentSpecimenGross(1), vbCr)
    End If



For Index = 0 to Ubound(currentSpecimenGross)
    'For Each Line in currentSpecimenGross
        If InStr(Index, "Number of pieces:") > 0 And InStr(Index, "[") = 0 Then
            'numberOfPieces, extraBlocks = Cint(Trim(Mid(Line, 17+InStr(Index, "Number of pieces:"))))
		numberOfPiecesText = Trim(Mid(Line, 17+InStr(Index, "Number of pieces:")))
		If Not IsNumeric(NumberOfPiecesText) Then Exit Sub
		numberOfPieces = Cint(numberOfPiecesText)
            Exit For
        End If
 Next

    'logic for Number of pieces in blocks, 5 pieces for each block
    If numberOfPieces < 6 Then
        extraBlocks = 0
    ElseIf numberOfPieces < 11 Then
        extraBlocks = 1
    ElseIf numberOfPieces < 16 Then
        extraBlocks = 2
    ElseIf numberOfPieces < 21 Then
        extraBlocks = 3
    ElseIf numberOfPieces < 26 Then
        extraBlocks = 4
    ElseIf numberOfPieces < 31 Then
        extraBlocks = 5
    End If

    dim alphabetArray, blockrange 
    alphabetArray = Array("A", "B", "C", "D", "E", "F")

    Set blockrange = wordObj.Selection.Range
    blockrange.End = blockrange.End + 10

    If InStr(blockrange, "in toto") > 0 And wordObj.Selection.Text = "[___]" Then 
        wordObj.Selection.Text = specimenNumber & "A"
        wordObj.Selection.Collapse 0
        wordObj.Selection.MoveRight
        If Not extraBlocks = 0 Then wordObj.Selection.Text = specimenNumber & AlphabetArray(extraBlocks) & "-"
    End If

End Sub
