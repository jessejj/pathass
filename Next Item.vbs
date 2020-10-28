Dim scriptVersion
scriptVersion = "Next Item Script 2020.10.28" 'by Jesse Pelley

'This script was written to help Pathologist Assistants describe specimens more effeciently while they dissect

pauseTime=250           'in milliseconds
cassetteKeyFill=True    'autofills the first block and guesses the last block in a range based on the number of pieces

Dim pauseSymbols(7)   'symbols that imply that the user does not want to move to the next item
pauseSymbols(0) = ","
pauseSymbols(1) = ";"
pauseSymbols(2) = "("
pauseSymbols(3) = "-"
pauseSymbols(4) = "?"
pauseSymbols(5) = "."
pauseSymbols(6) = ":"
pauseSymbols(7) = "/"

Dim pauseWords(4)
pauseWords(0) = " and"
pauseWords(1) = " by"
pauseWords(2) = " the"
pauseWords(3) = " to"
pauseWords(4) = " that"


Dim numberofBlanks, previousnumberofBlanks, advancedFields, CaseID

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

    previousnumberofBlanks = numberofBlanks
    numberofBlanks = countBlanks

    WScript.Sleep pauseTime

    If incompleteParagraph Then previousnumberofBlanks = previousnumberofBlanks - 1

    If numberofBlanks < previousnumberofBlanks Then
        nextItem
        Say "There are " & numberofBlanks & " blank fields left in the gross description"
        WaitMessage = True 'for when we need to wait again

    Else
        If WaitMessage = True Then
            Say "Waiting..."
            WaitMessage = False
        End If 
        waitLoop
    End If

    If cassetteKeyFill then blockAutofill

Loop Until numberofBlanks = 0

Say scriptVersion & " Finished. " & advancedFields & " fields cleared."
Set wordObj = Nothing
WScript.Sleep 3000

Function waitLoop

    Do 
        WScript.Sleep pauseTime 
        If InStr(wordObj.Selection.Text, "]") = 0 Then 'Fill in not present or filled out, so stop waiting
            Say "Resuming..."
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

    WScript.StdOut.Write Chr(7)     'beep
    End If

End Sub

Function incompleteParagraph
    dim lastCharacters, Index, lastWord
    Set lastCharacters = wordObj.Selection.Range
    lastCharacters.Start = lastCharacters.Start - 2

    For Index = 0 to Ubound(pauseSymbols)
        If InStr(lastCharacters, pauseSymbols(Index)) > 0 Then
            incompleteParagraph = True
            Say "Paused because of '" & pauseSymbols(Index) & "'"
            Exit For
        End If
    Next

    Set lastWord = wordObj.Selection.Range
    lastWord.Start = lastWord.Start - 5

    For Index = 0 to Ubound(pauseWords)
        If InStr(lastCharacters, pauseWords(Index)) > 0 Then
            incompleteParagraph = True
            Say "Paused because of '" & pauseWords(Index) & "'"
            Exit For
        End If
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

Function checkCaseID
    Dim prop
	For Each prop In wordObj.ActiveDocument.CustomDocumentProperties
		If prop.Name = "CaseID" Then
			checkCaseID = prop.Value
			Exit For
		End If
	Next

    If Not Len(checkCaseID) = 7 Then 
        Say "Bad CaseID. Is this a PowerPath report?"
        WScript.sleep 3000
        Set wordObj = Nothing
        WScript.Quit
    End If
End Function

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
        arrayofLines = Split(blockText, vbCr)
    Else
        currentSpecimenGross = Split(blockText, specimenNumber & ". The specimen")
        arrayofLines = Split(currentSpecimenGross(1), vbCr)
    End If

    For Index = 0 to Ubound(arrayofLines)

        If InStr(arrayofLines(Index), "Number of pieces:") > 0 And InStr(arrayofLines(Index), "[") = 0 Then
		numberOfPiecesText = Mid(arrayofLines(Index), 17+InStr(arrayofLines(Index), "Number of pieces:"))

        If WaitMessage = True Then Say "Number of pieces:" & numberOfPiecesText 

		If Not IsNumeric(NumberOfPiecesText) Then 
            If WaitMessage = True Then Say "Not Numeric"    
            numberOfPieces = 0
        Else 
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
