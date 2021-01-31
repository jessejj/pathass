Public Function checkReport()
	'On Error GoTo ErrorHandler
	'On Error Resume Next
	Dim wordObj, FSO As Object
	Dim report_text, Headers, missing_headers, missing_placeholders, user_message As String
	Dim prop, ReportTextArr, Laterality
	Dim Index As Integer
	Dim PowerPathReport, ProblemsFound As Boolean


	Set wordObj = GetObject(,"Word.Application")
	wordObj.StatusBar = "checking report..."
		If wordObj.Documents.Count = 0 Then Exit All

			For Each prop In wordObj.ActiveDocument.CustomDocumentProperties
				If prop.Name = "ReturnToPowerPath" Then
					PowerPathReport = True
					Exit For
				End If
			Next

	Dim mystringArr, AccessionNo, mystring

		Set FSO = CreateObject("Scripting.FileSystemObject")
			mystring = FSO.OpenTextFile(wordObj.ActiveDocument.FullName, 1, False).ReadAll
		Set FSO = Nothing

		mystringArr = Split(mystring, "Accession #:", 2)
		mystringArr = Split(mystringArr(1), "charrsid6755835 ", 2)
		mystringArr = Split(mystringArr(1), "\cell", 2)
		AccessionNo = mystringArr(0)
		AccessionNo = Replace(AccessionNo, vbCrLf, "")

	If InStr(AccessionNo, "MHS-")+InStr(AccessionNo, "SP-")+InStr(AccessionNo, "CCS-")+InStr(AccessionNo, "QCS-") = 0 Then
		wordObj.StatusBar = "Report Not Supported. Report check not available"
		Set wordObj = Nothing
		Exit function
	End if

	report_text = wordObj.ActiveDocument.Range.Text
	ReportTextArr = Split(wordObj.ActiveDocument.Range.Text, "CLINICAL HISTORY:", 2)


	If InStr(AccessionNo, "CCS-") > 0 Then 'It's a CCS Report
		Headers = Split("DIAGNOSIS:, SYNOPTIC REPORT:, GROSS DESCRIPTION:", ", ")
	Else 'It's either MHS or SP 
		Headers = Split("DIAGNOSIS:, MICROSCOPIC DESCRIPTION:, GROSS DESCRIPTION:, COMMENTS:", ", ")
	End If

	For Index = 0 to UBound(Headers)

		If InStr(report_text, Headers(Index)) = 0 Then
			If Not ProblemsFound Then ProblemsFound = True
			missing_headers = missing_headers & Headers(Index) & vbcrlf
		End If
		
	Next

Dim err, PatientName, PatientNameArr, LastName, FirstName, Initial1, Initial2, initials_text, FirstCheck, SecondCheck, InitialsCheckFailed, grossArr, InitialsMsg, SpecimenQuantity, specimen_source_text, specimenSiteArr, gross_description, Index2, Index3, SpecimenNumber, missing_laterality
Dim InitialCheckCounter as Integer

	'isolate gross description, formatting, split into lines
	gross_description = Mid(report_text,InStr(report_text, "GROSS DESCRIPTION:"))
	If InStr(gross_description, "MICROSCOPIC DESCRIPTION:") > 0 Then gross_description = Left(gross_description, InStr(gross_description, "MICROSCOPIC DESCRIPTION:"))
	gross_description = Replace(gross_description, "with the initials of ", "with the initials ")
	grossArr = Split(gross_description, vbCr)

    'Find Patient Name
    mystringArr = Split(mystring, "Patient:", 2)
    mystringArr = Split(mystringArr(1), "charrsid6755835 ", 2)
    mystringArr = Split(mystringArr(1), "\cell", 2)
    PatientName = mystringArr(0)
    PatientNameArr = Split(PatientName, ", ", 2)
    LastName = PatientNameArr(0)
    FirstName = PatientNameArr(1)
	Initial1 = Left(LastName, 1)
	Initial2 = Left(FirstName, 1)

    For Index = 1 To UBound(grossArr)

        If InStr(grossArr(Index), "initials") > 0 And InStr(grossArr(Index), "specimen is received") > 0 And InStr(grossArr(Index), "[") = 0 Then 'check this line

			InitialCheckCounter = InitialCheckCounter + 1

            initials_text = Mid(grossArr(Index),InStr(grossArr(Index), "initials")+9)
            initials_text = Left(initials_text, 6)

            FirstCheck = InStr(initials_text, Initial1)

			If Initial1 = Initial2 Then 'Remove the first one for the secondcheck
				SecondCheck = InStr(Replace(initials_text, Initial1, "",, 1), Initial2)
			Else
            	SecondCheck = InStr(initials_text, Initial2)
			End If

			If FirstCheck = 0 Or SecondCheck = 0 Or FirstCheck > SecondCheck Then
				InitialsCheckFailed = True
				ProblemsFound = True
				Exit For
			End If

        End If
    Next

    If InitialsCheckFailed Then
		Beep
	        With wordObj.Selection.Find
			    .Text = initials_text
			    .Forward = True
			    .Wrap = 1
			    .Format = False
			    .MatchCase = False
			    .MatchWholeWord = False
			    .MatchAllWordForms = False
			    .MatchSoundsLike = False
			    .MatchWildcards = False
			End With
			wordObj.Selection.Find.Execute
			wordObj.Selection.Expand 3
			Wait 1
			wordObj.Selection.Collapse
		InitialsMsg = InitialsMsg & "Please double check the container and review dictated initials." & vbCrLf & vbCrLf & "Patient name is " & PatientName & vbCrLf & vbCrLf & "Expected initials '" & Initial1 & Initial2 & "'."
    End If

	If InitialsCheckFailed Then 
			user_message = user_message & InitialsMsg & vbCrlf & vbcrlf
		Beep
		'WakeUp
		SetMicrophone 1
		Msgbox user_message, vbSystemModal, "VoiceOver - Report Check"
		AppActivate "Case"
	Else
		If err = 0 then 

			If InitialCheckCounter = 1 Then wordObj.StatusBar = "Dictated Patient Initials Checked on 1 line."
			If InitialCheckCounter > 1 Then wordObj.StatusBar = "Dictated Patient Initials Checked on " & InitialCheckCounter & " lines."

		End If
	End If

	'Get SpecimenQuantity from the Specimen source text
	If InStr(report_text, "CLINICAL HISTORY:") > 0 Then
		specimen_source_text = Left(report_text, InStr(report_text, "CLINICAL HISTORY:"))
	Else
		specimen_source_text = Left(report_text, InStr(report_text, "DIAGNOSIS:"))
	End If

	specimen_source_text = Mid(specimen_source_text, InStr(specimen_source_text, "SPECIMEN SOURCE:"))
	specimen_source_text = Replace(specimen_source_text, vbTab, "")
	specimen_source_text = Replace(specimen_source_text, "SPECIMEN SOURCE: ", "")

	SpecimenQuantity = 1
	If Not InStr(ReportTextArr(0),"SPECIMEN SOURCE:  	1.") > 0 Then
		SpecimenQuantity = 1
	Else
		For Index = 1 To 50
			If InStr(ReportTextArr(0), vbTab & Index & ". ") > 0 Then SpecimenQuantity = Index
		Next
	End If

	'If Not SpecimenQuantity = 1 Then specimenSiteArr = Split(specimen_source_text, vbCr)
	specimenSiteArr = Split(specimen_source_text, vbCr)

	Laterality = Split("left, right, superior, inferior, lateral, medial, anterior, ascending, descending, posterior, proximal, distal, base, apex", ", ")

	For Index = 0 to UBound(grossArr)

		If InStr(grossArr(Index), "initials") > 0 And InStr(grossArr(Index), "specimen is received") > 0 and InStr(grossArr(Index), "[") = 0 Then 'check this line

			'figure out which specimen it is
			If Not SpecimenQuantity = 1 Then
				For Index3 = 1 To SpecimenQuantity
					SpecimenNumber = Index3
					If InStr(grossArr(Index), Index3 & ". ") > 0 Then Exit For
					
				Next
			Else
				SpecimenNumber = 1
			End If

			For Index2 = 0 To UBound(Laterality)
				'for each laterality check to see if it's found in the specimen site, if it is then check for its presence in the grossArr
				If InStr(LCase(specimenSiteArr(SpecimenNumber-1)), Laterality(Index2)) > 0 Then
					If InStr(LCase(grossArr(Index)), Laterality(Index2)) = 0 Then
						If Not ProblemsFound Then ProblemsFound = True
						missing_laterality = missing_laterality & "Expected '" & Laterality(Index2) & "' for specimen " & SpecimenNumber & vbcrlf
					End If
				End If
			Next

		End If	

	Next

	dim missing_block
	

	If Len(missing_block) > 0 Then ProblemsFound = True

	If ProblemsFound Then 
		If Len(missing_laterality) > 0 Then 
			wordObj.StatusBar = "WARNING Please double check the specimen sites(s) for laterality: " & missing_laterality
		ElseIf Len(missing_block) > 0 Then
			wordObj.StatusBar = "WARNING your cassette key may be missing block(s): " & missing_block
		End if
	End If

	If InitialsCheckFailed Then 'log it into the QA database
		Dim objConnection, myRecordSet, ReportText, Initials, SaveTime, DatabasePath, SQL, adStateClosed

		Set objConnection = CreateObject("ADODB.Connection")
		Set myRecordSet = CreateObject("ADODB.Recordset")

		Set wordObj = GetObject(,"Word.Application")
		ReportText = wordObj.ActiveDocument.Range.Text
		ReportText = Replace(ReportText, Chr(39), Chr(34)) 'REQUIRED
		Initials = vbreadini(c_vbvouserini,"POWERPATH","MYINITIALS")
		SaveTime = Now()
		initials_text = Replace(initials_text, "'", " ") 'remove quotes for SQL
		initials_text = Replace(initials_text, Chr(34), " ") 'remove quotes for SQL
		ReportText = Replace(ReportText, "'", " ") 'remove quotes for SQL
		ReportText = Replace(ReportText, Chr(34), " ") 'remove quotes for SQL

		DatabasePath = "QA Statistics.accdb"


		If objConnection.State = adStateClosed Then
			objConnection.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & DatabasePath & ";"
			objConnection.Open
		End If

		myRecordSet.ActiveConnection = objConnection
		SQL = "INSERT INTO InitialsCheck ("
			SQL = SQL & "UserInitials"
			SQL = SQL & ",AccessionNo"
			SQL = SQL & ",ProblemLine"
			SQL = SQL & ",ExpectedInitials"
			SQL = SQL & ",ReportText"
			SQL = SQL & ",SaveTime"

			SQL = SQL & ") VALUES ("

			SQL = SQL & "'" & Initials & "',"
			SQL = SQL & "'" & AccessionNo & "',"
			SQL = SQL & "'" & initials_text & "',"			
			SQL = SQL & "'" & Initial1 & Initial2 & "',"
			SQL = SQL & "'" & ReportText & "',"
			SQL = SQL & "'" & SaveTime & "'"

		myRecordSet.Open SQL & ")"


		Set objConnection = Nothing
		Set myRecordSet = Nothing

	
	End If


	If InitialsCheckFailed Then Exit All


	exit function
	ErrorHandler:
	wordObj.StatusBar = "unspecified error occurred while checking report"


End Function