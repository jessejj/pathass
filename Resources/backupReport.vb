Function backupReport(ByVal CommandName As String)

	Dim tempObj, FSO, File As Object
	Dim Prop
	Dim BackupInstructions, BUFile, BackupText As String
	Dim MaxSize As Long
	Dim PowerPathDoc As Boolean

	BUFile = "Report Backups.txt" 'location of the backup file
	MaxSize = 2097152 'bytes, this is the MaxSize that will trigger the backup file to be renamed, replacing the previous 'old' file

	On Error GoTo ErrSub

		Set tempObj = GetObject(,"Word.Application")
		BackupText = tempObj.ActiveDocument.Range.Text

		For Each prop In tempObj.ActiveDocument.CustomDocumentProperties
			If prop.Name = "ReturnToPowerPath" Then
				PowerPathDoc = True
				Exit For
			End If
		Next

		Set tempObj = Nothing

	If Not PowerPathDoc Then Exit Function

BackupInstructions = _
"Instructions:" & vbCrLf & _
"Text from reports that are edited on this computer are added to the end of this file each time certain Voiceover commands are used." & vbCrLf & _
"This attempts to provide multiple chronological backups that could be used with the PowerPath AMP log to troubleshoot report changes or loss." & vbCrLf & _
"After the file size reaches " & Round(MaxSize / 1048576) & "MB it will be renamed as an 'Old' file then soon replaced." & vbCrLf & _
"No personally identifiable information is saved. It is up to the person dictating to identify their own work from this file."

    Set FSO = CreateObject("Scripting.FileSystemObject")

		If Not FSO.FileExists(BUFile) Then

			Set File = FSO.CreateTextFile(BUFile)
			File.Write BackupInstructions
			File.Close

		End If
		
	Set File = FSO.GetFile(BUFile)

	If File.DateLastModified > Now()-(1/24/60) Then Exit Function 'a backup was made within the last minute

	Set File = FSO.OpenTextFile(BUFile, 8) 'open for appending

	File.Write vbCrLf & "___________________________" & vbCrLf & _
		"Backup: " & Now & " by " & CommandName & " Command" & vbCrLf & _
		Replace(BackupText, vbCr, vbCrLf) 'formatting adjustment
	File.Close

	Set File = FSO.GetFile(BUFile)

	If File.Size > MaxSize Then

		FSO.CopyFile BUFile, Replace(BUFile, ".txt", " Old.txt" ), True
		FSO.DeleteFile BUFile

    End If

    Set File = Nothing
    Set FSO = Nothing

	ErrSub:
End Function