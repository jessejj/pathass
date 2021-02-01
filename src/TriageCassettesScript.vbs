'This script was written to change the exit bin for cassettes printed and to prioritize smaller batches

Say "Hello!"
Say "Checking folders..."

Set shell = CreateObject("Wscript.Shell")

CassetteTriageFolder = "C:\Cassette Triage\"
OutputFolder = "\\CLPATHIF01\DIS_SHARE\" 'backslash needed
DefaultFolder = "\\CLPATHIF01\DIS_SHARE\"

Set FSO = CreateObject("Scripting.FileSystemObject")

If Not FSO.FolderExists(CassetteTriageFolder) Then
    Say CassetteTriageFolder & " does not exist. Creating..."
    FSO.CreateFolder(CassetteTriageFolder)
Else
    Say CassetteTriageFolder & " already exists. Continuing..."
End If

Wscript.Sleep 1000
If Not WScript.Arguments(0) = "ExitBin1" Then ChangeExitBin = True
If WScript.Arguments(0) = "ExitBin2" Then ExitBin = 2
If WScript.Arguments(0) = "ExitBin3" Then ExitBin = 3

If ChangeExitBin Then Say "The cassettes will be printed to Exit Bin " & ExitBin
Wscript.Sleep 3000

shell.RegWrite "HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\IMPAC\PseudoDriver\CIV-GRO-CAS1\v1.0\Directory Path", DefaultFolder, "REG_SZ"

If shell.RegRead( "HKLM\SOFTWARE\Wow6432Node\IMPAC\PseudoDriver\CIV-GRO-CAS1\v1.0\Directory Path" ) = DefaultFolder Then
    shell.RegWrite "HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\IMPAC\PseudoDriver\CIV-GRO-CAS1\v1.0\Directory Path", CassetteTriageFolder, "REG_SZ" 
Else
    msgbox "Unknown folder set"
End If

Say "Changed DIS settings. Please remember to click the 'Stop' button to revert settings."
Say "Now ready!"

Do

    If FSO.FolderExists(CassetteTriageFolder) And FSO.FolderExists(OutputFolder) Then

        Set Folder = FSO.GetFolder(CassetteTriageFolder)

        Set Files = Folder.Files

        If Files.Count = 0 Then

            Wscript.Sleep 1000

        Else

            CassetteCount = 0
            TotalCassettes = Files.Count
            If TotalCassettes > 4 Then LargeBatch = True Else LargeBatch = False 

            For Each file in Files

                CassetteCount = CassetteCount + 1

                Set ReadFile = FSO.OpenTextFile(file.path, 1)
                    text = ReadFile.ReadAll
                    ReadFile.close

                    If instr(text, "<STORE>1<>") = 0 Then 
                        Say file.path & " does not appear to be a cassette text file. Ending script..."
                        Wscript.Sleep 5000
                        Exit Do
                    End If
                
                If ChangeExitBin Then

                    Set WriteFile = FSO.OpenTextFile(file.path, 2)
                        text = Replace(text, "<STORE>1<>", "<STORE>" & ExitBin & "<>")
                        WriteFile.Write text
                        WriteFile.close

                End If
                    
                    file.move OutputFolder
                    Say file.name & " sent to printer. " & Files.Count & " remain. Please don't stop script."
                    If CassetteCount = TotalCassettes Then Say "Now safe to stop script."
                
                If LargeBatch Then
                    For Index = 1 to 11 'seconds (change this value to control how long it takes for the queue to be processed)
                        Wscript.Sleep 1000
                    Next
                End If

            Next

        End If

    Else

        Say "Target folder(s) do not exist. Ending script..."
        Wscript.Sleep 5000
        Exit Do

    End If

Loop

shell.RegWrite "HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\IMPAC\PseudoDriver\CIV-GRO-CAS1\v1.0\Directory Path", DefaultFolder, "REG_SZ"
    
Set FSO = Nothing
Set shell = Nothing
 
Function Say(message)
    WScript.StdOut.Write message
    WScript.StdOut.WriteLine
End Function
