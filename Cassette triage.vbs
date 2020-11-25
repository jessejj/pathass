'This script was written to help pathology technicians print cassettes while prioritizing cassettes printed by PAs

Set shell = CreateObject("Wscript.Shell")

CassetteTriageFolder = "C:\Cassette Triage\"

Set FSO = CreateObject("Scripting.FileSystemObject")

If Not FSO.FolderExists(CassetteTriageFolder) Then
    FSO.CreateFolder(CassetteTriageFolder)
End If

OutputFolder = "\\CLPATHIF01\DIS_SHARE"
DefaultFolder = "\\CLPATHIF01\DIS_SHARE"

shell.RegWrite "HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\IMPAC\PseudoDriver\CIV-GRO-CAS1\v1.0\Directory Path", DefaultFolder, "REG_SZ"

If shell.RegRead( "HKLM\SOFTWARE\Wow6432Node\IMPAC\PseudoDriver\CIV-GRO-CAS1\v1.0\Directory Path" ) = DefaultFolder Then
    shell.RegWrite "HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\IMPAC\PseudoDriver\CIV-GRO-CAS1\v1.0\Directory Path", InputFolder, "REG_SZ" 
Else
    msgbox "Unknown folder set"
End If

Do

    If shell.AppActivate("AMP Grossing Station Warning") Then
        'if warning appears, dismiss it then start the script

        Wscript.Sleep 500 'delay required before dismissing prompt

        shell.Sendkeys "y", True
        shell.Sendkeys "~", True

        Wscript.Sleep 1000

        if shell.AppActivate("Confirm") then
            shell.Sendkeys "n", True
            shell.Sendkeys "~", True
            Wscript.Sleep 1000
        end if

        shell.AppActivate("PowerPath Advanced")

        shell.Sendkeys "%p", True

        Wscript.Sleep 1000

    End If

    If FSO.FolderExists(CassetteTriageFolder) and FSO.FolderExists(OutputFolder) Then

        Set Folder = FSO.GetFolder(CassetteTriageFolder)

        Set Files = Folder.Files

            If Files.Count > 0 Then

                For Each file in Files
                    filelist = filelist & file.name & vbcrlf

                    Set ReadFile = FSO.OpenTextFile(file.path, 1)
                        text = ReadFile.ReadAll
                        ReadFile.close
                    
                    Set WriteFile = FSO.OpenTextFile(file.path, 2)
                        text = Replace(text, "<STORE>1<>", "<STORE>2<>")
                        WriteFile.Write text
                        WriteFile.close
                        
                        file.move OutputFolder
                Say "Sent 1 cassette to labelase for bin 2"
                       
                
                    For Index = 1 to 10 'seconds (change this value to control how long it takes for the queue to be processed)
                        Wscript.Sleep 1000
                    
                        If shell.AppActivate("AMP Grossing Station Warning") Then
                            Wscript.Sleep 1000 
                            shell.Sendkeys "y", True
                            shell.Sendkeys "~", True

                        End If

                    Next

                Next
                
            Else

            Wscript.Sleep 1000

            End If

        Else

        msgbox "target folder(s) do not exist"
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
