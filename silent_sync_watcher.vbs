Set WshShell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

Function IsValidProjectDir(dirPath)
    IsValidProjectDir = False
    If dirPath <> "" Then
        If fso.FolderExists(dirPath) Then
            If fso.FileExists(dirPath & "\masterfile_template.xlsx") Or fso.FileExists(dirPath & "\Masterfile.xlsx") Or fso.FileExists(dirPath & "\app\cli\sync_vendors.py") Then
                IsValidProjectDir = True
            End If
        End If
    End If
End Function

projectDir = ""

' 1. Check directory where this script itself is located
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
If IsValidProjectDir(scriptDir) Then
    projectDir = scriptDir
End If

' 2. Check %APPDATA%\TaxiSoftware\project_dir.txt
If projectDir = "" Then
    appdata = WshShell.ExpandEnvironmentStrings("%APPDATA%")
    cfgFile = appdata & "\TaxiSoftware\project_dir.txt"
    If fso.FileExists(cfgFile) Then
        On Error Resume Next
        Set f = fso.OpenTextFile(cfgFile, 1)
        If Not f.AtEndOfStream Then
            savedDir = Trim(f.ReadLine())
            If IsValidProjectDir(savedDir) Then
                projectDir = savedDir
            End If
        End If
        f.Close
        On Error GoTo 0
    End If
End If

' 3. Check HKCU\Software\TaxiSoftware\ProjectPath
If projectDir = "" Then
    On Error Resume Next
    regDir = WshShell.RegRead("HKCU\Software\TaxiSoftware\ProjectPath")
    If IsValidProjectDir(regDir) Then
        projectDir = regDir
    End If
    On Error GoTo 0
End If

' 4. Check common locations relative to User Profile / Downloads / Desktop
If projectDir = "" Then
    userProfile = WshShell.ExpandEnvironmentStrings("%USERPROFILE%")
    candidates = Array( _
        userProfile & "\Downloads\Main Taxi\Taxi Software", _
        userProfile & "\Downloads\Taxi Software", _
        userProfile & "\Downloads\Main Taxi", _
        userProfile & "\Desktop\Taxi Software", _
        userProfile & "\Desktop\Main Taxi\Taxi Software", _
        userProfile & "\Documents\Taxi Software", _
        "C:\Taxi Software", _
        "D:\Taxi Software" _
    )
    For Each cand In candidates
        If IsValidProjectDir(cand) Then
            projectDir = cand
            Exit For
        End If
    Next
End If

' 5. If found, save to config and launch pythonw
If projectDir <> "" Then
    On Error Resume Next
    appdata = WshShell.ExpandEnvironmentStrings("%APPDATA%")
    cfgDir = appdata & "\TaxiSoftware"
    If Not fso.FolderExists(cfgDir) Then fso.CreateFolder(cfgDir)
    Set f = fso.CreateTextFile(cfgDir & "\project_dir.txt", True)
    f.WriteLine(projectDir)
    f.Close
    WshShell.RegWrite "HKCU\Software\TaxiSoftware\ProjectPath", projectDir, "REG_SZ"
    On Error GoTo 0

    WshShell.CurrentDirectory = projectDir

    pythonwExe = "C:\Python314\pythonw.exe"
    If Not fso.FileExists(pythonwExe) Then
        pythonwExe = "pythonw.exe"
    End If

    WshShell.Run """" & pythonwExe & """ -m app.cli.sync_vendors --watch", 0, False
End If


