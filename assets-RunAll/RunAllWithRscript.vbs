
Const strcAppName = "Run Rscript Cleanup"
'Const gstrLogHeader = """UserName"",""MachineName"",""Date Time"",""Local Drive"",""Local Drive Free"","" Local Drive Available"",""Local Drive Percent Available"","" Local Drive FileSystem"",""CPU Max Clock Speed"",""Machine Asset Tag"",""Available RAM Gb"",""Total RAM Gb"",""Horizontal Screen Resolution"","" Vertical Screen Resolution"","" Dpi Width"","" Dpi Height"",""OS Bit"",""Office Version"""
Const gstrLogHeader = """ProjectName"",""UserName"",""MachineName"",""Date Time"",""Local Drive"",""Local Drive Free"","" Local Drive Available"",""Local Drive Percent Available"","" Local Drive FileSystem"",""CPU Max Clock Speed"",""Machine Asset Tag"",""Available RAM Gb"",""Total RAM Gb"",""OS Bit"",""Office Version"""
'For Placing MsgBox on top
Const mdblcAllwaysOnTop = 262144

'------------for Progress bar in hta --------------
'Public w
'Public x
'Public y
Public MyTitle
'Public fIncrementOnHtmlWrite

'w = 100
'x = 0
'y = 100
MyTitle = " Progress"

'For pauses in code to allow for document updates after document finishes loading
'Public idTimer
'Initiate Actions, while allowing page to finalize
'idTimer = Window.setTimeout("ExecuteActions", 50, "VBScript")
    '===========================================================
    Sub ExecuteActions()
    '===========================================================
        'Clear Timer as we only need to ExecuteActions once
        Window.clearTimeout (idTimer)

        'Cleanup Local Variables
        Const HKEY_CURRENT_USER = &H80000001
        Const HKEY_LOCAL_MACHINE = &H80000002
        Const strcConfigIniFileName = "RunAll.ini"
        
        'Windows Folder Namespaces:
        'https://technet.microsoft.com/en-us/library/ee176604.aspx
        Const ssfMyDocuments = &H5  ' My Documents: C:\Users\%username%\Documents
        Const ssfStartup = &H7  ' Startup: C:\Users\%username%\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup
        Const ssfUserProfile = &H28 'User Profile: C:\Users\%username%
        Const ssfLocalAppdata = &H1C ' Local: C:\Users\%username%\AppData\Local  

        'Define Object Variables        
        Dim wshShell ' As IWshRuntimeLibrary.WshShell
        Set wshShell = CreateObject("WScript.Shell") ' New IWshRuntimeLibrary.wshShell   
        Dim wshExec ' As IWshRuntimeLibrary.WshExec
        Dim wshShortcut ' As IWshRuntimeLibrary.WshShortcut
        Dim objShell ' As Shell32.Shell
        Dim Folder ' As fso.Folder
        Dim fs ' As Scripting.FileSystemObject
        Dim objNameSpace
        Set fs = CreateObject("Scripting.FileSystemObject")
        Set objShell = CreateObject("Shell.Application")

        ' -------------------------------------------------------------------------
        ' GetIniValues
        ' -------------------------------------------------------------------------
        Dim strConfigIniFilePath
        strConfigIniFilePath = GetCurrentPath & "\" & strcConfigIniFileName
        Dim strLogPath
        strLogPath  = GetIniValue(strConfigIniFilePath, "Log", "LogLocation")
        Dim strProjectName
	strProjectName = GetIniValue(strConfigIniFilePath, "Config", "ProjectName")
        Dim strRunScriptName
        strRunScriptName = GetIniValue(strConfigIniFilePath, "Config", "RunScriptName")

        ' -------------------------------------------------------------------------
        ' Execute Scripts
        ' -------------------------------------------------------------------------
        'WriteHtmlLine "<h1>Execute Scripts</h1>"
        'WriteHtmlLine "<h2>" & strProjectName & "</h2>"

    Dim strRscriptExe
    Dim intPathNumber
    intPathNumber = 1
    Do 
        '"%userprofile%\Documents\R-3.5.1\bin\x64\Rscript.exe"   
        strRscriptExe =  wshShell.ExpandEnvironmentStrings( _
            GetIniValue( _
                strConfigIniFilePath, _
                "Rscript", _
                intPathNumber _
            ) _
        )
        intPathNumber = intPathNumber + 1
    Loop Until intPathNumber > 100 Or FileExists(strRscriptExe)
    'WriteHtmlLine "Found Rscript.exe at " & strRscriptExe & "<br/>"
    
    Dim strFolder
    strFolder = ParentFolder(GetCurrentPath)
    'WriteHtmlLine "Using script folder " & strFolder & "<br/>"
    Dim strCommand
    strCommand = """" & strRscriptExe & """ """ & strFolder & "\" & strRunScriptName & """"
    'WriteHtmlLine "Running command: " & strcommand  & "<br/>"
    on error resume next
    wshShell.Run strCommand, 0, True
    if err.number <> 0 then
        msgbox "Erorr:" & err.Number & vbcrlf & err.Description, ,  strProjectName
    end if
    ' ------------------------------------------------------------------------- 
    ' Logging use 
    ' ------------------------------------------------------------------------- 
    If FolderExists(ParentFolder(strLogPath)) Then
        'x = 85 
        'WriteHtmlLine "<h1>Logging Script Run</h1>" 
        'Gather Machine Information (calling WriteToLog enters the UserName, Machine Name, and ...)
        Dim strLogDetails
        'WriteHtmlLine _
        '    "<br/>" & "-------------Drives In Use-------------" & "<br/>" & _
        'GetMachineDrives & GetMachineNetworkDrives(False)
        'Add HDD Free Space and other machine information to a WipeLog on the Network
        'We each item on a seperate line to ensure that we have the correct number of commas
        strLogDetails = strProjectName
        strLogDetails = strLogDetails & ",""" 
        strLogDetails = strLogDetails & GetMachineLogicalDriveInformationForLog 'Local Drive,Local Drive Free, Local Drive Available ,Local Drive Percent Available, Local Drive FileSystem
        strLogDetails = strLogDetails & ",""" 
        strLogDetails = strLogDetails & GetMaxClockSpeed 'CPU Max Clock Speed
        strLogDetails = strLogDetails & """,""" 
        strLogDetails = strLogDetails & GetMachineAssetTag 'Machine Asset Tag
        strLogDetails = strLogDetails & """," 
        strLogDetails = strLogDetails & GetTotalAndAvailableRamForLog 'Available RAM Gb,Total RAM Gb
        strLogDetails = strLogDetails & ",""" 
        strLogDetails = strLogDetails & GetOSBitNumber 'OS Bit
        strLogDetails = strLogDetails & """,""" 
        strLogDetails = strLogDetails & GetOfficeVersion 'Office Version
        strLogDetails = strLogDetails & """" 
        'WriteHtmlLine "Writing LogDetails to Log Path:" & strLogPath & "<br/>" & gstrLogHeader &  "<br/>" & strLogDetails 
        WriteToLog strLogPath, strLogDetails             
    End If

    Dim strCompletionNotice
    strCompletionNotice = strProjectName  & " Complete"
    'x = 94
    'WriteHtmlLine "<h2>" & strCompletionNotice & "</h2>"
    'Close window in 5 seconds
    Dim intSecondsRemaining
    intSecondsRemaining = 5        
    'WriteHtmlLine "<br/> Closing window in "
    Do 
        'WriteHtmlLine intSecondsRemaining & "..." 
        WaitPingTimes 1
        intSecondsRemaining = intSecondsRemaining - 1
    Loop until intSecondsRemaining < 1
    Window.Close
End Sub 

    '===========================================================
    'reading value in Registry
    Function ReadingValueInReg(root, strKeyPath, strValueName, intValueType)
	on error resume next
        ' if an error occurs, then we just return null instead of reading, the calling code will handle that.
        Dim strComputer: strComputer = "."
        Dim oReg: Set oReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")
        Dim isCreateOK: isCreateOK = oReg.CreateKey(root, strKeyPath)
        Dim arrValues, rregResult, Value, i, strValue
        If isCreateOK = 0 Then
            Select Case intValueType
                Case 1
                    oReg.GetStringValue root, strKeyPath & "\", strValueName, arrValues
                    rregResult = arrValues
                Case 2
                    oReg.GetDWORDValue root, strKeyPath & "\", strValueName, arrValues
                    rregResult = arrValues
                Case 3
                    oReg.GetMultiStringValue root, strKeyPath & "\", strValueName, arrValues
                    For Each Value In arrValues
                        rregResult = rregResult & vbCrLf & Value
                    Next
                Case 4
                    oReg.GetExpandedStringValue root, strKeyPath & "\", strValueName, arrValues
                    rregResult = arrValues
                Case 5
                    oReg.GetBinaryValue root, strKeyPath & "\", strValueName, arrValues
                    For i = LBound(strValue) To UBound(strValue)
                        rregResult = rregResult & vbCrLf & strValue(i)
                    Next
                Case Else
                rregResult = ""
            End Select
        End If
        ReadingValueInReg = rregResult
exitFunction:
    End Function

    '===========================================================
    'Does this file exist?
    Function FileExists(strPath)
	on error resume next
        Dim ObjFSO
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	FileExists = objFSO.FileExists(strPath)
    End Function

    '===========================================================
    'Does this folder exist?
    Function FolderExists(strPath)
	on error resume next
        Dim ObjFSO
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	FolderExists = objFSO.FolderExists(strPath)
    End Function


    '===========================================================
    ' Get parent folder name
    Function ParentFolder(strPath)
	on error resume next
        Dim ObjFSO
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	ParentFolder = objFSO.GetParentFolderName(strPath)
    End Function

    '===========================================================
    'get office version
    Function GetOfficeVersion()
        'Constants for Registry Key Roots
        Const HKEY_CLASSES_ROOT = &H80000000
        Const HKEY_CURRENT_USER = &H80000001
        Const HKEY_LOCAL_MACHINE = &H80000002
        Const HKEY_USERS = &H80000003
        Const HKEY_CURRENT_CONFIG = &H80000005
    
        Dim officeVersion
        Dim excelFilePath
        Dim powerpntFilePath
        Dim puplishFilePath
        Dim outlookFilePath
        
        Dim subObjFSO: Set subObjFSO = CreateObject("Scripting.FileSystemObject")
        'On Error Resume Next
        Dim wordFilePath: wordFilePath = ReadingValueInReg(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Winword.exe", "", 1)
        If Trim(wordFilePath) <> "" Then
          If subObjFSO.FileExists(wordFilePath) Then
             officeVersion = subObjFSO.GetFileVersion(wordFilePath)
          End If
        Else
          excelFilePath = ReadingValueInReg(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe", "", 1)
          If Trim(excelFilePath) <> "" Then
             If subObjFSO.FileExists(excelFilePath) Then
                officeVersion = subObjFSO.GetFileVersion(excelFilePath)
             End If
          Else
             powerpntFilePath = ReadingValueInReg(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\powerpnt.exe", "", 1)
             If Trim(powerpntFilePath) <> "" Then
                If subObjFSO.FileExists(powerpntFilePath) Then
                    officeVersion = subObjFSO.GetFileVersion(powerpntFilePath)
                End If
             Else
                puplishFilePath = ReadingValueInReg(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\MSPUB.EXE", "", 1)
                If Trim(puplishFilePath) <> "" Then
                    If subObjFSO.FileExists(puplishFilePath) Then
                      officeVersion = subObjFSO.GetFileVersion(puplishFilePath)
                    End If
                Else
                    outlookFilePath = ReadingValueInReg(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.exe", "", 1)
                    If Trim(outlookFilePath) <> "" Then
                      If subObjFSO.FileExists(outlookFilePath) Then
                         officeVersion = subObjFSO.GetFileVersion(outlookFilePath)
                      End If
                    Else
                      If GetOSBitNumber() = 64 Then
                         wordFilePath = ReadingValueInReg(HKEY_LOCAL_MACHINE, "SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\App Paths\Winword.exe", "", 1)
                         If Trim(wordFilePath) <> "" Then
                            If subObjFSO.FileExists(wordFilePath) Then
                                officeVersion = subObjFSO.GetFileVersion(wordFilePath)
                            End If
                         Else
                            excelFilePath = ReadingValueInReg(HKEY_LOCAL_MACHINE, "SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\App Paths\excel.exe", "", 1)
                            If Trim(excelFilePath) <> "" Then
                                If subObjFSO.FileExists(excelFilePath) Then
                                  officeVersion = subObjFSO.GetFileVersion(excelFilePath)
                                End If
                            Else
                                powerpntFilePath = ReadingValueInReg(HKEY_LOCAL_MACHINE, "SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\App Paths\powerpnt.exe", "", 1)
                                If Trim(powerpntFilePath) <> "" Then
                                  If subObjFSO.FileExists(powerpntFilePath) Then
                                     officeVersion = subObjFSO.GetFileVersion(powerpntFilePath)
                                  End If
                                Else
                                  puplishFilePath = ReadingValueInReg(HKEY_LOCAL_MACHINE, "SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\App Paths\MSPUB.EXE", "", 1)
                                  If Trim(puplishFilePath) <> "" Then
                                     If subObjFSO.FileExists(puplishFilePath) Then
                                        officeVersion = subObjFSO.GetFileVersion(puplishFilePath)
                                     End If
                                  Else
                                     outlookFilePath = ReadingValueInReg(HKEY_LOCAL_MACHINE, "SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.exe", "", 1)
                                     If Trim(outlookFilePath) <> "" Then
                                        If subObjFSO.FileExists(outlookFilePath) Then
                                            officeVersion = subObjFSO.GetFileVersion(outlookFilePath)
                                        End If
                                     End If
                                  End If
                                End If
                            End If
                         End If
                      End If
                    End If
                End If
             End If
          End If
        End If
    
        On Error GoTo 0
    
        GetOfficeVersion = officeVersion
    End Function

    '===========================================================
    Function GetMachineAssetTag()
    On Error Resume Next
        'Returns null if not set
        Err.Clear
        Dim strComputer: strComputer = "."
        Dim objWMIServices: Set objWMIServices = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
        Dim colSMBIOS: Set colSMBIOS = objWMIServices.ExecQuery("Select SMBIOSAssetTag from Win32_SystemEnclosure")
        Dim objSMBIOS
        Dim strGetMachineAssetTag
        For Each objSMBIOS In colSMBIOS
            strGetMachineAssetTag = objSMBIOS.SMBIOSAssetTag
        Next
        If Err.Number <> 0 Then
            GetMachineAssetTag = "Err"
        Else
            GetMachineAssetTag = strGetMachineAssetTag
        End If
        Err.Clear
        On Error GoTo 0
    End Function
    
    '===========================================================
    Function GetOSBitNumber()
    On Error Resume Next
    Err.Clear
        Dim wshShell: Set wshShell = CreateObject("WScript.Shell")
        Dim KeyValue: KeyValue = "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment\PROCESSOR_ARCHITECTURE"
        Dim strVersion: strVersion = wshShell.RegRead(KeyValue)
        Dim osVersionBitNo
        If InStr(strVersion, "86") > 0 Then
            osVersionBitNo = 32
        Else
            osVersionBitNo = 64
        End If
        If Err.Number <> 0 Then
            GetOSBitNumber = 0
        Else
            GetOSBitNumber = osVersionBitNo
        End If
        Err.Clear
        On Error GoTo 0
        Set wshShell = Nothing
    End Function

    '===========================================================
    Function GetMaxClockSpeed()
    On Error Resume Next
        Err.Clear
        Dim strComputer: strComputer = "."
        Dim strGetMaxClockSpeed
        Dim objWMIService: Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
        Dim colItems, objItem: Set colItems = objWMIService.ExecQuery("Select MaxClockSpeed from Win32_Processor")
    
        For Each objItem In colItems
            strGetMaxClockSpeed =  objItem.MaxClockSpeed
        Next
        GetMaxClockSpeed = strGetMaxClockSpeed
        Err.Clear
        On Error GoTo 0
    End Function
    
    '===========================================================
    Function GetTotalAndAvailableRamForLog()
    On Error Resume Next
    Dim strLogonUser
    Dim strComputer
    Dim objWMIService
    Dim perfData
    Dim entry
    Dim strRam
    
        strComputer = "."
        Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
        
'        Set objWMIService = GetObject("winmgmts:" _
'        & "{impersonationLevel=impersonate}!\\" _
'        & strComputer & "\root\cimv2")
       
        'Get Available Ram
        Set perfData = objWMIService.ExecQuery _
            ("Select * from Win32_PerfFormattedData_PerfOS_Memory")
        For Each entry In perfData
            strRam = Round((entry.AvailableBytes / 2^30), 2)
        Next
	
    
        'Get Total Ram
        Set perfData = objWMIService.ExecQuery _
            ("Select * from Win32_Computersystem")
        For Each entry In perfData
            strRam = strRam & """,""" 
            strRam = strRam & Round((entry.TotalPhysicalMemory / 2^30), 2)
        Next
        GetTotalAndAvailableRamForLog = """" & strRam & """"
        Set objWMIService = Nothing
        'Set WS = Nothing
    End Function

    '===========================================================
    Sub WriteToLog(strFilePath, strText)
    On Error Resume Next
        If Len(strFilePath) > 1 Then
            'Failure to Log will not stop this script
            Const ForAppending = 8
            Dim objNetwork
            Dim fso
            Dim tf
            Set objNetwork = CreateObject("WScript.Network")
            Set fso = CreateObject("Scripting.FileSystemObject")
            If Not fso.FileExists(strFilePath) Then
                Set tf = fso.OpenTextFile(strFilePath, ForAppending, True)
                tf.WriteLine gstrLogHeader 
                ' """UserName"",""MachineName"",""Date Time"",""Local Drive"",""Local Drive Free"","" Local Drive Available"",""Local Drive Percent Available"","" Local Drive FileSystem"",""CPU Max Clock Speed"",""Machine Asset Tag"",""Available RAM Gb"",""Total RAM Gb"",""Horizontal Screen Resolution"","" Vertical Screen Resolution"","" Dpi Width"","" Dpi Height"",""OS Bit"",""Office Version"""
                tf.Close
            End If
            Set tf = fso.OpenTextFile(strFilePath, ForAppending, True)
            tf.WriteLine """" & GetMachineUserName & """,""" & GetMachineName & """,""" & Now() & """," & strText
            tf.Close
            Set fso = Nothing
            Set tf = Nothing
            Set objNetwork = Nothing
        End If
    End Sub

    '===========================================================
    Function GetIniValue(strFilePath, strSection, strKey)

        Const ForReading = 1
    
        Dim fso ' As Scripting.FileSystemObject
        Dim tf ' As Scripting.TextStream
        Dim strLine
        Dim nEqualPos
        Dim strLeftString

        GetIniValue = ""

        Set fso = CreateObject("Scripting.FileSystemObject") ' New Scripting.FileSystemObject

        If fso.FileExists(strFilePath) Then

            Set tf = fso.OpenTextFile(strFilePath, ForReading, False)

            Do While tf.AtEndOfStream = False

                ' Continue with next line
                strLine = Trim(tf.ReadLine)

                ' Check if section is found in the current line
                If LCase(strLine) = "[" & LCase(strSection) & "]" Then

                    Do While tf.AtEndOfStream = False

                        ' Continue with next line
                        strLine = Trim(tf.ReadLine)

                        ' Abort loop if next section is reached
                        If Left(strLine, 1) = "[" Then
                            Exit Do
                        End If

                        ' Find position of equal sign in the line
                        nEqualPos = InStr(1, strLine, "=", 1)
                        If nEqualPos > 0 Then
                            strLeftString = Trim(Left(strLine, nEqualPos - 1))
                            ' Check if item is found in the current line
                            If LCase(strLeftString) = LCase(strKey) Then
                                GetIniValue = Trim(Mid(strLine, nEqualPos + 1))
                                ' Abort loop when item is found
                                Exit Do
                            End If
                        End If

                    Loop

                    Exit Do

                End If

            Loop

            tf.Close

        End If

    End Function
      
    '===========================================================
    Function IIf(Expression, TruePart, FalsePart)
    'The Block if is faster and only evaluates one part at a time minimizing errors
        If Expression Then
            IIf = TruePart
        Else
            IIf = FalsePart
        End If
    End Function

    '===========================================================
    Sub WaitPingTimes(ByRef PingTimes)
        'Alternatively use: http://blogs.technet.com/b/heyscriptingguy/archive/2006/03/30/how-can-i-temporarily-pause-a-script-in-an-hta.aspx
        Dim wshShell
        Dim Command
        Set wshShell = CreateObject("WScript.Shell")
        Command = "cmd /c ping -n " & PingTimes + 1 & " 127.0.0.1"
        wshShell.Run Command, 0, True
        Set wshShell = Nothing
    End Sub


    '===========================================================    
    '----------------------------------------------------------------------
    ' - There are several methods to get the current directory of the HTA -
    '----------------------------------------------------------------------
    'From testing don't use the following
    'This method drops the server path for network paths
    'strPath =  Left(Document.Location.pathname, InStrRev(Document.Location.pathname, "\") - 1) & "\Main.hta"
    'This method works fine if the HTA directly executed,
    'but if explorer.exe or cscript.exe executes the hta this method returns the %windir% dictory
    'Dim objScripShell
    'Set objScripShell = CreateObject("WScript.Shell")
    'strPath =  objScripShell.CurrentDirectory & "\Main.hta"
    Function GetCurrentPath()

        Dim strPath
        strPath = UrlDecode(Document.Location.href)
        strPath = Left(strPath, InStrRev(strPath, "/") - 1)
        strPath = Replace(strPath, "/", "\")

        'URLs that begin with a drive letter will begin with 'file:\\\' Check this first
        If Left(strPath, 8) = "file:\\\" Then
        strPath = Right(strPath, Len(strPath) - 8)
        End If

        'URLs that begin with a server name will begin with 'file:'
        If Left(strPath, 5) = "file:" Then
        strPath = Right(strPath, Len(strPath) - 5)
        End If

        GetCurrentPath = strPath

    End Function
    
    '===========================================================
    'On error, cycling through three different methods to grab the computer name
    'from fastest to slowest (2000x, 150x, 1x)
    Function GetMachineName() 'As String
    On Error Resume Next
    Dim strComputerName 'As String
        If Err.Number <> 0 Or Len(strComputerName) = 0 Then
            Err.Clear
            Dim WSNetwork 'As Object:
            Set WSNetwork = CreateObject("WScript.Network")
            strComputerName = WSNetwork.ComputerName
            Set WSNetwork = Nothing
            If Err.Number <> 0 Or Len(strComputerName) = 0 Then
                'This is not the preferred method, 'As it is maintained by CIM
                'and WMI Collection is slow (only 100 queries/Second)
                Dim objColItem, objColItems, objWMIService
                Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
                Set objColItems = objWMIService.ExecQuery _
                    ("Select Name from Win32_Computersystem")
                For Each objColItem In objColItems
                    GetMachineName = objColItem.Name
                Next
                'Cleanup
                Set objColItem = Nothing: Set objColItems = Nothing: Set objWMIService = Nothing
            End If
        End If
        GetMachineName = strComputerName
    End Function

    '===========================================================
    Function GetMachineUserName() 'As String
    On Error Resume Next
        Dim WSNetwork 'As Object
        Set WSNetwork = CreateObject("WScript.Network")
        Dim strUserName 'As String
        strUserName = WSNetwork.UserName
    
        Set WSNetwork = Nothing
        If Err.Number <> 0 Or Len(strUserName) = 0 Then
            Dim objColItem 'As Object
            Dim objColItems 'As Object
            Dim objWMIService 'As Object
            Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
            Set objColItems = objWMIService.ExecQuery _
                ("Select Username from Win32_Computersystem")
            For Each objColItem In objColItems
                GetMachineUserName = objColItem.UserName
            Next
            Set objColItem = Nothing: Set objColItems = Nothing: Set objWMIService = Nothing
        Else
            GetMachineUserName = strUserName
        End If
    End Function

    '===========================================================
    Function GetMachineDomainName() 'As String
    On Error Resume Next
    Dim WSNetwork 'As Object
    Set WSNetwork = CreateObject("WScript.Network")
    Dim strUserDomain 'As String
    strUserDomain = WSNetwork.UserDomain
        Set WSNetwork = Nothing
        If Err.Number <> 0 Or Len(strUserDomain) = 0 Then
        Dim objColItem 'As Object
        Dim objColItems 'As Object
        Dim objWMIService 'As Object

            Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
            Set objColItems = objWMIService.ExecQuery _
                ("Select Username from Win32_Computersystem")
            For Each objColItem In objColItems
                Dim strUserName 'As String
                strUserName = objColItem.UserName
                GetMachineDomainName = Left(strUserName, InStr(1, strUserName, "\") - 1)
            Next
            Set objColItem = Nothing: Set objColItems = Nothing: Set objWMIService = Nothing
        Else
            GetMachineDomainName = strUserDomain
        End If
    End Function

    '===========================================================
    Function GetMachineDrives() 'As String
    Dim strDriveLetter 'As String
    Dim lngDriveLetterValue 'As Long ''As Long
    Dim fso 'As Object ' 'As FileSystemObject
    Dim fsoDrive 'As Object
    Dim strCurrentDrive 'As String
    Dim strConvertedDrive 'As String

        Set fso = CreateObject("Scripting.FileSystemObject")
        'We are looping from A to Z
        lngDriveLetterValue = 65 ' Letters A (65) to Z (90)
        Do While lngDriveLetterValue < 91
            strDriveLetter = Chr(lngDriveLetterValue)
            If fso.DriveExists(strDriveLetter) Then
                Set fsoDrive = fso.GetDrive(strDriveLetter)
                    strCurrentDrive = strDriveLetter & ": "
                    GetMachineDrives = GetMachineDrives & _
                        strCurrentDrive & ","
            End If
            lngDriveLetterValue = lngDriveLetterValue + 1
        Loop
        'Remove trailing ,
        If Len(strDriveLetter) <> 0 Then
            strDriveLetter = Left(strDriveLetter, Len(strDriveLetter) - 1)
        End If
        Set fso = Nothing
    End Function

    '===========================================================
    Function GetMachineLogicalDriveInformationForLog() ' As String
    ' This currently only returs the last local disk found, and will need to be rewritten if there are multiple local drives
    ' http://msdn.microsoft.com/en-us/library/aa394592(VS.85).aspx
	on error resume next
        Dim objWMI
        Dim colDisks
        Dim objDisk
        Dim strDiskDeviceId
        Dim strDiskFileSystem
        Dim dblDiskFreeSpaceGb
        Dim dblDiskAvailableSpaceGb
        Dim dblDiskPercentFree
            Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
            Set colDisks = objWMI.ExecQuery("Select * from Win32_LogicalDisk")
            For Each objDisk In colDisks
                With objDisk
                    If .DriveType = 3 Then 'Is a local Disk
                        strDiskDeviceId = strDiskDeviceId & .DeviceID
                        dblDiskFreeSpaceGb = .FreeSpace / 2 ^ 30
                        dblDiskAvailableSpaceGb = .Size / 2 ^ 30
                        dblDiskPercentFree = dblDiskFreeSpaceGb / dblDiskAvailableSpaceGb
                        strDiskFileSystem = .FileSystem
                    End If
                End With
            Next
            'We add items one line at a time to ensure the correct number of commas are returned
            Dim strVal
            strVal = """"
            strVal = strVal & strDiskDeviceId 
            strVal = strVal & ""","""
            strVal = strVal & Round(dblDiskFreeSpaceGb, 2)
            strVal = strVal & ""","""
            strVal = strVal & Round(dblDiskAvailableSpaceGb, 2)
            strVal = strVal & ""","""
            strVal = strVal & Round(dblDiskPercentFree, 2)
            strVal = strVal & """,""" 
            strVal = strVal & strDiskFileSystem 
            strVal = strVal & """"
            GetMachineLogicalDriveInformationForLog = strVal 

            'Cleanup
            Set objWMI = Nothing
            Set colDisks = Nothing
    End Function

    Function GetMachineNetworkDrives(fPrintersOnly) 'As String
    If Len(fPrintersOnly) = 0 Then
      fPrintersOnly = False
    End If
    'Enumerating network drives will also list Network that have been disconnected
    'but not yet forgotten, and network shares currently open
    Dim wshNetwork 'As Object
    Set wshNetwork = CreateObject("WScript.Network")
    Dim oDrives 'As Object
    Set oDrives = wshNetwork.EnumNetworkDrives
    Dim oPrinters 'As Object
    Set oPrinters = wshNetwork.EnumPrinterConnections
    Dim intDriveEnum 'As Integer
    Dim intPrinterEnum 'As Integer
        If Not fPrintersOnly Then
            GetMachineNetworkDrives = "<br/>" & "Network drive mappings:" & "<br/>"
            For intDriveEnum = 0 To oDrives.Count - 1 Step 2
                GetMachineNetworkDrives = GetMachineNetworkDrives & "Drive " & _
                    oDrives.Item(intDriveEnum) & "(" & oDrives.Item(intDriveEnum + 1) & ")" & "<br/>"
            Next
        End If
        GetMachineNetworkDrives = GetMachineNetworkDrives & "<br/>" & "Network printer mappings:" & "<br/>"
        For intPrinterEnum = 0 To oPrinters.Count - 1 Step 2
            GetMachineNetworkDrives = GetMachineNetworkDrives & """" & "Port " & _
                oPrinters.Item(intPrinterEnum) & "(" & oPrinters.Item(intPrinterEnum + 1) & ")" & "<br/>"
        Next
    End Function

    Sub BuildDir(strPath)
        Dim fso ' As Scripting.FileSystemObject
        Dim arryPaths
        Dim strBuiltPath
        Dim intDir
        
        Set fso = CreateObject("Scripting.FileSystemObject") ' New Scripting.FileSystemObject
        arryPaths = Split(strPath, "\")
        For intDir = 0 To UBound(arryPaths)
            strBuiltPath = strBuiltPath & arryPaths(intDir) & "\"
            If Not fso.FolderExists(strBuiltPath) Then
                fso.CreateFolder strBuiltPath
            End If
        Next
        
    End Sub
