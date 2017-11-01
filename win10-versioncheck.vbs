' Check Windows 10 version and launch update notice if needed
' Author: Mikko Järvinen <mikko.tapio.jarvinen@gmail.com>
' 
' ProductName                       ReleaseId   Notice   End of support
' -------------------------------------------------------------------
' Windows 10 Enterprise 2015 LTSB   1507        No       N/A
' Windows 10 Enterprise             1507        Yes      9.5.2017
' Windows 10 Enterprise             1511        Yes      10.10.2017
' Windows 10 Enterprise 2016 LTSB   1607        No       N/A
' Windows 10 Enterprise             1607        No       ?
' Windows 10 Enterprise             1703        No       ?
'
'
' "Enterprise" can be also "Education"
'
' If ProductName ends with "LTSB" we are on a long term servicing channel
' If ProductName begins with "Windows 10" and does not end with "LTSB" we are on semi-annual channel

Option Explicit

Const HKLM = &H80000002

Dim objShell
Dim objFileSystem
Dim strRegPath
Dim strProductName
Dim strReleaseId
Dim blnShowNotice
Dim blnUpgradeExists
Dim intReturnCode
Dim strCurrentDir
Dim strCommandLineArgument
Dim strCommandLine

Set objShell = CreateObject("Wscript.Shell")

Set objFileSystem = CreateObject("Scripting.FileSystemObject")

strCurrentDir = objFileSystem.GetParentFolderName(WScript.ScriptFullName)

' Get Windows 10 version information

strRegPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion"

strProductName = GetStringValue(HKLM, strRegPath, "ProductName" ,64)
strReleaseId = GetStringValue(HKLM, strRegPath, "ReleaseId" ,64)

If IsNull(strProductName) Then
  WScript.Quit
End If

If IsNull(strReleaseId) Then
  WScript.Quit
End If

blnShowNotice = false
blnUpgradeExists = false

' *** Windows 10 Version Check routine. Edit when needed. ***

' Test for LTSB versions (long-term servicing channel)
If Right(strProductName, 4) = "LTSB" Then
  Select Case Right(strProductName, 9)
    Case "2015 LTSB"        ' 2015 LTSB
      blnShowNotice = false
      strCommandLineArgument = "DD.MM.YYYY"
    Case "2016 LTSB"        ' 2016 LTSB
      blnShowNotice = false
      strCommandLineArgument = "DD.MM.YYYY"
    Case Else               ' Unknown LTSB
      blnShowNotice = false
      strCommandLineArgument = "DD.MM.YYYY"
  End Select
ElseIf Left(strProductName, 10) = "Windows 10" Then ' No LTSB version, test for Windows 10 
  Select Case strReleaseId                          ' We have Windows 10 (semi-annual-channel), find out version and configure notifications
    Case "1507"             ' 1507
      blnShowNotice = true
      strCommandLineArgument = "9.5.2017"
    Case "1511"             ' 1511
      blnShowNotice = true
      strCommandLineArgument = "10.10.2017"
    Case "1607"             ' 1607
      blnShowNotice = true
      strCommandLineArgument = "1.4.2018"
    Case "1703"             ' 1703
      blnShowNotice = false
      strCommandLineArgument = "DD.MM.YYYY"
    Case Else               ' Unknown ReleaseId
      blnShowNotice = false
  End Select
End If

If blnShowNotice = true Then ' We need to show a notification

  If Windows10UpgradeAvailable Then  ' Is there Windows 10 Upgrade available in Software Center
    blnUpgradeExists = true
  End If

  If blnUpgradeExists = true Then   ' Show EOS info with upgrade option (computers in personal use)
    strCommandLine = """%systemroot%\System32\mshta.exe"" """ & strCurrentDir & "\Windows10-EOS-warning.hta"" """ & strCommandLineArgument & """"
    intReturnCode = objShell.Run(strCommandLine, 1, False) ' False: don't wait for the program to finish
  Else   ' Show EOS info without upgrade option (shared computers) *** TBI ***
'    strCommandLine = """%systemroot%\System32\mshta.exe"" """ & strCurrentDir & "\Windows10-EOS-inform.hta"" """ & strCommandLineArgument & """"
'    intReturnCode = objShell.Run(strCommandLine, 1, False) ' False: don't wait for the program to finish
  End If
End If

WScript.Quit(0) ' All Done

' *** Functions ***

' Find out if there are Windows 10 Upgrades Task Sequence available in Software Center
Function Windows10UpgradeAvailable
  Dim strComputer
  Dim objWMIService
  Dim colPrograms
  Dim objProgram

  Windows10UpgradeAvailable = false
  strComputer = "."
  Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\ccm\ClientSDK")
  If Err.Number <> 0 Then
    wScript.Quit
  End If
  Err.Clear
  Set colPrograms = objWMIService.ExecQuery("Select * from CCM_Program where ProgramId=""*""") ' Find available Task Sequences
  If Err.Number = 0 Then
    On Error Goto 0 ' Disable Error Handling

    For Each objProgram in colPrograms
      If Left(objProgram.FullName, 18) = "Windows 10 Upgrade" Then    ' Match the beginning of strings like "Windows 10 Upgrade (1703)"
        Windows10UpgradeAvailable = true                              ' We have an upgrade available!
      End If
    Next
  End If
End Function

Function GetStringValue(hDefKey,sSubKeyName,sValueName,intRegProv)
  Dim inParams
  Dim outParams
  Dim objCtx,objLocator,objServices,objStdRegProv

  Set objCtx = CreateObject("WbemScripting.SWbemNamedValueSet")
  objCtx.Add "__ProviderArchitecture", intRegProv
  Set objLocator = CreateObject("Wbemscripting.SWbemLocator")
  Set objServices = objLocator.ConnectServer("","root\default","","",,,,objCtx)
  Set objStdRegProv = objServices.Get("StdRegProv") 
  Set inParams = objStdRegProv.Methods_("GetStringValue").Inparameters
  inParams.hDefKey = hDefKey
  inParams.sSubkeyname = sSubKeyName
  inParams.sValuename = sValueName
  Set outParams = objStdRegProv.ExecMethod_("GetStringValue",inParams,,objCtx)
  GetStringValue = outParams.SValue

  Set objCtx = Nothing
  Set objLocator = Nothing
  Set objServices = Nothing
  Set objStdRegProv = Nothing
End Function