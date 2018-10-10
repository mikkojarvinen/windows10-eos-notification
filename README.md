# windows10-eos-notification
Show users a notification when a Windows 10 version is reaching it's EOS date and there is a Windows 10 feature update Task Sequence available in the Software Center.

Launch ```win10-versioncheck.vbs``` using login script, scheduled task (good triggers: logon, user unlock) or other similar method.
Edit the script and ```Windows10-EOS-warning.hta``` file as needed.

## Setting up a scheduled task for notifications using group policies

To be clear, this is just one way to make the clients run the script.

It seems to be next to impossible to use group policy preferences to set up a scheduled task for a user but it is easy to do this with schtasks.exe and an .xml file. So the wild and crazy solution is to use group policy preferences to create an immediate scheduled task which will run the schtasks.exe with .xml to set up the initial notification scheduled task for the users.

### Immediate scheduled task creation using group policy preferences

**Task**

* Name: Create Scheduled Task for Windows 10 End of Support Notification
* Author:
* Description:  Creates an Immediate Scheduled Task which will create a Scheduled Task "Windows 10 End of Support Notification"   
* Run only when user is logged on:  S4U
* UserId:  NT AUTHORITY\System
* Run with highest privileges:  LeastPrivilege
* Hidden:  No
* Configure for:  1.3
* Enabled:  Yes

**Triggers**

* At task creation/modification     
* Enabled:  Yes 

**Actions**

Start a program
* Program/script: ```schtasks.exe```   
* Arguments: ```/create /tn "Windows 10 End of Support Notification" /xml "\\unc-path\Win10EOS.xml"```

### Example .xml file for schtasks.exe
```xml
<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.2" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo>
    <Date>2017-10-01T16:00:00.0000000</Date>
    <Author>DOMAIN\author</Author>
    <URI>\Windows 10 End of Support Notification</URI>
  </RegistrationInfo>
  <Triggers>
    <LogonTrigger>
      <Enabled>true</Enabled>
    </LogonTrigger>
    <SessionStateChangeTrigger>
      <Enabled>true</Enabled>
      <StateChange>SessionUnlock</StateChange>
    </SessionStateChangeTrigger>
  </Triggers>
  <Principals>
    <Principal id="Author">
      <GroupId>S-1-5-32-545</GroupId>
      <RunLevel>LeastPrivilege</RunLevel>
    </Principal>
  </Principals>
  <Settings>
    <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>
    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
    <StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>
    <AllowHardTerminate>true</AllowHardTerminate>
    <StartWhenAvailable>false</StartWhenAvailable>
    <RunOnlyIfNetworkAvailable>true</RunOnlyIfNetworkAvailable>
    <IdleSettings>
      <StopOnIdleEnd>true</StopOnIdleEnd>
      <RestartOnIdle>false</RestartOnIdle>
    </IdleSettings>
    <AllowStartOnDemand>true</AllowStartOnDemand>
    <Enabled>true</Enabled>
    <Hidden>false</Hidden>
    <RunOnlyIfIdle>false</RunOnlyIfIdle>
    <WakeToRun>false</WakeToRun>
    <ExecutionTimeLimit>PT0S</ExecutionTimeLimit>
    <Priority>7</Priority>
  </Settings>
  <Actions>
    <Exec>
      <Command>wscript.exe</Command>
      <Arguments>/b "\\unc-path\win10-versioncheck.vbs"</Arguments>
    </Exec>
  </Actions>
</Task>
```
