pushd %~dp0
powercfg /sleepstudy /verbose /duration 14
powercfg /batteryreport /duration 14
xcopy c:\windows\minidump\*.*
netsh wlan sh wlanreport 
xcopy c:\ProgramData\Microsoft\Windows\WlanReport\*.*
netsh wlan sh i >> %Computername%_wlan_status.txt
wevtutil epl System %Computername%_System_Event_Log.evtx
wevtutil epl Application %Computername%_Application_Event_Log.evtx
popd
