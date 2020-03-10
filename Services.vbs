'------------------------------------------
' Elevate this script before invoking it.
' 25.2.2011 FNL
'-----------------------------------------
Set WSHShell    = CreateObject("WScript.Shell")
Set objFSO      = CreateObject("Scripting.FileSystemObject")
Set RegExp      = CreateObject("VBScript.RegExp")
RegExp.Global   = True
RegExp.Pattern  = "([0-9]+)\.(.*)"
MinOS           = RegExp.Replace( "6.0.6000"  , "$1" )
CurOS           = RegExp.Replace( OSNumber( ) , "$1" )
' WScript.Echo TypeName(CurOS) '
MinOS = CInt( MinOS ) 'Converting string to integer
CurOS = CInt( CurOS ) 'Converting string to integer
If CurOS >= MinOS Then
    bElevate = False
    If WScript.Arguments.Count > 0 Then If WScript.Arguments(WScript.Arguments.Count-1) <> "|" Then bElevate = True
    If bElevate Or WScript.Arguments.Count = 0 Then ElevateUAC
'********************************************
' Your script goes here                     *
'********************************************
    If OSName( ) = "XP" Then
        ManualServices = Array( _ 
            "TapiSrv" , "dmadmin" , "EventSystem" , "NtLmSsp" , "UPS" , "TermService" , "Wmi" , "HTTPFilter" , "RasMan" , _ 
            "MSIServer" , "WmiApSrv" , "RasAuto" , "SCardSvr" , "stisvc" , "AppMgmt" , "Apple Mobile Device" _
        )
        AutoServices = Array( _
            "ShellHWDetection" , "AudioSrv" , "Themes" , "TrkWks" , "CryptSvc" , "RpcSs" , "SENS" , "dmserver" , "Dhcp" , _ 
            "Dnscache" , "ERSvc" , "Eventlog" , "PlugPlay" , "PolicyAgent" , "SamSs" , "winmgmt" , "Spooler" , "WZCSVC" , _ 
            "DcomLaunch" , "LmHosts" _
        )
    End If
    If OSName( ) = "2003/XP64" Then

    End If

    If ( ( OSName( ) = "Vista" ) And (OSName( ) = "Vista SP1" ) And (OSName( ) = "Vista SP2" ) ) Then

    End If
    If OSName( ) = "7" Then
        AutoServices = Array( _
            "AudioSrv" , "RpcSs" , "gpsvc" , "EventLog" , "MMCSS" , "NlaSvc" , "DPS" , "Spooler" , "ProfSvc" , "TrkWks" , _ 
            "AudioEndpointBuilder" , "ShellHWDetection" , "Winmgmt" , "Themes" , "SamSs" , "SENS" , "CryptSvc" , "Dhcp" , _ 
            "DcomLaunch" , "EventSystem" , "Wlansvc" , "PlugPlay" , "RpcEptMapper" , "Power" , "sppsvc" , "PSI_SVC_2_x64" , _ 
            "Apple Mobile Device" _ 
        )
        ManualServices = Array( _
            "ProtectedStorage" , "Wecsvc" , "THREADORDER" , "WPDBusEnum" , "WmiApSrv" , "RasAuto" , "RasMan" , "Netman" , _ 
            "pla" , "PNRPAutoReg" , "AppMgmt" , "PolicyAgent" , "SCPolicySvc" , "sppuinotify" , "RpcLocator" , "KeyIso" , _ 
            "dot3svc" , "EapHost" , "SensrSvc" , "wudfsvc" , "WwanSvc" , "SCardSvr" , "SstpSvc" , "bthserv" , "PNRPsvc" , _ 
            "defragsvc" , "lltdsvc" , "SNMPTRAP" , "p2pimsvc" , "IKEEXT" , "SDRSVC" , "IPBusEnum" , "nsi" , _ 
            "MSiSCSI" , "napagent" , "netprofm" , "WcsPlugInService" , "AppIDSvc" , "W32Time" , "TBS" , "vds" , "KtmRm" , _ 
            "MSDTC" , "UmRdpService" , "HomeGroupProvider", "HomeGroupListener" , "VaultSvc" , "stisvc" , "NativeWifiP" , _ 
            "rdpbus" ,"Appinfo" , "swprv" , "BFE" , "vwififlt" _
        )
   End If

    If ( ( OSName( ) = "8" ) And (OSName( ) = "8.1" ) ) Then

    End If

    If OSName( ) = "10" Then
        AutoServices = Array( _
            "Winmgmt" , "WlanSvc" , "Wcmsvc" , "WpnService" , , "AudioEndpointBuilder" , "wscsvc" , "NlaSvc" , "sppsvc" , _ 
            "Audiosrv" , "BTDevManager" , "BrokerInfrastructure" , "CoreMessagingRegistrar" , "RpcEptMapper" , "CDPSvc" , _ 
            "ShellHWDetection" , "SystemEventsBroker" , "Themes" , "TrkWks" , "UserManager" , "SgrmBroker" , "iphlpsvc" , _ 
            "DcomLaunch" , "DeviceAssociationService" , "Dnscache" , "DispBrokerDesktopSvc" , "MapsBroker" , "EventLog" , _ 
            "EventSystem" , "mpssvc" , "gpsvc" , "SENS" , "Dhcp" , "ProfSvc" , "stisvc" , "RpcSs" , "WSearch" , "Power" , _ 
            "osrss" , "nsi" , "DPS" , "BFE" , "LSM" , "SamSs" , "DoSvc" , "Apple Mobile Device" , "PSI_SVC_2_x64" _
        )
        ManualServices = Array( _
            "diagnosticshub.standardcollector.service" , "DisplayEnhancementService" , "DmEnrollmentSvc" , "WpcMonSvc" , _ 
            "AppIDSvc" , "Appinfo" , "AppReadiness" , "autotimesvc" , "DevQueryBroker" , "UmRdpService" , "RetailDemo" , _ 
            "DsmSvc" , "dmwappushservice" , "Eaphost" , "EFS" , "embeddedmode" , "EntAppSvc" , "SCardSvr" , "spectrum" , _ 
            "FrameServer" , "IKEEXT" , "IpxlatCfgSvc" , "KtmRm" , "LicenseManager" , "lltdsvc" , "netprofm" , "KeyIso" , _ 
            "MSiSCSI" , "msiserver" , "NaturalAuthentication" , "NcaSvc" , "NcbService" , "NcdAutoSetup" , "defragsvc" , _ 
            "NetSetupSvc" , "NgcCtnrSvc" , "NgcSvc" , "p2pimsvc" , "p2psvc" , "perceptionsimulation" , "ALG" , "MSDTC" , _ 
            "PerfHost" , "PhoneSvc" , "pla" , "RpcLocator" , "SensorDataService" , "SensorService" , "fdPHost" , "Fax" , _ 
            "PlugPlay" , "PNRPAutoReg" , "PNRPsvc" , "PolicyAgent" , "WerSvc" , "SSDPSRV" , "DeviceInstall" , "LxpSvc" , _ 
            "ScDeviceEnum" , "SCPolicySvc" , "SDRSVC" , "seclogon" , "SecurityHealthService" , "SEMgrSvc" , "SNMPTRAP" , _ 
            "SstpSvc" , "StateRepository" , "TabletInputService" , "vmicvmsession" , "svsvc" , "swprv" , "TermService" , _ 
            "TieringEngineService" , "TimeBrokerSvc" , "TroubleshootingSvc" , "TrustedInstaller" , "WiaRpc" , "PcaSvc" , _ 
            "vmicheartbeat" , "VacSvc" , "VaultSvc" , "vmickvpexchange" , "vmicrdv" , "vmicshutdown" , "WFDSConMgrSvc" , _ 
            "w3logsvc" , "WaaSMedicSvc" , "WalletService" , "WAS" , "wbengine" , "WbioSrvc" , "wcncsvc" , "WEPHOSTSVC" , _ 
            "wlidsvc" , "wlpasvc" , "WManSvc" , "wmiApSrv" , "WPDBusEnum" , "TapiSrv" , "VSS" , "StorSvc" , "upnphost" , _ 
            "BTAGService" , "CertPropSvc" , "CryptSvc" , "dot3svc" , "lmhosts" , "WdiServiceHost" , "wisvc" , "camsvc" , _ 
            "SharedAccess" , "SharedRealitySvc" , "smphost" , "ClipSVC" , "diagsvc" , "SmsRouter" , "RmSvc" , "BDESVC" , _ 
            "WdiSystemHost" , "WdNisSvc" , "Wecsvc" , "wercplsupport" , "WinHttpAutoProxySvc" , "WwanSvc" , "SensrSvc" , _ 
            "HvHost" , "vmictimesync" , "Netman" , "ose64" _ 
        )
    End If
    DisableServices = Array( _ 
        "HidServ" , "hkmsvc" , "Netlogon" , "upnphost" , "NetDDEdsdm" , "CliqzMaintenance" , "NetDDE" , "Messenger" , _ 
        "Schedule" , "seclogon" , "SharedAccess" , "srservice" , "SSDPSRV" , "SysmonLog" , "wuauserv" , "WebClient" , _ 
        "SwPrv" , "mnmsrvc" , "BITS" , "W32Time" , "WmdmPmSN" , "RpcLocator" , "TlntSvr" , "ImapiService" , "MSDTC" , _ 
        "WMPNetworkSvc" , "CiSvc" , "FastUserSwitchingCompatibility" , "COMSysApp" , "NvNetworkService" , "Alerter" , _ 
        "LanmanServer" , "AllJoyn Router Service" , "Application Layer Gateway Service" , "DNSClient" , "DiagTrack" , _ 
        "Remote Desktop Services UserMode Port Director" , "Retail Demo Service" , "HomeGroup Listener" , "AppXSvc" , _ 
        "LanmanWorkstation" , "NvStreamNetworkSvc" , "RemoteAccess" , "ALG" , "Nla" , "VSS" , "Browser" , "gupdate" , _ 
        "Remote Access Connection Manager" , "Internet Connection Sharing (ICS)" , "SessionEnv" , "GamesAppService" , _ 
        "BranchCache" , "Certificate Propagation" , "Downloaded Maps Manager" , "Sensor Data Service" , "SNMP Trap" , _ 
        "Hyper-V Data Exchange Service" , "Bluetooth Support Service" , "Sensor Monitoring Service" , "TokenBroker" , _ 
        "Hyper-V Heartbeat Service" , "Microsoft iSCSI Initiator Service" , "GoogleChromeElevationService" , "cphs" , _ 
        "Hyper-V Guest Service Interface" , "Internet Explorer ETW Collector Service" , "IP Helper" , "jhi_service" , _ 
        "Hyper-V Remote Desktop Virtualization Service" , "Remote Access Auto Connection Manager" , "Offline Files" , _
        "Hyper-V Time Synchronization Service" , "Hyper-V VM Session Service" , "odserv" , "Print Spooler" , "RSVP" , _ 
        "Hyper-V Volume Shadow Copy Requestor" , "Windows Biometric Service" , "SecurityHealthService" , "gupdatem" , _ 
        "Hyper-V Guest Shutdown Service" , "Windows Remote Management (WS-Management)" , "BTDevManager" , "ehRecvr" , _ 
        "Microsoft Windows SMS Router Service" , "Printer Extensions and Notifications" , "UevAgentService" , "Avg" , _ 
        "UserDataSvc" , "tiledatamodelsvc" , "Touch Keyboard and Handwriting Panel Service" , "DsSvc" , "ePowerSvc" , _ 
        "Updater Service" , "BBUpdate" , "Intel(R) TPM Provisioning Service" , "igfxCUIService2.0.0.0" , "TVESched" , _ 
        "WsAppService" , "RetailDemo" , "Intel(R) Capability Licensing Service TCP IP Interface" , "DevQueryBroker" , _ 
        "WpnService" , "Storage Service" , "Geolocation Service" , "aspnet_state" , "NMIndexingService" , "Mcx2Svc" , _ 
        "clr_optimization_v4.0.30319_64" , "clr_optimization_v4.0.30319_32" , "WarpJITSvc" , "Greg_Service" , "MDM" , _ 
        "clr_optimization_v2.0.50727_64" , "clr_optimization_v2.0.50727_32" , "AMD External Events Utility" , "ose" , _
        "AgereModemAudio" , "AMD FUEL Service" , "NcbService" , "CDPUserSvc" , "esifsvc" , "CoreMessagingRegistrar" , _ 
        "Microsoft Office Groove Audit Service" , "DcpSvc" , "Partner Service" , "ServiceLayer" , "SynTPEnhService" , _ 
        "UnistoreSvc" , "dmwappushservice" , "PimIndexMaintenanceSvc" , "QWAVE" , "MWLService" , "lfsvc" , "rpcapd" , _ 
        "AdobeFlashPlayerUpdateSvc" , "ESLoadService" , "BBSvc" , "nvsvc" , "AdobeARMservice" , "TVECapSvc" , "LMS" , _ 
        "TuneUp.UtilitiesSvc" , "NTISchedulerSvc" , "WatAdminSvc" , "HP Comm Recover" , "HPWMISVC" , "NTIBackupSvc" , _ 
        "MapsBroker" , "vmicguestinterface" , "WinHttpAutoProxySvc" , "GamesAppIntegrationService" , "hpqcaslwmiex" , _ 
        "AJRouter" , "WdNisSvc" , "dbupdatem" , "Remote Desktop Services" , "HPJumpStartBridge" , "GraphicsPerfSvc" , _ 
        "ClickToRunSvc" , "UsoSvc" , "fhsvc" , "RichVideo" , "MessagingService" , "IAANTMON" , "PhoneSvc" , "idsvc" , _ 
        "Remote Desktop Configuration" , "Remote Procedure Call (RPC) Locator" , "VIAKaraokeService" , "OneSyncSvc" , _ 
        "NtmsSvc" , "ProtectedStorage" , "RemoteRegistry" , "xmlprov" , "ClipSrv" , "LicenseManager" , "Superfetch" , _ 
        "Windows Mobile Hotspot Service" , "Windows Media Player Network Sharing Service" , "RDSessMgr" , "wlidsvc" , _ 
        "CursorMania_7lService" , "lmhosts" , "TeamViewer7" , "Dnscache" , "SysMain" , "ezGOSvc" , "vds" , "stisvc" , _
        "Microsoft (R) Diagnostics Hub Standard Collector Service" , "Sensor Service" , "Smart Card Removal Policy" , _ 
        "HomeGroup Provider" , "helpsvc" , "lanmanserver" , "iphlpsvc" , "RealNetworks Downloader Resolver Service" , _ 
        "Update diamondata" , "Pml Driver HPZ12" , "Net Driver HPZ12" , "CscService" , "avastm" , "Util diamondata" , _ 
        "WinDefend" , "WSearch" , "avast" , "Audiosrv" , "eventlog" , "DragonUpdater" , "WebClient" , "SkypeUpdate" , _ 
        "WmdmPmSp" , "EPSON_PM_RPCV4_01" , "EPSON_EB_RPCV4_01" , "UxSms" , "PeerDistSvc" , "TermService" , "fsssvc" , _ 
        "UNS" , "Util Primary Result" , "TosCoSrv" , "ehSched" , "ST2012_Svc" , "wcncsvc" , "AxInstSV" , "FDResPub" , _ 
        "BDESVC" , "CertPropSvc" , "FontCache3.0.0.0" , "UmRdpService" , "WinRM" , "PcaSvc" , "EFS" , "AeLookupSvc" , _ 
        "IEEtwCollectorService" , "Origin Client Service" , "Origin Web Helper Service" , "BthAvctpSvc" , "bthserv" , _ 
        "workfolderssvc" , "AESMService" , "shpamsvc" , "vmicvss" , "DSAUpdateService" , "FontCache" , "DSAService" , _ 
        "SetupARService" , "IAStorDataMgrSvc" , "InstallService" , "MozillaMaintenance" , "PushToInstall" , "swprv" , _ 
        "Windows Remediation Service" , "ProtonVPN Service" , "ssh-agent" , "GamesAppServices" , "NTI IScheduleSvc" , _ 
        "RtkAudioService" , "RtkBtManServ" , "Bonjour Service" , "Windows Search" , "SharedRealitySvc" , "dbupdate" , _ 
        "Xbox Live Game Save" , "Xbox Live Auth Manager" , "XboxNetApiSvc" , "XblAuthManager" , "RasMan" , "WPCSvc" , _ 
        "XboxGipSvc" , "XblGameSave" , "WalletService" , "wercplsupport" , "RasAuto" , "NlaSvc" , "ProfSvc" , "Fax" , _ 
        "HPSupportSolutionsFrameworkService" , "wscsvc" , "wisvc" , "AppIDSvc" , "cplspcon" , "DusmSvc" , "Appinfo" , _ 
        "AGMService" , "AGSService" , "brave" , "bravem" , "tzautoupdate" , "SeagateDashboardService" , "IPBusEnum" , _ 
        "NetMsmqActivator" , "NetPipeActivator" , "NetTcpActivator" , "NetTcpPortSharing" , "wbengine" , "PnkBstrA" , _ 
        "Smart Card Device Enumeration Service" , "Windows Connect Now - Config Registrar" , "NvTelemetryContainer" , _ 
        "Intel(R) Capability Licensing Service TCP IP Interface" , "Steam Client Service" , "RpcEptMapper" , "NetBT" , _
        "NVDisplay.ContainerLocalSystem" , "Update service" , "EasyAntiCheat" , "afunix" , "bam" , "lltdio" , "Vid" , _
        "NdisVirtualBus" , "NativeWifiP" , "storqosflt" , "npsvctrig" , "NetBIOS" , "Psched" , "icssvc" , "AppHost" , _
        "NvContainerNetworkService" , "NvContainerLocalSystem" , "NovabenchService" , "WCAssistantService" , "rdbss" , _
        "rdpbus" , "SgrmAgent" , "spaceport" , "vwififlt" , "AppHostSvc" , "ssh-agent" , "ETDService" , "rspndr" , _
        "igfxCUIService1.0.0.0" , "MsLldp" _ 
    )
' osppsvc   = Office Software Protection Platform
' ose64     = Office 64 Source Engine
' AVG Antivirus , avgbIDSAgent , VmbService, 
' "NativeWifiP" "rdpbus" , vwififlt  "FontCache" ,

    ' Chamar e correr de acordo com o OS exepto os DisableServices
    For each Service in DisableServices
        Call UpdateServices( Service , 4 ) 
    Next
    For each Service in ManualServices
        Call UpdateServices( Service , 3 )
    Next
    For each Service in AutoServices
        Call UpdateServices( Service , 2 )
    Next
End If
Function UpdateServices( RegServ , UserData )
    Const HKEY_LOCAL_MACHINE    = &H80000002
    Set WshShell = Wscript.CreateObject( "WScript.Shell" )
    RegRoot      = "HKEY_LOCAL_MACHINE"
    RegPath      = "\SYSTEM\CurrentControlSet\services\"
    RegName      = "Start"
    'RegServ2     = RegServ
    RegServ      = RegServ & "\"
    FullPath     = RegRoot & RegPath & RegServ
    'If RegKeyExists( FullPath ) = True Then
        'WScript.Echo " exist " & RegKeyExists( FullPath )
        If RegKeyExists( FullPath & RegName ) = True Then
            RegData = WshShell.RegRead( FullPath & RegName )
            If IsNumeric( RegData ) Then
            'WshShell.run("net stop " & RegServ)
                If Not RegData = UserData Then
                    WshShell.RegWrite ""&FullPath&RegName&"" , UserData , "REG_DWORD"
                End If
            End If
        End If
    'End If
    'Set WshShell = Nothing
End Function
'****************************************************************************************
' https://www.tek-tips.com/viewthread.cfm?qid=1025447                                   *
' https://www.itninja.com/question/detecting-If-a-registry-key-is-present-or-not        *
'****************************************************************************************
Function RegKeyExists( strKey )
    Dim bExists
    ssig = "Unable to open registry key"
    Set WshShell = Wscript.CreateObject( "WScript.Shell" )
    on error resume next
    present = WshShell.RegRead( strKey )
    If err.number <> 0 Then
        If right( strKey , 1 ) = "\" Then    'strKey is a registry key
            ' Assign a value using the single-line form of syntax. 
            If instr( 1 , err.description , ssig , 1 ) <> 0 Then bExists = True Else bExists = False 
        else    'strKey is a registry valuename
            bExists = False
        End If
        err.clear
    else
        bExists = True
    End If
    on error goto 0
    RegKeyExists = bExists
End Function
'*************************************
' Achar o nome da versao do windows  *
'*************************************
Function OSName( )
    ' https://www.gaijin.at/en/infos/windows-version-numbers
    ' https://en.wikipedia.org/wiki/Ver_(command)'
    set WshShell        = Wscript.CreateObject( "WScript.Shell" )
    Set getOSVersion    = WshShell.exec( "%comspec% /c ver" )
    version             = getOSVersion.stdout.readall
    GetName             = "Unknown" 
    versions            = Array( _
        "1.04"       , "2.11" , "3" , "3.10.528" , "3.11" , "3.5.807" , "3.51.1057" , "4.0.1381" , "4.0.950" , _ 
        "4.1.1998"   , "4.1.2222"   , "4.90.3000"  , "5.0.2195"   , "5.1.2600"   , "5.2.3790"   , "6.0.6000" , _ 
        "6.0.6001"   , "6.0.6002"   , "6.1.7600"   , "6.1.7601"   , "6.2.9200"   , "6.3.9200"   , "6.3.9600" , _
        "10.0.10240" , "10.0.10586" , "10.0.14393" , "10.0.15063" , "10.0.16299" , "10.0.17134" , "10.0.17763" _
    )
    names               = Array( _
        "1" , "2" , "3" , "3.1" , "3.11" , "3.5" , "3.51" , "4.0" , "95" , "98" , "98 SE" , "ME" , "2000" , "XP" , _ 
        "2003/XP64" , "Vista" , "Vista SP1" , "Vista SP2" , "7" , "7" , "8" , "8.1" , "8.1" , "10" , "10" , "10" , _ 
        "10" , "10" , "10" , "10" _
    )
    For i = 0 To UBound( versions )
        If InStr( version , versions(i) ) <> 0 Then GetName = names(i) End If
    Next
    OSName = GetName
End Function
'**************************************
' Achar o numero da versao do windows *
'**************************************
Function OSNumber()
    'How to find operating system name without using WMI in VBScript?
    set WshShell        = Wscript.CreateObject( "WScript.Shell" )
    Set getOSVersion    = WshShell.exec( "%comspec% /c ver" )
    version             = getOSVersion.stdout.readall
    GetOS               = "Unknown"
    ' Array com todas as versÃµes do windows
    versions = Array( _
        "1.04" , "2.11" , "3" , "3.10.528" , "3.11" , "3.5.807"     , "3.51.1057"   , "4.0.1381"    , "4.0.950"  , _ 
        "4.1.1998"   , "4.1.2222"   , "4.90.3000"   , "5.0.2195"    , "5.1.2600"    , "5.2.3790"    , "6.0.6000" , _ 
        "6.0.6001"   , "6.0.6002"   , "6.1.7600"    , "6.1.7601"    , "6.2.9200"    , "6.3.9200"    , "6.3.9600" , _
        "10.0.10240" , "10.0.10586" , "10.0.14393"  , "10.0.15063"  , "10.0.16299"  , "10.0.17134"  , "10.0.17763" _
    )
    For i = 0 To UBound(versions)
        If InStr( version , versions(i) ) <> 0 Then GetOS = versions(i) End If
    Next
    OSNumber = GetOS
End Function
'****************************************************************************************
' Run this script under elevated privileges                                             *
' https://superuser.com/questions/1236466/unable-to-run-vbscript-file-on-windows-10     *
'****************************************************************************************
Sub ElevateUAC
    sParms = " |"
    If WScript.Arguments.Count > 0 Then
        For i  = WScript.Arguments.Count-1 To 0 Step -1
        sParms = " " & WScript.Arguments(i) & sParms
        Next
    End If
    Set oShell = CreateObject( "Shell.Application" )
    oShell.ShellExecute "wscript.exe" , WScript.ScriptFullName & sParms , , "runas" , 1
    Set oShell = Nothing
    WScript.Quit
End Sub

' 1. AllJoyn Router Service
' 2. Application Layer Gateway Service
' 3. Bluetooth Support Service
' 4. BranchCache
' 5. Certificate Propagation
' 6. DNSClient
' 7. Downloaded Maps Manager
' 8. Geolocation Service
' 9. HomeGroup Listener
' 10. HomeGroup Provider
' 11. Hyper-V Data Exchange Service
' 12. Hyper-V Guest Service Interface
' 13. Hyper-V Guest Shutdown Service
' 14. Hyper-V Heartbeat Service
' 15. Hyper-V Remote Desktop Virtualization Service
' 16. Hyper-V Time Synchronization Service
' 17. Hyper-V VM Session Service
' 18. Hyper-V Volume Shadow Copy Requestor
' 19. Internet Connection Sharing (ICS)
' 20. Internet Explorer ETW Collector Service
' 21. IP Helper
' 22. Microsoft (R) Diagnostics Hub Standard Collector Service
' 23. Microsoft iSCSI Initiator Service
' 24. Microsoft Windows SMS Router Service
' 25. Netlogon
' 26. Offline Files
' 27. Print Spooler
' 28. Printer Extensions and Notifications
' 29. Remote Access Auto Connection Manager
' 30. Remote Access Connection Manager
' 31. Remote Desktop Configuration
' 32. Remote Desktop Services
' 33. Remote Desktop Services UserMode Port Director
' 34. Remote Procedure Call (RPC) Locator
' 35. Retail Demo Service
' 36. Sensor Data Service
' 37. Sensor Monitoring Service
' 38. Sensor Service
' 39. Smart Card Device Enumeration Service
' 40. Smart Card Removal Policy
' 41. SNMP Trap
' 42. Storage Service
' 43. Superfetch
' 44. Touch Keyboard and Handwriting Panel Service
' 45. Windows Biometric Service
' 46. Windows Connect Now - Config Registrar
' 47. Windows Media Player Network Sharing Service
' 48. Windows Mobile Hotspot Service
' 49. Windows Remote Management (WS-Management)
' 50. Windows Search
' 51. Xbox Live Auth Manager
' 52. Xbox Live Game Save
' 53. XboxNetApiSvc
' CNG Key Isolation
' Diagnostic Tracking services - (Windows Phone home tracking service aka Telemetry)
' Remote Registry
' Routing and remote access
' Server
' Smart Card
' SSDP Discovery
' Superfetch - (If you have an SSD, other wise leave it on)
' Windows Connect now - (only if you use a fixed LAN and not Wireless, otherwise leave it on)
' Windows Defender - (unless you rely on this as an antivirus)
' Windows Font Cache service - (if you have an SSD, but not necessary)
' LanmanWorkstation