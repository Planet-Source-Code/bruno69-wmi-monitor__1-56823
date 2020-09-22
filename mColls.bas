Attribute VB_Name = "mColls"
Option Explicit

Public Function col_Win32(ByVal sClass As String) As Collection
   
   Select Case sClass
      Case "1349Controller": Set col_Win32 = Controller1349
      Case "Account": Set col_Win32 = Account
      Case "AutochkSetting": Set col_Win32 = AutochkSetting
      Case "BaseBoard": Set col_Win32 = BaseBoard
      Case "BaseService": Set col_Win32 = BaseService
      Case "Battery": Set col_Win32 = Battery
      Case "Binary": Set col_Win32 = Binary
      Case "BindImageAction": Set col_Win32 = BindImageAction
      Case "Bios": Set col_Win32 = Bios
      Case "BootConfiguration": Set col_Win32 = BootConfiguration
      Case "Bus": Set col_Win32 = Bus
      Case "CDROMDrive": Set col_Win32 = CDROMDrive
      Case "COMApplication": Set col_Win32 = COMApplication
      Case "COMClass": Set col_Win32 = COMClass
      Case "COMSetting": Set col_Win32 = COMSetting
      Case "CacheMemory": Set col_Win32 = CacheMemory
      Case "ClassInfoAction": Set col_Win32 = ClassInfoAction
      Case "ClassicCOMClass": Set col_Win32 = ClassicCOMClass
      Case "ClassicCOMClassSetting": Set col_Win32 = ClassicCOMClassSetting
      Case "CodecFile": Set col_Win32 = CodecFile
      Case "CommandLineAccess": Set col_Win32 = CommandLineAccess
      Case "ComponentCategory": Set col_Win32 = ComponentCategory
      Case "ComputerShutdownEvent": Set col_Win32 = ComputerShutdownEvent
      Case "ComputerSystem": Set col_Win32 = ComputerSystem
      Case "ComputerSystemEvent": Set col_Win32 = ComputerSystemEvent
      Case "ComputerSystemProduct": Set col_Win32 = ComputerSystemProduct
      Case "Condition": Set col_Win32 = Condition
      Case "CreateFolderAction": Set col_Win32 = CreateFolderAction
      Case "CurrentProbe": Set col_Win32 = CurrentProbe
      Case "CurrentTime": Set col_Win32 = CurrentTime
      Case "DCOMApplication": Set col_Win32 = DCOMApplication
      Case "DMAChannel": Set col_Win32 = DMAChannel
      Case "Directory": Set col_Win32 = Directory
      Case "DiskDrive": Set col_Win32 = DiskDrive
      Case "DiskPartition": Set col_Win32 = DiskPartition
      Case "DisplayConfiguration": Set col_Win32 = DisplayConfiguration
      Case "DisplayControllerConfiguration": Set col_Win32 = DisplayControllerConfiguration
      Case "DriverVXD": Set col_Win32 = DriverVXD
      Case "DuplicateFileAction": Set col_Win32 = DuplicateFileAction
      Case "Environment": Set col_Win32 = Environment
      Case "EnvironmentSpecification": Set col_Win32 = EnvironmentSpecification
      Case "ExtensionInfoAction": Set col_Win32 = ExtensionInfoAction
      Case "Fan": Set col_Win32 = Fan
      Case "FileSpecification": Set col_Win32 = FileSpecification
      Case "FloppyController": Set col_Win32 = FloppyController
      Case "FloppyDrive": Set col_Win32 = FloppyDrive
      Case "FontInfoAction": Set col_Win32 = FontInfoAction
      Case "Group": Set col_Win32 = Group
      Case "HeatPipe": Set col_Win32 = HeatPipe
      Case "IDEController": Set col_Win32 = IDEController
      Case "IP4PersistedRouteTable": Set col_Win32 = IP4PersistedRouteTable
      Case "IP4RouteTable": Set col_Win32 = IP4RouteTable
      Case "IRQResource": Set col_Win32 = IRQResource
      Case "InfraredDevice": Set col_Win32 = InfraredDevice
      Case "IniFileSpecification": Set col_Win32 = IniFileSpecification
      Case "JobObjectStatus": Set col_Win32 = JobObjectStatus
      Case "Keyboard": Set col_Win32 = Keyboard
      Case "LaunchCondition": Set col_Win32 = LaunchCondition
      Case "LoadOrderGroup": Set col_Win32 = LoadOrderGroup
      Case "LocalTime": Set col_Win32 = LocalTime
      Case "LogicalDisk": Set col_Win32 = LogicalDisk
      Case "LogicalFileSecuritySetting": Set col_Win32 = LogicalFileSecuritySetting
      Case "LogicalMemoryConfiguration": Set col_Win32 = LogicalMemoryConfiguration
      Case "LogicalProgramGroup": Set col_Win32 = LogicalProgramGroup
      Case "LogicalProgramGroupItem": Set col_Win32 = LogicalProgramGroupItem
      Case "LogicalShareSetting": Set col_Win32 = LogicalShareSetting
      Case "LogonSession": Set col_Win32 = LogonSession
      Case "MIMEInfoAction": Set col_Win32 = MIMEInfoAction
      Case "MSIResource": Set col_Win32 = MSIResource
      Case "MappedLogicalDisk": Set col_Win32 = MappedLogicalDisk
      Case "MemoryArray": Set col_Win32 = MemoryArray
      Case "MemoryDevice": Set col_Win32 = MemoryDevice
      Case "ModuleLoadTrace": Set col_Win32 = ModuleLoadTrace
      Case "MotherboardDevice": Set col_Win32 = MotherboardDevice
      Case "MoveFileAction": Set col_Win32 = MoveFileAction
      Case "NTDomain": Set col_Win32 = NTDomain
      Case "NTEventLogFile": Set col_Win32 = NTEventLogFile
      Case "NTLogEvent": Set col_Win32 = NTLogEvent
      Case "NamedJobObject": Set col_Win32 = NamedJobObject
      Case "NamedJobObjectActgInfo": Set col_Win32 = NamedJobObjectActgInfo
      Case "NamedJobObjectLimitSetting": Set col_Win32 = NamedJobObjectLimitSetting
      Case "NetworkAdapter": Set col_Win32 = NetworkAdapter
      Case "NetworkAdapterConfiguration": Set col_Win32 = NetworkAdapterConfiguration
      Case "NetworkClient": Set col_Win32 = NetworkClient
      Case "NetworkConnection": Set col_Win32 = NetworkConnection
      Case "NetworkLoginProfile": Set col_Win32 = NetworkLoginProfile
      Case "NetworkProtocol": Set col_Win32 = NetworkProtocol
      Case "ODBCAttribute": Set col_Win32 = ODBCAttribute
      Case "ODBCDataSourceSpecification": Set col_Win32 = ODBCDataSourceSpecification
      Case "ODBCDriverSpecification": Set col_Win32 = ODBCDriverSpecification
      Case "ODBCSourceAttribute": Set col_Win32 = ODBCSourceAttribute
      Case "ODBCTranslatorSpecification": Set col_Win32 = ODBCTranslatorSpecification
      Case "OSRecoveryConfiguration": Set col_Win32 = OSRecoveryConfiguration
      Case "OnBoardDevice": Set col_Win32 = OnBoardDevice
      Case "OperatingSystem": Set col_Win32 = OperatingSystem
      Case "PCMCIAController": Set col_Win32 = PCMCIAController
      Case "POTSModem": Set col_Win32 = POTSModem
      Case "PageFile": Set col_Win32 = PageFile
      Case "PageFileSetting": Set col_Win32 = PageFileSetting
      Case "PageFileUsage": Set col_Win32 = PageFileUsage
      Case "ParallelPort": Set col_Win32 = ParallelPort
      Case "Patch": Set col_Win32 = Patch
      Case "PatchPackage": Set col_Win32 = PatchPackage
      Case "Perf": Set col_Win32 = Perf
      Case "PerfFormattedData": Set col_Win32 = PerfFormattedData
      Case "PerfFormattedData_ASP_ActiveServerPages": Set col_Win32 = PerfFormattedData_ASP_ActiveServerPages
      Case "PerfFormattedData_ContentFilter_IndexingServiceFilter": Set col_Win32 = PerfFormattedData_ContentFilter_IndexingServiceFilter
      Case "PerfFormattedData_ContentIndex_IndexingService": Set col_Win32 = PerfFormattedData_ContentIndex_IndexingService
      Case "PerfFormattedData_ISAPISearch_HttpIndexingService": Set col_Win32 = PerfFormattedData_ISAPISearch_HttpIndexingService
      Case "PerfFormattedData_InetInfo_InternetInformationServicesGlobal": Set col_Win32 = PerfFormattedData_InetInfo_InternetInformationServicesGlobal
      Case "PerfFormattedData_MSDTC_DistributedTransactionCoordinator": Set col_Win32 = PerfFormattedData_MSDTC_DistributedTransactionCoordinator
      Case "PerfFormattedData_NTFSDRV_SMTPNTFSStoreDriver": Set col_Win32 = PerfFormattedData_NTFSDRV_SMTPNTFSStoreDriver
      Case "PerfFormattedData_PSched_PSchedFlow": Set col_Win32 = PerfFormattedData_PSched_PSchedFlow
      Case "PerfFormattedData_PSched_PSchedPipe": Set col_Win32 = PerfFormattedData_PSched_PSchedPipe
      Case "PerfFormattedData_PerfDisk_LogicalDisk": Set col_Win32 = PerfFormattedData_PerfDisk_LogicalDisk
      Case "PerfFormattedData_PerfDisk_PhysicalDisk": Set col_Win32 = PerfFormattedData_PerfDisk_PhysicalDisk
      Case "PerfFormattedData_PerfNet_Browser": Set col_Win32 = PerfFormattedData_PerfNet_Browser
      Case "PerfFormattedData_PerfNet_Redirector": Set col_Win32 = PerfFormattedData_PerfNet_Redirector
      Case "PerfFormattedData_PerfNet_Server": Set col_Win32 = PerfFormattedData_PerfNet_Server
      Case "PerfFormattedData_PerfNet_ServerWorkQueues": Set col_Win32 = PerfFormattedData_PerfNet_ServerWorkQueues
      Case "PerfFormattedData_PerfOS_Cache": Set col_Win32 = PerfFormattedData_PerfOS_Cache
      Case "PerfFormattedData_PerfOS_Memory": Set col_Win32 = PerfFormattedData_PerfOS_Memory
      Case "PerfFormattedData_PerfOS_Objects": Set col_Win32 = PerfFormattedData_PerfOS_Objects
      Case "PerfFormattedData_PerfOS_PagingFile": Set col_Win32 = PerfFormattedData_PerfOS_PagingFile
      Case "PerfFormattedData_PerfOS_Processor": Set col_Win32 = PerfFormattedData_PerfOS_Processor
      Case "PerfFormattedData_PerfOS_System": Set col_Win32 = PerfFormattedData_PerfOS_System
      Case "PerfFormattedData_PerfProc_FullImage_Costly": Set col_Win32 = PerfFormattedData_PerfProc_FullImage_Costly
      Case "PerfFormattedData_PerfProc_Image_Costly": Set col_Win32 = PerfFormattedData_PerfProc_Image_Costly
      Case "PerfFormattedData_PerfProc_JobObject": Set col_Win32 = PerfFormattedData_PerfProc_JobObject
      Case "PerfFormattedData_PerfProc_JobObjectDetails": Set col_Win32 = PerfFormattedData_PerfProc_JobObjectDetails
      Case "PerfFormattedData_PerfProc_Process": Set col_Win32 = PerfFormattedData_PerfProc_Process
      Case "PerfFormattedData_PerfProc_ProcessAddressSpace_Costly": Set col_Win32 = PerfFormattedData_PerfProc_ProcessAddressSpace_Costly
      Case "PerfFormattedData_PerfProc_Thread": Set col_Win32 = PerfFormattedData_PerfProc_Thread
      Case "PerfFormattedData_PerfProc_ThreadDetails_Costly": Set col_Win32 = PerfFormattedData_PerfProc_ThreadDetails_Costly
      Case "PerfFormattedData_RSVP_ACSRSVPInterfaces": Set col_Win32 = PerfFormattedData_RSVP_ACSRSVPInterfaces
      Case "PerfFormattedData_RSVP_ACSRSVPService": Set col_Win32 = PerfFormattedData_RSVP_ACSRSVPService
      Case "PerfFormattedData_RemoteAccess_RASPort": Set col_Win32 = PerfFormattedData_RemoteAccess_RASPort
      Case "PerfFormattedData_RemoteAccess_RASTotal": Set col_Win32 = PerfFormattedData_RemoteAccess_RASTotal
      Case "PerfFormattedData_SMTPSVC_SMTPServer": Set col_Win32 = PerfFormattedData_SMTPSVC_SMTPServer
      Case "PerfFormattedData_Spooler_PrintQueue": Set col_Win32 = PerfFormattedData_Spooler_PrintQueue
      Case "PerfFormattedData_TapiSrv_Telephony": Set col_Win32 = PerfFormattedData_TapiSrv_Telephony
      Case "PerfFormattedData_Tcpip_ICMP": Set col_Win32 = PerfFormattedData_Tcpip_ICMP
      Case "PerfFormattedData_Tcpip_IP": Set col_Win32 = PerfFormattedData_Tcpip_IP
      Case "PerfFormattedData_Tcpip_NBTConnection": Set col_Win32 = PerfFormattedData_Tcpip_NBTConnection
      Case "PerfFormattedData_Tcpip_NetworkInterface": Set col_Win32 = PerfFormattedData_Tcpip_NetworkInterface
      Case "PerfFormattedData_Tcpip_TCP": Set col_Win32 = PerfFormattedData_Tcpip_TCP
      Case "PerfFormattedData_Tcpip_UDP": Set col_Win32 = PerfFormattedData_Tcpip_UDP
      Case "PerfFormattedData_TermService_TerminalServices": Set col_Win32 = PerfFormattedData_TermService_TerminalServices
      Case "PerfFormattedData_TermService_TerminalServicesSession": Set col_Win32 = PerfFormattedData_TermService_TerminalServicesSession
      Case "PerfFormattedData_W3SVC_WebService": Set col_Win32 = PerfFormattedData_W3SVC_WebService
      Case "PerfRawData": Set col_Win32 = PerfRawData
      Case "PerfRawData_ASP_ActiveServerPages": Set col_Win32 = PerfRawData_ASP_ActiveServerPages
      Case "PerfRawData_ContentFilter_IndexingServiceFilter": Set col_Win32 = PerfRawData_ContentFilter_IndexingServiceFilter
      Case "PerfRawData_ContentIndex_IndexingService": Set col_Win32 = PerfRawData_ContentIndex_IndexingService
      Case "PerfRawData_ISAPISearch_HttpIndexingService": Set col_Win32 = PerfRawData_ISAPISearch_HttpIndexingService
      Case "PerfRawData_InetInfo_InternetInformationServicesGlobal": Set col_Win32 = PerfRawData_InetInfo_InternetInformationServicesGlobal
      Case "PerfRawData_MSDTC_DistributedTransactionCoordinator": Set col_Win32 = PerfRawData_MSDTC_DistributedTransactionCoordinator
      Case "PerfRawData_NTFSDRV_SMTPNTFSStoreDriver": Set col_Win32 = PerfRawData_NTFSDRV_SMTPNTFSStoreDriver
      Case "PerfRawData_PSched_PSchedFlow": Set col_Win32 = PerfRawData_PSched_PSchedFlow
      Case "PerfRawData_PSched_PSchedPipe": Set col_Win32 = PerfRawData_PSched_PSchedPipe
      Case "PerfRawData_PerfDisk_LogicalDisk": Set col_Win32 = PerfRawData_PerfDisk_LogicalDisk
      Case "PerfRawData_PerfDisk_PhysicalDisk": Set col_Win32 = PerfRawData_PerfDisk_PhysicalDisk
      Case "PerfRawData_PerfNet_Browser": Set col_Win32 = PerfRawData_PerfNet_Browser
      Case "PerfRawData_PerfNet_Redirector": Set col_Win32 = PerfRawData_PerfNet_Redirector
      Case "PerfRawData_PerfNet_Server": Set col_Win32 = PerfRawData_PerfNet_Server
      Case "PerfRawData_PerfNet_ServerWorkQueues": Set col_Win32 = PerfRawData_PerfNet_ServerWorkQueues
      Case "PerfRawData_PerfOS_Cache": Set col_Win32 = PerfRawData_PerfOS_Cache
      Case "PerfRawData_PerfOS_Memory": Set col_Win32 = PerfRawData_PerfOS_Memory
      Case "PerfRawData_PerfOS_Objects": Set col_Win32 = PerfRawData_PerfOS_Objects
      Case "PerfRawData_PerfOS_PagingFile": Set col_Win32 = PerfRawData_PerfOS_PagingFile
      Case "PerfRawData_PerfOS_Processor": Set col_Win32 = PerfRawData_PerfOS_Processor
      Case "PerfRawData_PerfOS_System": Set col_Win32 = PerfRawData_PerfOS_System
      Case "PerfRawData_PerfProc_FullImage_Costly": Set col_Win32 = PerfRawData_PerfProc_FullImage_Costly
      Case "PerfRawData_PerfProc_Image_Costly": Set col_Win32 = PerfRawData_PerfProc_Image_Costly
      Case "PerfRawData_PerfProc_JobObject": Set col_Win32 = PerfRawData_PerfProc_JobObject
      Case "PerfRawData_PerfProc_JobObjectDetails": Set col_Win32 = PerfRawData_PerfProc_JobObjectDetails
      Case "PerfRawData_PerfProc_Process": Set col_Win32 = PerfRawData_PerfProc_Process
      Case "PerfRawData_PerfProc_ProcessAddressSpace_Costly": Set col_Win32 = PerfRawData_PerfProc_ProcessAddressSpace_Costly
      Case "PerfRawData_PerfProc_Thread": Set col_Win32 = PerfRawData_PerfProc_Thread
      Case "PerfRawData_PerfProc_ThreadDetails_Costly": Set col_Win32 = PerfRawData_PerfProc_ThreadDetails_Costly
      Case "PerfRawData_RSVP_ACSRSVPInterfaces": Set col_Win32 = PerfRawData_RSVP_ACSRSVPInterfaces
      Case "PerfRawData_RSVP_ACSRSVPService": Set col_Win32 = PerfRawData_RSVP_ACSRSVPService
      Case "PerfRawData_RemoteAccess_RASPort": Set col_Win32 = PerfRawData_RemoteAccess_RASPort
      Case "PerfRawData_RemoteAccess_RASTotal": Set col_Win32 = PerfRawData_RemoteAccess_RASTotal
      Case "PerfRawData_SMTPSVC_SMTPServer": Set col_Win32 = PerfRawData_SMTPSVC_SMTPServer
      Case "PerfRawData_Spooler_PrintQueue": Set col_Win32 = PerfRawData_Spooler_PrintQueue
      Case "PerfRawData_TapiSrv_Telephony": Set col_Win32 = PerfRawData_TapiSrv_Telephony
      Case "PerfRawData_Tcpip_ICMP": Set col_Win32 = PerfRawData_Tcpip_ICMP
      Case "PerfRawData_Tcpip_IP": Set col_Win32 = PerfRawData_Tcpip_IP
      Case "PerfRawData_Tcpip_NBTConnection": Set col_Win32 = PerfRawData_Tcpip_NBTConnection
      Case "PerfRawData_Tcpip_NetworkInterface": Set col_Win32 = PerfRawData_Tcpip_NetworkInterface
      Case "PerfRawData_Tcpip_TCP": Set col_Win32 = PerfRawData_Tcpip_TCP
      Case "PerfRawData_Tcpip_UDP": Set col_Win32 = PerfRawData_Tcpip_UDP
      Case "PerfRawData_TermService_TerminalServices": Set col_Win32 = PerfRawData_TermService_TerminalServices
      Case "PerfRawData_TermService_TerminalServicesSession": Set col_Win32 = PerfRawData_TermService_TerminalServicesSession
      Case "PerfRawData_W3SVC_WebService": Set col_Win32 = PerfRawData_W3SVC_WebService
      Case "PhysicalMedia": Set col_Win32 = PhysicalMedia
      Case "PhysicalMemory": Set col_Win32 = PhysicalMemory
      Case "PhysicalMemoryArray": Set col_Win32 = PhysicalMemoryArray
      Case "PingStatus": Set col_Win32 = PingStatus
      Case "PnPEntity": Set col_Win32 = PnPEntity
      Case "PnPSignedDriver": Set col_Win32 = PnPSignedDriver
      Case "PointingDevice": Set col_Win32 = PointingDevice
      Case "PortConnector": Set col_Win32 = PortConnector
      Case "PortResource": Set col_Win32 = PortResource
      Case "PortableBattery": Set col_Win32 = PortableBattery
      Case "PowerManagementEvent": Set col_Win32 = PowerManagementEvent
      Case "PrintJob": Set col_Win32 = PrintJob
      Case "Printer": Set col_Win32 = Printer
      Case "PrinterConfiguration": Set col_Win32 = PrinterConfiguration
      Case "PrinterDriver": Set col_Win32 = PrinterDriver
      Case "PrivilegesStatus": Set col_Win32 = PrivilegesStatus
      Case "Process": Set col_Win32 = Process
      Case "ProcessStartTrace": Set col_Win32 = ProcessStartTrace
      Case "ProcessStartup": Set col_Win32 = ProcessStartup
      Case "ProcessStopTrace": Set col_Win32 = ProcessStopTrace
      Case "ProcessTrace": Set col_Win32 = ProcessTrace
      Case "Processor": Set col_Win32 = Processor
      Case "Product": Set col_Win32 = Product
      Case "ProgIDSpecification": Set col_Win32 = ProgIDSpecification
      Case "ProgramGroup": Set col_Win32 = ProgramGroup
      Case "ProgramGroupOrItem": Set col_Win32 = ProgramGroupOrItem
      Case "Property": Set col_Win32 = Property
      Case "Proxy": Set col_Win32 = Proxy
      Case "PublishComponentAction": Set col_Win32 = PublishComponentAction
      Case "QuickFixEngineering": Set col_Win32 = QuickFixEngineering
      Case "QuotaSetting": Set col_Win32 = QuotaSetting
      Case "Refrigeration": Set col_Win32 = Refrigeration
      Case "Registry": Set col_Win32 = Registry
      Case "RegistryAction": Set col_Win32 = RegistryAction
      Case "RemoveFileAction": Set col_Win32 = RemoveFileAction
      Case "RemoveIniAction": Set col_Win32 = RemoveIniAction
      Case "ReserveCost": Set col_Win32 = ReserveCost
      Case "SCSIController": Set col_Win32 = SCSIController
      Case "SID": Set col_Win32 = SID
      Case "SMBIOSMemory": Set col_Win32 = SMBIOSMemory
      Case "ScheduledJob": Set col_Win32 = ScheduledJob
      Case "SecuritySetting": Set col_Win32 = SecuritySetting
      Case "SelfRegModuleAction": Set col_Win32 = SelfRegModuleAction
      Case "SerialPort": Set col_Win32 = SerialPort
      Case "SerialPortConfiguration": Set col_Win32 = SerialPortConfiguration
      Case "ServerConnection": Set col_Win32 = ServerConnection
      Case "ServerSession": Set col_Win32 = ServerSession
      Case "Service": Set col_Win32 = Service
      Case "ServiceControl": Set col_Win32 = ServiceControl
      Case "ServiceSpecification": Set col_Win32 = ServiceSpecification
      Case "Session": Set col_Win32 = Session
      Case "ShadowContext": Set col_Win32 = ShadowContext
      Case "ShadowCopy": Set col_Win32 = ShadowCopy
      Case "ShadowProvider": Set col_Win32 = ShadowProvider
      Case "Share": Set col_Win32 = Share
      Case "ShortcutAction": Set col_Win32 = ShortcutAction
      Case "ShortcutFile": Set col_Win32 = ShortcutFile
      Case "SoftwareElement": Set col_Win32 = SoftwareElement
      Case "SoftwareElementCondition": Set col_Win32 = SoftwareElementCondition
      Case "SoftwareFeature": Set col_Win32 = SoftwareFeature
      Case "SoundDevice": Set col_Win32 = SoundDevice
      Case "StartupCommand": Set col_Win32 = StartupCommand
      Case "SystemAccount": Set col_Win32 = SystemAccount
      Case "SystemConfigurationChangeEvent": Set col_Win32 = SystemConfigurationChangeEvent
      Case "SystemEnclosure": Set col_Win32 = SystemEnclosure
      Case "SystemMemoryResource": Set col_Win32 = SystemMemoryResource
      Case "SystemSlot": Set col_Win32 = SystemSlot
      Case "SystemTrace": Set col_Win32 = SystemTrace
      Case "TCPIPPrinterPort": Set col_Win32 = TCPIPPrinterPort
      Case "TapeDrive": Set col_Win32 = TapeDrive
      Case "TemperatureProbe": Set col_Win32 = TemperatureProbe
      Case "Thread": Set col_Win32 = Thread
      Case "ThreadStartTrace": Set col_Win32 = ThreadStartTrace
      Case "ThreadStopTrace": Set col_Win32 = ThreadStopTrace
      Case "ThreadTrace": Set col_Win32 = ThreadTrace
      Case "TimeZone": Set col_Win32 = TimeZone
      Case "Trustee": Set col_Win32 = Trustee
      Case "TypeLibraryAction": Set col_Win32 = TypeLibraryAction
      Case "USBController": Set col_Win32 = USBController
      Case "USBHub": Set col_Win32 = USBHub
      Case "UTCTime": Set col_Win32 = UTCTime
      Case "UninterruptiblePowerSupply": Set col_Win32 = UninterruptiblePowerSupply
      Case "UserAccount": Set col_Win32 = UserAccount
      Case "VideoConfiguration": Set col_Win32 = VideoConfiguration
      Case "VideoController": Set col_Win32 = VideoController
      Case "VoltageProbe": Set col_Win32 = VoltageProbe
      Case "Volume": Set col_Win32 = Volume
      Case "VolumeChangeEvent": Set col_Win32 = VolumeChangeEvent
      Case "WMISetting": Set col_Win32 = WMISetting
      Case "WindowsProductActivation": Set col_Win32 = WindowsProductActivation
   End Select
End Function

Private Function Account() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Caption"
      .Add "Description"
      .Add "Domain"
      .Add "InstallDate"
      .Add "LocalAccount"
      .Add "Name"
      .Add "SID"
      .Add "SIDType"
      .Add "Status"
   End With
   Set Account = c
End Function

Private Function Battery() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Availability"
      .Add "BatteryRechargeTime"
      .Add "BatteryStatus"
      .Add "Caption"
      .Add "Chemistry"
      .Add "ConfigManagerErrorCode"
      .Add "ConfigManagerUserConfig"
      .Add "CreationClassName"
      .Add "Description"
      .Add "DesignCapacity"
      .Add "DesignVoltage"
      .Add "DeviceID"
      .Add "ErrorCleared"
      .Add "ErrorDescription"
      .Add "EstimatedChargeRemaining"
      .Add "EstimatedRunTime"
      .Add "ExpectedBatteryLife"
      .Add "ExpectedLife"
      .Add "FullChargeCapacity"
      .Add "InstallDate"
      .Add "LastErrorCode"
      .Add "MaxRechargeTime"
      .Add "Name"
      .Add "PNPDeviceID"
      .Add "PowerManagementCapabilities[]"
      .Add "PowerManagementSupported"
      .Add "SmartBatteryVersion"
      .Add "Status"
      .Add "StatusInfo"
      .Add "SystemCreationClassName"
      .Add "SystemName"
      .Add "TimeOnBattery"
      .Add "TimeToFullCharge"
   End With
   Set Battery = c
End Function
Private Function Bios() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "BiosCharacteristics[]"
      .Add "BIOSVersion[]"
      .Add "BuildNumber"
      .Add "Caption"
      .Add "CodeSet"
      .Add "CurrentLanguage"
      .Add "Description"
      .Add "IdentificationCode"
      .Add "InstallableLanguages"
      .Add "InstallDate"
      .Add "LanguageEdition"
      .Add "ListOfLanguages[]"
      .Add "Manufacturer"
      .Add "Name"
      .Add "OtherTargetOS"
      .Add "PrimaryBIOS"
      .Add "ReleaseDate"
      .Add "SerialNumber"
      .Add "SMBIOSBIOSVersion"
      .Add "SMBIOSMajorVersion"
      .Add "SMBIOSMinorVersion"
      .Add "SMBIOSPresent"
      .Add "SoftwareElementID"
      .Add "SoftwareElementState"
      .Add "Status"
      .Add "TargetOperatingSystem"
      .Add "Version"
   End With
   Set Bios = c
End Function
Private Function Service() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "AcceptPause"
      .Add "AcceptStop"
      .Add "Caption"
      .Add "CheckPoint"
      .Add "CreationClassName"
      .Add "Description"
      .Add "DesktopInteract"
      .Add "DisplayName"
      .Add "ErrorControl"
      .Add "ExitCode"
      .Add "InstallDate"
      .Add "Name"
      .Add "PathName"
      .Add "ProcessId"
      .Add "ServiceSpecificExitCode"
      .Add "ServiceType"
      .Add "Started"
      .Add "StartMode"
      .Add "StartName"
      .Add "State"
      .Add "Status"
      .Add "SystemCreationClassName"
      .Add "SystemName"
      .Add "TagId"
      .Add "WaitHint"
   End With
   Set Service = c
End Function
Private Function ODBCDataSourceSpecification() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Caption"
      .Add "CheckID"
      .Add "CheckMode"
      .Add "DataSource"
      .Add "Description"
      .Add "DriverDescription"
      .Add "Name"
      .Add "Registration"
      .Add "SoftwareElementID"
      .Add "SoftwareElementState"
      .Add "TargetOperatingSystem"
      .Add "Version"
   End With
   Set ODBCDataSourceSpecification = c
End Function
Private Function ODBCAttribute() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Attribute"
      .Add "Caption"
      .Add "Description"
      .Add "Driver"
      .Add "SettingID"
      .Add "Value"
   End With
   Set ODBCAttribute = c
End Function
Private Function Share() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "AccessMask"
      .Add "AllowMaximum"
      .Add "Caption"
      .Add "Description"
      .Add "InstallDate"
      .Add "MaximumAllowed"
      .Add "Name"
      .Add "Path"
      .Add "Status"
      .Add "Type"
   End With
   Set Share = c
End Function
Private Function ScheduledJob() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Caption"
      .Add "Command"
      .Add "DaysOfMonth"
      .Add "DaysOfWeek"
      .Add "Description"
      .Add "ElapsedTime"
      .Add "InstallDate"
      .Add "InteractWithDesktop"
      .Add "JobId"
      .Add "JobStatus"
      .Add "Name"
      .Add "Notify"
      .Add "Owner"
      .Add "Priority"
      .Add "RunRepeatedly"
      .Add "StartTime"
      .Add "Status"
      .Add "TimeSubmitted"
      .Add "UntilTime"
   End With
   Set ScheduledJob = c
End Function
Private Function UserAccount() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "AccountType"
      .Add "Caption"
      .Add "Description"
      .Add "Disabled"
      .Add "Domain"
      .Add "FullName"
      .Add "InstallDate"
      .Add "LocalAccount"
      .Add "Lockout"
      .Add "Name"
      .Add "PasswordChangeable"
      .Add "PasswordExpires"
      .Add "PasswordRequired"
      .Add "SID"
      .Add "SIDType"
      .Add "Status"
   End With
   Set UserAccount = c
End Function
Private Function CodecFile() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "AccessMask"
      .Add "Archive"
      .Add "Caption"
      .Add "Compressed"
      .Add "CompressionMethod"
      .Add "CreationClassName"
      .Add "CreationDate"
      .Add "CSCreationClassName"
      .Add "CSName"
      .Add "Description"
      .Add "Drive"
      .Add "EightDotThreeFileName"
      .Add "Encrypted"
      .Add "EncryptionMethod"
      .Add "Extension"
      .Add "FileName"
      .Add "FileSize"
      .Add "FileType"
      .Add "FSCreationClassName"
      .Add "FSName"
      .Add "Group"
      .Add "Hidden"
      .Add "InstallDate"
      .Add "InUseCount"
      .Add "LastAccessed"
      .Add "LastModified"
      .Add "Manufacturer"
      .Add "Name"
      .Add "Path"
      .Add "Readable"
      .Add "Status"
      .Add "System"
      .Add "Version"
      .Add "Writeable"
   End With
   Set CodecFile = c
End Function
Private Function Directory() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "AccessMask[]"
      .Add "Archive"
      .Add "Caption"
      .Add "Compressed"
      .Add "CompressionMethod"
      .Add "CreationClassName"
      .Add "CreationDate"
      .Add "CSCreationClassName"
      .Add "CSName"
      .Add "Description"
      .Add "Drive"
      .Add "EightDotThreeFileName"
      .Add "Encrypted"
      .Add "EncryptionMethod"
      .Add "Extension"
      .Add "FileName"
      .Add "FileSize"
      .Add "FileType"
      .Add "FSCreationClassName"
      .Add "FSName"
      .Add "Hidden"
      .Add "InstallDate"
      .Add "InUseCount"
      .Add "LastAccessed"
      .Add "LastModified"
      .Add "Name"
      .Add "Path"
      .Add "Readable"
      .Add "Status"
      .Add "System"
      .Add "Writeable"
   End With
   Set Directory = c
End Function
Private Function DiskDrive() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Availability"
      .Add "BytesPerSector"
      .Add "Capabilities[]"
      .Add "CapabilityDescriptions[]"
      .Add "Caption"
      .Add "CompressionMethod"
      .Add "ConfigManagerErrorCode"
      .Add "ConfigManagerUserConfig"
      .Add "CreationClassName"
      .Add "DefaultBlockSize"
      .Add "Description"
      .Add "DeviceID"
      .Add "ErrorCleared"
      .Add "ErrorDescription"
      .Add "ErrorMethodology"
      .Add "Index"
      .Add "InstallDate"
      .Add "InterfaceType"
      .Add "LastErrorCode"
      .Add "Manufacturer"
      .Add "MaxBlockSize"
      .Add "MaxMediaSize"
      .Add "MediaLoaded"
      .Add "MediaType"
      .Add "MinBlockSize"
      .Add "Model"
      .Add "Name"
      .Add "NeedsCleaning"
      .Add "NumberOfMediaSupported"
      .Add "Partitions"
      .Add "PNPDeviceID"
      .Add "PowerManagementCapabilities[]"
      .Add "PowerManagementSupported"
      .Add "SCSIBus"
      .Add "SCSILogicalUnit"
      .Add "SCSIPort"
      .Add "SCSITargetId"
      .Add "SectorsPerTrack"
      .Add "Signature"
      .Add "Size"
      .Add "Status"
      .Add "StatusInfo"
      .Add "SystemCreationClassName"
      .Add "SystemName"
      .Add "TotalCylinders"
      .Add "TotalHeads"
      .Add "TotalSectors"
      .Add "TotalTracks"
      .Add "TracksPerCylinder"
   End With
   Set DiskDrive = c
End Function
Private Function DiskPartition() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Access"
      .Add "Availability"
      .Add "BlockSize"
      .Add "Bootable"
      .Add "BootPartition"
      .Add "Caption"
      .Add "ConfigManagerErrorCode"
      .Add "ConfigManagerUserConfig"
      .Add "CreationClassName"
      .Add "Description"
      .Add "DeviceID"
      .Add "DiskIndex"
      .Add "ErrorCleared"
      .Add "ErrorDescription"
      .Add "ErrorMethodology"
      .Add "HiddenSectors"
      .Add "Index"
      .Add "InstallDate"
      .Add "LastErrorCode"
      .Add "Name"
      .Add "NumberOfBlocks"
      .Add "PNPDeviceID"
      .Add "PowerManagementCapabilities[]"
      .Add "PowerManagementSupported"
      .Add "PrimaryPartition"
      .Add "Purpose"
      .Add "RewritePartition"
      .Add "Size"
      .Add "StartingOffset"
      .Add "Status"
      .Add "StatusInfo"
      .Add "SystemCreationClassName"
      .Add "SystemName"
      .Add "Type"
   End With
   Set DiskPartition = c
End Function
Private Function DisplayConfiguration() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "BitsPerPel"
      .Add "Caption"
      .Add "Description"
      .Add "DeviceName"
      .Add "DisplayFlags"
      .Add "DisplayFrequency"
      .Add "DitherType"
      .Add "DriverVersion"
      .Add "ICMIntent"
      .Add "ICMMethod"
      .Add "LogPixels"
      .Add "PelsHeight"
      .Add "PelsWidth"
      .Add "SettingID"
      .Add "SpecificationVersion"
   End With
   Set DisplayConfiguration = c
End Function
Private Function DisplayControllerConfiguration() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "BitsPerPixel"
      .Add "Caption"
      .Add "ColorPlanes"
      .Add "Description"
      .Add "DeviceEntriesInAColorTable"
      .Add "DeviceSpecificPens"
      .Add "HorizontalResolution"
      .Add "Name"
      .Add "RefreshRate"
      .Add "ReservedSystemPaletteEntries"
      .Add "SettingID"
      .Add "SystemPaletteEntries"
      .Add "VerticalResolution"
      .Add "VideoMode"
   End With
   Set DisplayControllerConfiguration = c
End Function
Private Function DMAChannel() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "AddressSize"
      .Add "Availability"
      .Add "BurstMode"
      .Add "ByteMode"
      .Add "Caption"
      .Add "ChannelTiming"
      .Add "CreationClassName"
      .Add "CSCreationClassName"
      .Add "CSName"
      .Add "Description"
      .Add "DMAChannel"
      .Add "InstallDate"
      .Add "MaxTransferSize"
      .Add "Name"
      .Add "Port"
      .Add "Status"
      .Add "TransferWidths[]"
      .Add "TypeCTiming"
      .Add "WordMode"
   End With
   Set DMAChannel = c
End Function
Private Function DriverVXD() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "BuildNumber"
      .Add "Caption"
      .Add "CodeSet"
      .Add "Control"
      .Add "Description"
      .Add "DeviceDescriptorBlock"
      .Add "IdentificationCode"
      .Add "InstallDate"
      .Add "LanguageEdition"
      .Add "Manufacturer"
      .Add "Name"
      .Add "OtherTargetOS"
      .Add "PM_API"
      .Add "SerialNumber"
      .Add "ServiceTableSize"
      .Add "SoftwareElementID"
      .Add "SoftwareElementState"
      .Add "Status"
      .Add "TargetOperatingSystem"
      .Add "V86_API"
      .Add "Version"
   End With
   Set DriverVXD = c
End Function
Private Function DuplicateFileAction() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "ActionID"
      .Add "Caption"
      .Add "DeleteAfterCopy"
      .Add "Description"
      .Add "Destination"
      .Add "Direction"
      .Add "FileKey"
      .Add "Name"
      .Add "SoftwareElementID"
      .Add "SoftwareElementState"
      .Add "Source"
      .Add "TargetOperatingSystem"
      .Add "Version"
   End With
   Set DuplicateFileAction = c
End Function
Private Function Environment() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Caption"
      .Add "Description"
      .Add "InstallDate"
      .Add "Name"
      .Add "Status"
      .Add "SystemVariable"
      .Add "UserName"
      .Add "VariableValue"
   End With
   Set Environment = c
End Function
Private Function EnvironmentSpecification() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Caption"
      .Add "CheckID"
      .Add "CheckMode"
      .Add "Description"
      .Add "Environment"
      .Add "Name"
      .Add "SoftwareElementID"
      .Add "SoftwareElementState"
      .Add "TargetOperatingSystem"
      .Add "Value"
      .Add "Version"
   End With
   Set EnvironmentSpecification = c
End Function
Private Function ExtensionInfoAction() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "ActionID"
      .Add "Argument"
      .Add "Caption"
      .Add "Command"
      .Add "Description"
      .Add "Direction"
      .Add "Extension"
      .Add "MIME"
      .Add "Name"
      .Add "ProgID"
      .Add "ShellNew"
      .Add "ShellNewValue"
      .Add "SoftwareElementID"
      .Add "SoftwareElementState"
      .Add "TargetOperatingSystem"
      .Add "Verb"
      .Add "Version"
   End With
   Set ExtensionInfoAction = c
End Function
Private Function Fan() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "ActiveCooling"
      .Add "Availability"
      .Add "Caption"
      .Add "ConfigManagerErrorCode"
      .Add "ConfigManagerUserConfig"
      .Add "CreationClassName"
      .Add "Description"
      .Add "DesiredSpeed"
      .Add "DeviceID"
      .Add "ErrorCleared"
      .Add "ErrorDescription"
      .Add "InstallDate"
      .Add "LastErrorCode"
      .Add "Name"
      .Add "PNPDeviceID"
      .Add "PowerManagementCapabilities[]"
      .Add "PowerManagementSupported"
      .Add "Status"
      .Add "StatusInfo"
      .Add "SystemCreationClassName"
      .Add "SystemName"
      .Add "VariableSpeed"
   End With
   Set Fan = c
End Function
Private Function FileSpecification() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Attributes"
      .Add "Caption"
      .Add "CheckID"
      .Add "CheckMode"
      .Add "CheckSum"
      .Add "CRC1"
      .Add "CRC2"
      .Add "CreateTimeStamp"
      .Add "Description"
      .Add "FileID"
      .Add "FileSize"
      .Add "Language"
      .Add "MD5CheckSum"
      .Add "Name"
      .Add "Sequence"
      .Add "SoftwareElementID"
      .Add "SoftwareElementState"
      .Add "TargetOperatingSystem"
      .Add "Version"
   End With
   Set FileSpecification = c
End Function
Private Function FloppyController() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Availability"
      .Add "Caption"
      .Add "ConfigManagerErrorCode"
      .Add "ConfigManagerUserConfig"
      .Add "CreationClassName"
      .Add "Description"
      .Add "DeviceID"
      .Add "ErrorCleared"
      .Add "ErrorDescription"
      .Add "InstallDate"
      .Add "LastErrorCode"
      .Add "Manufacturer"
      .Add "MaxNumberControlled"
      .Add "Name"
      .Add "PNPDeviceID"
      .Add "PowerManagementCapabilities[]"
      .Add "PowerManagementSupported"
      .Add "ProtocolSupported"
      .Add "Status"
      .Add "StatusInfo"
      .Add "SystemCreationClassName"
      .Add "SystemName"
      .Add "TimeOfLastReset"
   End With
   Set FloppyController = c
End Function
Private Function FloppyDrive() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Availability"
      .Add "Capabilities[]"
      .Add "CapabilityDescriptions[]"
      .Add "Caption"
      .Add "CompressionMethod"
      .Add "ConfigManagerErrorCode"
      .Add "ConfigManagerUserConfig"
      .Add "CreationClassName"
      .Add "DefaultBlockSize"
      .Add "Description"
      .Add "DeviceID"
      .Add "ErrorCleared"
      .Add "ErrorDescription"
      .Add "ErrorMethodology"
      .Add "InstallDate"
      .Add "LastErrorCode"
      .Add "Manufacturer"
      .Add "MaxBlockSize"
      .Add "MaxMediaSize"
      .Add "MinBlockSize"
      .Add "Name"
      .Add "NeedsCleaning"
      .Add "NumberOfMediaSupported"
      .Add "PNPDeviceID"
      .Add "PowerManagementCapabilities[]"
      .Add "PowerManagementSupported"
      .Add "Status"
      .Add "StatusInfo"
      .Add "SystemCreationClassName"
      .Add "SystemName"
   End With
   Set FloppyDrive = c
End Function
Private Function FontInfoAction() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "ActionID"
      .Add "Caption"
      .Add "Description"
      .Add "Direction"
      .Add "File"
      .Add "FontTitle"
      .Add "Name"
      .Add "SoftwareElementID"
      .Add "SoftwareElementState"
      .Add "TargetOperatingSystem"
      .Add "Version"
   End With
   Set FontInfoAction = c
End Function
Private Function Group() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Caption"
      .Add "Description"
      .Add "Domain"
      .Add "InstallDate"
      .Add "LocalAccount"
      .Add "Name"
      .Add "SID"
      .Add "SIDType"
      .Add "Status"
   End With
   Set Group = c
End Function
Private Function HeatPipe() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "ActiveCooling"
      .Add "Availability"
      .Add "Caption"
      .Add "ConfigManagerErrorCode"
      .Add "ConfigManagerUserConfig"
      .Add "CreationClassName"
      .Add "Description"
      .Add "DeviceID"
      .Add "ErrorCleared"
      .Add "ErrorDescription"
      .Add "InstallDate"
      .Add "LastErrorCode"
      .Add "Name"
      .Add "PNPDeviceID"
      .Add "PowerManagementCapabilities[]"
      .Add "PowerManagementSupported"
      .Add "Status"
      .Add "StatusInfo"
      .Add "SystemCreationClassName"
      .Add "SystemName"
   End With
   Set HeatPipe = c
End Function
Private Function IDEController() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Availability"
      .Add "Caption"
      .Add "ConfigManagerErrorCode"
      .Add "ConfigManagerUserConfig"
      .Add "CreationClassName"
      .Add "Description"
      .Add "DeviceID"
      .Add "ErrorCleared"
      .Add "ErrorDescription"
      .Add "InstallDate"
      .Add "LastErrorCode"
      .Add "Manufacturer"
      .Add "MaxNumberControlled"
      .Add "Name"
      .Add "PNPDeviceID"
      .Add "PowerManagementCapabilities[]"
      .Add "PowerManagementSupported"
      .Add "ProtocolSupported"
      .Add "Status"
      .Add "StatusInfo"
      .Add "SystemCreationClassName"
      .Add "SystemName"
      .Add "TimeOfLastReset"
   End With
   Set IDEController = c
End Function
Private Function InfraredDevice() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Availability"
      .Add "Caption"
      .Add "ConfigManagerErrorCode"
      .Add "ConfigManagerUserConfig"
      .Add "CreationClassName"
      .Add "Description"
      .Add "DeviceID"
      .Add "ErrorCleared"
      .Add "ErrorDescription"
      .Add "InstallDate"
      .Add "LastErrorCode"
      .Add "Manufacturer"
      .Add "MaxNumberControlled"
      .Add "Name"
      .Add "PNPDeviceID"
      .Add "PowerManagementCapabilities[]"
      .Add "PowerManagementSupported"
      .Add "ProtocolSupported"
      .Add "Status"
      .Add "StatusInfo"
      .Add "SystemCreationClassName"
      .Add "SystemName"
      .Add "TimeOfLastReset"
   End With
   Set InfraredDevice = c
End Function
Private Function IniFileSpecification() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Action"
      .Add "Caption"
      .Add "CheckID"
      .Add "CheckMode"
      .Add "CheckSum"
      .Add "CRC1"
      .Add "CRC2"
      .Add "CreateTimeStamp"
      .Add "Description"
      .Add "FileSize"
      .Add "IniFile"
      .Add "Key"
      .Add "MD5Checksum"
      .Add "Name"
      .Add "Section"
      .Add "SoftwareElementID"
      .Add "SoftwareElementState"
      .Add "TargetOperatingSystem"
      .Add "Value"
      .Add "Version"
   End With
   Set IniFileSpecification = c
End Function
Private Function IP4PersistedRouteTable() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Caption"
      .Add "Description"
      .Add "Destination"
      .Add "InstallDate"
      .Add "Mask"
      .Add "Metric1"
      .Add "Name"
      .Add "NextHop"
      .Add "Status"
   End With
   Set IP4PersistedRouteTable = c
End Function
Private Function IP4RouteTable() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Age"
      .Add "Caption"
      .Add "Description"
      .Add "Destination"
      .Add "Information"
      .Add "InstallDate"
      .Add "InterfaceIndex"
      .Add "Mask"
      .Add "Metric1"
      .Add "Metric2"
      .Add "Metric3"
      .Add "Metric4"
      .Add "Metric5"
      .Add "Name"
      .Add "NextHop"
      .Add "Protocol"
      .Add "Status"
      .Add "Type"
   End With
   Set IP4RouteTable = c
End Function
Private Function IRQResource() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Availability"
      .Add "Caption"
      .Add "CreationClassName"
      .Add "CSCreationClassName"
      .Add "CSName"
      .Add "Description"
      .Add "Hardware"
      .Add "InstallDate"
      .Add "IRQNumber"
      .Add "Name"
      .Add "Shareable"
      .Add "Status"
      .Add "TriggerLevel"
      .Add "TriggerType"
      .Add "Vector"
   End With
   Set IRQResource = c
End Function
Private Function JobObjectStatus() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "AdditionalDescription"
      .Add "Description"
      .Add "Operation"
      .Add "ParameterInfo"
      .Add "ProviderName"
      .Add "StatusCode"
      .Add "Win32ErrorCode"
   End With
   Set JobObjectStatus = c
End Function
Private Function Keyboard() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Availability"
      .Add "Caption"
      .Add "ConfigManagerErrorCode"
      .Add "ConfigManagerUserConfig"
      .Add "CreationClassName"
      .Add "Description"
      .Add "DeviceID"
      .Add "ErrorCleared"
      .Add "ErrorDescription"
      .Add "InstallDate"
      .Add "IsLocked"
      .Add "LastErrorCode"
      .Add "Layout"
      .Add "Name"
      .Add "NumberOfFunctionKeys"
      .Add "Password"
      .Add "PNPDeviceID"
      .Add "PowerManagementCapabilities[]"
      .Add "PowerManagementSupported"
      .Add "Status"
      .Add "StatusInfo"
      .Add "SystemCreationClassName"
      .Add "SystemName"
   End With
   Set Keyboard = c
End Function
Private Function LaunchCondition() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Caption"
      .Add "CheckID"
      .Add "CheckMode"
      .Add "Condition"
      .Add "Description"
      .Add "Name"
      .Add "SoftwareElementID"
      .Add "SoftwareElementState"
      .Add "TargetOperatingSystem"
      .Add "Version"
   End With
   Set LaunchCondition = c
End Function
Private Function LoadOrderGroup() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Caption"
      .Add "Description"
      .Add "DriverEnabled"
      .Add "GroupOrder"
      .Add "InstallDate"
      .Add "Name"
      .Add "Status"
   End With
   Set LoadOrderGroup = c
End Function
Private Function LocalTime() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Day"
      .Add "DayOfWeek"
      .Add "Hour"
      .Add "Milliseconds"
      .Add "Minute"
      .Add "Month"
      .Add "Quarter"
      .Add "Second"
      .Add "WeekInMonth"
      .Add "Year"
   End With
   Set LocalTime = c
End Function
Private Function LogicalDisk() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Access"
      .Add "Availability"
      .Add "BlockSize"
      .Add "Caption"
      .Add "Compressed"
      .Add "ConfigManagerErrorCode"
      .Add "ConfigManagerUserConfig"
      .Add "CreationClassName"
      .Add "Description"
      .Add "DeviceID"
      .Add "DriveType"
      .Add "ErrorCleared"
      .Add "ErrorDescription"
      .Add "ErrorMethodology"
      .Add "FileSystem"
      .Add "FreeSpace"
      .Add "InstallDate"
      .Add "LastErrorCode"
      .Add "MaximumComponentLength"
      .Add "MediaType"
      .Add "Name"
      .Add "NumberOfBlocks"
      .Add "PNPDeviceID"
      .Add "PowerManagementCapabilities[]"
      .Add "PowerManagementSupported"
      .Add "ProviderName"
      .Add "Purpose"
      .Add "QuotasDisabled"
      .Add "QuotasIncomplete"
      .Add "QuotasRebuilding"
      .Add "Size"
      .Add "Status"
      .Add "StatusInfo"
      .Add "SupportsDiskQuotas"
      .Add "SupportsFileBasedCompression"
      .Add "SystemCreationClassName"
      .Add "SystemName"
      .Add "VolumeDirty"
      .Add "VolumeName"
      .Add "VolumeSerialNumber"
   End With
   Set LogicalDisk = c
End Function
Private Function LogicalFileSecuritySetting() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Caption"
      .Add "ControlFlags"
      .Add "Description"
      .Add "OwnerPermissions"
      .Add "Path"
      .Add "SettingID"
   End With
   Set LogicalFileSecuritySetting = c
End Function
Private Function LogicalMemoryConfiguration() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "AvailableVirtualMemory"
      .Add "Caption"
      .Add "Description"
      .Add "Name"
      .Add "SettingID"
      .Add "TotalPageFileSpace"
      .Add "TotalPhysicalMemory"
      .Add "TotalVirtualMemory"
   End With
   Set LogicalMemoryConfiguration = c
End Function
Private Function LogicalProgramGroup() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Caption"
      .Add "Description"
      .Add "GroupName"
      .Add "InstallDate"
      .Add "Name"
      .Add "Status"
      .Add "UserName"
   End With
   Set LogicalProgramGroup = c
End Function
Private Function LogicalProgramGroupItem() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Caption"
      .Add "Description"
      .Add "InstallDate"
      .Add "Name"
      .Add "Status"
   End With
   Set LogicalProgramGroupItem = c
End Function
Private Function LogicalShareSetting() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Caption"
      .Add "ControlFlags"
      .Add "Description"
      .Add "Name"
      .Add "SettingID"
   End With
   Set LogicalShareSetting = c
End Function
Private Function LogonSession() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "AuthenticationPackage"
      .Add "Caption"
      .Add "Description"
      .Add "InstallDate"
      .Add "LogonId"
      .Add "LogonType"
      .Add "Name"
      .Add "StartTime"
      .Add "Status"
   End With
   Set LogonSession = c
End Function
Private Function Process() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Caption"
      .Add "CommandLine"
      .Add "CreationClassName"
      .Add "CreationDate"
      .Add "CSCreationClassName"
      .Add "CSName"
      .Add "Description"
      .Add "ExecutablePath"
      .Add "ExecutionState"
      .Add "Handle"
      .Add "HandleCount"
      .Add "InstallDate"
      .Add "KernelModeTime"
      .Add "MaximumWorkingSetSize"
      .Add "MinimumWorkingSetSize"
      .Add "Name"
      .Add "OSCreationClassName"
      .Add "OSName"
      .Add "OtherOperationCount"
      .Add "OtherTransferCount"
      .Add "PageFaults"
      .Add "PageFileUsage"
      .Add "ParentProcessId"
      .Add "PeakPageFileUsage"
      .Add "PeakVirtualSize"
      .Add "PeakWorkingSetSize"
      .Add "Priority"
      .Add "PrivatePageCount"
      .Add "ProcessId"
      .Add "QuotaNonPagedPoolUsage"
      .Add "QuotaPagedPoolUsage"
      .Add "QuotaPeakNonPagedPoolUsage"
      .Add "QuotaPeakPagedPoolUsage"
      .Add "ReadOperationCount"
      .Add "ReadTransferCount"
      .Add "SessionId"
      .Add "Status"
      .Add "TerminationDate"
      .Add "ThreadCount"
      .Add "UserModeTime"
      .Add "VirtualSize"
      .Add "WindowsVersion"
      .Add "WorkingSetSize"
      .Add "WriteOperationCount"
      .Add "WriteTransferCount"
   End With
   Set Process = c
End Function
Private Function MappedLogicalDisk() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Access"
      .Add "Availability"
      .Add "BlockSize"
      .Add "Caption"
      .Add "Compressed"
      .Add "ConfigManagerErrorCode"
      .Add "ConfigManagerUserConfig"
      .Add "CreationClassName"
      .Add "Description"
      .Add "DeviceID"
      .Add "ErrorCleared"
      .Add "ErrorDescription"
      .Add "ErrorMethodology"
      .Add "FileSystem"
      .Add "FreeSpace"
      .Add "InstallDate"
      .Add "LastErrorCode"
      .Add "MaximumComponentLength"
      .Add "Name"
      .Add "NumberOfBlocks"
      .Add "PNPDeviceID"
      .Add "PowerManagementCapabilities"
      .Add "PowerManagementSupported"
      .Add "ProviderName"
      .Add "Purpose"
      .Add "QuotasDisabled"
      .Add "QuotasIncomplete"
      .Add "QuotasRebuilding"
      .Add "SessionID"
      .Add "Size"
      .Add "Status"
      .Add "StatusInfo"
      .Add "SupportsDiskQuotas"
      .Add "SupportsFileBasedCompression"
      .Add "SystemCreationClassName"
      .Add "SystemName"
      .Add "VolumeName"
      .Add "VolumeSerialNumber"
   End With
   Set MappedLogicalDisk = c
End Function
Private Function MemoryArray() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Access"
      .Add "AdditionalErrorData[]"
      .Add "Availability"
      .Add "BlockSize"
      .Add "Caption"
      .Add "ConfigManagerErrorCode"
      .Add "ConfigManagerUserConfig"
      .Add "CorrectableError"
      .Add "CreationClassName"
      .Add "Description"
      .Add "DeviceID"
      .Add "EndingAddress"
      .Add "ErrorAccess"
      .Add "ErrorAddress"
      .Add "ErrorCleared"
      .Add "ErrorData[]"
      .Add "ErrorDataOrder"
      .Add "ErrorDescription"
      .Add "ErrorGranularity"
      .Add "ErrorInfo"
      .Add "ErrorMethodology"
      .Add "ErrorResolution"
      .Add "ErrorTime"
      .Add "ErrorTransferSize"
      .Add "InstallDate"
      .Add "LastErrorCode"
      .Add "Name"
      .Add "NumberOfBlocks"
      .Add "OtherErrorDescription"
      .Add "PNPDeviceID"
      .Add "PowerManagementCapabilities[]"
      .Add "PowerManagementSupported"
      .Add "Purpose"
      .Add "StartingAddress"
      .Add "Status"
      .Add "StatusInfo"
      .Add "SystemCreationClassName"
      .Add "SystemLevelAddress"
      .Add "SystemName"
   End With
   Set MemoryArray = c
End Function
Private Function MemoryDevice() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Access"
      .Add "AdditionalErrorData[]"
      .Add "Availability"
      .Add "BlockSize"
      .Add "Caption"
      .Add "ConfigManagerErrorCode"
      .Add "ConfigManagerUserConfig"
      .Add "CorrectableError"
      .Add "CreationClassName"
      .Add "Description"
      .Add "DeviceID"
      .Add "EndingAddress"
      .Add "ErrorAccess"
      .Add "ErrorAddress"
      .Add "ErrorCleared"
      .Add "ErrorData[]"
      .Add "ErrorDataOrder"
      .Add "ErrorDescription"
      .Add "ErrorGranularity"
      .Add "ErrorInfo"
      .Add "ErrorMethodology"
      .Add "ErrorResolution"
      .Add "ErrorTime"
      .Add "ErrorTransferSize"
      .Add "InstallDate"
      .Add "LastErrorCode"
      .Add "Name"
      .Add "NumberOfBlocks"
      .Add "OtherErrorDescription"
      .Add "PNPDeviceID"
      .Add "PowerManagementCapabilities[]"
      .Add "PowerManagementSupported"
      .Add "Purpose"
      .Add "StartingAddress"
      .Add "Status"
      .Add "StatusInfo"
      .Add "SystemCreationClassName"
      .Add "SystemLevelAddress"
      .Add "SystemName"
   End With
   Set MemoryDevice = c
End Function
Private Function MIMEInfoAction() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "ActionID"
      .Add "Caption"
      .Add "CLSID"
      .Add "ContentType"
      .Add "Description"
      .Add "Direction"
      .Add "Extension"
      .Add "Name"
      .Add "SoftwareElementID"
      .Add "SoftwareElementState"
      .Add "TargetOperatingSystem"
      .Add "Version"
   End With
   Set MIMEInfoAction = c
End Function
Private Function ModuleLoadTrace() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "FileName"
      .Add "ImageBase"
      .Add "ImageSize"
      .Add "ProcessID"
      .Add "SECURITY_DESCRIPTOR[]"
      .Add "TIME_CREATED"
   End With
   Set ModuleLoadTrace = c
End Function
Private Function MotherboardDevice() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Availability"
      .Add "Caption"
      .Add "ConfigManagerErrorCode"
      .Add "ConfigManagerUserConfig"
      .Add "CreationClassName"
      .Add "Description"
      .Add "DeviceID"
      .Add "ErrorCleared"
      .Add "ErrorDescription"
      .Add "InstallDate"
      .Add "LastErrorCode"
      .Add "Name"
      .Add "PNPDeviceID"
      .Add "PowerManagementCapabilities[]"
      .Add "PowerManagementSupported"
      .Add "PrimaryBusType"
      .Add "RevisionNumber"
      .Add "SecondaryBusType"
      .Add "Status"
      .Add "StatusInfo"
      .Add "SystemCreationClassName"
      .Add "SystemName"
   End With
   Set MotherboardDevice = c
End Function
Private Function MoveFileAction() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "ActionID"
      .Add "Caption"
      .Add "Description"
      .Add "DestFolder"
      .Add "DestName"
      .Add "Direction"
      .Add "FileKey"
      .Add "Name"
      .Add "Options"
      .Add "SoftwareElementID"
      .Add "SoftwareElementState"
      .Add "SourceFolder"
      .Add "SourceName"
      .Add "TargetOperatingSystem"
      .Add "Version"
   End With
   Set MoveFileAction = c
End Function
Private Function MSIResource() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Caption"
      .Add "Description"
      .Add "SettingID"
   End With
   Set MSIResource = c
End Function
Private Function NamedJobObject() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "BasicUIRestrictions"
      .Add "Caption"
      .Add "CollectionID"
      .Add "Description"
   End With
   Set NamedJobObject = c
End Function
Private Function NamedJobObjectActgInfo() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "ActiveProcesses"
      .Add "Caption"
      .Add "Description"
      .Add "Name"
      .Add "OtherOperationCount"
      .Add "OtherTransferCount"
      .Add "PeakJobMemoryUsed"
      .Add "PeakProcessMemoryUsed"
      .Add "ReadOperationCount"
      .Add "ReadTransferCount"
      .Add "ThisPeriodTotalKernelTime"
      .Add "ThisPeriodTotalUserTime"
      .Add "TotalKernelTime"
      .Add "TotalPageFaultCount"
      .Add "TotalProcesses"
      .Add "TotalTerminatedProcesses"
      .Add "TotalUserTime"
      .Add "WriteOperationCount"
      .Add "WriteTransferCount"
   End With
   Set NamedJobObjectActgInfo = c
End Function
Private Function NamedJobObjectLimitSetting() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "ActiveProcessLimit"
      .Add "Affinity"
      .Add "Caption"
      .Add "Description"
      .Add "JobMemoryLimit"
      .Add "LimitFlags"
      .Add "MaximumWorkingSetSize"
      .Add "MinimumWorkingSetSize"
      .Add "PerJobUserTimeLimit"
      .Add "PerProcessUserTimeLimit"
      .Add "PriorityClass"
      .Add "ProcessMemoryLimit"
      .Add "SchedulingClass"
      .Add "SettingID"
   End With
   Set NamedJobObjectLimitSetting = c
End Function
Private Function NetworkAdapter() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "AdapterType"
      .Add "AdapterTypeID"
      .Add "AutoSense"
      .Add "Availability"
      .Add "Caption"
      .Add "ConfigManagerErrorCode"
      .Add "ConfigManagerUserConfig"
      .Add "CreationClassName"
      .Add "Description"
      .Add "DeviceID"
      .Add "ErrorCleared"
      .Add "ErrorDescription"
      .Add "Index"
      .Add "InstallDate"
      .Add "Installed"
      .Add "InterfaceIndex"
      .Add "LastErrorCode"
      .Add "MACAddress"
      .Add "Manufacturer"
      .Add "MaxNumberControlled"
      .Add "MaxSpeed"
      .Add "Name"
      .Add "NetConnectionID"
      .Add "NetConnectionStatus"
      .Add "NetworkAddresses[]"
      .Add "PermanentAddress"
      .Add "PNPDeviceID"
      .Add "PowerManagementCapabilities[]"
      .Add "PowerManagementSupported"
      .Add "ProductName"
      .Add "ServiceName"
      .Add "Speed"
      .Add "Status"
      .Add "StatusInfo"
      .Add "SystemCreationClassName"
      .Add "SystemName"
      .Add "TimeOfLastReset"
   End With
   Set NetworkAdapter = c
End Function
Private Function NetworkAdapterConfiguration() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "ArpAlwaysSourceRoute"
      .Add "ArpUseEtherSNAP"
      .Add "Caption"
      .Add "DatabasePath"
      .Add "DeadGWDetectEnabled"
      .Add "DefaultIPGateway[]"
      .Add "DefaultTOS"
      .Add "DefaultTTL"
      .Add "Description"
      .Add "DHCPEnabled"
      .Add "DHCPLeaseExpires"
      .Add "DHCPLeaseObtained"
      .Add "DHCPServer"
      .Add "DNSDomain"
      .Add "DNSDomainSuffixSearchOrder[]"
      .Add "DNSEnabledForWINSResolution"
      .Add "DNSHostName"
      .Add "DNSServerSearchOrder[]"
      .Add "DomainDNSRegistrationEnabled"
      .Add "ForwardBufferMemory"
      .Add "FullDNSRegistrationEnabled"
      .Add "GatewayCostMetric[]"
      .Add "IGMPLevel"
      .Add "Index"
      .Add "InterfaceIndex"
      .Add "IPAddress[]"
      .Add "IPConnectionMetric"
      .Add "IPEnabled"
      .Add "IPFilterSecurityEnabled"
      .Add "IPPortSecurityEnabled"
      .Add "IPSecPermitIPProtocols[]"
      .Add "IPSecPermitTCPPorts[]"
      .Add "IPSecPermitUDPPorts[]"
      .Add "IPSubnet[]"
      .Add "IPUseZeroBroadcast"
      .Add "IPXAddress"
      .Add "IPXEnabled"
      .Add "IPXFrameType[]"
      .Add "IPXMediaType"
      .Add "IPXNetworkNumber[]"
      .Add "IPXVirtualNetNumber"
      .Add "KeepAliveInterval"
      .Add "KeepAliveTime"
      .Add "MACAddress"
      .Add "MTU"
      .Add "NumForwardPackets"
      .Add "PMTUBHDetectEnabled"
      .Add "PMTUDiscoveryEnabled"
      .Add "ServiceName"
      .Add "SettingID"
      .Add "TcpipNetbiosOptions"
      .Add "TcpMaxConnectRetransmissions"
      .Add "TcpMaxDataRetransmissions"
      .Add "TcpNumConnections"
      .Add "TcpUseRFC1122UrgentPointer"
      .Add "TcpWindowSize"
      .Add "WINSEnableLMHostsLookup"
      .Add "WINSHostLookupFile"
      .Add "WINSPrimaryServer"
      .Add "WINSScopeID"
      .Add "WINSSecondaryServer"
   End With
   Set NetworkAdapterConfiguration = c
End Function
Private Function NetworkClient() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Caption"
      .Add "Description"
      .Add "InstallDate"
      .Add "Manufacturer"
      .Add "Name"
      .Add "Status"
   End With
   Set NetworkClient = c
End Function
Private Function NetworkConnection() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "AccessMask[]"
      .Add "Caption"
      .Add "Comment"
      .Add "ConnectionState"
      .Add "ConnectionType"
      .Add "Description"
      .Add "DisplayType"
      .Add "InstallDate"
      .Add "LocalName"
      .Add "Name"
      .Add "Persistent"
      .Add "ProviderName"
      .Add "RemoteName"
      .Add "RemotePath"
      .Add "ResourceType"
      .Add "Status"
      .Add "UserName"
   End With
   Set NetworkConnection = c
End Function
Private Function NetworkLoginProfile() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "AccountExpires"
      .Add "AuthorizationFlags"
      .Add "BadPasswordCount"
      .Add "Caption"
      .Add "CodePage"
      .Add "Comment"
      .Add "CountryCode"
      .Add "Description"
      .Add "Flags"
      .Add "FullName"
      .Add "HomeDirectory"
      .Add "HomeDirectoryDrive"
      .Add "LastLogoff"
      .Add "LastLogon"
      .Add "LogonHours"
      .Add "LogonServer"
      .Add "MaximumStorage"
      .Add "Name"
      .Add "NumberOfLogons"
      .Add "Parameters"
      .Add "PasswordAge"
      .Add "PasswordExpires"
      .Add "PrimaryGroupId"
      .Add "Privileges"
      .Add "Profile"
      .Add "ScriptPath"
      .Add "SettingID"
      .Add "UnitsPerWeek"
      .Add "UserComment"
      .Add "UserId"
      .Add "UserType"
      .Add "Workstations"
   End With
   Set NetworkLoginProfile = c
End Function
Private Function NetworkProtocol() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Caption"
      .Add "ConnectionlessService"
      .Add "Description"
      .Add "GuaranteesDelivery"
      .Add "GuaranteesSequencing"
      .Add "InstallDate"
      .Add "MaximumAddressSize"
      .Add "MaximumMessageSize"
      .Add "MessageOriented"
      .Add "MinimumAddressSize"
      .Add "Name"
      .Add "PseudoStreamOriented"
      .Add "Status"
      .Add "SupportsBroadcasting"
      .Add "SupportsConnectData"
      .Add "SupportsDisconnectData"
      .Add "SupportsEncryption"
      .Add "SupportsExpeditedData"
      .Add "SupportsFragmentation"
      .Add "SupportsGracefulClosing"
      .Add "SupportsGuaranteedBandwidth"
      .Add "SupportsMulticasting"
      .Add "SupportsQualityofService"
   End With
   Set NetworkProtocol = c
End Function
Private Function NTDomain() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Caption"
      .Add "ClientSiteName"
      .Add "CreationClassName"
      .Add "DcSiteName"
      .Add "Description"
      .Add "DNSForestName"
      .Add "DomainControllerAddress"
      .Add "DomainControllerAddressType"
      .Add "DomainControllerName"
      .Add "DomainGUID"
      .Add "DomainName"
      .Add "DSDirectoryServiceFlag"
      .Add "DSDnsControllerFlag"
      .Add "DSDnsDomainFlag"
      .Add "DSDnsForestFlag"
      .Add "DSGlobalCatalogFlag"
      .Add "DSKerberosDistributionCenterFlag"
      .Add "DSPrimaryDomainControllerFlag"
      .Add "DSTimeServiceFlag"
      .Add "DSWritableFlag"
      .Add "InstallDate"
      .Add "Name"
      .Add "NameFormat"
      .Add "PrimaryOwnerContact"
      .Add "PrimaryOwnerName"
      .Add "Roles"
      .Add "Status"
   End With
   Set NTDomain = c
End Function
Private Function NTEventLogFile() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "AccessMask[]"
      .Add "Archive"
      .Add "Caption"
      .Add "Compressed"
      .Add "CompressionMethod"
      .Add "CreationClassName"
      .Add "CreationDate"
      .Add "CSCreationClassName"
      .Add "CSName"
      .Add "Description"
      .Add "Drive"
      .Add "EightDotThreeFileName"
      .Add "Encrypted"
      .Add "EncryptionMethod"
      .Add "Extension"
      .Add "FileName"
      .Add "FileSize"
      .Add "FileType"
      .Add "FSCreationClassName"
      .Add "FSName"
      .Add "Hidden"
      .Add "InstallDate"
      .Add "InUseCount"
      .Add "LastAccessed"
      .Add "LastModified"
      .Add "LogfileName"
      .Add "Manufacturer"
      .Add "MaxFileSize"
      .Add "Name"
      .Add "NumberOfRecords"
      .Add "OverwriteOutDated"
      .Add "OverWritePolicy"
      .Add "Path"
      .Add "Readable"
      .Add "Sources[]"
      .Add "Status"
      .Add "System"
      .Add "Version"
      .Add "Writeable"
   End With
   Set NTEventLogFile = c
End Function
Private Function NTLogEvent() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Category"
      .Add "CategoryString"
      .Add "ComputerName"
      .Add "Data[]"
      .Add "EventCode"
      .Add "EventIdentifier"
      .Add "EventType"
      .Add "InsertionStrings[]"
      .Add "Logfile"
      .Add "Message"
      .Add "RecordNumber"
      .Add "SourceName"
      .Add "TimeGenerated"
      .Add "TimeWritten"
      .Add "Type"
      .Add "User"
   End With
   Set NTLogEvent = c
End Function
Private Function ODBCDriverSpecification() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Caption"
      .Add "CheckID"
      .Add "CheckMode"
      .Add "Description"
      .Add "Driver"
      .Add "File"
      .Add "Name"
      .Add "SetupFile"
      .Add "SoftwareElementID"
      .Add "SoftwareElementState"
      .Add "TargetOperatingSystem"
      .Add "Version"
   End With
   Set ODBCDriverSpecification = c
End Function
Private Function ODBCSourceAttribute() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Attribute"
      .Add "Caption"
      .Add "DataSource"
      .Add "Description"
      .Add "SettingID"
      .Add "Value"
   End With
   Set ODBCSourceAttribute = c
End Function
Private Function ODBCTranslatorSpecification() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Caption"
      .Add "CheckID"
      .Add "CheckMode"
      .Add "Description"
      .Add "File"
      .Add "Name"
      .Add "SetupFile"
      .Add "SoftwareElementID"
      .Add "SoftwareElementState"
      .Add "TargetOperatingSystem"
      .Add "Translator"
      .Add "Version"
   End With
   Set ODBCTranslatorSpecification = c
End Function
Private Function OnBoardDevice() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Caption"
      .Add "CreationClassName"
      .Add "Description"
      .Add "DeviceType"
      .Add "Enabled"
      .Add "HotSwappable"
      .Add "InstallDate"
      .Add "Manufacturer"
      .Add "Model"
      .Add "Name"
      .Add "OtherIdentifyingInfo"
      .Add "PartNumber"
      .Add "PoweredOn"
      .Add "Removable"
      .Add "Replaceable"
      .Add "SerialNumber"
      .Add "SKU"
      .Add "Status"
      .Add "Tag"
      .Add "Version"
   End With
   Set OnBoardDevice = c
End Function
Private Function OperatingSystem() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "BootDevice"
      .Add "BuildNumber"
      .Add "BuildType"
      .Add "Caption"
      .Add "CodeSet"
      .Add "CountryCode"
      .Add "CreationClassName"
      .Add "CSCreationClassName"
      .Add "CSDVersion"
      .Add "CSName"
      .Add "CurrentTimeZone"
      .Add "Debug"
      .Add "Description"
      .Add "Distributed"
      .Add "EncryptionLevel"
      .Add "ForegroundApplicationBoost"
      .Add "FreePhysicalMemory"
      .Add "FreeSpaceInPagingFiles"
      .Add "FreeVirtualMemory"
      .Add "InstallDate"
      .Add "LargeSystemCache"
      .Add "LastBootUpTime"
      .Add "LocalDateTime"
      .Add "Locale"
      .Add "Manufacturer"
      .Add "MaxNumberOfProcesses"
      .Add "MaxProcessMemorySize"
      .Add "Name"
      .Add "NumberOfLicensedUsers"
      .Add "NumberOfProcesses"
      .Add "NumberOfUsers"
      .Add "Organization"
      .Add "OSLanguage"
      .Add "OSProductSuite"
      .Add "OSType"
      .Add "OtherTypeDescription"
      .Add "PlusProductID"
      .Add "PlusVersionNumber"
      .Add "Primary"
      .Add "ProductType"
      .Add "QuantumLength"
      .Add "QuantumType"
      .Add "RegisteredUser"
      .Add "SerialNumber"
      .Add "ServicePackMajorVersion"
      .Add "ServicePackMinorVersion"
      .Add "SizeStoredInPagingFiles"
      .Add "Status"
      .Add "SuiteMask"
      .Add "SystemDevice"
      .Add "SystemDirectory"
      .Add "SystemDrive"
      .Add "TotalSwapSpaceSize"
      .Add "TotalVirtualMemorySize"
      .Add "TotalVisibleMemorySize"
      .Add "Version"
      .Add "WindowsDirectory"
   End With
   Set OperatingSystem = c
End Function
Private Function OSRecoveryConfiguration() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "AutoReboot"
      .Add "Caption"
      .Add "DebugFilePath"
      .Add "DebugInfoType"
      .Add "Description"
      .Add "ExpandedDebugFilePath"
      .Add "ExpandedMiniDumpDirectory"
      .Add "KernelDumpOnly"
      .Add "MiniDumpDirectory"
      .Add "Name"
      .Add "OverwriteExistingDebugFile"
      .Add "SendAdminAlert"
      .Add "SettingID"
      .Add "WriteDebugInfo"
      .Add "WriteToSystemLog"
   End With
   Set OSRecoveryConfiguration = c
End Function
Private Function PageFile() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "AccessMask[]"
      .Add "Archive"
      .Add "Caption"
      .Add "Compressed"
      .Add "CompressionMethod"
      .Add "CreationClassName"
      .Add "CreationDate"
      .Add "CSCreationClassName"
      .Add "CSName"
      .Add "Description"
      .Add "Drive"
      .Add "EightDotThreeFileName"
      .Add "Encrypted"
      .Add "EncryptionMethod"
      .Add "Extension"
      .Add "FileName"
      .Add "FileSize"
      .Add "FileType"
      .Add "FreeSpace"
      .Add "FSCreationClassName"
      .Add "FSName"
      .Add "Hidden"
      .Add "InitialSize"
      .Add "InstallDate"
      .Add "InUseCount"
      .Add "LastAccessed"
      .Add "LastModified"
      .Add "Manufacturer"
      .Add "MaximumSize"
      .Add "Name"
      .Add "Path"
      .Add "Readable"
      .Add "Status"
      .Add "System"
      .Add "Version"
      .Add "Writeable"
   End With
   Set PageFile = c
End Function
Private Function PageFileSetting() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Caption"
      .Add "Description"
      .Add "InitialSize"
      .Add "MaximumSize"
      .Add "Name"
      .Add "SettingID"
   End With
   Set PageFileSetting = c
End Function
Private Function PageFileUsage() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "AllocatedBaseSize"
      .Add "Caption"
      .Add "CurrentUsage"
      .Add "Description"
      .Add "InstallDate"
      .Add "Name"
      .Add "PeakUsage"
      .Add "Status"
      .Add "TempPageFile"
   End With
   Set PageFileUsage = c
End Function
Private Function ParallelPort() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Availability"
      .Add "Capabilities[]"
      .Add "CapabilityDescriptions[]"
      .Add "Caption"
      .Add "ConfigManagerErrorCode"
      .Add "ConfigManagerUserConfig"
      .Add "CreationClassName"
      .Add "Description"
      .Add "DeviceID"
      .Add "DMASupport"
      .Add "ErrorCleared"
      .Add "ErrorDescription"
      .Add "InstallDate"
      .Add "LastErrorCode"
      .Add "MaxNumberControlled"
      .Add "Name"
      .Add "OSAutoDiscovered"
      .Add "PNPDeviceID"
      .Add "PowerManagementCapabilities[]"
      .Add "PowerManagementSupported"
      .Add "ProtocolSupported"
      .Add "Status"
      .Add "StatusInfo"
      .Add "SystemCreationClassName"
      .Add "SystemName"
      .Add "TimeOfLastReset"
   End With
   Set ParallelPort = c
End Function
Private Function Patch() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Attributes"
      .Add "Caption"
      .Add "Description"
      .Add "File"
      .Add "PatchSize"
      .Add "ProductCode"
      .Add "Sequence"
      .Add "SettingID"
   End With
   Set Patch = c
End Function
Private Function PatchPackage() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Caption"
      .Add "Description"
      .Add "PatchID"
      .Add "ProductCode"
      .Add "SettingID"
   End With
   Set PatchPackage = c
End Function
Private Function PCMCIAController() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Availability"
      .Add "Caption"
      .Add "ConfigManagerErrorCode"
      .Add "ConfigManagerUserConfig"
      .Add "CreationClassName"
      .Add "Description"
      .Add "DeviceID"
      .Add "ErrorCleared"
      .Add "ErrorDescription"
      .Add "InstallDate"
      .Add "LastErrorCode"
      .Add "Manufacturer"
      .Add "MaxNumberControlled"
      .Add "Name"
      .Add "PNPDeviceID"
      .Add "PowerManagementCapabilities[]"
      .Add "PowerManagementSupported"
      .Add "ProtocolSupported"
      .Add "Status"
      .Add "StatusInfo"
      .Add "SystemCreationClassName"
      .Add "SystemName"
      .Add "TimeOfLastReset"
   End With
   Set PCMCIAController = c
End Function
Private Function Perf() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Caption"
      .Add "Description"
      .Add "Frequency_Object"
      .Add "Frequency_PerfTime"
      .Add "Frequency_Sys100NS"
      .Add "Name"
      .Add "Timestamp_Object"
      .Add "Timestamp_PerfTime"
      .Add "Timestamp_Sys100NS"
   End With
   Set Perf = c
End Function
Private Function PerfFormattedData() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Caption"
      .Add "Description"
      .Add "Frequency_Object"
      .Add "Frequency_PerfTime"
      .Add "Frequency_Sys100NS"
      .Add "Name"
      .Add "Timestamp_Object"
      .Add "Timestamp_PerfTime"
      .Add "Timestamp_Sys100NS"
   End With
   Set PerfFormattedData = c
End Function
Private Function PerfFormattedData_ASP_ActiveServerPages() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Caption"
      .Add "DebuggingRequests"
      .Add "Description"
      .Add "ErrorsDuringScriptRuntime"
      .Add "ErrorsFromASPPreprocessor"
      .Add "ErrorsFromScriptCompilers"
      .Add "ErrorsPerSec"
      .Add "Frequency_Object"
      .Add "Frequency_PerfTime"
      .Add "Frequency_Sys100NS"
      .Add "InMemoryTemplateCacheHitRate"
      .Add "InMemoryTemplatesCached"
      .Add "Name"
      .Add "RequestBytesInTotal"
      .Add "RequestBytesOutTotal"
      .Add "RequestExecutionTime"
      .Add "RequestsDisconnected"
      .Add "RequestsExecuting"
      .Add "RequestsFailedTotal"
      .Add "RequestsNotAuthorized"
      .Add "RequestsNotFound"
      .Add "RequestsPerSec"
      .Add "RequestsQueued"
      .Add "RequestsRejected"
      .Add "RequestsSucceeded"
      .Add "RequestsTimedOut"
      .Add "RequestsTotal"
      .Add "RequestWaitTime"
      .Add "ScriptEnginesCached"
      .Add "SessionDuration"
      .Add "SessionsCurrent"
      .Add "SessionsTimedOut"
      .Add "SessionsTotal"
      .Add "TemplateCacheHitRate"
      .Add "TemplateNotifications"
      .Add "TemplatesCached"
      .Add "Timestamp_Object"
      .Add "Timestamp_PerfTime"
      .Add "Timestamp_Sys100NS"
      .Add "TransactionsAborted"
      .Add "TransactionsCommitted"
      .Add "TransactionsPending"
      .Add "TransactionsPerSec"
      .Add "TransactionsTotal"
   End With
   Set PerfFormattedData_ASP_ActiveServerPages = c
End Function
Private Function PerfFormattedData_ContentFilter_IndexingServiceFilter() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "BindingTimeSec"
      .Add "Caption"
      .Add "Description"
      .Add "Frequency_Object"
      .Add "Frequency_PerfTime"
      .Add "Frequency_Sys100NS"
      .Add "IndexingSpeedMBPerHr"
      .Add "Name"
      .Add "Timestamp_Object"
      .Add "Timestamp_PerfTime"
      .Add "Timestamp_Sys100NS"
      .Add "TotalIndexingSpeedMBPerHr"
   End With
   Set PerfFormattedData_ContentFilter_IndexingServiceFilter = c
End Function
Private Function PerfFormattedData_ContentIndex_IndexingService() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Caption"
      .Add "DeferredForIndexing"
      .Add "Description"
      .Add "FilesToBeIndexed"
      .Add "Frequency_Object"
      .Add "Frequency_PerfTime"
      .Add "Frequency_Sys100NS"
      .Add "IndexSizeMB"
      .Add "MergeProgress"
      .Add "Name"
      .Add "NumberDocumentsIndexed"
      .Add "RunningQueries"
      .Add "SavedIndexes"
      .Add "Timestamp_Object"
      .Add "Timestamp_PerfTime"
      .Add "Timestamp_Sys100NS"
      .Add "TotalNumberDocuments"
      .Add "TotalNumberOfQueries"
      .Add "UniqueKeys"
      .Add "WordLists"
   End With
   Set PerfFormattedData_ContentIndex_IndexingService = c
End Function
Private Function PerfFormattedData_InetInfo_InternetInformationServicesGlobal() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "ActiveFlushedEntries"
      .Add "BLOBCacheFlushes"
      .Add "BLOBCacheHits"
      .Add "BLOBCacheHitsPercent"
      .Add "BLOBCacheMisses"
      .Add "Caption"
      .Add "CurrentBLOBsCached"
      .Add "CurrentBlockedAsyncIORequests"
      .Add "CurrentFileCacheMemoryUsage"
      .Add "CurrentFilesCached"
      .Add "CurrentURIsCached"
      .Add "Description"
      .Add "FileCacheFlushes"
      .Add "FileCacheHits"
      .Add "FileCacheHitsPercent"
      .Add "FileCacheMisses"
      .Add "Frequency_Object"
      .Add "Frequency_PerfTime"
      .Add "Frequency_Sys100NS"
      .Add "MaximumFileCacheMemoryUsage"
      .Add "MeasuredAsyncIOBandwidthUsage"
      .Add "Name"
      .Add "Timestamp_Object"
      .Add "Timestamp_PerfTime"
      .Add "Timestamp_Sys100NS"
      .Add "TotalAllowedAsyncIORequests"
      .Add "TotalBLOBsCached"
      .Add "TotalBlockedAsyncIORequests"
      .Add "TotalFilesCached"
      .Add "TotalFlushedBLOBs"
      .Add "TotalFlushedFiles"
      .Add "TotalFlushedURIs"
      .Add "TotalRejectedAsyncIORequests"
      .Add "TotalURIsCached"
      .Add "URICacheFlushes"
      .Add "URICacheHits"
      .Add "URICacheHitsPercent"
      .Add "URICacheMisses"
   End With
   Set PerfFormattedData_InetInfo_InternetInformationServicesGlobal = c
End Function
Private Function PerfFormattedData_ISAPISearch_HttpIndexingService() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "ActiveQueries"
      .Add "CacheItems"
      .Add "Caption"
      .Add "CurrentRequestsQueued"
      .Add "Description"
      .Add "Frequency_Object"
      .Add "Frequency_PerfTime"
      .Add "Frequency_Sys100NS"
      .Add "Name"
      .Add "PercentCacheHits"
      .Add "PercentCacheMisses"
      .Add "QueriesPerMinute"
      .Add "Timestamp_Object"
      .Add "Timestamp_PerfTime"
      .Add "Timestamp_Sys100NS"
      .Add "TotalQueries"
      .Add "TotalRequestsRejected"
   End With
   Set PerfFormattedData_ISAPISearch_HttpIndexingService = c
End Function
Private Function PerfFormattedData_MSDTC_DistributedTransactionCoordinator() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "AbortedTransactions"
      .Add "AbortedTransactionsPerSec"
      .Add "ActiveTransactions"
      .Add "ActiveTransactionsMaximum"
      .Add "Caption"
      .Add "CommittedTransactions"
      .Add "CommittedTransactionsPerSec"
      .Add "Description"
      .Add "ForceAbortedTransactions"
      .Add "ForceCommittedTransactions"
      .Add "Frequency_Object"
      .Add "Frequency_PerfTime"
      .Add "Frequency_Sys100NS"
      .Add "InDoubtTransactions"
      .Add "Name"
      .Add "ResponseTimeAverage"
      .Add "ResponseTimeMaximum"
      .Add "ResponseTimeMinimum"
      .Add "Timestamp_Object"
      .Add "Timestamp_PerfTime"
      .Add "Timestamp_Sys100NS"
      .Add "TransactionsPerSec"
   End With
   Set PerfFormattedData_MSDTC_DistributedTransactionCoordinator = c
End Function
Private Function PerfFormattedData_NTFSDRV_SMTPNTFSStoreDriver() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Caption"
      .Add "Description"
      .Add "Frequency_Object"
      .Add "Frequency_PerfTime"
      .Add "Frequency_Sys100NS"
      .Add "MessagesAllocated"
      .Add "MessagesDeleted"
      .Add "MessagesEnumerated"
      .Add "MessagesInTheQueueDirectory"
      .Add "Name"
      .Add "OpenMessageBodies"
      .Add "OpenMessageStreams"
      .Add "Timestamp_Object"
      .Add "Timestamp_PerfTime"
      .Add "Timestamp_Sys100NS"
   End With
   Set PerfFormattedData_NTFSDRV_SMTPNTFSStoreDriver = c
End Function
Private Function PerfFormattedData_PerfDisk_LogicalDisk() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "AvgDiskBytesPerRead"
      .Add "AvgDiskBytesPerTransfer"
      .Add "AvgDiskBytesPerWrite"
      .Add "AvgDiskQueueLength"
      .Add "AvgDiskReadQueueLength"
      .Add "AvgDiskSecPerRead"
      .Add "AvgDiskSecPerTransfer"
      .Add "AvgDiskSecPerWrite"
      .Add "AvgDiskWriteQueueLength"
      .Add "Caption"
      .Add "CurrentDiskQueueLength"
      .Add "Description"
      .Add "DiskBytesPerSec"
      .Add "DiskReadBytesPerSec"
      .Add "DiskReadsPerSec"
      .Add "DiskTransfersPerSec"
      .Add "DiskWriteBytesPerSec"
      .Add "DiskWritesPerSec"
      .Add "FreeMegabytes"
      .Add "Frequency_Object"
      .Add "Frequency_PerfTime"
      .Add "Frequency_Sys100NS"
      .Add "Name"
      .Add "PercentDiskReadTime"
      .Add "PercentDiskTime"
      .Add "PercentDiskWriteTime"
      .Add "PercentFreeSpace"
      .Add "PercentIdleTime"
      .Add "SplitIOPerSec"
      .Add "Timestamp_Object"
      .Add "Timestamp_PerfTime"
      .Add "Timestamp_Sys100NS"
   End With
   Set PerfFormattedData_PerfDisk_LogicalDisk = c
End Function
Private Function PerfFormattedData_PerfDisk_PhysicalDisk() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "AvgDiskBytesPerRead"
      .Add "AvgDiskBytesPerTransfer"
      .Add "AvgDiskBytesPerWrite"
      .Add "AvgDiskQueueLength"
      .Add "AvgDiskReadQueueLength"
      .Add "AvgDiskSecPerRead"
      .Add "AvgDiskSecPerTransfer"
      .Add "AvgDiskSecPerWrite"
      .Add "AvgDiskWriteQueueLength"
      .Add "Caption"
      .Add "CurrentDiskQueueLength"
      .Add "Description"
      .Add "DiskBytesPerSec"
      .Add "DiskReadBytesPerSec"
      .Add "DiskReadsPerSec"
      .Add "DiskTransfersPerSec"
      .Add "DiskWriteBytesPerSec"
      .Add "DiskWritesPerSec"
      .Add "Frequency_Object"
      .Add "Frequency_PerfTime"
      .Add "Frequency_Sys100NS"
      .Add "Name"
      .Add "PercentDiskReadTime"
      .Add "PercentDiskTime"
      .Add "PercentDiskWriteTime"
      .Add "PercentIdleTime"
      .Add "SplitIOPerSec"
      .Add "Timestamp_Object"
      .Add "Timestamp_PerfTime"
      .Add "Timestamp_Sys100NS"
   End With
   Set PerfFormattedData_PerfDisk_PhysicalDisk = c
End Function
Private Function PerfFormattedData_PerfNet_Browser() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "AnnouncementsDomainPerSec"
      .Add "AnnouncementsServerPerSec"
      .Add "AnnouncementsTotalPerSec"
      .Add "Caption"
      .Add "Description"
      .Add "DuplicateMasterAnnouncements"
      .Add "ElectionPacketsPerSec"
      .Add "EnumerationsDomainPerSec"
      .Add "EnumerationsOtherPerSec"
      .Add "EnumerationsServerPerSec"
      .Add "EnumerationsTotalPerSec"
      .Add "Frequency_Object"
      .Add "Frequency_PerfTime"
      .Add "Frequency_Sys100NS"
      .Add "IllegalDatagramsPerSec"
      .Add "MailslotAllocationsFailed"
      .Add "MailslotOpensFailedPerSec"
      .Add "MailslotReceivesFailed"
      .Add "MailslotWritesFailed"
      .Add "MailslotWritesPerSec"
      .Add "MissedMailslotDatagrams"
      .Add "MissedServerAnnouncements"
      .Add "MissedServerListRequests"
      .Add "Name"
      .Add "ServerAnnounceAllocationsFailedPerSec"
      .Add "ServerListRequestsPerSec"
      .Add "Timestamp_Object"
      .Add "Timestamp_PerfTime"
      .Add "Timestamp_Sys100NS"
   End With
   Set PerfFormattedData_PerfNet_Browser = c
End Function












Private Function PerfFormattedData_PerfNet_Redirector() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "ServerDisconnects"
      .Add "BytesReceivedPerSec"
      .Add "BytesTotalPerSec"
      .Add "BytesTransmittedPerSec"
      .Add "Caption"
      .Add "ConnectsCore"
      .Add "ConnectsLanManager20"
      .Add "ConnectsLanManager21"
      .Add "ConnectsWindowsNT"
      .Add "CurrentCommands"
      .Add "Description"
      .Add "FileDataOperationsPerSec"
      .Add "FileReadOperationsPerSec"
      .Add "FileWriteOperationsPerSec"
      .Add "Frequency_Object"
      .Add "Frequency_PerfTime"
      .Add "Frequency_Sys100NS"
      .Add "Name"
      .Add "NetworkErrorsPerSec"
      .Add "PacketsPerSec"
      .Add "PacketsReceivedPerSec"
      .Add "PacketsTransmittedPerSec"
      .Add "ReadBytesCachePerSec"
      .Add "ReadBytesNetworkPerSec"
      .Add "ReadBytesNonPagingPerSec"
      .Add "ReadBytesPagingPerSec"
      .Add "ReadOperationsRandomPerSec"
      .Add "ReadPacketsPerSec"
      .Add "ReadPacketsSmallPerSec"
      .Add "ReadsDeniedPerSec"
      .Add "ReadsLargePerSec"
      .Add "ServerReconnects"
      .Add "ServerSessions"
      .Add "ServerSessionsHung"
      .Add "Timestamp_Object"
      .Add "Timestamp_PerfTime"
      .Add "Timestamp_Sys100NS"
      .Add "WriteBytesCachePerSec"
      .Add "WriteBytesNetworkPerSec"
      .Add "WriteBytesNonPagingPerSec"
      .Add "WriteBytesPagingPerSec"
      .Add "WriteOperationsRandomPerSec"
      .Add "WritePacketsPerSec"
      .Add "WritePacketsSmallPerSec"
      .Add "WritesDeniedPerSec"
      .Add "WritesLargePerSec"
   End With
   Set PerfFormattedData_PerfNet_Redirector = c
End Function
Private Function PerfFormattedData_PerfNet_Server() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "BlockingRequestsRejected"
      .Add "BytesReceivedPerSec"
      .Add "BytesTotalPerSec"
      .Add "BytesTransmittedPerSec"
      .Add "Caption"
      .Add "ContextBlocksQueuedPerSec"
      .Add "Description"
      .Add "ErrorsAccessPermissions"
      .Add "ErrorsGrantedAccess"
      .Add "ErrorsLogon"
      .Add "ErrorsSystem"
      .Add "FileDirectorySearches"
      .Add "FilesOpen"
      .Add "FilesOpenedTotal"
      .Add "Frequency_Object"
      .Add "Frequency_PerfTime"
      .Add "Frequency_Sys100NS"
      .Add "LogonPerSec"
      .Add "LogonTotal"
      .Add "Name"
      .Add "PoolNonPagedBytes"
      .Add "PoolNonPagedFailures"
      .Add "PoolNonPagedPeak"
      .Add "PoolPagedBytes"
      .Add "PoolPagedFailures"
      .Add "PoolPagedPeak"
      .Add "ServerSessions"
      .Add "SessionsErroredOut"
      .Add "SessionsForcedOff"
      .Add "SessionsLoggedOff"
      .Add "SessionsTimedOut"
      .Add "Timestamp_Object"
      .Add "Timestamp_PerfTime"
      .Add "Timestamp_Sys100NS"
      .Add "WorkItemShortages"
   End With
   Set PerfFormattedData_PerfNet_Server = c
End Function
Private Function PerfFormattedData_PerfNet_ServerWorkQueues() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "ActiveThreads"
      .Add "AvailableThreads"
      .Add "AvailableWorkItems"
      .Add "BorrowedWorkItems"
      .Add "BytesReceivedPerSec"
      .Add "BytesSentPerSec"
      .Add "BytesTransferredPerSec"
      .Add "Caption"
      .Add "ContextBlocksQueuedPerSec"
      .Add "CurrentClients"
      .Add "Description"
      .Add "Frequency_Object"
      .Add "Frequency_PerfTime"
      .Add "Frequency_Sys100NS"
      .Add "Name"
      .Add "QueueLength"
      .Add "ReadBytesPerSec"
      .Add "ReadOperationsPerSec"
      .Add "Timestamp_Object"
      .Add "Timestamp_PerfTime"
      .Add "Timestamp_Sys100NS"
      .Add "TotalBytesPerSec"
      .Add "TotalOperationsPerSec"
      .Add "WorkItemShortages"
      .Add "WriteBytesPerSec"
      .Add "WriteOperationsPerSec"
   End With
   Set PerfFormattedData_PerfNet_ServerWorkQueues = c
End Function
Private Function PerfFormattedData_PerfOS_Cache() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "AsyncCopyReadsPerSec"
      .Add "AsyncDataMapsPerSec"
      .Add "AsyncFastReadsPerSec"
      .Add "AsyncMDLReadsPerSec"
      .Add "AsyncPinReadsPerSec"
      .Add "Caption"
      .Add "CopyReadHitsPercent"
      .Add "CopyReadsPerSec"
      .Add "DataFlushesPerSec"
      .Add "DataFlushPagesPerSec"
      .Add "DataMapHitsPercent"
      .Add "DataMapPinsPerSec"
      .Add "DataMapsPerSec"
      .Add "Description"
      .Add "FastReadNotPossiblesPerSec"
      .Add "FastReadResourceMissesPerSec"
      .Add "FastReadsPerSec"
      .Add "Frequency_Object"
      .Add "Frequency_PerfTime"
      .Add "Frequency_Sys100NS"
      .Add "LazyWriteFlushesPerSec"
      .Add "LazyWritePagesPerSec"
      .Add "MDLReadHitsPercent"
      .Add "MDLReadsPerSec"
      .Add "Name"
      .Add "PinReadHitsPercent"
      .Add "PinReadsPerSec"
      .Add "ReadAheadsPerSec"
      .Add "SyncCopyReadsPerSec"
      .Add "SyncDataMapsPerSec"
      .Add "SyncFastReadsPerSec"
      .Add "SyncMDLReadsPerSec"
      .Add "SyncPinReadsPerSec"
      .Add "Timestamp_Object"
      .Add "Timestamp_PerfTime"
      .Add "Timestamp_Sys100NS"
   End With
   Set PerfFormattedData_PerfOS_Cache = c
End Function
Private Function PerfFormattedData_PerfOS_Memory() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "AvailableBytes"
      .Add "AvailableKBytes"
      .Add "AvailableMBytes"
      .Add "CacheBytes"
      .Add "CacheBytesPeak"
      .Add "CacheFaultsPerSec"
      .Add "Caption"
      .Add "CommitLimit"
      .Add "CommittedBytes"
      .Add "DemandZeroFaultsPerSec"
      .Add "Description"
      .Add "FreeSystemPageTableEntries"
      .Add "Frequency_Object"
      .Add "Frequency_PerfTime"
      .Add "Frequency_Sys100NS"
      .Add "Name"
      .Add "PageFaultsPerSec"
      .Add "PageReadsPerSec"
      .Add "PagesInputPerSec"
      .Add "PagesOutputPerSec"
      .Add "PagesPerSec"
      .Add "PageWritesPerSec"
      .Add "PercentCommittedBytesInUse"
      .Add "PoolNonpagedAllocs"
      .Add "PoolNonpagedBytes"
      .Add "PoolPagedAllocs"
      .Add "PoolPagedBytes"
      .Add "PoolPagedResidentBytes"
      .Add "SystemCacheResidentBytes"
      .Add "SystemCodeResidentBytes"
      .Add "SystemCodeTotalBytes"
      .Add "SystemDriverResidentBytes"
      .Add "SystemDriverTotalBytes"
      .Add "Timestamp_Object"
      .Add "Timestamp_PerfTime"
      .Add "Timestamp_Sys100NS"
      .Add "TransitionFaultsPerSec"
      .Add "WriteCopiesPerSec"

   End With
   Set PerfFormattedData_PerfOS_Memory = c
End Function
Private Function PerfFormattedData_PerfOS_Objects() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Caption"
      .Add "Description"
      .Add "Events"
      .Add "Frequency_Object"
      .Add "Frequency_PerfTime"
      .Add "Frequency_Sys100NS"
      .Add "Mutexes"
      .Add "Name"
      .Add "Processes"
      .Add "Sections"
      .Add "Semaphores"
      .Add "Threads"
      .Add "Timestamp_Object"
      .Add "Timestamp_PerfTime"
      .Add "Timestamp_Sys100NS"

   End With
   Set PerfFormattedData_PerfOS_Objects = c
End Function
Private Function PerfFormattedData_PerfOS_PagingFile() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Caption"
      .Add "Description"
      .Add "Frequency_Object"
      .Add "Frequency_PerfTime"
      .Add "Frequency_Sys100NS"
      .Add "Name"
      .Add "PercentUsage"
      .Add "PercentUsagePeak"
      .Add "Timestamp_Object"
      .Add "Timestamp_PerfTime"
      .Add "Timestamp_Sys100NS"

   End With
   Set PerfFormattedData_PerfOS_PagingFile = c
End Function
Private Function PerfFormattedData_PerfOS_Processor() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "C1TransitionsPerSec"
      .Add "C2TransitionsPerSec"
      .Add "C3TransitionsPerSec"
      .Add "Caption"
      .Add "Description"
      .Add "DPCRate"
      .Add "DPCsQueuedPerSec"
      .Add "Frequency_Object"
      .Add "Frequency_PerfTime"
      .Add "Frequency_Sys100NS"
      .Add "InterruptsPerSec"
      .Add "Name"
      .Add "PercentC1Time"
      .Add "PercentC2Time"
      .Add "PercentC3Time"
      .Add "PercentDPCTime"
      .Add "PercentIdleTime"
      .Add "PercentInterruptTime"
      .Add "PercentPrivilegedTime"
      .Add "PercentProcessorTime"
      .Add "PercentUserTime"
      .Add "Timestamp_Object"
      .Add "Timestamp_PerfTime"
      .Add "Timestamp_Sys100NS"

   End With
   Set PerfFormattedData_PerfOS_Processor = c
End Function
Private Function PerfFormattedData_PerfOS_System() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "AlignmentFixupsPerSec"
      .Add "Caption"
      .Add "ContextSwitchesPerSec"
      .Add "Description"
      .Add "ExceptionDispatchesPerSec"
      .Add "FileControlBytesPerSec"
      .Add "FileControlOperationsPerSec"
      .Add "FileDataOperationsPerSec"
      .Add "FileReadBytesPerSec"
      .Add "FileReadOperationsPerSec"
      .Add "FileWriteBytesPerSec"
      .Add "FileWriteOperationsPerSec"
      .Add "FloatingEmulationsPerSec"
      .Add "Frequency_Object"
      .Add "Frequency_PerfTime"
      .Add "Frequency_Sys100NS"
      .Add "Name"
      .Add "PercentRegistryQuotaInUse"
      .Add "Processes"
      .Add "ProcessorQueueLength"
      .Add "SystemCallsPerSec"
      .Add "SystemUpTime"
      .Add "Threads"
      .Add "Timestamp_Object"
      .Add "Timestamp_PerfTime"
      .Add "Timestamp_Sys100NS"

   End With
   Set PerfFormattedData_PerfOS_System = c
End Function
Private Function PerfFormattedData_PerfProc_FullImage_Costly() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Caption"
      .Add "Description"
      .Add "ExecReadOnly"
      .Add "ExecReadPerWrite"
      .Add "Executable"
      .Add "ExecWriteCopy"
      .Add "Frequency_Object"
      .Add "Frequency_PerfTime"
      .Add "Frequency_Sys100NS"
      .Add "Name"
      .Add "NoAccess"
      .Add "ReadOnly"
      .Add "ReadPerWrite"
      .Add "Timestamp_Object"
      .Add "Timestamp_PerfTime"
      .Add "Timestamp_Sys100NS"
      .Add "WriteCopy"

   End With
   Set PerfFormattedData_PerfProc_FullImage_Costly = c
End Function
Private Function PerfFormattedData_PerfProc_Image_Costly() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Caption"
      .Add "Description"
      .Add "ExecReadOnly"
      .Add "ExecReadPerWrite"
      .Add "Executable"
      .Add "ExecWriteCopy"
      .Add "Frequency_Object"
      .Add "Frequency_PerfTime"
      .Add "Frequency_Sys100NS"
      .Add "Name"
      .Add "NoAccess"
      .Add "ReadOnly"
      .Add "ReadPerWrite"
      .Add "Timestamp_Object"
      .Add "Timestamp_PerfTime"
      .Add "Timestamp_Sys100NS"
      .Add "WriteCopy"

   End With
   Set PerfFormattedData_PerfProc_Image_Costly = c
End Function
Private Function PerfFormattedData_PerfProc_JobObject() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Caption"
      .Add "CurrentPercentKernelModeTime"
      .Add "CurrentPercentProcessorTime"
      .Add "CurrentPercentUserModeTime"
      .Add "Description"
      .Add "Frequency_Object"
      .Add "Frequency_PerfTime"
      .Add "Frequency_Sys100NS"
      .Add "Name"
      .Add "PagesPerSec"
      .Add "ProcessCountActive"
      .Add "ProcessCountTerminated"
      .Add "ProcessCountTotal"
      .Add "ThisPeriodmSecKernelMode"
      .Add "ThisPeriodmSecProcessor"
      .Add "ThisPeriodmSecUserMode"
      .Add "Timestamp_Object"
      .Add "Timestamp_PerfTime"
      .Add "Timestamp_Sys100NS"
      .Add "TotalmSecKernelMode"
      .Add "TotalmSecProcessor"
      .Add "TotalmSecUserMode"

   End With
   Set PerfFormattedData_PerfProc_JobObject = c
End Function
Private Function PerfFormattedData_PerfProc_JobObjectDetails() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Caption"
      .Add "CreatingProcessID"
      .Add "Description"
      .Add "ElapsedTime"
      .Add "Frequency_Object"
      .Add "Frequency_PerfTime"
      .Add "Frequency_Sys100NS"
      .Add "HandleCount"
      .Add "IDProcess"
      .Add "IODataBytesPerSec"
      .Add "IODataOperationsPerSec"
      .Add "IOOtherBytesPerSec"
      .Add "IOOtherOperationsPerSec"
      .Add "IOReadBytesPerSec"
      .Add "IOReadOperationsPerSec"
      .Add "IOWriteBytesPerSec"
      .Add "IOWriteOperationsPerSec"
      .Add "Name"
      .Add "PageFaultsPerSec"
      .Add "PageFileBytes"
      .Add "PageFileBytesPeak"
      .Add "PercentPrivilegedTime"
      .Add "PercentProcessorTime"
      .Add "PercentUserTime"
      .Add "PoolNonpagedBytes"
      .Add "PoolPagedBytes"
      .Add "PriorityBase"
      .Add "PrivateBytes"
      .Add "ThreadCount"
      .Add "Timestamp_Object"
      .Add "Timestamp_PerfTime"
      .Add "Timestamp_Sys100NS"
      .Add "VirtualBytes"
      .Add "VirtualBytesPeak"
      .Add "WorkingSet"
      .Add "WorkingSetPeak"

   End With
   Set PerfFormattedData_PerfProc_JobObjectDetails = c
End Function
Private Function PerfFormattedData_PerfProc_Process() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Caption"
      .Add "CreatingProcessID"
      .Add "Description"
      .Add "ElapsedTime"
      .Add "Frequency_Object"
      .Add "Frequency_PerfTime"
      .Add "Frequency_Sys100NS"
      .Add "HandleCount"
      .Add "IDProcess"
      .Add "IODataBytesPerSec"
      .Add "IODataOperationsPerSec"
      .Add "IOOtherBytesPerSec"
      .Add "IOOtherOperationsPerSec"
      .Add "IOReadBytesPerSec"
      .Add "IOReadOperationsPerSec"
      .Add "IOWriteBytesPerSec"
      .Add "IOWriteOperationsPerSec"
      .Add "Name"
      .Add "PageFaultsPerSec"
      .Add "PageFileBytes"
      .Add "PageFileBytesPeak"
      .Add "PercentPrivilegedTime"
      .Add "PercentProcessorTime"
      .Add "PercentUserTime"
      .Add "PoolNonpagedBytes"
      .Add "PoolPagedBytes"
      .Add "PriorityBase"
      .Add "PrivateBytes"
      .Add "ThreadCount"
      .Add "Timestamp_Object"
      .Add "Timestamp_PerfTime"
      .Add "Timestamp_Sys100NS"
      .Add "VirtualBytes"
      .Add "VirtualBytesPeak"
      .Add "WorkingSet"
      .Add "WorkingSetPeak"

   End With
   Set PerfFormattedData_PerfProc_Process = c
End Function
Private Function PerfFormattedData_PerfProc_ProcessAddressSpace_Costly() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "BytesFree"
      .Add "BytesImageFree"
      .Add "BytesImageReserved"
      .Add "BytesReserved"
      .Add "Caption"
      .Add "Description"
      .Add "Frequency_Object"
      .Add "Frequency_PerfTime"
      .Add "Frequency_Sys100NS"
      .Add "IDProcess"
      .Add "ImageSpaceExecReadOnly"
      .Add "ImageSpaceExecReadPerWrite"
      .Add "ImageSpaceExecutable"
      .Add "ImageSpaceExecWriteCopy"
      .Add "ImageSpaceNoAccess"
      .Add "ImageSpaceReadOnly"
      .Add "ImageSpaceReadPerWrite"
      .Add "ImageSpaceWriteCopy"
      .Add "MappedSpaceExecReadOnly"
      .Add "MappedSpaceExecReadPerWrite"
      .Add "MappedSpaceExecutable"
      .Add "MappedSpaceExecWriteCopy"
      .Add "MappedSpaceNoAccess"
      .Add "MappedSpaceReadOnly"
      .Add "MappedSpaceReadPerWrite"
      .Add "MappedSpaceWriteCopy"
      .Add "Name"
      .Add "ReservedSpaceExecReadOnly"
      .Add "ReservedSpaceExecReadPerWrite"
      .Add "ReservedSpaceExecutable"
      .Add "ReservedSpaceExecWriteCopy"
      .Add "ReservedSpaceNoAccess"
      .Add "ReservedSpaceReadOnly"
      .Add "ReservedSpaceReadPerWrite"
      .Add "ReservedSpaceWriteCopy"
      .Add "Timestamp_Object"
      .Add "Timestamp_PerfTime"
      .Add "Timestamp_Sys100NS"
      .Add "UnassignedSpaceExecReadOnly"
      .Add "UnassignedSpaceExecReadPerWrite"
      .Add "UnassignedSpaceExecutable"
      .Add "UnassignedSpaceExecWriteCopy"
      .Add "UnassignedSpaceNoAccess"
      .Add "UnassignedSpaceReadOnly"
      .Add "UnassignedSpaceReadPerWrite"
      .Add "UnassignedSpaceWriteCopy"

   End With
   Set PerfFormattedData_PerfProc_ProcessAddressSpace_Costly = c
End Function
Private Function Bus() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Availability"
      .Add "BusNum"
      .Add "BusType"
      .Add "Caption"
      .Add "ConfigManagerErrorCode"
      .Add "ConfigManagerUserConfig"
      .Add "CreationClassName"
      .Add "Description"
      .Add "DeviceID"
      .Add "ErrorCleared"
      .Add "ErrorDescription"
      .Add "InstallDate"
      .Add "LastErrorCode"
      .Add "Name"
      .Add "PNPDeviceID"
      .Add "PowerManagementCapabilities[]"
      .Add "PowerManagementSupported"
      .Add "Status"
      .Add "StatusInfo"
      .Add "SystemCreationClassName"
      .Add "SystemName"
   End With
   Set Bus = c
End Function
Private Function PerfFormattedData_PerfProc_Thread() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Caption"
      .Add "ContextSwitchesPerSec"
      .Add "Description"
      .Add "ElapsedTime"
      .Add "Frequency_Object"
      .Add "Frequency_PerfTime"
      .Add "Frequency_Sys100NS"
      .Add "IDProcess"
      .Add "IDThread"
      .Add "Name"
      .Add "PercentPrivilegedTime"
      .Add "PercentProcessorTime"
      .Add "PercentUserTime"
      .Add "PriorityBase"
      .Add "PriorityCurrent"
      .Add "StartAddress"
      .Add "ThreadState"
      .Add "ThreadWaitReason"
      .Add "Timestamp_Object"
      .Add "Timestamp_PerfTime"
      .Add "Timestamp_Sys100NS"

   End With
   Set PerfFormattedData_PerfProc_Thread = c
End Function
Private Function PerfFormattedData_PerfProc_ThreadDetails_Costly() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Caption"
      .Add "Description"
      .Add "Frequency_Object"
      .Add "Frequency_PerfTime"
      .Add "Frequency_Sys100NS"
      .Add "Name"
      .Add "Timestamp_Object"
      .Add "Timestamp_PerfTime"
      .Add "Timestamp_Sys100NS"
      .Add "UserPC"

   End With
   Set PerfFormattedData_PerfProc_ThreadDetails_Costly = c
End Function
Private Function PerfFormattedData_PSched_PSchedFlow() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "AveragePacketsInNetCard"
      .Add "AveragePacketsInSequencer"
      .Add "AveragePacketsInShaper"
      .Add "BytesScheduled"
      .Add "BytesScheduledPerSec"
      .Add "BytesTransmitted"
      .Add "BytesTransmittedPerSec"
      .Add "Caption"
      .Add "Description"
      .Add "Frequency_Object"
      .Add "Frequency_PerfTime"
      .Add "Frequency_Sys100NS"
      .Add "MaximumPacketsInNetCard"
      .Add "MaxPacketsInSequencer"
      .Add "MaxPacketsInShaper"
      .Add "Name"
      .Add "NonConformingPacketsScheduled"
      .Add "NonConformingPacketsScheduledPerSec"
      .Add "NonConformingPacketsTransmitted"
      .Add "NonConformingPacketsTransmittedPerSec"
      .Add "PacketsDropped"
      .Add "PacketsDroppedPerSec"
      .Add "PacketsScheduled"
      .Add "PacketsScheduledPerSec"
      .Add "PacketsTransmitted"
      .Add "PacketsTransmittedPerSec"
      .Add "Timestamp_Object"
      .Add "Timestamp_PerfTime"
      .Add "Timestamp_Sys100NS"

   End With
   Set PerfFormattedData_PSched_PSchedFlow = c
End Function
Private Function PerfFormattedData_PSched_PSchedPipe() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "AveragePacketsInNetCard"
      .Add "AveragePacketsInSequencer"
      .Add "AveragePacketsInShaper"
      .Add "Caption"
      .Add "Description"
      .Add "FlowModsRejected"
      .Add "FlowsClosed"
      .Add "FlowsModified"
      .Add "FlowsOpened"
      .Add "FlowsRejected"
      .Add "Frequency_Object"
      .Add "Frequency_PerfTime"
      .Add "Frequency_Sys100NS"
      .Add "MaxPacketsInNetCard"
      .Add "MaxPacketsInSequencer"
      .Add "MaxPacketsInShaper"
      .Add "MaxSimultaneousFlows"
      .Add "Name"
      .Add "NonConformingPacketsScheduled"
      .Add "NonConformingPacketsScheduledPerSec"
      .Add "NonConformingPacketsTransmitted"
      .Add "NonConformingPacketsTransmittedPerSec"
      .Add "OutOfPackets"
      .Add "Timestamp_Object"
      .Add "Timestamp_PerfTime"
      .Add "Timestamp_Sys100NS"

   End With
   Set PerfFormattedData_PSched_PSchedPipe = c
End Function
Private Function PerfFormattedData_RemoteAccess_RASPort() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "AlignmentErrors"
      .Add "BufferOverrunErrors"
      .Add "BytesReceived"
      .Add "BytesReceivedPerSec"
      .Add "BytesTransmitted"
      .Add "BytesTransmittedPerSec"
      .Add "Caption"
      .Add "CRCErrors"
      .Add "Description"
      .Add "FramesReceived"
      .Add "FramesReceivedPerSec"
      .Add "FramesTransmitted"
      .Add "FramesTransmittedPerSec"
      .Add "Frequency_Object"
      .Add "Frequency_PerfTime"
      .Add "Frequency_Sys100NS"
      .Add "Name"
      .Add "PercentCompressionIn"
      .Add "PercentCompressionOut"
      .Add "SerialOverrunErrors"
      .Add "TimeoutErrors"
      .Add "Timestamp_Object"
      .Add "Timestamp_PerfTime"
      .Add "Timestamp_Sys100NS"
      .Add "TotalErrors"
      .Add "TotalErrorsPerSec"

   End With
   Set PerfFormattedData_RemoteAccess_RASPort = c
End Function
Private Function PerfFormattedData_RemoteAccess_RASTotal() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "AlignmentErrors"
      .Add "BufferOverrunErrors"
      .Add "BytesReceived"
      .Add "BytesReceivedPerSec"
      .Add "BytesTransmitted"
      .Add "BytesTransmittedPerSec"
      .Add "Caption"
      .Add "CRCErrors"
      .Add "Description"
      .Add "FramesReceived"
      .Add "FramesReceivedPerSec"
      .Add "FramesTransmitted"
      .Add "FramesTransmittedPerSec"
      .Add "Frequency_Object"
      .Add "Frequency_PerfTime"
      .Add "Frequency_Sys100NS"
      .Add "Name"
      .Add "PercentCompressionIn"
      .Add "PercentCompressionOut"
      .Add "SerialOverrunErrors"
      .Add "TimeoutErrors"
      .Add "Timestamp_Object"
      .Add "Timestamp_PerfTime"
      .Add "Timestamp_Sys100NS"
      .Add "TotalConnections"
      .Add "TotalErrors"
      .Add "TotalErrorsPerSec"

   End With
   Set PerfFormattedData_RemoteAccess_RASTotal = c
End Function
Private Function PerfFormattedData_RSVP_ACSRSVPInterfaces() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "AdmittedBandwidth"
      .Add "BlockedRESVs"
      .Add "Caption"
      .Add "Description"
      .Add "Frequency_Object"
      .Add "Frequency_PerfTime"
      .Add "Frequency_Sys100NS"
      .Add "GeneralFailures"
      .Add "MaximumAdmittedBandwidth"
      .Add "Name"
      .Add "NumberOfActiveFlows"
      .Add "NumberOfIncomingMessagesDropped"
      .Add "NumberOfOutGoingMessagesDropped"
      .Add "PATHERRMessagesReceived"
      .Add "PATHERRMessagesSent"
      .Add "PATHMessagesReceived"
      .Add "PATHMessagesSent"
      .Add "PATHStateBlockTimeouts"
      .Add "PATHTEARMessagesReceived"
      .Add "PATHTEARMessagesSent"
      .Add "PolicyControlFailures"
      .Add "ReceiveMessagesErrorsBigMessages"
      .Add "ReceiveMessagesErrorsNoMemory"
      .Add "ResourceControlFailures"
      .Add "RESVCONFIRMMessagesReceived"
      .Add "RESVCONFIRMMessagesSent"
      .Add "RESVERRMessagesReceived"
      .Add "RESVERRMessagesSent"
      .Add "RESVMessagesReceived"
      .Add "RESVMessagesSent"
      .Add "RESVStateBlockTimeouts"
      .Add "RESVTEARMessagesReceived"
      .Add "RESVTEARMessagesSent"
      .Add "SendMessagesErrorsBigMessages"
      .Add "SendMessagesErrorsNoMemory"
      .Add "SignalingBytesReceived"
      .Add "SignalingBytesSent"
      .Add "Timestamp_Object"
      .Add "Timestamp_PerfTime"
      .Add "Timestamp_Sys100NS"

   End With
   Set PerfFormattedData_RSVP_ACSRSVPInterfaces = c
End Function
Private Function PerfFormattedData_RSVP_ACSRSVPService() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "BytesInQoSNotifications"
      .Add "Caption"
      .Add "Description"
      .Add "FailedQoSRequests"
      .Add "FailedQoSSends"
      .Add "Frequency_Object"
      .Add "Frequency_PerfTime"
      .Add "Frequency_Sys100NS"
      .Add "Name"
      .Add "NetworkInterfaces"
      .Add "NetworkSockets"
      .Add "QoSEnabledReceivers"
      .Add "QoSEnabledSenders"
      .Add "QoSNotifications"
      .Add "QoSSockets"
      .Add "RSVPSessions"
      .Add "Timers"
      .Add "Timestamp_Object"
      .Add "Timestamp_PerfTime"
      .Add "Timestamp_Sys100NS"

   End With
   Set PerfFormattedData_RSVP_ACSRSVPService = c
End Function
Private Function PerfFormattedData_SMTPSVC_SMTPServer() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "AvgRecipientsPerMsgReceived"
      .Add "AvgRecipientsPerMsgSent"
      .Add "AvgRetriesPerMsgDelivered"
      .Add "AvgRetriesPerMsgSent"
      .Add "BadmailedMessagesBadPickupFile"
      .Add "BadmailedMessagesGeneralFailure"
      .Add "BadmailedMessagesHopCountExceeded"
      .Add "BadmailedMessagesNDRofDSN"
      .Add "BadmailedMessagesNoRecipients"
      .Add "BadmailedMessagesTriggeredViaEvent"
      .Add "BytesReceivedPerSec"
      .Add "BytesReceivedTotal"
      .Add "BytesSentPerSec"
      .Add "BytesSentTotal"
      .Add "BytesTotal"
      .Add "BytesTotalPerSec"
      .Add "Caption"
      .Add "CatAddressLookupCompletions"
      .Add "CatAddressLookupCompletionsPerSec"
      .Add "CatAddressLookups"
      .Add "CatAddressLookupsNotFound"
      .Add "CatAddressLookupsPerSec"
      .Add "CatCategorizationsCompleted"
      .Add "CatCategorizationsCompletedPerSec"
      .Add "CatCategorizationsCompletedSuccessfully"
      .Add "CatCategorizationsFailedDSConnectionFailure"
      .Add "CatCategorizationsFailedDSLogonFailure"
      .Add "CatCategorizationsFailedNonRetryableError"
      .Add "CatCategorizationsFailedOutOfMemory"
      .Add "CatCategorizationsFailedRetryableError"
      .Add "CatCategorizationsFailedSinkRetryableError"
      .Add "CatCategorizationsInProgress"
      .Add "CategorizerQueueLength"
      .Add "CatLDAPBindFailures"
      .Add "CatLDAPBinds"
      .Add "CatLDAPConnectionFailures"
      .Add "CatLDAPConnections"
      .Add "CatLDAPConnectionsCurrentlyOpen"
      .Add "CatLDAPGeneralCompletionFailures"
      .Add "CatLDAPPagedSearches"
      .Add "CatLDAPPagedSearchesCompleted"
      .Add "CatLDAPPagedSearchFailures"
      .Add "CatLDAPSearchCompletionFailures"
      .Add "CatLDAPSearches"
      .Add "CatLDAPSearchesAbandoned"
      .Add "CatLDAPSearchesCompleted"
      .Add "CatLDAPSearchesCompletedPerSec"
      .Add "CatLDAPSearchesPendingCompletion"
      .Add "CatLDAPSearchesPerSec"
      .Add "CatLDAPSearchFailures"
      .Add "CatMailMsgDuplicateCollisions"
      .Add "CatMessagesAborted"
      .Add "CatMessagesBifurcated"
      .Add "CatMessagesCategorized"
      .Add "CatMessagesSubmitted"
      .Add "CatMessagesSubmittedPerSec"
      .Add "CatRecipientsAfterCategorization"
      .Add "CatRecipientsBeforeCategorization"
      .Add "CatRecipientsInCategorization"
      .Add "CatRecipientsNDRdAmbiguousAddress"
      .Add "CatRecipientsNDRdByCategorizer"
      .Add "CatRecipientsNDRdForwardingLoop"
      .Add "CatRecipientsNDRdIllegalAddress"
      .Add "CatRecipientsNDRdSinkRecipErrors"
      .Add "CatRecipientsNDRdUnresolved"
      .Add "CatSendersUnresolved"
      .Add "CatSendersWithAmbiguousAddresses"
      .Add "ConnectionErrorsPerSec"
      .Add "CurrentMessagesInLocalDelivery"
      .Add "Description"
      .Add "DirectoryDropsPerSec"
      .Add "DirectoryDropsTotal"
      .Add "DNSQueriesPerSec"
      .Add "DNSQueriesTotal"
      .Add "ETRNMessagesPerSec"
      .Add "ETRNMessagesTotal"
      .Add "Frequency_Object"
      .Add "Frequency_PerfTime"
      .Add "Frequency_Sys100NS"
      .Add "InboundConnectionsCurrent"
      .Add "InboundConnectionsTotal"
      .Add "LocalQueueLength"
      .Add "LocalRetryQueueLength"
      .Add "MessageBytesReceivedPerSec"
      .Add "MessageBytesReceivedTotal"
      .Add "MessageBytesSentPerSec"
      .Add "MessageBytesSentTotal"
      .Add "MessageBytesTotal"
      .Add "MessageBytesTotalPerSec"
      .Add "MessageDeliveryRetries"
      .Add "MessagesCurrentlyUndeliverable"
      .Add "MessagesDeliveredPerSec"
      .Add "MessagesDeliveredTotal"
      .Add "MessageSendRetries"
      .Add "MessagesPendingRouting"
      .Add "MessagesReceivedPerSec"
      .Add "MessagesReceivedTotal"
      .Add "MessagesRefusedForAddressObjects"
      .Add "MessagesRefusedForMailObjects"
      .Add "MessagesRefusedForSize"
      .Add "MessagesSentPerSec"
      .Add "MessagesSentTotal"
      .Add "Name"
      .Add "NDRsGenerated"
      .Add "NumberOfMailFilesOpen"
      .Add "NumberOfQueueFilesOpen"
      .Add "OutboundConnectionsCurrent"
      .Add "OutboundConnectionsRefused"
      .Add "OutboundConnectionsTotal"
      .Add "PercentRecipientsLocal"
      .Add "PercentRecipientsRemote"
      .Add "PickupDirectoryMessagesRetrievedPerSec"
      .Add "PickupDirectoryMessagesRetrievedTotal"
      .Add "RemoteQueueLength"
      .Add "RemoteRetryQueueLength"
      .Add "RoutingTableLookupsPerSec"
      .Add "RoutingTableLookupsTotal"
      .Add "Timestamp_Object"
      .Add "Timestamp_PerfTime"
      .Add "Timestamp_Sys100NS"
      .Add "TotalConnectionErrors"
      .Add "TotalDSNFailures"
      .Add "TotalMessagesSubmitted"

   End With
   Set PerfFormattedData_SMTPSVC_SMTPServer = c
End Function
Private Function PerfFormattedData_Spooler_PrintQueue() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "AddNetworkPrinterCalls"
      .Add "BytesPrintedPerSec"
      .Add "Caption"
      .Add "Description"
      .Add "EnumerateNetworkPrinterCalls"
      .Add "Frequency_Object"
      .Add "Frequency_PerfTime"
      .Add "Frequency_Sys100NS"
      .Add "JobErrors"
      .Add "Jobs"
      .Add "JobsSpooling"
      .Add "MaxJobsSpooling"
      .Add "MaxReferences"
      .Add "Name"
      .Add "NotReadyErrors"
      .Add "OutOfPaperErrors"
      .Add "References"
      .Add "Timestamp_Object"
      .Add "Timestamp_PerfTime"
      .Add "Timestamp_Sys100NS"
      .Add "TotalJobsPrinted"
      .Add "TotalPagesPrinted"

   End With
   Set PerfFormattedData_Spooler_PrintQueue = c
End Function
Private Function PerfFormattedData_TapiSrv_Telephony() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "ActiveLines"
      .Add "ActiveTelephones"
      .Add "Caption"
      .Add "ClientApps"
      .Add "CurrentIncomingCalls"
      .Add "CurrentOutgoingCalls"
      .Add "Description"
      .Add "Frequency_Object"
      .Add "Frequency_PerfTime"
      .Add "Frequency_Sys100NS"
      .Add "IncomingCallsPerSec"
      .Add "Lines"
      .Add "Name"
      .Add "OutgoingCallsPerSec"
      .Add "TelephoneDevices"
      .Add "Timestamp_Object"
      .Add "Timestamp_PerfTime"
      .Add "Timestamp_Sys100NS"

   End With
   Set PerfFormattedData_TapiSrv_Telephony = c
End Function
Private Function PerfFormattedData_Tcpip_ICMP() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Caption"
      .Add "Description"
      .Add "Frequency_Object"
      .Add "Frequency_PerfTime"
      .Add "Frequency_Sys100NS"
      .Add "MessagesOutboundErrors"
      .Add "MessagesPerSec"
      .Add "MessagesReceivedErrors"
      .Add "MessagesReceivedPerSec"
      .Add "MessagesSentPerSec"
      .Add "Name"
      .Add "ReceivedAddressMask"
      .Add "ReceivedAddressMaskReply"
      .Add "ReceivedDestUnreachable"
      .Add "ReceivedEchoPerSec"
      .Add "ReceivedEchoReplyPerSec"
      .Add "ReceivedParameterProblem"
      .Add "ReceivedRedirectPerSec"
      .Add "ReceivedSourceQuench"
      .Add "ReceivedTimeExceeded"
      .Add "ReceivedTimestampPerSec"
      .Add "ReceivedTimestampReplyPerSec"
      .Add "SentAddressMask"
      .Add "SentAddressMaskReply"
      .Add "SentDestinationUnreachable"
      .Add "SentEchoPerSec"
      .Add "SentEchoReplyPerSec"
      .Add "SentParameterProblem"
      .Add "SentRedirectPerSec"
      .Add "SentSourceQuench"
      .Add "SentTimeExceeded"
      .Add "SentTimestampPerSec"
      .Add "SentTimestampReplyPerSec"
      .Add "Timestamp_Object"
      .Add "Timestamp_PerfTime"
      .Add "Timestamp_Sys100NS"

   End With
   Set PerfFormattedData_Tcpip_ICMP = c
End Function
Private Function PerfFormattedData_Tcpip_IP() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Caption"
      .Add "DatagramsForwardedPerSec"
      .Add "DatagramsOutboundDiscarded"
      .Add "DatagramsOutboundNoRoute"
      .Add "DatagramsPerSec"
      .Add "DatagramsReceivedAddressErrors"
      .Add "DatagramsReceivedDeliveredPerSec"
      .Add "DatagramsReceivedDiscarded"
      .Add "DatagramsReceivedHeaderErrors"
      .Add "DatagramsReceivedPerSec"
      .Add "DatagramsReceivedUnknownProtocol"
      .Add "DatagramsSentPerSec"
      .Add "Description"
      .Add "FragmentationFailures"
      .Add "FragmentedDatagramsPerSec"
      .Add "FragmentReassemblyFailures"
      .Add "FragmentsCreatedPerSec"
      .Add "FragmentsReassembledPerSec"
      .Add "FragmentsReceivedPerSec"
      .Add "Frequency_Object"
      .Add "Frequency_PerfTime"
      .Add "Frequency_Sys100NS"
      .Add "Name"
      .Add "Timestamp_Object"
      .Add "Timestamp_PerfTime"
      .Add "Timestamp_Sys100NS"

   End With
   Set PerfFormattedData_Tcpip_IP = c
End Function
Private Function PerfFormattedData_Tcpip_NBTConnection() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "BytesReceivedPerSec"
      .Add "BytesSentPerSec"
      .Add "BytesTotalPerSec"
      .Add "Caption"
      .Add "Description"
      .Add "Frequency_Object"
      .Add "Frequency_PerfTime"
      .Add "Frequency_Sys100NS"
      .Add "Name"
      .Add "Timestamp_Object"
      .Add "Timestamp_PerfTime"
      .Add "Timestamp_Sys100NS"

   End With
   Set PerfFormattedData_Tcpip_NBTConnection = c
End Function
Private Function PerfFormattedData_Tcpip_NetworkInterface() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "BytesReceivedPerSec"
      .Add "BytesSentPerSec"
      .Add "BytesTotalPerSec"
      .Add "Caption"
      .Add "CurrentBandwidth"
      .Add "Description"
      .Add "Frequency_Object"
      .Add "Frequency_PerfTime"
      .Add "Frequency_Sys100NS"
      .Add "Name"
      .Add "OutputQueueLength"
      .Add "PacketsOutboundDiscarded"
      .Add "PacketsOutboundErrors"
      .Add "PacketsPerSec"
      .Add "PacketsReceivedDiscarded"
      .Add "PacketsReceivedErrors"
      .Add "PacketsReceivedNonUnicastPerSec"
      .Add "PacketsReceivedPerSec"
      .Add "PacketsReceivedUnicastPerSec"
      .Add "PacketsReceivedUnknown"
      .Add "PacketsSentNonUnicastPerSec"
      .Add "PacketsSentPerSec"
      .Add "PacketsSentUnicastPerSec"
      .Add "Timestamp_Object"
      .Add "Timestamp_PerfTime"
      .Add "Timestamp_Sys100NS"

   End With
   Set PerfFormattedData_Tcpip_NetworkInterface = c
End Function
Private Function PerfFormattedData_Tcpip_TCP() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Caption"
      .Add "ConnectionFailures"
      .Add "ConnectionsActive"
      .Add "ConnectionsEstablished"
      .Add "ConnectionsPassive"
      .Add "ConnectionsReset"
      .Add "Description"
      .Add "Frequency_Object"
      .Add "Frequency_PerfTime"
      .Add "Frequency_Sys100NS"
      .Add "Name"
      .Add "SegmentsPerSec"
      .Add "SegmentsReceivedPerSec"
      .Add "SegmentsRetransmittedPerSec"
      .Add "SegmentsSentPerSec"
      .Add "Timestamp_Object"
      .Add "Timestamp_PerfTime"
      .Add "Timestamp_Sys100NS"

   End With
   Set PerfFormattedData_Tcpip_TCP = c
End Function
Private Function PerfFormattedData_Tcpip_UDP() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Caption"
      .Add "DatagramsNoPortPerSec"
      .Add "DatagramsPerSec"
      .Add "DatagramsReceivedErrors"
      .Add "DatagramsReceivedPerSec"
      .Add "DatagramsSentPerSec"
      .Add "Description"
      .Add "Frequency_Object"
      .Add "Frequency_PerfTime"
      .Add "Frequency_Sys100NS"
      .Add "Name"
      .Add "Timestamp_Object"
      .Add "Timestamp_PerfTime"
      .Add "Timestamp_Sys100NS"

   End With
   Set PerfFormattedData_Tcpip_UDP = c
End Function
Private Function PerfFormattedData_TermService_TerminalServices() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "ActiveSessions"
      .Add "Caption"
      .Add "Description"
      .Add "Frequency_Object"
      .Add "Frequency_PerfTime"
      .Add "Frequency_Sys100NS"
      .Add "InactiveSessions"
      .Add "Name"
      .Add "Timestamp_Object"
      .Add "Timestamp_PerfTime"
      .Add "Timestamp_Sys100NS"
      .Add "TotalSessions"

   End With
   Set PerfFormattedData_TermService_TerminalServices = c
End Function
Private Function PerfFormattedData_TermService_TerminalServicesSession() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Caption"
      .Add "Description"
      .Add "Frequency_Object"
      .Add "Frequency_PerfTime"
      .Add "Frequency_Sys100NS"
      .Add "HandleCount"
      .Add "InputAsyncFrameError"
      .Add "InputAsyncOverflow"
      .Add "InputAsyncOverrun"
      .Add "InputAsyncParityError"
      .Add "InputBytes"
      .Add "InputCompressedBytes"
      .Add "InputCompressFlushes"
      .Add "InputCompressionRatio"
      .Add "InputErrors"
      .Add "InputFrames"
      .Add "InputTimeouts"
      .Add "InputTransportErrors"
      .Add "InputWaitForOutBuf"
      .Add "InputWdBytes"
      .Add "InputWdFrames"
      .Add "Name"
      .Add "OutputAsyncFrameError"
      .Add "OutputAsyncOverflow"
      .Add "OutputAsyncOverrun"
      .Add "OutputAsyncParityError"
      .Add "OutputBytes"
      .Add "OutputCompressedBytes"
      .Add "OutputCompressFlushes"
      .Add "OutputCompressionRatio"
      .Add "OutputErrors"
      .Add "OutputFrames"
      .Add "OutputTimeouts"
      .Add "OutputTransportErrors"
      .Add "OutputWaitForOutBuf"
      .Add "OutputWdBytes"
      .Add "OutputWdFrames"
      .Add "PageFaultsPerSec"
      .Add "PageFileBytes"
      .Add "PageFileBytesPeak"
      .Add "PercentPrivilegedTime"
      .Add "PercentProcessorTime"
      .Add "PercentUserTime"
      .Add "PoolNonpagedBytes"
      .Add "PoolPagedBytes"
      .Add "PrivateBytes"
      .Add "ProtocolBitmapCacheHitRatio"
      .Add "ProtocolBitmapCacheHits"
      .Add "ProtocolBitmapCacheReads"
      .Add "ProtocolBrushCacheHitRatio"
      .Add "ProtocolBrushCacheHits"
      .Add "ProtocolBrushCacheReads"
      .Add "ProtocolGlyphCacheHitRatio"
      .Add "ProtocolGlyphCacheHits"
      .Add "ProtocolGlyphCacheReads"
      .Add "ProtocolSaveScreenBitmapCacheHitRatio"
      .Add "ProtocolSaveScreenBitmapCacheHits"
      .Add "ProtocolSaveScreenBitmapCacheReads"
      .Add "ThreadCount"
      .Add "Timestamp_Object"
      .Add "Timestamp_PerfTime"
      .Add "Timestamp_Sys100NS"
      .Add "TotalAsyncFrameError"
      .Add "TotalAsyncOverflow"
      .Add "TotalAsyncOverrun"
      .Add "TotalAsyncParityError"
      .Add "TotalBytes"
      .Add "TotalCompressedBytes"
      .Add "TotalCompressFlushes"
      .Add "TotalCompressionRatio"
      .Add "TotalErrors"
      .Add "TotalFrames"
      .Add "TotalProtocolCacheHitRatio"
      .Add "TotalProtocolCacheHits"
      .Add "TotalProtocolCacheReads"
      .Add "TotalTimeouts"
      .Add "TotalTransportErrors"
      .Add "TotalWaitForOutBuf"
      .Add "TotalWdBytes"
      .Add "TotalWdFrames"
      .Add "VirtualBytes"
      .Add "VirtualBytesPeak"
      .Add "WorkingSet"
      .Add "WorkingSetPeak"

   End With
   Set PerfFormattedData_TermService_TerminalServicesSession = c
End Function
Private Function PerfFormattedData_W3SVC_WebService() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "AnonymousUsersPerSec"
      .Add "BytesReceivedPerSec"
      .Add "BytesSentPerSec"
      .Add "BytesTotalPerSec"
      .Add "Caption"
      .Add "CGIRequestsPerSec"
      .Add "ConnectionAttemptsPerSec"
      .Add "CopyRequestsPerSec"
      .Add "CurrentAnonymousUsers"
      .Add "CurrentBlockedAsyncIORequests"
      .Add "CurrentCGIRequests"
      .Add "CurrentConnections"
      .Add "CurrentISAPIExtensionRequests"
      .Add "CurrentNonAnonymousUsers"
      .Add "DeleteRequestsPerSec"
      .Add "Description"
      .Add "FilesPerSec"
      .Add "FilesReceivedPerSec"
      .Add "FilesSentPerSec"
      .Add "Frequency_Object"
      .Add "Frequency_PerfTime"
      .Add "Frequency_Sys100NS"
      .Add "GetRequestsPerSec"
      .Add "HeadRequestsPerSec"
      .Add "ISAPIExtensionRequestsPerSec"
      .Add "LockedErrorsPerSec"
      .Add "LockRequestsPerSec"
      .Add "LogonAttemptsPerSec"
      .Add "MaximumAnonymousUsers"
      .Add "MaximumCGIRequests"
      .Add "MaximumConnections"
      .Add "MaximumISAPIExtensionRequests"
      .Add "MaximumNonAnonymousUsers"
      .Add "MeasuredAsyncIOBandwidthUsage"
      .Add "MkcolRequestsPerSec"
      .Add "MoveRequestsPerSec"
      .Add "Name"
      .Add "NonAnonymousUsersPerSec"
      .Add "NotFoundErrorsPerSec"
      .Add "OptionsRequestsPerSec"
      .Add "OtherRequestMethodsPerSec"
      .Add "PostRequestsPerSec"
      .Add "PropfindRequestsPerSec"
      .Add "ProppatchRequestsPerSec"
      .Add "PutRequestsPerSec"
      .Add "SearchRequestsPerSec"
      .Add "ServiceUptime"
      .Add "Timestamp_Object"
      .Add "Timestamp_PerfTime"
      .Add "Timestamp_Sys100NS"
      .Add "TotalAllowedAsyncIORequests"
      .Add "TotalAnonymousUsers"
      .Add "TotalBlockedAsyncIORequests"
      .Add "TotalCGIRequests"
      .Add "TotalConnectionAttemptsAllInstances"
      .Add "TotalCopyRequests"
      .Add "TotalDeleteRequests"
      .Add "TotalFilesReceived"
      .Add "TotalFilesSent"
      .Add "TotalFilesTransferred"
      .Add "TotalGetRequests"
      .Add "TotalHeadRequests"
      .Add "TotalISAPIExtensionRequests"
      .Add "TotalLockedErrors"
      .Add "TotalLockRequests"
      .Add "TotalLogonAttempts"
      .Add "TotalMethodRequests"
      .Add "TotalMethodRequestsPerSec"
      .Add "TotalMkcolRequests"
      .Add "TotalMoveRequests"
      .Add "TotalNonAnonymousUsers"
      .Add "TotalNotFoundErrors"
      .Add "TotalOptionsRequests"
      .Add "TotalOtherRequestMethods"
      .Add "TotalPostRequests"
      .Add "TotalPropfindRequests"
      .Add "TotalProppatchRequests"
      .Add "TotalPutRequests"
      .Add "TotalRejectedAsyncIORequests"
      .Add "TotalSearchRequests"
      .Add "TotalTraceRequests"
      .Add "TotalUnlockRequests"
      .Add "TraceRequestsPerSec"
      .Add "UnlockRequestsPerSec"

   End With
   Set PerfFormattedData_W3SVC_WebService = c
End Function
Private Function PerfRawData() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Caption"
      .Add "Description"
      .Add "Frequency_Object"
      .Add "Frequency_PerfTime"
      .Add "Frequency_Sys100NS"
      .Add "Name"
      .Add "Timestamp_Object"
      .Add "Timestamp_PerfTime"
      .Add "Timestamp_Sys100NS"

   End With
   Set PerfRawData = c
End Function
Private Function PerfRawData_ASP_ActiveServerPages() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Caption"
      .Add "DebuggingRequests"
      .Add "Description"
      .Add "ErrorsDuringScriptRuntime"
      .Add "ErrorsFromASPPreprocessor"
      .Add "ErrorsFromScriptCompilers"
      .Add "ErrorsPerSec"
      .Add "Frequency_Object"
      .Add "Frequency_PerfTime"
      .Add "Frequency_Sys100NS"
      .Add "InMemoryTemplateCacheHitRate"
      .Add "InMemoryTemplateCacheHitRate_Base"
      .Add "InMemoryTemplatesCached"
      .Add "Name"
      .Add "RequestBytesInTotal"
      .Add "RequestBytesOutTotal"
      .Add "RequestExecutionTime"
      .Add "RequestsDisconnected"
      .Add "RequestsExecuting"
      .Add "RequestsFailedTotal"
      .Add "RequestsNotAuthorized"
      .Add "RequestsNotFound"
      .Add "RequestsPerSec"
      .Add "RequestsQueued"
      .Add "RequestsRejected"
      .Add "RequestsSucceeded"
      .Add "RequestsTimedOut"
      .Add "RequestsTotal"
      .Add "RequestWaitTime"
      .Add "ScriptEnginesCached"
      .Add "SessionDuration"
      .Add "SessionsCurrent"
      .Add "SessionsTimedOut"
      .Add "SessionsTotal"
      .Add "TemplateCacheHitRate"
      .Add "TemplateCacheHitRate_Base"
      .Add "TemplateNotifications"
      .Add "TemplatesCached"
      .Add "Timestamp_Object"
      .Add "Timestamp_PerfTime"
      .Add "Timestamp_Sys100NS"
      .Add "TransactionsAborted"
      .Add "TransactionsCommitted"
      .Add "TransactionsPending"
      .Add "TransactionsPerSec"
      .Add "TransactionsTotal"

   End With
   Set PerfRawData_ASP_ActiveServerPages = c
End Function
Private Function PerfRawData_ContentFilter_IndexingServiceFilter() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "BindingTimeSec"
      .Add "Caption"
      .Add "Description"
      .Add "Frequency_Object"
      .Add "Frequency_PerfTime"
      .Add "Frequency_Sys100NS"
      .Add "IndexingSpeedMBPerHr"
      .Add "Name"
      .Add "Timestamp_Object"
      .Add "Timestamp_PerfTime"
      .Add "Timestamp_Sys100NS"
      .Add "TotalIndexingSpeedMBPerHr"

   End With
   Set PerfRawData_ContentFilter_IndexingServiceFilter = c
End Function
Private Function Processor() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "AddressWidth"
      .Add "Architecture"
      .Add "Availability"
      .Add "Caption"
      .Add "ConfigManagerErrorCode"
      .Add "ConfigManagerUserConfig"
      .Add "CpuStatus"
      .Add "CreationClassName"
      .Add "CurrentClockSpeed"
      .Add "CurrentVoltage"
      .Add "DataWidth"
      .Add "Description"
      .Add "DeviceID"
      .Add "ErrorCleared"
      .Add "ErrorDescription"
      .Add "ExtClock"
      .Add "Family"
      .Add "InstallDate"
      .Add "L2CacheSize"
      .Add "L2CacheSpeed"
      .Add "LastErrorCode"
      .Add "Level"
      .Add "LoadPercentage"
      .Add "Manufacturer"
      .Add "MaxClockSpeed"
      .Add "Name"
      .Add "OtherFamilyDescription"
      .Add "PNPDeviceID"
      .Add "PowerManagementCapabilities[]"
      .Add "PowerManagementSupported"
      .Add "ProcessorId"
      .Add "ProcessorType"
      .Add "Revision"
      .Add "Role"
      .Add "SocketDesignation"
      .Add "Status"
      .Add "StatusInfo"
      .Add "Stepping"
      .Add "SystemCreationClassName"
      .Add "SystemName"
      .Add "UniqueId"
      .Add "UpgradeMethod"
      .Add "Version"
      .Add "VoltageCaps"
   End With
   Set Processor = c
End Function
Private Function ComponentCategory() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Caption"
      .Add "CategoryId"
      .Add "Description"
      .Add "InstallDate"
      .Add "Name"
      .Add "Status"
   End With
   Set ComponentCategory = c
End Function
Private Function COMClass() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Caption"
      .Add "Description"
      .Add "InstallDate"
      .Add "Name"
      .Add "Status"
   End With
   Set COMClass = c
End Function
Private Function CDROMDrive() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Availability"
      .Add "Capabilities[]"
      .Add "CapabilityDescriptions[]"
      .Add "Caption"
      .Add "CompressionMethod"
      .Add "ConfigManagerErrorCode"
      .Add "ConfigManagerUserConfig"
      .Add "CreationClassName"
      .Add "DefaultBlockSize"
      .Add "Description"
      .Add "DeviceID"
      .Add "Drive"
      .Add "DriveIntegrity"
      .Add "ErrorCleared"
      .Add "ErrorDescription"
      .Add "ErrorMethodology"
      .Add "FileSystemFlags"
      .Add "FileSystemFlagsEx"
      .Add "Id"
      .Add "InstallDate"
      .Add "LastErrorCode"
      .Add "Manufacturer"
      .Add "MaxBlockSize"
      .Add "MaximumComponentLength"
      .Add "MaxMediaSize"
      .Add "MediaLoaded"
      .Add "MediaType"
      .Add "MfrAssignedRevisionLevel"
      .Add "MinBlockSize"
      .Add "Name"
      .Add "NeedsCleaning"
      .Add "NumberOfMediaSupported"
      .Add "PNPDeviceID"
      .Add "PowerManagementCapabilities[]"
      .Add "PowerManagementSupported"
      .Add "RevisionLevel"
      .Add "SCSIBus"
      .Add "SCSILogicalUnit"
      .Add "SCSIPort"
      .Add "SCSITargetId"
      .Add "Size"
      .Add "Status"
      .Add "StatusInfo"
      .Add "SystemCreationClassName"
      .Add "SystemName"
      .Add "TransferRate"
      .Add "VolumeName"
      .Add "VolumeSerialNumber"
   End With
   Set CDROMDrive = c
End Function
Private Function ClassicCOMClassSetting() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "AppID"
      .Add "AutoConvertToClsid"
      .Add "AutoTreatAsClsid"
      .Add "Caption"
      .Add "ComponentId"
      .Add "Control"
      .Add "DefaultIcon"
      .Add "Description"
      .Add "InprocHandler"
      .Add "InprocHandler32"
      .Add "InprocServer"
      .Add "InprocServer32"
      .Add "Insertable"
      .Add "JavaClass"
      .Add "LocalServer"
      .Add "LocalServer32"
      .Add "LongDisplayName"
      .Add "ProgId"
      .Add "SettingID"
      .Add "ShortDisplayName"
      .Add "ThreadingModel"
      .Add "ToolBoxBitmap32"
      .Add "TreatAsClsid"
      .Add "TypeLibraryId"
      .Add "Version"
      .Add "VersionIndependentProgId"
   End With
   Set ClassicCOMClassSetting = c
End Function
Private Function COMApplication() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Caption"
      .Add "Description"
      .Add "InstallDate"
      .Add "Name"
      .Add "Status"
   End With
   Set COMApplication = c
End Function
Private Function CommandLineAccess() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Caption"
      .Add "CommandLine"
      .Add "CreationClassName"
      .Add "Description"
      .Add "InstallDate"
      .Add "Name"
      .Add "Status"
      .Add "SystemCreationClassName"
      .Add "SystemName"
      .Add "Type"
   End With
   Set CommandLineAccess = c
End Function
Private Function BindImageAction() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "ActionID"
      .Add "Caption"
      .Add "Description"
      .Add "Direction"
      .Add "File"
      .Add "Name"
      .Add "Path"
      .Add "SoftwareElementID"
      .Add "SoftwareElementState"
      .Add "TargetOperatingSystem"
      .Add "Version"
   End With
   Set BindImageAction = c
End Function
Private Function Condition() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Caption"
      .Add "CheckID"
      .Add "CheckMode"
      .Add "Condition"
      .Add "Description"
      .Add "Feature"
      .Add "Level"
      .Add "Name"
      .Add "SoftwareElementID"
      .Add "SoftwareElementState"
      .Add "TargetOperatingSystem"
      .Add "Version"
   End With
   Set Condition = c
End Function
Private Function COMSetting() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Caption"
      .Add "Description"
      .Add "SettingID"
   End With
   Set COMSetting = c
End Function
Private Function ComputerSystem() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "AdminPasswordStatus"
      .Add "AutomaticResetBootOption"
      .Add "AutomaticResetCapability"
      .Add "BootOptionOnLimit"
      .Add "BootOptionOnWatchDog"
      .Add "BootROMSupported"
      .Add "BootupState"
      .Add "Caption"
      .Add "ChassisBootupState"
      .Add "CreationClassName"
      .Add "CurrentTimeZone"
      .Add "DaylightInEffect"
      .Add "Description"
      .Add "DNSHostName"
      .Add "Domain"
      .Add "DomainRole"
      .Add "EnableDaylightSavingsTime"
      .Add "FrontPanelResetStatus"
      .Add "InfraredSupported"
      .Add "InitialLoadInfo"
      .Add "InstallDate"
      .Add "KeyboardPasswordStatus"
      .Add "LastLoadInfo"
      .Add "Manufacturer"
      .Add "Model"
      .Add "Name"
      .Add "NameFormat"
      .Add "NetworkServerModeEnabled"
      .Add "NumberOfProcessors"
      .Add "OEMLogoBitmap[]"
      .Add "OEMStringArray[]"
      .Add "PartOfDomain"
      .Add "PauseAfterReset"
      .Add "PowerManagementCapabilities[]"
      .Add "PowerManagementSupported"
      .Add "PowerOnPasswordStatus"
      .Add "PowerState"
      .Add "PowerSupplyState"
      .Add "PrimaryOwnerContact"
      .Add "PrimaryOwnerName"
      .Add "ResetCapability"
      .Add "ResetCount"
      .Add "ResetLimit"
      .Add "Roles[]"
      .Add "Status"
      .Add "SupportContactDescription[]"
      .Add "SystemStartupDelay"
      .Add "SystemStartupOptions[]"
      .Add "SystemStartupSetting"
      .Add "SystemType"
      .Add "ThermalState"
      .Add "TotalPhysicalMemory"
      .Add "UserName"
      .Add "WakeUpType"
      .Add "Workgroup"
   End With
   Set ComputerSystem = c
End Function
Private Function ComputerShutdownEvent() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "MachineName"
      .Add "SECURITY_DESCRIPTOR[]"
      .Add "TIME_CREATED"
      .Add "Type"
   End With
   Set ComputerShutdownEvent = c
End Function
Private Function ComputerSystemProduct() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Caption"
      .Add "Description"
      .Add "IdentifyingNumber"
      .Add "Name"
      .Add "SKUNumber"
      .Add "UUID"
      .Add "Vendor"
      .Add "Version"
   End With
   Set ComputerSystemProduct = c
End Function
Private Function AutochkSetting() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Caption"
      .Add "Description"
      .Add "SettingID"
      .Add "UserInputDelay"
   End With
   Set AutochkSetting = c
End Function
Private Function Controller1349() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Availability"
      .Add "Caption"
      .Add "ConfigManagerErrorCode"
      .Add "ConfigManagerUserConfig"
      .Add "CreationClassName"
      .Add "Description"
      .Add "DeviceID"
      .Add "ErrorCleared"
      .Add "ErrorDescription"
      .Add "InstallDate"
      .Add "LastErrorCode"
      .Add "Manufacturer"
      .Add "MaxNumberControlled"
      .Add "Name"
      .Add "PNPDeviceID"
      .Add "PowerManagementCapabilities[]"
      .Add "PowerManagementSupported"
      .Add "ProtocolSupported"
      .Add "Status"
      .Add "StatusInfo"
      .Add "SystemCreationClassName"
      .Add "SystemName"
      .Add "TimeOfLastReset"
   End With
   Set Controller1349 = c
End Function
Private Function ComputerSystemEvent() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "MachineName"
      .Add "SECURITY_DESCRIPTOR[]"
      .Add "TIME_CREATED"
   End With
   Set ComputerSystemEvent = c
End Function
Private Function DCOMApplication() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "AppID"
      .Add "Caption"
      .Add "Description"
      .Add "InstallDate"
      .Add "Name"
      .Add "Status"
   End With
   Set DCOMApplication = c
End Function
Private Function CurrentTime() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Day"
      .Add "DayOfWeek"
      .Add "Hour"
      .Add "Milliseconds"
      .Add "Minute"
      .Add "Month"
      .Add "Quarter"
      .Add "Second"
      .Add "WeekInMonth"
      .Add "Year"
   End With
   Set CurrentTime = c
End Function
Private Function BootConfiguration() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "BootDirectory"
      .Add "Caption"
      .Add "ConfigurationPath"
      .Add "Description"
      .Add "LastDrive"
      .Add "Name"
      .Add "ScratchDirectory"
      .Add "SettingID"
      .Add "TempDirectory"
   End With
   Set BootConfiguration = c
End Function
Private Function Binary() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Caption"
      .Add "Data"
      .Add "Description"
      .Add "Name"
      .Add "ProductCode"
      .Add "SettingID"
   End With
   Set Binary = c
End Function
Private Function BaseBoard() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Caption"
      .Add "ConfigOptions[]"
      .Add "CreationClassName"
      .Add "Depth"
      .Add "Description"
      .Add "Height"
      .Add "HostingBoard"
      .Add "HotSwappable"
      .Add "InstallDate"
      .Add "Manufacturer"
      .Add "Model"
      .Add "Name"
      .Add "OtherIdentifyingInfo"
      .Add "PartNumber"
      .Add "PoweredOn"
      .Add "Product"
      .Add "Removable"
      .Add "Replaceable"
      .Add "RequirementsDescription"
      .Add "RequiresDaughterBoard"
      .Add "SerialNumber"
      .Add "SKU"
      .Add "SlotLayout"
      .Add "SpecialRequirements"
      .Add "Status"
      .Add "Tag"
      .Add "Version"
      .Add "Weight"
      .Add "Width"
   End With
   Set BaseBoard = c
End Function
Private Function BaseService() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "AcceptPause"
      .Add "AcceptStop"
      .Add "Caption"
      .Add "CreationClassName"
      .Add "Description"
      .Add "DesktopInteract"
      .Add "DisplayName"
      .Add "ErrorControl"
      .Add "ExitCode"
      .Add "InstallDate"
      .Add "Name"
      .Add "PathName"
      .Add "ServiceSpecificExitCode"
      .Add "ServiceType"
      .Add "Started"
      .Add "StartMode"
      .Add "StartName"
      .Add "State"
      .Add "Status"
      .Add "SystemCreationClassName"
      .Add "SystemName"
      .Add "TagId"
   End With
   Set BaseService = c
End Function
Private Function CacheMemory() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Access"
      .Add "AdditionalErrorData[]"
      .Add "Associativity"
      .Add "Availability"
      .Add "BlockSize"
      .Add "CacheSpeed"
      .Add "CacheType"
      .Add "Caption"
      .Add "ConfigManagerErrorCode"
      .Add "ConfigManagerUserConfig"
      .Add "CorrectableError"
      .Add "CreationClassName"
      .Add "CurrentSRAM[]"
      .Add "Description"
      .Add "DeviceID"
      .Add "EndingAddress"
      .Add "ErrorAccess"
      .Add "ErrorAddress"
      .Add "ErrorCleared"
      .Add "ErrorCorrectType"
      .Add "ErrorData[]"
      .Add "ErrorDataOrder"
      .Add "ErrorDescription"
      .Add "ErrorInfo"
      .Add "ErrorMethodology"
      .Add "ErrorResolution"
      .Add "ErrorTime"
      .Add "ErrorTransferSize"
      .Add "FlushTimer"
      .Add "InstallDate"
      .Add "InstalledSize"
      .Add "LastErrorCode"
      .Add "Level"
      .Add "LineSize"
      .Add "Location"
      .Add "MaxCacheSize"
      .Add "Name"
      .Add "NumberOfBlocks"
      .Add "OtherErrorDescription"
      .Add "PNPDeviceID"
      .Add "PowerManagementCapabilities[]"
      .Add "PowerManagementSupported"
      .Add "Purpose"
      .Add "ReadPolicy"
      .Add "ReplacementPolicy"
      .Add "StartingAddress"
      .Add "Status"
      .Add "StatusInfo"
      .Add "SupportedSRAM[]"
      .Add "SystemCreationClassName"
      .Add "SystemLevelAddress"
      .Add "SystemName"
      .Add "WritePolicy"
   End With
   Set CacheMemory = c
End Function
Private Function ClassicCOMClass() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Caption"
      .Add "ComponentId"
      .Add "Description"
      .Add "InstallDate"
      .Add "Name"
      .Add "Status"
   End With
   Set ClassicCOMClass = c
End Function
Private Function ClassInfoAction() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "ActionID"
      .Add "AppID"
      .Add "Argument"
      .Add "Caption"
      .Add "CLSID"
      .Add "Context"
      .Add "DefInprocHandler"
      .Add "Description"
      .Add "Direction"
      .Add "FileTypeMask"
      .Add "Insertable"
      .Add "Name"
      .Add "ProgID"
      .Add "RemoteName"
      .Add "SoftwareElementID"
      .Add "SoftwareElementState"
      .Add "TargetOperatingSystem"
      .Add "Version"
      .Add "VIProgID"
   End With
   Set ClassInfoAction = c
End Function
Private Function CreateFolderAction() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "ActionID"
      .Add "Caption"
      .Add "Description"
      .Add "Direction"
      .Add "DirectoryName"
      .Add "Name"
      .Add "SoftwareElementID"
      .Add "SoftwareElementState"
      .Add "TargetOperatingSystem"
      .Add "Version"
   End With
   Set CreateFolderAction = c
End Function
Private Function CurrentProbe() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Accuracy"
      .Add "Availability"
      .Add "Caption"
      .Add "ConfigManagerErrorCode"
      .Add "ConfigManagerUserConfig"
      .Add "CreationClassName"
      .Add "CurrentReading"
      .Add "Description"
      .Add "DeviceID"
      .Add "ErrorCleared"
      .Add "ErrorDescription"
      .Add "InstallDate"
      .Add "IsLinear"
      .Add "LastErrorCode"
      .Add "LowerThresholdCritical"
      .Add "LowerThresholdFatal"
      .Add "LowerThresholdNonCritical"
      .Add "MaxReadable"
      .Add "MinReadable"
      .Add "Name"
      .Add "NominalReading"
      .Add "NormalMax"
      .Add "NormalMin"
      .Add "PNPDeviceID"
      .Add "PowerManagementCapabilities[]"
      .Add "PowerManagementSupported"
      .Add "Resolution"
      .Add "Status"
      .Add "StatusInfo"
      .Add "SystemCreationClassName"
      .Add "SystemName"
      .Add "Tolerance"
      .Add "UpperThresholdCritical"
      .Add "UpperThresholdFatal"
      .Add "UpperThresholdNonCritical"
   End With
   Set CurrentProbe = c
End Function




Private Function PerfRawData_ContentIndex_IndexingService() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Caption"
      .Add "DeferredForIndexing"
      .Add "Description"
      .Add "FilesToBeIndexed"
      .Add "Frequency_Object"
      .Add "Frequency_PerfTime"
      .Add "Frequency_Sys100NS"
      .Add "IndexSizeMB"
      .Add "MergeProgress"
      .Add "Name"
      .Add "NumberDocumentsIndexed"
      .Add "RunningQueries"
      .Add "SavedIndexes"
      .Add "Timestamp_Object"
      .Add "Timestamp_PerfTime"
      .Add "Timestamp_Sys100NS"
      .Add "UniqueKeys"
      .Add "WordLists"
   End With
   Set PerfRawData_ContentIndex_IndexingService = c
End Function


Private Function SoftwareElement() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Attributes"
      .Add "BuildNumber"
      .Add "Caption"
      .Add "CodeSet"
      .Add "Description"
      .Add "IdentificationCode"
      .Add "InstallDate"
      .Add "InstallState"
      .Add "LanguageEdition"
      .Add "Manufacturer"
      .Add "Name"
      .Add "OtherTargetOS"
      .Add "Path"
      .Add "SerialNumber"
      .Add "SoftwareElementID"
      .Add "SoftwareElementState"
      .Add "Status"
      .Add "TargetOperatingSystem"
      .Add "Version"
   End With
   Set SoftwareElement = c
End Function
Private Function SoftwareElementCondition() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Caption"
      .Add "CheckID"
      .Add "CheckMode"
      .Add "Condition"
      .Add "Description"
      .Add "Name"
      .Add "SoftwareElementID"
      .Add "SoftwareElementState"
      .Add "TargetOperatingSystem"
      .Add "Version"
   End With
   Set SoftwareElementCondition = c
End Function
Private Function SoftwareFeature() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Accesses"
      .Add "Attributes"
      .Add "Caption"
      .Add "Description"
      .Add "IdentifyingNumber"
      .Add "InstallDate"
      .Add "InstallState"
      .Add "LastUse"
      .Add "Name"
      .Add "ProductName"
      .Add "Status"
      .Add "Vendor"
      .Add "Version"
   End With
   Set SoftwareFeature = c
End Function
Private Function SoundDevice() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Availability"
      .Add "Caption"
      .Add "ConfigManagerErrorCode"
      .Add "ConfigManagerUserConfig"
      .Add "CreationClassName"
      .Add "Description"
      .Add "DeviceID"
      .Add "DMABufferSize"
      .Add "ErrorCleared"
      .Add "ErrorDescription"
      .Add "InstallDate"
      .Add "LastErrorCode"
      .Add "Manufacturer"
      .Add "MPU401Address"
      .Add "Name"
      .Add "PNPDeviceID"
      .Add "PowerManagementCapabilities[]"
      .Add "PowerManagementSupported"
      .Add "ProductName"
      .Add "Status"
      .Add "StatusInfo"
      .Add "SystemCreationClassName"
      .Add "SystemName"
   End With
   Set SoundDevice = c
End Function
Private Function PerfRawData_PerfDisk_PhysicalDisk() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "AvgDiskBytesPerRead"
      .Add "AvgDiskBytesPerRead_Base"
      .Add "AvgDiskBytesPerTransfer"
      .Add "AvgDiskBytesPerTransfer_Base"
      .Add "AvgDiskBytesPerWrite"
      .Add "AvgDiskBytesPerWrite_Base"
      .Add "AvgDiskQueueLength"
      .Add "AvgDiskReadQueueLength"
      .Add "AvgDiskSecPerRead"
      .Add "AvgDiskSecPerRead_Base"
      .Add "AvgDiskSecPerTransfer"
      .Add "AvgDiskSecPerTransfer_Base"
      .Add "AvgDiskSecPerWrite"
      .Add "AvgDiskSecPerWrite_Base"
      .Add "AvgDiskWriteQueueLength"
      .Add "Caption"
      .Add "CurrentDiskQueueLength"
      .Add "Description"
      .Add "DiskBytesPerSec"
      .Add "DiskReadBytesPerSec"
      .Add "DiskReadsPerSec"
      .Add "DiskTransfersPerSec"
      .Add "DiskWriteBytesPerSec"
      .Add "DiskWritesPerSec"
      .Add "Frequency_Object"
      .Add "Frequency_PerfTime"
      .Add "Frequency_Sys100NS"
      .Add "Name"
      .Add "PercentDiskReadTime"
      .Add "PercentDiskReadTime_Base"
      .Add "PercentDiskTime"
      .Add "PercentDiskTime_Base"
      .Add "PercentDiskWriteTime"
      .Add "PercentDiskWriteTime_Base"
      .Add "PercentFreeSpace_Base"
      .Add "PercentIdleTime"
      .Add "PercentIdleTime_Base"
      .Add "SplitIOPerSec"
      .Add "Timestamp_Object"
      .Add "Timestamp_PerfTime"
      .Add "Timestamp_Sys100NS"
   End With
   Set PerfRawData_PerfDisk_PhysicalDisk = c
End Function
Private Function PerfRawData_PerfNet_Browser() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "AnnouncementsDomainPerSec"
      .Add "AnnouncementsServerPerSec"
      .Add "AnnouncementsTotalPerSec"
      .Add "Caption"
      .Add "Description"
      .Add "DuplicateMasterAnnouncements"
      .Add "ElectionPacketsPerSec"
      .Add "EnumerationsDomainPerSec"
      .Add "EnumerationsOtherPerSec"
      .Add "EnumerationsServerPerSec"
      .Add "EnumerationsTotalPerSec"
      .Add "Frequency_Object"
      .Add "Frequency_PerfTime"
      .Add "Frequency_Sys100NS"
      .Add "IllegalDatagramsPerSec"
      .Add "MailslotAllocationsFailed"
      .Add "MailslotOpensFailedPerSec"
      .Add "MailslotReceivesFailed"
      .Add "MailslotWritesFailed"
      .Add "MailslotWritesPerSec"
      .Add "MissedMailslotDatagrams"
      .Add "MissedServerAnnouncements"
      .Add "MissedServerListRequests"
      .Add "Name"
      .Add "ServerAnnounceAllocationsFailedPerSec"
      .Add "ServerListRequestsPerSec"
      .Add "Timestamp_Object"
      .Add "Timestamp_PerfTime"
      .Add "Timestamp_Sys100NS"
   End With
   Set PerfRawData_PerfNet_Browser = c
End Function
Private Function PerfRawData_PerfNet_Redirector() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "BytesReceivedPerSec"
      .Add "BytesTotalPerSec"
      .Add "BytesTransmittedPerSec"
      .Add "Caption"
      .Add "ConnectsCore"
      .Add "ConnectsLanManager20"
      .Add "ConnectsLanManager21"
      .Add "ConnectsWindowsNT"
      .Add "CurrentCommands"
      .Add "Description"
      .Add "FileDataOperationsPerSec"
      .Add "FileReadOperationsPerSec"
      .Add "FileWriteOperationsPerSec"
      .Add "Frequency_Object"
      .Add "Frequency_PerfTime"
      .Add "Frequency_Sys100NS"
      .Add "Name"
      .Add "NetworkErrorsPerSec"
      .Add "PacketsPerSec"
      .Add "PacketsReceivedPerSec"
      .Add "PacketsTransmittedPerSec"
      .Add "ReadBytesCachePerSec"
      .Add "ReadBytesNetworkPerSec"
      .Add "ReadBytesNonPagingPerSec"
      .Add "ReadBytesPagingPerSec"
      .Add "ReadOperationsRandomPerSec"
      .Add "ReadPacketsPerSec"
      .Add "ReadPacketsSmallPerSec"
      .Add "ReadsDeniedPerSec"
      .Add "ReadsLargePerSec"
      .Add "ServerDisconnects"
      .Add "ServerReconnects"
      .Add "ServerSessions"
      .Add "ServerSessionsHung"
      .Add "Timestamp_Object"
      .Add "Timestamp_PerfTime"
      .Add "Timestamp_Sys100NS"
      .Add "WriteBytesCachePerSec"
      .Add "WriteBytesNetworkPerSec"
      .Add "WriteBytesNonPagingPerSec"
      .Add "WriteBytesPagingPerSec"
      .Add "WriteDeniedPerSec"
      .Add "WriteLargePerSec"
      .Add "WriteOperationsRandomPerSec"
      .Add "WritePacketsPerSec"
      .Add "WritePacketsSmallPerSec"
   End With
   Set PerfRawData_PerfNet_Redirector = c
End Function
Private Function PerfRawData_PerfNet_Server() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "BlockingRequestsRejected"
      .Add "BytesReceivedPerSec"
      .Add "BytesTotalPerSec"
      .Add "BytesTransmittedPerSec"
      .Add "Caption"
      .Add "ContextBlocksQueuedPerSec"
      .Add "Description"
      .Add "ErrorsAccessPermissions"
      .Add "ErrorsGrantedAccess"
      .Add "ErrorsLogon"
      .Add "ErrorsSystem"
      .Add "FileDirectorySearches"
      .Add "FilesOpen"
      .Add "FilesOpenedTotal"
      .Add "Frequency_Object"
      .Add "Frequency_PerfTime"
      .Add "Frequency_Sys100NS"
      .Add "LogonPerSec"
      .Add "LogonTotal"
      .Add "Name"
      .Add "PoolNonPagedBytes"
      .Add "PoolNonPagedFailures"
      .Add "PoolNonPagedPeak"
      .Add "PoolPagedBytes"
      .Add "PoolPagedFailures"
      .Add "PoolPagedPeak"
      .Add "ServerSessions"
      .Add "SessionsErroredOut"
      .Add "SessionsForcedOff"
      .Add "SessionsLoggedOff"
      .Add "SessionsTimedOut"
      .Add "Timestamp_Object"
      .Add "Timestamp_PerfTime"
      .Add "Timestamp_Sys100NS"
      .Add "WorkItemShortages"
   End With
   Set PerfRawData_PerfNet_Server = c
End Function
Private Function PerfRawData_PerfNet_ServerWorkQueues() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "ActiveThreads"
      .Add "AvailableThreads"
      .Add "AvailableWorkItems"
      .Add "BorrowedWorkItems"
      .Add "BytesReceivedPerSec"
      .Add "BytesSentPerSec"
      .Add "BytesTransferredPerSec"
      .Add "Caption"
      .Add "ContextBlocksQueuedPerSec"
      .Add "CurrentClients"
      .Add "Description"
      .Add "Frequency_Object"
      .Add "Frequency_PerfTime"
      .Add "Frequency_Sys100NS"
      .Add "Name"
      .Add "QueueLength"
      .Add "ReadBytesPerSec"
      .Add "ReadOperationsPerSec"
      .Add "Timestamp_Object"
      .Add "Timestamp_PerfTime"
      .Add "Timestamp_Sys100NS"
      .Add "TotalBytesPerSec"
      .Add "TotalOperationsPerSec"
      .Add "WorkItemShortages"
      .Add "WriteBytesPerSec"
      .Add "WriteOperationsPerSec"
   End With
   Set PerfRawData_PerfNet_ServerWorkQueues = c
End Function
Private Function PerfRawData_PerfOS_Cache() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "AsyncCopyReadsPerSec"
      .Add "AsyncDataMapsPerSec"
      .Add "AsyncFastReadsPerSec"
      .Add "AsyncMDLReadsPerSec"
      .Add "AsyncPinReadsPerSec"
      .Add "Caption"
      .Add "CopyReadHitsPercent"
      .Add "CopyReadHitsPercent_Base"
      .Add "CopyReadsPerSec"
      .Add "DataFlushesPerSec"
      .Add "DataFlushPagesPerSec"
      .Add "DataMapHitsPercent"
      .Add "DataMapHitsPercent_Base"
      .Add "DataMapPinsPerSec"
      .Add "DataMapPinsPerSec_Base"
      .Add "DataMapsPerSec"
      .Add "Description"
      .Add "FastReadNotPossiblesPerSec"
      .Add "FastReadResourceMissesPerSec"
      .Add "FastReadsPerSec"
      .Add "Frequency_Object"
      .Add "Frequency_PerfTime"
      .Add "Frequency_Sys100NS"
      .Add "LazyWriteFlushesPerSec"
      .Add "LazyWritePagesPerSec"
      .Add "MDLReadHitsPercent"
      .Add "MDLReadHitsPercent_Base"
      .Add "MDLReadsPerSec"
      .Add "Name"
      .Add "PinReadHitsPercent"
      .Add "PinReadHitsPercent_Base"
      .Add "PinReadsPerSec"
      .Add "ReadAheadsPerSec"
      .Add "SyncCopyReadsPerSec"
      .Add "SyncDataMapsPerSec"
      .Add "SyncFastReadsPerSec"
      .Add "SyncMDLReadsPerSec"
      .Add "SyncPinReadsPerSec"
      .Add "Timestamp_Object"
      .Add "Timestamp_PerfTime"
      .Add "Timestamp_Sys100NS"
   End With
   Set PerfRawData_PerfOS_Cache = c
End Function
Private Function PerfRawData_PerfOS_Memory() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "AvailableBytes"
      .Add "AvailableKBytes"
      .Add "AvailableMBytes"
      .Add "CacheBytes"
      .Add "CacheBytesPeak"
      .Add "CacheFaultsPerSec"
      .Add "Caption"
      .Add "CommitLimit"
      .Add "CommittedBytes"
      .Add "DemandZeroFaultsPerSec"
      .Add "Description"
      .Add "FreeSystemPageTableEntries"
      .Add "Frequency_Object"
      .Add "Frequency_PerfTime"
      .Add "Frequency_Sys100NS"
      .Add "Name"
      .Add "PageFaultsPerSec"
      .Add "PageReadsPerSec"
      .Add "PagesInputPerSec"
      .Add "PagesOutputPerSec"
      .Add "PagesPerSec"
      .Add "PageWritesPerSec"
      .Add "PercentCommittedBytesInUse"
      .Add "PercentCommittedBytesInUse_Base"
      .Add "PoolNonpagedAllocs"
      .Add "PoolNonpagedBytes"
      .Add "PoolPagedAllocs"
      .Add "PoolPagedBytes"
      .Add "PoolPagedResidentBytes"
      .Add "SystemCacheResidentBytes"
      .Add "SystemCodeResidentBytes"
      .Add "SystemCodeTotalBytes"
      .Add "SystemDriverResidentBytes"
      .Add "SystemDriverTotalBytes"
      .Add "Timestamp_Object"
      .Add "Timestamp_PerfTime"
      .Add "Timestamp_Sys100NS"
      .Add "TransitionFaultsPerSec"
      .Add "WriteCopiesPerSec"
   End With
   Set PerfRawData_PerfOS_Memory = c
End Function
Private Function PerfRawData_PerfOS_Objects() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Caption"
      .Add "Description"
      .Add "Events"
      .Add "Frequency_Object"
      .Add "Frequency_PerfTime"
      .Add "Frequency_Sys100NS"
      .Add "Mutexes"
      .Add "Name"
      .Add "Processes"
      .Add "Sections"
      .Add "Semaphores"
      .Add "Threads"
      .Add "Timestamp_Object"
      .Add "Timestamp_PerfTime"
      .Add "Timestamp_Sys100NS"
   End With
   Set PerfRawData_PerfOS_Objects = c
End Function
Private Function PerfRawData_PerfOS_PagingFile() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Caption"
      .Add "Description"
      .Add "Frequency_Object"
      .Add "Frequency_PerfTime"
      .Add "Frequency_Sys100NS"
      .Add "Name"
      .Add "PercentUsage"
      .Add "PercentUsage_Base"
      .Add "PercentUsagePeak"
      .Add "PercentUsagePeak_Base"
      .Add "Timestamp_Object"
      .Add "Timestamp_PerfTime"
      .Add "Timestamp_Sys100NS"
   End With
   Set PerfRawData_PerfOS_PagingFile = c
End Function
Private Function PerfRawData_PerfOS_Processor() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "C1TransitionsPerSec"
      .Add "C2TransitionsPerSec"
      .Add "C3TransitionsPerSec"
      .Add "Caption"
      .Add "Description"
      .Add "DPCRate"
      .Add "DPCsQueuedPerSec"
      .Add "Frequency_Object"
      .Add "Frequency_PerfTime"
      .Add "Frequency_Sys100NS"
      .Add "InterruptsPerSec"
      .Add "Name"
      .Add "PercentC1Time"
      .Add "PercentC2Time"
      .Add "PercentC3Time"
      .Add "PercentDPCTime"
      .Add "PercentIdleTime"
      .Add "PercentInterruptTime"
      .Add "PercentPrivilegedTime"
      .Add "PercentProcessorTime"
      .Add "PercentUserTime"
      .Add "Timestamp_Object"
      .Add "Timestamp_PerfTime"
      .Add "Timestamp_Sys100NS"
   End With
   Set PerfRawData_PerfOS_Processor = c
End Function
Private Function PerfRawData_PerfOS_System() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "AlignmentFixupsPerSec"
      .Add "Caption"
      .Add "ContextSwitchesPerSec"
      .Add "Description"
      .Add "ExceptionDispatchesPerSec"
      .Add "FileControlBytesPerSec"
      .Add "FileControlOperationsPerSec"
      .Add "FileDataOperationsPerSec"
      .Add "FileReadBytesPerSec"
      .Add "FileReadOperationsPerSec"
      .Add "FileWriteBytesPerSec"
      .Add "FileWriteOperationsPerSec"
      .Add "FloatingEmulationsPerSec"
      .Add "Frequency_Object"
      .Add "Frequency_PerfTime"
      .Add "Frequency_Sys100NS"
      .Add "Name"
      .Add "PercentRegistryQuotaInUse"
      .Add "PercentRegistryQuotaInUse_Base"
      .Add "Processes"
      .Add "ProcessorQueueLength"
      .Add "SystemCallsPerSec"
      .Add "SystemUpTime"
      .Add "Threads"
      .Add "Timestamp_Object"
      .Add "Timestamp_PerfTime"
      .Add "Timestamp_Sys100NS"
   End With
   Set PerfRawData_PerfOS_System = c
End Function
Private Function PerfRawData_PerfProc_FullImage_Costly() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Caption"
      .Add "Description"
      .Add "ExecReadOnly"
      .Add "ExecReadPerWrite"
      .Add "Executable"
      .Add "ExecWriteCopy"
      .Add "Frequency_Object"
      .Add "Frequency_PerfTime"
      .Add "Frequency_Sys100NS"
      .Add "Name"
      .Add "NoAccess"
      .Add "ReadOnly"
      .Add "ReadPerWrite"
      .Add "Timestamp_Object"
      .Add "Timestamp_PerfTime"
      .Add "Timestamp_Sys100NS"
      .Add "WriteCopy"
   End With
   Set PerfRawData_PerfProc_FullImage_Costly = c
End Function
Private Function PerfRawData_PerfProc_Image_Costly() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Caption"
      .Add "Description"
      .Add "ExecReadOnly"
      .Add "ExecReadPerWrite"
      .Add "Executable"
      .Add "ExecWriteCopy"
      .Add "Frequency_Object"
      .Add "Frequency_PerfTime"
      .Add "Frequency_Sys100NS"
      .Add "Name"
      .Add "NoAccess"
      .Add "ReadOnly"
      .Add "ReadPerWrite"
      .Add "Timestamp_Object"
      .Add "Timestamp_PerfTime"
      .Add "Timestamp_Sys100NS"
      .Add "WriteCopy"
   End With
   Set PerfRawData_PerfProc_Image_Costly = c
End Function
Private Function PerfRawData_PerfProc_JobObject() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Caption"
      .Add "CurrentPercentKernelModeTime"
      .Add "CurrentPercentProcessorTime"
      .Add "CurrentPercentUserModeTime"
      .Add "Description"
      .Add "Frequency_Object"
      .Add "Frequency_PerfTime"
      .Add "Frequency_Sys100NS"
      .Add "Name"
      .Add "PagesPerSec"
      .Add "ProcessCountActive"
      .Add "ProcessCountTerminated"
      .Add "ProcessCountTotal"
      .Add "ThisPeriodmSecKernelMode"
      .Add "ThisPeriodmSecProcessor"
      .Add "ThisPeriodmSecUserMode"
      .Add "Timestamp_Object"
      .Add "Timestamp_PerfTime"
      .Add "Timestamp_Sys100NS"
      .Add "TotalmSecKernelMode"
      .Add "TotalmSecProcessor"
      .Add "TotalmSecUserMode"
   End With
   Set PerfRawData_PerfProc_JobObject = c
End Function
Private Function PerfRawData_PerfProc_JobObjectDetails() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Caption"
      .Add "CreatingProcessID"
      .Add "Description"
      .Add "ElapsedTime"
      .Add "Frequency_Object"
      .Add "Frequency_PerfTime"
      .Add "Frequency_Sys100NS"
      .Add "HandleCount"
      .Add "IDProcess"
      .Add "IODataBytesPerSec"
      .Add "IODataOperationsPerSec"
      .Add "IOOtherBytesPerSec"
      .Add "IOOtherOperationsPerSec"
      .Add "IOReadBytesPerSec"
      .Add "IOReadOperationsPerSec"
      .Add "IOWriteBytesPerSec"
      .Add "IOWriteOperationsPerSec"
      .Add "Name"
      .Add "PageFaultsPerSec"
      .Add "PageFileBytes"
      .Add "PageFileBytesPeak"
      .Add "PercentPrivilegedTime"
      .Add "PercentProcessorTime"
      .Add "PercentUserTime"
      .Add "PoolNonpagedBytes"
      .Add "PoolPagedBytes"
      .Add "PriorityBase"
      .Add "PrivateBytes"
      .Add "ThreadCount"
      .Add "Timestamp_Object"
      .Add "Timestamp_PerfTime"
      .Add "Timestamp_Sys100NS"
      .Add "VirtualBytes"
      .Add "VirtualBytesPeak"
      .Add "WorkingSet"
      .Add "WorkingSetPeak"
   End With
   Set PerfRawData_PerfProc_JobObjectDetails = c
End Function
Private Function StartupCommand() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Caption"
      .Add "Command"
      .Add "Description"
      .Add "Location"
      .Add "Name"
      .Add "SettingID"
      .Add "User"
   End With
   Set StartupCommand = c
End Function
Private Function PerfRawData_PerfProc_Process() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Caption"
      .Add "CreatingProcessID"
      .Add "Description"
      .Add "ElapsedTime"
      .Add "Frequency_Object"
      .Add "Frequency_PerfTime"
      .Add "Frequency_Sys100NS"
      .Add "HandleCount"
      .Add "IDProcess"
      .Add "IODataBytesPerSec"
      .Add "IODataOperationsPerSec"
      .Add "IOOtherBytesPerSec"
      .Add "IOOtherOperationsPerSec"
      .Add "IOReadBytesPerSec"
      .Add "IOReadOperationsPerSec"
      .Add "IOWriteBytesPerSec"
      .Add "IOWriteOperationsPerSec"
      .Add "Name"
      .Add "PageFaultsPerSec"
      .Add "PageFileBytes"
      .Add "PageFileBytesPeak"
      .Add "PercentPrivilegedTime"
      .Add "PercentProcessorTime"
      .Add "PercentUserTime"
      .Add "PoolNonpagedBytes"
      .Add "PoolPagedBytes"
      .Add "PriorityBase"
      .Add "PrivateBytes"
      .Add "ThreadCount"
      .Add "Timestamp_Object"
      .Add "Timestamp_PerfTime"
      .Add "Timestamp_Sys100NS"
      .Add "VirtualBytes"
      .Add "VirtualBytesPeak"
      .Add "WorkingSet"
      .Add "WorkingSetPeak"
   End With
   Set PerfRawData_PerfProc_Process = c
End Function
Private Function PerfRawData_PerfProc_ProcessAddressSpace_Costly() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "BytesFree"
      .Add "BytesImageFree"
      .Add "BytesImageReserved"
      .Add "BytesReserved"
      .Add "Caption"
      .Add "Description"
      .Add "Frequency_Object"
      .Add "Frequency_PerfTime"
      .Add "Frequency_Sys100NS"
      .Add "IDProcess"
      .Add "ImageSpaceExecReadOnly"
      .Add "ImageSpaceExecReadPerWrite"
      .Add "ImageSpaceExecutable"
      .Add "ImageSpaceExecWriteCopy"
      .Add "ImageSpaceNoAccess"
      .Add "ImageSpaceReadOnly"
      .Add "ImageSpaceReadPerWrite"
      .Add "ImageSpaceWriteCopy"
      .Add "MappedSpaceExecReadOnly"
      .Add "MappedSpaceExecReadPerWrite"
      .Add "MappedSpaceExecutable"
      .Add "MappedSpaceExecWriteCopy"
      .Add "MappedSpaceNoAccess"
      .Add "MappedSpaceReadOnly"
      .Add "MappedSpaceReadPerWrite"
      .Add "MappedSpaceWriteCopy"
      .Add "Name"
      .Add "ReservedSpaceExecReadOnly"
      .Add "ReservedSpaceExecReadPerWrite"
      .Add "ReservedSpaceExecutable"
      .Add "ReservedSpaceExecWriteCopy"
      .Add "ReservedSpaceNoAccess"
      .Add "ReservedSpaceReadOnly"
      .Add "ReservedSpaceReadPerWrite"
      .Add "ReservedSpaceWriteCopy"
      .Add "Timestamp_Object"
      .Add "Timestamp_PerfTime"
      .Add "Timestamp_Sys100NS"
      .Add "UnassignedSpaceExecReadOnly"
      .Add "UnassignedSpaceExecReadPerWrite"
      .Add "UnassignedSpaceExecutable"
      .Add "UnassignedSpaceExecWriteCopy"
      .Add "UnassignedSpaceNoAccess"
      .Add "UnassignedSpaceReadOnly"
      .Add "UnassignedSpaceReadPerWrite"
      .Add "UnassignedSpaceWriteCopy"
   End With
   Set PerfRawData_PerfProc_ProcessAddressSpace_Costly = c
End Function
Private Function PerfRawData_PerfProc_Thread() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Caption"
      .Add "ContextSwitchesPerSec"
      .Add "Description"
      .Add "ElapsedTime"
      .Add "Frequency_Object"
      .Add "Frequency_PerfTime"
      .Add "Frequency_Sys100NS"
      .Add "IDProcess"
      .Add "IDThread"
      .Add "Name"
      .Add "PercentPrivilegedTime"
      .Add "PercentProcessorTime"
      .Add "PercentUserTime"
      .Add "PriorityBase"
      .Add "PriorityCurrent"
      .Add "StartAddress"
      .Add "ThreadState"
      .Add "ThreadWaitReason"
      .Add "Timestamp_Object"
      .Add "Timestamp_PerfTime"
      .Add "Timestamp_Sys100NS"
   End With
   Set PerfRawData_PerfProc_Thread = c
End Function
Private Function PerfRawData_PerfProc_ThreadDetails_Costly() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Caption"
      .Add "Description"
      .Add "Frequency_Object"
      .Add "Frequency_PerfTime"
      .Add "Frequency_Sys100NS"
      .Add "Name"
      .Add "Timestamp_Object"
      .Add "Timestamp_PerfTime"
      .Add "Timestamp_Sys100NS"
      .Add "UserPC"
   End With
   Set PerfRawData_PerfProc_ThreadDetails_Costly = c
End Function
Private Function PerfRawData_PSched_PSchedFlow() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "AveragePacketsInNetCard"
      .Add "AveragePacketsInSequencer"
      .Add "AveragePacketsInShaper"
      .Add "BytesScheduled"
      .Add "BytesScheduledPerSec"
      .Add "BytesTransmitted"
      .Add "BytesTransmittedPerSec"
      .Add "Caption"
      .Add "Description"
      .Add "Frequency_Object"
      .Add "Frequency_PerfTime"
      .Add "Frequency_Sys100NS"
      .Add "MaximumPacketsInNetCard"
      .Add "MaxPacketsInSequencer"
      .Add "MaxPacketsInShaper"
      .Add "Name"
      .Add "NonConformingPacketsScheduled"
      .Add "NonConformingPacketsScheduledPerSec"
      .Add "NonConformingPacketsTransmitted"
      .Add "NonConformingPacketsTransmittedPerSec"
      .Add "PacketsDropped"
      .Add "PacketsDroppedPerSec"
      .Add "PacketsScheduled"
      .Add "PacketsScheduledPerSec"
      .Add "PacketsTransmitted"
      .Add "PacketsTransmittedPerSec"
      .Add "Timestamp_Object"
      .Add "Timestamp_PerfTime"
      .Add "Timestamp_Sys100NS"
   End With
   Set PerfRawData_PSched_PSchedFlow = c
End Function
Private Function PerfRawData_PSched_PSchedPipe() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "AveragePacketsInNetCard"
      .Add "AveragePacketsInSequencer"
      .Add "AveragePacketsInShaper"
      .Add "Caption"
      .Add "Description"
      .Add "FlowModsRejected"
      .Add "FlowsClosed"
      .Add "FlowsModified"
      .Add "FlowsOpened"
      .Add "FlowsRejected"
      .Add "Frequency_Object"
      .Add "Frequency_PerfTime"
      .Add "Frequency_Sys100NS"
      .Add "MaxPacketsInNetCard"
      .Add "MaxPacketsInSequencer"
      .Add "MaxPacketsInShaper"
      .Add "MaxSimultaneousFlows"
      .Add "Name"
      .Add "NonConformingPacketsScheduled"
      .Add "NonConformingPacketsScheduledPerSec"
      .Add "NonConformingPacketsTransmitted"
      .Add "NonConformingPacketsTransmittedPerSec"
      .Add "OutOfPackets"
      .Add "Timestamp_Object"
      .Add "Timestamp_PerfTime"
      .Add "Timestamp_Sys100NS"
   End With
   Set PerfRawData_PSched_PSchedPipe = c
End Function
Private Function PerfRawData_RemoteAccess_RASPort() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "AlignmentErrors"
      .Add "BufferOverrunErrors"
      .Add "BytesReceived"
      .Add "BytesReceivedPerSec"
      .Add "BytesTransmitted"
      .Add "BytesTransmittedPerSec"
      .Add "Caption"
      .Add "CRCErrors"
      .Add "Description"
      .Add "FramesReceived"
      .Add "FramesReceivedPerSec"
      .Add "FramesTransmitted"
      .Add "FramesTransmittedPerSec"
      .Add "Frequency_Object"
      .Add "Frequency_PerfTime"
      .Add "Frequency_Sys100NS"
      .Add "Name"
      .Add "PercentCompressionIn"
      .Add "PercentCompressionOut"
      .Add "SerialOverrunErrors"
      .Add "TimeoutErrors"
      .Add "Timestamp_Object"
      .Add "Timestamp_PerfTime"
      .Add "Timestamp_Sys100NS"
      .Add "TotalErrors"
      .Add "TotalErrorsPerSec"
   End With
   Set PerfRawData_RemoteAccess_RASPort = c
End Function
Private Function PerfRawData_RemoteAccess_RASTotal() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "AlignmentErrors"
      .Add "BufferOverrunErrors"
      .Add "BytesReceived"
      .Add "BytesReceivedPerSec"
      .Add "BytesTransmitted"
      .Add "BytesTransmittedPerSec"
      .Add "Caption"
      .Add "CRCErrors"
      .Add "Description"
      .Add "FramesReceived"
      .Add "FramesReceivedPerSec"
      .Add "FramesTransmitted"
      .Add "FramesTransmittedPerSec"
      .Add "Frequency_Object"
      .Add "Frequency_PerfTime"
      .Add "Frequency_Sys100NS"
      .Add "Name"
      .Add "PercentCompressionIn"
      .Add "PercentCompressionOut"
      .Add "SerialOverrunErrors"
      .Add "TimeoutErrors"
      .Add "Timestamp_Object"
      .Add "Timestamp_PerfTime"
      .Add "Timestamp_Sys100NS"
      .Add "TotalConnections"
      .Add "TotalErrors"
      .Add "TotalErrorsPerSec"
   End With
   Set PerfRawData_RemoteAccess_RASTotal = c
End Function
Private Function PerfRawData_RSVP_ACSRSVPInterfaces() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "AdmittedBandwidth"
      .Add "BlockedRESVs"
      .Add "Caption"
      .Add "Description"
      .Add "Frequency_Object"
      .Add "Frequency_PerfTime"
      .Add "Frequency_Sys100NS"
      .Add "GeneralFailures"
      .Add "MaximumAdmittedBandwidth"
      .Add "Name"
      .Add "NumberOfActiveFlows"
      .Add "NumberOfIncomingMessagesDropped"
      .Add "NumberOfOutGoingMessagesDropped"
      .Add "PATHERRMessagesReceived"
      .Add "PATHERRMessagesSent"
      .Add "PATHMessagesReceived"
      .Add "PATHMessagesSent"
      .Add "PATHStateBlockTimeouts"
      .Add "PATHTEARMessagesReceived"
      .Add "PATHTEARMessagesSent"
      .Add "PolicyControlFailures"
      .Add "ReceiveMessagesErrorsBigMessages"
      .Add "ReceiveMessagesErrorsNoMemory"
      .Add "ResourceControlFailures"
      .Add "RESVCONFIRMMessagesReceived"
      .Add "RESVCONFIRMMessagesSent"
      .Add "RESVERRMessagesReceived"
      .Add "RESVERRMessagesSent"
      .Add "RESVMessagesReceived"
      .Add "RESVMessagesSent"
      .Add "RESVStateBlockTimeouts"
      .Add "RESVTEARMessagesReceived"
      .Add "RESVTEARMessagesSent"
      .Add "SendMessagesErrorsBigMessages"
      .Add "SendMessagesErrorsNoMemory"
      .Add "SignalingBytesReceived"
      .Add "SignalingBytesSent"
      .Add "Timestamp_Object"
      .Add "Timestamp_PerfTime"
      .Add "Timestamp_Sys100NS"
   End With
   Set PerfRawData_RSVP_ACSRSVPInterfaces = c
End Function
Private Function PerfRawData_RSVP_ACSRSVPService() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "BytesInQoSNotifications"
      .Add "Caption"
      .Add "Description"
      .Add "FailedQoSRequests"
      .Add "FailedQoSSends"
      .Add "Frequency_Object"
      .Add "Frequency_PerfTime"
      .Add "Frequency_Sys100NS"
      .Add "Name"
      .Add "NetworkInterfaces"
      .Add "NetworkSockets"
      .Add "QoSEnabledReceivers"
      .Add "QoSEnabledSenders"
      .Add "QoSNotifications"
      .Add "QoSSockets"
      .Add "RSVPSessions"
      .Add "Timers"
      .Add "Timestamp_Object"
      .Add "Timestamp_PerfTime"
      .Add "Timestamp_Sys100NS"
   End With
   Set PerfRawData_RSVP_ACSRSVPService = c
End Function
Private Function PerfRawData_SMTPSVC_SMTPServer() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "AvgRecipientsPerMsgReceived"
      .Add "AvgRecipientsPerMsgReceived_Base"
      .Add "AvgRecipientsPerMsgSent"
      .Add "AvgRecipientsPerMsgSent_Base"
      .Add "AvgRetriesPerMsgDelivered"
      .Add "AvgRetriesPerMsgDelivered_Base"
      .Add "AvgRetriesPerMsgSent"
      .Add "AvgRetriesPerMsgSent_Base"
      .Add "BadmailedMessagesBadPickupFile"
      .Add "BadmailedMessagesGeneralFailure"
      .Add "BadmailedMessagesHopCountExceeded"
      .Add "BadmailedMessagesNDRofDSN"
      .Add "BadmailedMessagesNoRecipients"
      .Add "BadmailedMessagesTriggeredViaEvent"
      .Add "BytesReceivedPerSec"
      .Add "BytesReceivedTotal"
      .Add "BytesSentPerSec"
      .Add "BytesSentTotal"
      .Add "BytesTotal"
      .Add "BytesTotalPerSec"
      .Add "Caption"
      .Add "CatAddressLookupCompletions"
      .Add "CatAddressLookupCompletionsPerSec"
      .Add "CatAddressLookups"
      .Add "CatAddressLookupsNotFound"
      .Add "CatAddressLookupsPerSec"
      .Add "CatCategorizationsCompleted"
      .Add "CatCategorizationsCompletedPerSec"
      .Add "CatCategorizationsCompletedSuccessfully"
      .Add "CatCategorizationsFailedDSConnectionFailure"
      .Add "CatCategorizationsFailedDSLogonFailure"
      .Add "CatCategorizationsFailedNonRetryableError"
      .Add "CatCategorizationsFailedOutOfMemory"
      .Add "CatCategorizationsFailedRetryableError"
      .Add "CatCategorizationsFailedSinkRetryableError"
      .Add "CatCategorizationsInProgress"
      .Add "CategorizerQueueLength"
      .Add "CatLDAPBindFailures"
      .Add "CatLDAPBinds"
      .Add "CatLDAPConnectionFailures"
      .Add "CatLDAPConnections"
      .Add "CatLDAPConnectionsCurrentlyOpen"
      .Add "CatLDAPGeneralCompletionFailures"
      .Add "CatLDAPPagedSearches"
      .Add "CatLDAPPagedSearchesCompleted"
      .Add "CatLDAPPagedSearchFailures"
      .Add "CatLDAPSearchCompletionFailures"
      .Add "CatLDAPSearches"
      .Add "CatLDAPSearchesAbandoned"
      .Add "CatLDAPSearchesCompleted"
      .Add "CatLDAPSearchesCompletedPerSec"
      .Add "CatLDAPSearchesPendingCompletion"
      .Add "CatLDAPSearchesPerSec"
      .Add "CatLDAPSearchFailures"
      .Add "CatMailMsgDuplicateCollisions"
      .Add "CatMessagesAborted"
      .Add "CatMessagesBifurcated"
      .Add "CatMessagesCategorized"
      .Add "CatMessagesSubmitted"
      .Add "CatMessagesSubmittedPerSec"
      .Add "CatRecipientsAfterCategorization"
      .Add "CatRecipientsBeforeCategorization"
      .Add "CatRecipientsInCategorization"
      .Add "CatRecipientsNDRdAmbiguousAddress"
      .Add "CatRecipientsNDRdByCategorizer"
      .Add "CatRecipientsNDRdForwardingLoop"
      .Add "CatRecipientsNDRdIllegalAddress"
      .Add "CatRecipientsNDRdSinkRecipErrors"
      .Add "CatRecipientsNDRdUnresolved"
      .Add "CatSendersUnresolved"
      .Add "CatSendersWithAmbiguousAddresses"
      .Add "ConnectionErrorsPerSec"
      .Add "CurrentMessagesInLocalDelivery"
      .Add "Description"
      .Add "DirectoryDropsPerSec"
      .Add "DirectoryDropsTotal"
      .Add "DNSQueriesPerSec"
      .Add "DNSQueriesTotal"
      .Add "ETRNMessagesPerSec"
      .Add "ETRNMessagesTotal"
      .Add "Frequency_Object"
      .Add "Frequency_PerfTime"
      .Add "Frequency_Sys100NS"
      .Add "InboundConnectionsCurrent"
      .Add "InboundConnectionsTotal"
      .Add "LocalQueueLength"
      .Add "LocalRetryQueueLength"
      .Add "MessageBytesReceivedPerSec"
      .Add "MessageBytesReceivedTotal"
      .Add "MessageBytesSentPerSec"
      .Add "MessageBytesSentTotal"
      .Add "MessageBytesTotal"
      .Add "MessageBytesTotalPerSec"
      .Add "MessageDeliveryRetries"
      .Add "MessagesCurrentlyUndeliverable"
      .Add "MessagesDeliveredPerSec"
      .Add "MessagesDeliveredTotal"
      .Add "MessageSendRetries"
      .Add "MessagesPendingRouting"
      .Add "MessagesReceivedPerSec"
      .Add "MessagesReceivedTotal"
      .Add "MessagesRefusedForAddressObjects"
      .Add "MessagesRefusedForMailObjects"
      .Add "MessagesRefusedForSize"
      .Add "MessagesSentPerSec"
      .Add "MessagesSentTotal"
      .Add "Name"
      .Add "NDRsGenerated"
      .Add "NumberOfMailFilesOpen"
      .Add "NumberOfQueueFilesOpen"
      .Add "OutboundConnectionsCurrent"
      .Add "OutboundConnectionsRefused"
      .Add "OutboundConnectionsTotal"
      .Add "PercentRecipientsLocal"
      .Add "PercentRecipientsLocal_Base"
      .Add "PercentRecipientsRemote"
      .Add "PercentRecipientsRemote_Base"
      .Add "PickupDirectoryMessagesRetrievedPerSec"
      .Add "PickupDirectoryMessagesRetrievedTotal"
      .Add "RemoteQueueLength"
      .Add "RemoteRetryQueueLength"
      .Add "RoutingTableLookupsPerSec"
      .Add "RoutingTableLookupsTotal"
      .Add "Timestamp_Object"
      .Add "Timestamp_PerfTime"
      .Add "Timestamp_Sys100NS"
      .Add "TotalConnectionErrors"
      .Add "TotalDSNFailures"
      .Add "TotalMessagesSubmitted"
   End With
   Set PerfRawData_SMTPSVC_SMTPServer = c
End Function
Private Function PerfRawData_Spooler_PrintQueue() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "AddNetworkPrinterCalls"
      .Add "BytesPrintedPerSec"
      .Add "Caption"
      .Add "Description"
      .Add "EnumerateNetworkPrinterCalls"
      .Add "Frequency_Object"
      .Add "Frequency_PerfTime"
      .Add "Frequency_Sys100NS"
      .Add "JobErrors"
      .Add "Jobs"
      .Add "JobsSpooling"
      .Add "MaxJobsSpooling"
      .Add "MaxReferences"
      .Add "Name"
      .Add "NotReadyErrors"
      .Add "OutOfPaperErrors"
      .Add "References"
      .Add "Timestamp_Object"
      .Add "Timestamp_PerfTime"
      .Add "Timestamp_Sys100NS"
      .Add "TotalJobsPrinted"
      .Add "TotalPagesPrinted"
   End With
   Set PerfRawData_Spooler_PrintQueue = c
End Function
Private Function PerfRawData_TapiSrv_Telephony() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "ActiveLines"
      .Add "ActiveTelephones"
      .Add "Caption"
      .Add "ClientApps"
      .Add "CurrentIncomingCalls"
      .Add "CurrentOutgoingCalls"
      .Add "Description"
      .Add "Frequency_Object"
      .Add "Frequency_PerfTime"
      .Add "Frequency_Sys100NS"
      .Add "IncomingCallsPerSec"
      .Add "Lines"
      .Add "Name"
      .Add "OutgoingCallsPerSec"
      .Add "TelephoneDevices"
      .Add "Timestamp_Object"
      .Add "Timestamp_PerfTime"
      .Add "Timestamp_Sys100NS"
   End With
   Set PerfRawData_TapiSrv_Telephony = c
End Function
Private Function PerfRawData_Tcpip_ICMP() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Caption"
      .Add "Description"
      .Add "Frequency_Object"
      .Add "Frequency_PerfTime"
      .Add "Frequency_Sys100NS"
      .Add "MessagesOutboundErrors"
      .Add "MessagesPerSec"
      .Add "MessagesReceivedErrors"
      .Add "MessagesReceivedPerSec"
      .Add "MessagesSentPerSec"
      .Add "Name"
      .Add "ReceivedAddressMask"
      .Add "ReceivedAddressMaskReply"
      .Add "ReceivedDestUnreachable"
      .Add "ReceivedEchoPerSec"
      .Add "ReceivedEchoReplyPerSec"
      .Add "ReceivedParameterProblem"
      .Add "ReceivedRedirectPerSec"
      .Add "ReceivedSourceQuench"
      .Add "ReceivedTimeExceeded"
      .Add "ReceivedTimestampPerSec"
      .Add "ReceivedTimestampReplyPerSec"
      .Add "SentAddressMask"
      .Add "SentAddressMaskReply"
      .Add "SentDestinationUnreachable"
      .Add "SentEchoPerSec"
      .Add "SentEchoReplyPerSec"
      .Add "SentParameterProblem"
      .Add "SentRedirectPerSec"
      .Add "SentSourceQuench"
      .Add "SentTimeExceeded"
      .Add "SentTimestampPerSec"
      .Add "SentTimestampReplyPerSec"
      .Add "Timestamp_Object"
      .Add "Timestamp_PerfTime"
      .Add "Timestamp_Sys100NS"
   End With
   Set PerfRawData_Tcpip_ICMP = c
End Function
Private Function PerfRawData_Tcpip_IP() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Caption"
      .Add "DatagramsForwardedPerSec"
      .Add "DatagramsOutboundDiscarded"
      .Add "DatagramsOutboundNoRoute"
      .Add "DatagramsPerSec"
      .Add "DatagramsReceivedAddressErrors"
      .Add "DatagramsReceivedDeliveredPerSec"
      .Add "DatagramsReceivedDiscarded"
      .Add "DatagramsReceivedHeaderErrors"
      .Add "DatagramsReceivedPerSec"
      .Add "DatagramsReceivedUnknownProtocol"
      .Add "DatagramsSentPerSec"
      .Add "Description"
      .Add "FragmentationFailures"
      .Add "FragmentedDatagramsPerSec"
      .Add "FragmentReassemblyFailures"
      .Add "FragmentsCreatedPerSec"
      .Add "FragmentsReassembledPerSec"
      .Add "FragmentsReceivedPerSec"
      .Add "Frequency_Object"
      .Add "Frequency_PerfTime"
      .Add "Frequency_Sys100NS"
      .Add "Name"
      .Add "Timestamp_Object"
      .Add "Timestamp_PerfTime"
      .Add "Timestamp_Sys100NS"
   End With
   Set PerfRawData_Tcpip_IP = c
End Function
Private Function PerfRawData_Tcpip_NBTConnection() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "BytesReceivedPerSec"
      .Add "BytesSentPerSec"
      .Add "BytesTotalPerSec"
      .Add "Caption"
      .Add "Description"
      .Add "Frequency_Object"
      .Add "Frequency_PerfTime"
      .Add "Frequency_Sys100NS"
      .Add "Name"
      .Add "Timestamp_Object"
      .Add "Timestamp_PerfTime"
      .Add "Timestamp_Sys100NS"
   End With
   Set PerfRawData_Tcpip_NBTConnection = c
End Function
Private Function PerfRawData_PerfDisk_LogicalDisk() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "AvgDiskBytesPerRead"
      .Add "AvgDiskBytesPerRead_Base"
      .Add "AvgDiskBytesPerTransfer"
      .Add "AvgDiskBytesPerTransfer_Base"
      .Add "AvgDiskBytesPerWrite"
      .Add "AvgDiskBytesPerWrite_Base"
      .Add "AvgDiskQueueLength"
      .Add "AvgDiskReadQueueLength"
      .Add "AvgDiskSecPerRead"
      .Add "AvgDiskSecPerRead_Base"
      .Add "AvgDiskSecPerTransfer"
      .Add "AvgDiskSecPerTransfer_Base"
      .Add "AvgDiskSecPerWrite"
      .Add "AvgDiskSecPerWrite_Base"
      .Add "AvgDiskWriteQueueLength"
      .Add "Caption"
      .Add "CurrentDiskQueueLength"
      .Add "Description"
      .Add "DiskBytesPerSec"
      .Add "DiskReadBytesPerSec"
      .Add "DiskReadsPerSec"
      .Add "DiskTransfersPerSec"
      .Add "DiskWriteBytesPerSec"
      .Add "DiskWritesPerSec"
      .Add "FreeMegabytes"
      .Add "Frequency_Object"
      .Add "Frequency_PerfTime"
      .Add "Frequency_Sys100NS"
      .Add "Name"
      .Add "PercentDiskReadTime"
      .Add "PercentDiskReadTime_Base"
      .Add "PercentDiskTime"
      .Add "PercentDiskTime_Base"
      .Add "PercentDiskWriteTime"
      .Add "PercentDiskWriteTime_Base"
      .Add "PercentFreeSpace"
      .Add "PercentFreeSpace_Base"
      .Add "PercentIdleTime"
      .Add "PercentIdleTime_Base"
      .Add "SplitIOPerSec"
      .Add "Timestamp_Object"
      .Add "Timestamp_PerfTime"
      .Add "Timestamp_Sys100NS"
   End With
   Set PerfRawData_PerfDisk_LogicalDisk = c
End Function
Private Function PerfRawData_NTFSDRV_SMTPNTFSStoreDriver() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Caption"
      .Add "Description"
      .Add "Frequency_Object"
      .Add "Frequency_PerfTime"
      .Add "Frequency_Sys100NS"
      .Add "MessagesAllocated"
      .Add "MessagesDeleted"
      .Add "MessagesEnumerated"
      .Add "MessagesInTheQueueDirectory"
      .Add "Name"
      .Add "OpenMessageBodies"
      .Add "OpenMessageStreams"
      .Add "Timestamp_Object"
      .Add "Timestamp_PerfTime"
      .Add "Timestamp_Sys100NS"
   End With
   Set PerfRawData_NTFSDRV_SMTPNTFSStoreDriver = c
End Function
Private Function PerfRawData_InetInfo_InternetInformationServicesGlobal() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "ActiveFlushedEntries"
      .Add "BLOBCacheFlushes"
      .Add "BLOBCacheHits"
      .Add "BLOBCacheHitsPercent"
      .Add "BLOBCacheHitsPercent_Base"
      .Add "BLOBCacheMisses"
      .Add "Caption"
      .Add "CurrentBLOBsCached"
      .Add "CurrentBlockedAsyncIORequests"
      .Add "CurrentFileCacheMemoryUsage"
      .Add "CurrentFilesCached"
      .Add "CurrentURIsCached"
      .Add "Description"
      .Add "FileCacheFlushes"
      .Add "FileCacheHits"
      .Add "FileCacheHitsPercent"
      .Add "FileCacheHitsPercent_Base"
      .Add "FileCacheMisses"
      .Add "Frequency_Object"
      .Add "Frequency_PerfTime"
      .Add "Frequency_Sys100NS"
      .Add "MaximumFileCacheMemoryUsage"
      .Add "MeasuredAsyncIOBandwidthUsage"
      .Add "Name"
      .Add "Timestamp_Object"
      .Add "Timestamp_PerfTime"
      .Add "Timestamp_Sys100NS"
      .Add "TotalAllowedAsyncIORequests"
      .Add "TotalBLOBsCached"
      .Add "TotalBlockedAsyncIORequests"
      .Add "TotalFilesCached"
      .Add "TotalFlushedBLOBs"
      .Add "TotalFlushedFiles"
      .Add "TotalFlushedURIs"
      .Add "TotalRejectedAsyncIORequests"
      .Add "TotalURIsCached"
      .Add "URICacheFlushes"
      .Add "URICacheHits"
      .Add "URICacheHitsPercent"
      .Add "URICacheHitsPercent_Base"
      .Add "URICacheMisses"
   End With
   Set PerfRawData_InetInfo_InternetInformationServicesGlobal = c
End Function
Private Function PerfRawData_MSDTC_DistributedTransactionCoordinator() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "AbortedTransactions"
      .Add "AbortedTransactionsPerSec"
      .Add "ActiveTransactions"
      .Add "ActiveTransactionsMaximum"
      .Add "Caption"
      .Add "CommittedTransactions"
      .Add "CommittedTransactionsPerSec"
      .Add "Description"
      .Add "ForceAbortedTransactions"
      .Add "ForceCommittedTransactions"
      .Add "Frequency_Object"
      .Add "Frequency_PerfTime"
      .Add "Frequency_Sys100NS"
      .Add "InDoubtTransactions"
      .Add "Name"
      .Add "ResponseTimeAverage"
      .Add "ResponseTimeMaximum"
      .Add "ResponseTimeMinimum"
      .Add "Timestamp_Object"
      .Add "Timestamp_PerfTime"
      .Add "Timestamp_Sys100NS"
      .Add "TransactionsPerSec"
   End With
   Set PerfRawData_MSDTC_DistributedTransactionCoordinator = c
End Function
Private Function PerfRawData_ISAPISearch_HttpIndexingService() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "ActiveQueries"
      .Add "CacheItems"
      .Add "Caption"
      .Add "CurrentRequestsQueued"
      .Add "Description"
      .Add "Frequency_Object"
      .Add "Frequency_PerfTime"
      .Add "Frequency_Sys100NS"
      .Add "Name"
      .Add "PercentCacheHits"
      .Add "PercentCacheHits_Base"
      .Add "PercentCacheMisses"
      .Add "PercentCacheMisses_Base"
      .Add "QueriesPerMinute"
      .Add "Timestamp_Object"
      .Add "Timestamp_PerfTime"
      .Add "Timestamp_Sys100NS"
      .Add "TotalQueries"
      .Add "TotalRequestsRejected"
   End With
   Set PerfRawData_ISAPISearch_HttpIndexingService = c
End Function
Private Function PerfRawData_Tcpip_NetworkInterface() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "BytesReceivedPerSec"
      .Add "BytesSentPerSec"
      .Add "BytesTotalPerSec"
      .Add "Caption"
      .Add "CurrentBandwidth"
      .Add "Description"
      .Add "Frequency_Object"
      .Add "Frequency_PerfTime"
      .Add "Frequency_Sys100NS"
      .Add "Name"
      .Add "OutputQueueLength"
      .Add "PacketsOutboundDiscarded"
      .Add "PacketsOutboundErrors"
      .Add "PacketsPerSec"
      .Add "PacketsReceivedDiscarded"
      .Add "PacketsReceivedErrors"
      .Add "PacketsReceivedNonUnicastPerSec"
      .Add "PacketsReceivedPerSec"
      .Add "PacketsReceivedUnicastPerSec"
      .Add "PacketsReceivedUnknown"
      .Add "PacketsSentNonUnicastPerSec"
      .Add "PacketsSentPerSec"
      .Add "PacketsSentUnicastPerSec"
      .Add "Timestamp_Object"
      .Add "Timestamp_PerfTime"
      .Add "Timestamp_Sys100NS"
   End With
   Set PerfRawData_Tcpip_NetworkInterface = c
End Function
Private Function PerfRawData_Tcpip_TCP() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Caption"
      .Add "ConnectionFailures"
      .Add "ConnectionsActive"
      .Add "ConnectionsEstablished"
      .Add "ConnectionsPassive"
      .Add "ConnectionsReset"
      .Add "Description"
      .Add "Frequency_Object"
      .Add "Frequency_PerfTime"
      .Add "Frequency_Sys100NS"
      .Add "Name"
      .Add "SegmentsPerSec"
      .Add "SegmentsReceivedPerSec"
      .Add "SegmentsRetransmittedPerSec"
      .Add "SegmentsSentPerSec"
      .Add "Timestamp_Object"
      .Add "Timestamp_PerfTime"
      .Add "Timestamp_Sys100NS"
   End With
   Set PerfRawData_Tcpip_TCP = c
End Function
Private Function PerfRawData_Tcpip_UDP() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Caption"
      .Add "DatagramsNoPortPerSec"
      .Add "DatagramsPerSec"
      .Add "DatagramsReceivedErrors"
      .Add "DatagramsReceivedPerSec"
      .Add "DatagramsSentPerSec"
      .Add "Description"
      .Add "Frequency_Object"
      .Add "Frequency_PerfTime"
      .Add "Frequency_Sys100NS"
      .Add "Name"
      .Add "Timestamp_Object"
      .Add "Timestamp_PerfTime"
      .Add "Timestamp_Sys100NS"
   End With
   Set PerfRawData_Tcpip_UDP = c
End Function
Private Function PerfRawData_TermService_TerminalServices() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "ActiveSessions"
      .Add "Caption"
      .Add "Description"
      .Add "Frequency_Object"
      .Add "Frequency_PerfTime"
      .Add "Frequency_Sys100NS"
      .Add "InactiveSessions"
      .Add "Name"
      .Add "Timestamp_Object"
      .Add "Timestamp_PerfTime"
      .Add "Timestamp_Sys100NS"
      .Add "TotalSessions"
   End With
   Set PerfRawData_TermService_TerminalServices = c
End Function
Private Function PerfRawData_TermService_TerminalServicesSession() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Caption"
      .Add "Description"
      .Add "Frequency_Object"
      .Add "Frequency_PerfTime"
      .Add "Frequency_Sys100NS"
      .Add "HandleCount"
      .Add "InputAsyncFrameError"
      .Add "InputAsyncOverflow"
      .Add "InputAsyncOverrun"
      .Add "InputAsyncParityError"
      .Add "InputBytes"
      .Add "InputCompressedBytes"
      .Add "InputCompressFlushes"
      .Add "InputCompressionRatio"
      .Add "InputErrors"
      .Add "InputFrames"
      .Add "InputTimeouts"
      .Add "InputTransportErrors"
      .Add "InputWaitForOutBuf"
      .Add "InputWdBytes"
      .Add "InputWdFrames"
      .Add "Name"
      .Add "OutputAsyncFrameError"
      .Add "OutputAsyncOverflow"
      .Add "OutputAsyncOverrun"
      .Add "OutputAsyncParityError"
      .Add "OutputBytes"
      .Add "OutputCompressedBytes"
      .Add "OutputCompressFlushes"
      .Add "OutputCompressionRatio"
      .Add "OutputErrors"
      .Add "OutputFrames"
      .Add "OutputTimeouts"
      .Add "OutputTransportErrors"
      .Add "OutputWaitForOutBuf"
      .Add "OutputWdBytes"
      .Add "OutputWdFrames"
      .Add "PageFaultsPerSec"
      .Add "PageFileBytes"
      .Add "PageFileBytesPeak"
      .Add "PercentPrivilegedTime"
      .Add "PercentProcessorTime"
      .Add "PercentUserTime"
      .Add "PoolNonpagedBytes"
      .Add "PoolPagedBytes"
      .Add "PrivateBytes"
      .Add "ProtocolBitmapCacheHitRatio"
      .Add "ProtocolBitmapCacheHits"
      .Add "ProtocolBitmapCacheReads"
      .Add "ProtocolBrushCacheHitRatio"
      .Add "ProtocolBrushCacheHits"
      .Add "ProtocolBrushCacheReads"
      .Add "ProtocolGlyphCacheHitRatio"
      .Add "ProtocolGlyphCacheHits"
      .Add "ProtocolGlyphCacheReads"
      .Add "ProtocolSaveScreenBitmapCacheHitRatio"
      .Add "ProtocolSaveScreenBitmapCacheHits"
      .Add "ProtocolSaveScreenBitmapCacheReads"
      .Add "ThreadCount"
      .Add "Timestamp_Object"
      .Add "Timestamp_PerfTime"
      .Add "Timestamp_Sys100NS"
      .Add "TotalAsyncFrameError"
      .Add "TotalAsyncOverflow"
      .Add "TotalAsyncOverrun"
      .Add "TotalAsyncParityError"
      .Add "TotalBytes"
      .Add "TotalCompressedBytes"
      .Add "TotalCompressFlushes"
      .Add "TotalCompressionRatio"
      .Add "TotalErrors"
      .Add "TotalFrames"
      .Add "TotalProtocolCacheHitRatio"
      .Add "TotalProtocolCacheHits"
      .Add "TotalProtocolCacheReads"
      .Add "TotalTimeouts"
      .Add "TotalTransportErrors"
      .Add "TotalWaitForOutBuf"
      .Add "TotalWdBytes"
      .Add "TotalWdFrames"
      .Add "VirtualBytes"
      .Add "VirtualBytesPeak"
      .Add "WorkingSet"
      .Add "WorkingSetPeak"
   End With
   Set PerfRawData_TermService_TerminalServicesSession = c
End Function
Private Function PerfRawData_W3SVC_WebService() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "AnonymousUsersPerSec"
      .Add "BytesReceivedPerSec"
      .Add "BytesSentPerSec"
      .Add "BytesTotalPerSec"
      .Add "Caption"
      .Add "CGIRequestsPerSec"
      .Add "ConnectionAttemptsPerSec"
      .Add "CopyRequestsPerSec"
      .Add "CurrentAnonymousUsers"
      .Add "CurrentBlockedAsyncIORequests"
      .Add "CurrentCGIRequests"
      .Add "CurrentConnections"
      .Add "CurrentISAPIExtensionRequests"
      .Add "CurrentNonAnonymousUsers"
      .Add "DeleteRequestsPerSec"
      .Add "Description"
      .Add "FilesPerSec"
      .Add "FilesReceivedPerSec"
      .Add "FilesSentPerSec"
      .Add "Frequency_Object"
      .Add "Frequency_PerfTime"
      .Add "Frequency_Sys100NS"
      .Add "GetRequestsPerSec"
      .Add "HeadRequestsPerSec"
      .Add "ISAPIExtensionRequestsPerSec"
      .Add "LockedErrorsPerSec"
      .Add "LockRequestsPerSec"
      .Add "LogonAttemptsPerSec"
      .Add "MaximumAnonymousUsers"
      .Add "MaximumCGIRequests"
      .Add "MaximumConnections"
      .Add "MaximumISAPIExtensionRequests"
      .Add "MaximumNonAnonymousUsers"
      .Add "MeasuredAsyncIOBandwidthUsage"
      .Add "MkcolRequestsPerSec"
      .Add "MoveRequestsPerSec"
      .Add "Name"
      .Add "NonAnonymousUsersPerSec"
      .Add "NotFoundErrorsPerSec"
      .Add "OptionsRequestsPerSec"
      .Add "OtherRequestMethodsPerSec"
      .Add "PostRequestsPerSec"
      .Add "PropfindRequestsPerSec"
      .Add "ProppatchRequestsPerSec"
      .Add "PutRequestsPerSec"
      .Add "SearchRequestsPerSec"
      .Add "ServiceUptime"
      .Add "Timestamp_Object"
      .Add "Timestamp_PerfTime"
      .Add "Timestamp_Sys100NS"
      .Add "TotalAllowedAsyncIORequests"
      .Add "TotalAnonymousUsers"
      .Add "TotalBlockedAsyncIORequests"
      .Add "TotalCGIRequests"
      .Add "TotalConnectionAttemptsAllInstances"
      .Add "TotalCopyRequests"
      .Add "TotalDeleteRequests"
      .Add "TotalFilesReceived"
      .Add "TotalFilesSent"
      .Add "TotalFilesTransferred"
      .Add "TotalGetRequests"
      .Add "TotalHeadRequests"
      .Add "TotalISAPIExtensionRequests"
      .Add "TotalLockedErrors"
      .Add "TotalLockRequests"
      .Add "TotalLogonAttempts"
      .Add "TotalMethodRequests"
      .Add "TotalMethodRequestsPerSec"
      .Add "TotalMkcolRequests"
      .Add "TotalMoveRequests"
      .Add "TotalNonAnonymousUsers"
      .Add "TotalNotFoundErrors"
      .Add "TotalOptionsRequests"
      .Add "TotalOtherRequestMethods"
      .Add "TotalPostRequests"
      .Add "TotalPropfindRequests"
      .Add "TotalProppatchRequests"
      .Add "TotalPutRequests"
      .Add "TotalRejectedAsyncIORequests"
      .Add "TotalSearchRequests"
      .Add "TotalTraceRequests"
      .Add "TotalUnlockRequests"
      .Add "TraceRequestsPerSec"
      .Add "UnlockRequestsPerSec"
   End With
   Set PerfRawData_W3SVC_WebService = c
End Function
Private Function PhysicalMedia() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Capacity"
      .Add "Caption"
      .Add "CleanerMedia"
      .Add "CreationClassName"
      .Add "Desctription"
      .Add "HotSwappable"
      .Add "InstallDate"
      .Add "Manufacturer"
      .Add "MediaDescription"
      .Add "MediaType"
      .Add "Model"
      .Add "Name"
      .Add "OtherIdentifyingInfo"
      .Add "PartNumber"
      .Add "PoweredOn"
      .Add "Removable"
      .Add "Replacable"
      .Add "SerialNumber = NULL"
      .Add "SKU"
      .Add "Status"
      .Add "Tag = NULL"
      .Add "Version"
      .Add "WriteProtectOn"
   End With
   Set PhysicalMedia = c
End Function
Private Function PhysicalMemory() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "BankLabel"
      .Add "Capacity"
      .Add "Caption"
      .Add "CreationClassName"
      .Add "DataWidth"
      .Add "Description"
      .Add "DeviceLocator"
      .Add "FormFactor"
      .Add "HotSwappable"
      .Add "InstallDate"
      .Add "InterleaveDataDepth"
      .Add "InterleavePosition"
      .Add "Manufacturer"
      .Add "MemoryType"
      .Add "Model"
      .Add "Name"
      .Add "OtherIdentifyingInfo"
      .Add "PartNumber"
      .Add "PositionInRow"
      .Add "PoweredOn"
      .Add "Removable"
      .Add "Replaceable"
      .Add "SerialNumber"
      .Add "SKU"
      .Add "Speed"
      .Add "Status"
      .Add "Tag"
      .Add "TotalWidth"
      .Add "TypeDetail"
      .Add "Version"
   End With
   Set PhysicalMemory = c
End Function
Private Function PhysicalMemoryArray() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Caption"
      .Add "CreationClassName"
      .Add "Depth"
      .Add "Description"
      .Add "Height"
      .Add "HotSwappable"
      .Add "InstallDate"
      .Add "Location"
      .Add "Manufacturer"
      .Add "MaxCapacity"
      .Add "MemoryDevices"
      .Add "MemoryErrorCorrection"
      .Add "Model"
      .Add "Name"
      .Add "OtherIdentifyingInfo"
      .Add "PartNumber"
      .Add "PoweredOn"
      .Add "Removable"
      .Add "Replaceable"
      .Add "SerialNumber"
      .Add "SKU"
      .Add "Status"
      .Add "Tag"
      .Add "Use"
      .Add "Version"
      .Add "Weight"
      .Add "Width"
   End With
   Set PhysicalMemoryArray = c
End Function
Private Function PingStatus() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Address"
      .Add "BufferSize = 32"
      .Add "NoFragmentation = FALSE"
      .Add "PrimaryAddressResolutionStatus"
      .Add "ProtocolAddress = """
      .Add "ProtocolAddressResolved = """
      .Add "RecordRoute = 0"
      .Add "ReplyInconsistency"
      .Add "ReplySize"
      .Add "ResolveAddressNames = FALSE"
      .Add "ResponseTime"
      .Add "ResponseTimeToLive"
      .Add "RouteRecord[]"
      .Add "RouteRecordResolved[]"
      .Add "SourceRoute = """
      .Add "SourceRouteType = 0"
      .Add "StatusCode"
      .Add "Timeout = 4000"
      .Add "TimeStampRecord[]"
      .Add "TimeStampRecordAddress[]"
      .Add "TimeStampRecordAddressResolved[]"
      .Add "TimeStampRoute = 0"
      .Add "TimeToLive = 80"
      .Add "TypeofService = 0"
   End With
   Set PingStatus = c
End Function
Private Function PnPEntity() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Availability"
      .Add "Caption"
      .Add "ClassGuid"
      .Add "ConfigManagerErrorCode"
      .Add "ConfigManagerUserConfig"
      .Add "CreationClassName"
      .Add "Description"
      .Add "DeviceID"
      .Add "ErrorCleared"
      .Add "ErrorDescription"
      .Add "InstallDate"
      .Add "LastErrorCode"
      .Add "Manufacturer"
      .Add "Name"
      .Add "PNPDeviceID"
      .Add "PowerManagementCapabilities[]"
      .Add "PowerManagementSupported"
      .Add "Service"
      .Add "Status"
      .Add "StatusInfo"
      .Add "SystemCreationClassName"
      .Add "SystemName"
   End With
   Set PnPEntity = c
End Function
Private Function PnPSignedDriver() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "ClassGuid"
      .Add "CompatID"
      .Add "Description"
      .Add "DeviceClass"
      .Add "DeviceID"
      .Add "DeviceName"
      .Add "DevLoader"
      .Add "DriverDate"
      .Add "DriverName"
      .Add "DriverVersion"
      .Add "FriendlyName"
      .Add "HardWareID"
      .Add "InfName"
      .Add "InstallDate"
      .Add "IsSigned"
      .Add "Location"
      .Add "Manufacturer"
      .Add "Name"
      .Add "PDO"
      .Add "ProviderName"
      .Add "Signer"
      .Add "Started"
      .Add "StartMode"
      .Add "Status"
      .Add "SystemCreationClassName"
      .Add "SystemName"
   End With
   Set PnPSignedDriver = c
End Function
Private Function PointingDevice() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Availability"
      .Add "Caption"
      .Add "ConfigManagerErrorCode"
      .Add "ConfigManagerUserConfig"
      .Add "CreationClassName"
      .Add "Description"
      .Add "DeviceID"
      .Add "DeviceInterface"
      .Add "DoubleSpeedThreshold"
      .Add "ErrorCleared"
      .Add "ErrorDescription"
      .Add "Handedness"
      .Add "HardwareType"
      .Add "InfFileName"
      .Add "InfSection"
      .Add "InstallDate"
      .Add "IsLocked"
      .Add "LastErrorCode"
      .Add "Manufacturer"
      .Add "Name"
      .Add "NumberOfButtons"
      .Add "PNPDeviceID"
      .Add "PointingType"
      .Add "PowerManagementCapabilities[]"
      .Add "PowerManagementSupported"
      .Add "QuadSpeedThreshold"
      .Add "Resolution"
      .Add "SampleRate"
      .Add "Status"
      .Add "StatusInfo"
      .Add "Synch"
      .Add "SystemCreationClassName"
      .Add "SystemName"
   End With
   Set PointingDevice = c
End Function
Private Function PortableBattery() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Availability"
      .Add "BatteryRechargeTime"
      .Add "BatteryStatus"
      .Add "CapacityMultiplier"
      .Add "Caption"
      .Add "Chemistry"
      .Add "ConfigManagerErrorCode"
      .Add "ConfigManagerUserConfig"
      .Add "CreationClassName"
      .Add "Description"
      .Add "DesignCapacity"
      .Add "DesignVoltage"
      .Add "DeviceID"
      .Add "ErrorCleared"
      .Add "ErrorDescription"
      .Add "EstimatedChargeRemaining"
      .Add "EstimatedRunTime"
      .Add "ExpectedBatteryLife"
      .Add "ExpectedLife"
      .Add "FullChargeCapacity"
      .Add "InstallDate"
      .Add "LastErrorCode"
      .Add "Location"
      .Add "ManufactureDate"
      .Add "Manufacturer"
      .Add "MaxBatteryError"
      .Add "MaxRechargeTime"
      .Add "Name"
      .Add "PNPDeviceID"
      .Add "PowerManagementCapabilities[]"
      .Add "PowerManagementSupported"
      .Add "SmartBatteryVersion"
      .Add "Status"
      .Add "StatusInfo"
      .Add "SystemCreationClassName"
      .Add "SystemName"
      .Add "TimeOnBattery"
      .Add "TimeToFullCharge"
   End With
   Set PortableBattery = c
End Function
Private Function PortConnector() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Caption"
      .Add "ConnectorPinout"
      .Add "ConnectorType[]"
      .Add "CreationClassName"
      .Add "Description"
      .Add "ExternalReferenceDesignator"
      .Add "InstallDate"
      .Add "InternalReferenceDesignator"
      .Add "Manufacturer"
      .Add "Model"
      .Add "Name"
      .Add "OtherIdentifyingInfo"
      .Add "PartNumber"
      .Add "PortType"
      .Add "PoweredOn"
      .Add "SerialNumber"
      .Add "SKU"
      .Add "Status"
      .Add "Tag"
      .Add "Version"
   End With
   Set PortConnector = c
End Function
Private Function PortResource() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Alias"
      .Add "Caption"
      .Add "CreationClassName"
      .Add "CSCreationClassName"
      .Add "CSName"
      .Add "Description"
      .Add "EndingAddress"
      .Add "InstallDate"
      .Add "Name"
      .Add "StartingAddress"
      .Add "Status"
   End With
   Set PortResource = c
End Function
Private Function POTSModem() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "AnswerMode"
      .Add "AttachedTo"
      .Add "Availability"
      .Add "BlindOff"
      .Add "BlindOn"
      .Add "Caption"
      .Add "CompatibilityFlags"
      .Add "CompressionInfo"
      .Add "CompressionOff"
      .Add "CompressionOn"
      .Add "ConfigManagerErrorCode"
      .Add "ConfigManagerUserConfig"
      .Add "ConfigurationDialog"
      .Add "CountriesSupported[]"
      .Add "CountrySelected"
      .Add "CreationClassName"
      .Add "CurrentPasswords[]"
      .Add "DCB[]"
      .Add "Default[]"
      .Add "Description"
      .Add "DeviceID"
      .Add "DeviceLoader"
      .Add "DeviceType"
      .Add "DialType"
      .Add "DriverDate"
      .Add "ErrorCleared"
      .Add "ErrorControlForced"
      .Add "ErrorControlInfo"
      .Add "ErrorControlOff"
      .Add "ErrorControlOn"
      .Add "ErrorDescription"
      .Add "FlowControlHard"
      .Add "FlowControlOff"
      .Add "FlowControlSoft"
      .Add "InactivityScale"
      .Add "InactivityTimeout"
      .Add "Index"
      .Add "InstallDate"
      .Add "LastErrorCode"
      .Add "MaxBaudRateToPhone"
      .Add "MaxBaudRateToSerialPort"
      .Add "MaxNumberOfPasswords"
      .Add "Model"
      .Add "ModemInfPath"
      .Add "ModemInfSection"
      .Add "ModulationBell"
      .Add "ModulationCCITT"
      .Add "ModulationScheme"
      .Add "Name"
      .Add "PNPDeviceID"
      .Add "PortSubClass"
      .Add "PowerManagementCapabilities[]"
      .Add "PowerManagementSupported"
      .Add "Prefix"
      .Add "Properties[]"
      .Add "ProviderName"
      .Add "Pulse"
      .Add "Reset"
      .Add "ResponsesKeyName"
      .Add "RingsBeforeAnswer"
      .Add "SpeakerModeDial"
      .Add "SpeakerModeOff"
      .Add "SpeakerModeOn"
      .Add "SpeakerModeSetup"
      .Add "SpeakerVolumeHigh"
      .Add "SpeakerVolumeInfo"
      .Add "SpeakerVolumeLow"
      .Add "SpeakerVolumeMed"
      .Add "Status"
      .Add "StatusInfo"
      .Add "StringFormat"
      .Add "SupportsCallback"
      .Add "SupportsSynchronousConnect"
      .Add "SystemCreationClassName"
      .Add "SystemName"
      .Add "Terminator"
      .Add "TimeOfLastReset"
      .Add "Tone"
      .Add "VoiceSwitchFeature"
   End With
   Set POTSModem = c
End Function
Private Function PowerManagementEvent() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "EventType"
      .Add "OEMEventCode"
   End With
   Set PowerManagementEvent = c
End Function
Private Function Printer() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Attributes"
      .Add "Availability"
      .Add "AvailableJobSheets[]"
      .Add "AveragePagesPerMinute"
      .Add "Capabilities[]"
      .Add "CapabilityDescriptions[]"
      .Add "Caption"
      .Add "CharSetsSupported[]"
      .Add "Comment"
      .Add "ConfigManagerErrorCode"
      .Add "ConfigManagerUserConfig"
      .Add "CreationClassName"
      .Add "CurrentCapabilities[]"
      .Add "CurrentCharSet"
      .Add "CurrentLanguage"
      .Add "CurrentMimeType"
      .Add "CurrentNaturalLanguage"
      .Add "CurrentPaperType"
      .Add "Default"
      .Add "DefaultCapabilities[]"
      .Add "DefaultCopies"
      .Add "DefaultLanguage"
      .Add "DefaultMimeType"
      .Add "DefaultNumberUp"
      .Add "DefaultPaperType"
      .Add "DefaultPriority"
      .Add "Description"
      .Add "DetectedErrorState"
      .Add "DeviceID"
      .Add "Direct"
      .Add "DoCompleteFirst"
      .Add "DriverName"
      .Add "EnableBIDI"
      .Add "EnableDevQueryPrint"
      .Add "ErrorCleared"
      .Add "ErrorDescription"
      .Add "ErrorInformation[]"
      .Add "ExtendedDetectedErrorState"
      .Add "ExtendedPrinterStatus"
      .Add "Hidden"
      .Add "HorizontalResolution"
      .Add "InstallDate"
      .Add "JobCountSinceLastReset"
      .Add "KeepPrintedJobs"
      .Add "LanguagesSupported[]"
      .Add "LastErrorCode"
      .Add "Local"
      .Add "Location"
      .Add "MarkingTechnology"
      .Add "MaxCopies"
      .Add "MaxNumberUp"
      .Add "MaxSizeSupported"
      .Add "MimeTypesSupported[]"
      .Add "Name"
      .Add "NaturalLanguagesSupported[]"
      .Add "Network"
      .Add "PaperSizesSupported[]"
      .Add "PaperTypesAvailable[]"
      .Add "Parameters"
      .Add "PNPDeviceID"
      .Add "PortName"
      .Add "PowerManagementCapabilities[]"
      .Add "PowerManagementSupported"
      .Add "PrinterPaperNames[]"
      .Add "PrinterState"
      .Add "PrinterStatus"
      .Add "PrintJobDataType"
      .Add "PrintProcessor"
      .Add "Priority"
      .Add "Published"
      .Add "Queued"
      .Add "RawOnly"
      .Add "SeparatorFile"
      .Add "ServerName"
      .Add "Shared"
      .Add "ShareName"
      .Add "SpoolEnabled"
      .Add "StartTime"
      .Add "Status"
      .Add "StatusInfo"
      .Add "SystemCreationClassName"
      .Add "SystemName"
      .Add "TimeOfLastReset"
      .Add "UntilTime"
      .Add "VerticalResolution"
      .Add "WorkOffline"
   End With
   Set Printer = c
End Function
Private Function PrinterConfiguration() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "BitsPerPel"
      .Add "Caption"
      .Add "Collate"
      .Add "Color"
      .Add "Copies"
      .Add "Description"
      .Add "DeviceName"
      .Add "DisplayFlags"
      .Add "DisplayFrequency"
      .Add "DitherType"
      .Add "DriverVersion"
      .Add "Duplex"
      .Add "FormName"
      .Add "HorizontalResolution"
      .Add "ICMIntent"
      .Add "ICMMethod"
      .Add "LogPixels"
      .Add "MediaType"
      .Add "Name"
      .Add "Orientation"
      .Add "PaperLength"
      .Add "PaperSize"
      .Add "PaperWidth"
      .Add "PelsHeight"
      .Add "PelsWidth"
      .Add "PrintQuality"
      .Add "Scale"
      .Add "SettingID"
      .Add "SpecificationVersion"
      .Add "TTOption"
      .Add "VerticalResolution"
      .Add "XResolution"
      .Add "YResolution"
   End With
   Set PrinterConfiguration = c
End Function
Private Function PrinterDriver() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Caption"
      .Add "ConfigFile"
      .Add "CreationClassName"
      .Add "DataFile"
      .Add "DefaultDataType"
      .Add "DependentFiles[]"
      .Add "Description"
      .Add "DriverPath"
      .Add "FilePath"
      .Add "HelpFile"
      .Add "InfName"
      .Add "InstallDate"
      .Add "MonitorName"
      .Add "Name"
      .Add "OEMUrl"
      .Add "Started"
      .Add "StartMode"
      .Add "Status"
      .Add "SupportedPlatform"
      .Add "SystemCreationClassName"
      .Add "SystemName"
      .Add "Version"
   End With
   Set PrinterDriver = c
End Function
Private Function PrintJob() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Caption"
      .Add "DataType"
      .Add "Description"
      .Add "Document"
      .Add "DriverName"
      .Add "ElapsedTime"
      .Add "HostPrintQueue"
      .Add "InstallDate"
      .Add "JobId"
      .Add "JobStatus"
      .Add "Name"
      .Add "Notify"
      .Add "Owner"
      .Add "PagesPrinted"
      .Add "Parameters"
      .Add "PrintProcessor"
      .Add "Priority"
      .Add "Size"
      .Add "StartTime"
      .Add "Status"
      .Add "StatusMask"
      .Add "TimeSubmitted"
      .Add "TotalPages"
      .Add "UntilTime"
   End With
   Set PrintJob = c
End Function
Private Function PrivilegesStatus() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Description"
      .Add "Operation"
      .Add "ParameterInfo"
      .Add "PrivilegesNotHeld[]"
      .Add "PrivilegesRequired[]"
      .Add "ProviderName"
      .Add "StatusCode"
   End With
   Set PrivilegesStatus = c
End Function

Private Function ProcessStartTrace() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "PageDirectoryBase"
      .Add "ParentProcessName"
      .Add "ProcessID"
      .Add "ProcessName"
      .Add "SECURITY_DESCRIPTOR[]"
      .Add "SessionID"
      .Add "Sid[]"
      .Add "TIME_CREATED"
   End With
   Set ProcessStartTrace = c
End Function
Private Function ProcessStartup() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "CreateFlags"
      .Add "EnvironmentVariables[]"
      .Add "ErrorMode"
      .Add "FillAttribute"
      .Add "PriorityClass"
      .Add "ShowWindow"
      .Add "Title"
      .Add "WinstationDesktop"
      .Add "X"
      .Add "XCountChars"
      .Add "XSize"
      .Add "Y"
      .Add "YCountChars"
      .Add "YSize"
   End With
   Set ProcessStartup = c
End Function

Private Function ProcessStopTrace() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "PageDirectoryBase"
      .Add "ParentProcessID"
      .Add "ProcessID"
      .Add "ProcessName"
      .Add "SECURITY_DESCRIPTOR[]"
      .Add "SessionID"
      .Add "Sid[]"
      .Add "TIME_CREATED"
   End With
   Set ProcessStopTrace = c
End Function
Private Function ProcessTrace() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "PageDirectoryBase"
      .Add "ParentProcessID"
      .Add "ProcessID"
      .Add "ProcessName"
      .Add "SECURITY_DESCRIPTOR[]"
      .Add "SessionID"
      .Add "Sid[]"
      .Add "TIME_CREATED"
   End With
   Set ProcessTrace = c
End Function

Private Function Product() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Caption"
      .Add "Description"
      .Add "IdentifyingNumber"
      .Add "InstallDate"
      .Add "InstallDate2"
      .Add "InstallLocation"
      .Add "InstallState"
      .Add "Name"
      .Add "PackageCache"
      .Add "SKUNumber"
      .Add "Vendor"
      .Add "Version"
   End With
   Set Product = c
End Function
Private Function ProgIDSpecification() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Caption"
      .Add "CheckID"
      .Add "CheckMode"
      .Add "Description"
      .Add "Name"
      .Add "Parent"
      .Add "ProgID"
      .Add "SoftwareElementID"
      .Add "SoftwareElementState"
      .Add "TargetOperatingSystem"
      .Add "Version"
   End With
   Set ProgIDSpecification = c
End Function

Private Function ProgramGroup() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Caption"
      .Add "Description"
      .Add "GroupName"
      .Add "Name"
      .Add "SettingID"
      .Add "UserName"
   End With
   Set ProgramGroup = c
End Function
Private Function ProgramGroupOrItem() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Caption"
      .Add "Description"
      .Add "InstallDate"
      .Add "Name"
      .Add "Status"
   End With
   Set ProgramGroupOrItem = c
End Function

Private Function Property() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Caption"
      .Add "Description"
      .Add "ProductCode"
      .Add "Property"
      .Add "SettingID"
      .Add "Value"
   End With
   Set Property = c
End Function
Private Function Proxy() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Caption"
      .Add "Description"
      .Add "ProxyPortNumber"
      .Add "ProxyServer"
      .Add "ServerName"
      .Add "SettingID"
   End With
   Set Proxy = c
End Function

Private Function PublishComponentAction() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "ActionID"
      .Add "AppData"
      .Add "Caption"
      .Add "ComponentID"
      .Add "Description"
      .Add "Direction"
      .Add "Name"
      .Add "Qual"
      .Add "SoftwareElementID"
      .Add "SoftwareElementState"
      .Add "TargetOperatingSystem"
      .Add "Version"
   End With
   Set PublishComponentAction = c
End Function
Private Function QuickFixEngineering() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Caption"
      .Add "CSName"
      .Add "Description"
      .Add "FixComments"
      .Add "HotFixID"
      .Add "InstallDate"
      .Add "InstalledBy"
      .Add "InstalledOn"
      .Add "Name"
      .Add "ServicePackInEffect"
      .Add "Status"
   End With
   Set QuickFixEngineering = c
End Function

Private Function QuotaSetting() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Caption"
      .Add "DefaultLimit"
      .Add "DefaultWarningLimit"
      .Add "Description"
      .Add "ExceededNotification"
      .Add "SettingID"
      .Add "State"
      .Add "VolumePath"
      .Add "WarningExceededNotification"
   End With
   Set QuotaSetting = c
End Function
Private Function Refrigeration() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "ActiveCooling"
      .Add "Availability"
      .Add "Caption"
      .Add "ConfigManagerErrorCode"
      .Add "ConfigManagerUserConfig"
      .Add "CreationClassName"
      .Add "Description"
      .Add "DeviceID"
      .Add "ErrorCleared"
      .Add "ErrorDescription"
      .Add "InstallDate"
      .Add "LastErrorCode"
      .Add "Name"
      .Add "PNPDeviceID"
      .Add "PowerManagementCapabilities[]"
      .Add "PowerManagementSupported"
      .Add "Status"
      .Add "StatusInfo"
      .Add "SystemCreationClassName"
      .Add "SystemName"
   End With
   Set Refrigeration = c
End Function

Private Function Registry() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Caption"
      .Add "CurrentSize"
      .Add "Description"
      .Add "InstallDate"
      .Add "MaximumSize"
      .Add "Name"
      .Add "ProposedSize"
      .Add "Status"
   End With
   Set Registry = c
End Function
Private Function RegistryAction() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "ActionID"
      .Add "Caption"
      .Add "Description"
      .Add "Direction"
      .Add "EntryName"
      .Add "EntryValue"
      .Add "Key"
      .Add "Name"
      .Add "Registry"
      .Add "Root"
      .Add "SoftwareElementID"
      .Add "SoftwareElementState"
      .Add "TargetOperatingSystem"
      .Add "Version"
   End With
   Set RegistryAction = c
End Function

Private Function SecuritySetting() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Caption"
      .Add "ControlFlags"
      .Add "Description"
      .Add "SettingID"
   End With
   Set SecuritySetting = c
End Function
Private Function SelfRegModuleAction() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "ActionID"
      .Add "Caption"
      .Add "Cost"
      .Add "Description"
      .Add "Direction"
      .Add "File"
      .Add "Name"
      .Add "SoftwareElementID"
      .Add "SoftwareElementState"
      .Add "TargetOperatingSystem"
      .Add "Version"
   End With
   Set SelfRegModuleAction = c
End Function


Private Function SerialPort() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Availability"
      .Add "Binary"
      .Add "Capabilities[]"
      .Add "CapabilityDescriptions[]"
      .Add "Caption"
      .Add "ConfigManagerErrorCode"
      .Add "ConfigManagerUserConfig"
      .Add "CreationClassName"
      .Add "Description"
      .Add "DeviceID"
      .Add "ErrorCleared"
      .Add "ErrorDescription"
      .Add "InstallDate"
      .Add "LastErrorCode"
      .Add "MaxBaudRate"
      .Add "MaximumInputBufferSize"
      .Add "MaximumOutputBufferSize"
      .Add "MaxNumberControlled"
      .Add "Name"
      .Add "OSAutoDiscovered"
      .Add "PNPDeviceID"
      .Add "PowerManagementCapabilities[]"
      .Add "PowerManagementSupported"
      .Add "ProtocolSupported"
      .Add "ProviderType"
      .Add "SettableBaudRate"
      .Add "SettableDataBits"
      .Add "SettableFlowControl"
      .Add "SettableParity"
      .Add "SettableParityCheck"
      .Add "SettableRLSD"
      .Add "SettableStopBits"
      .Add "Status"
      .Add "StatusInfo"
      .Add "Supports16BitMode"
      .Add "SupportsDTRDSR"
      .Add "SupportsElapsedTimeouts"
      .Add "SupportsIntTimeouts"
      .Add "SupportsParityCheck"
      .Add "SupportsRLSD"
      .Add "SupportsRTSCTS"
      .Add "SupportsSpecialCharacters"
      .Add "SupportsXOnXOff"
      .Add "SupportsXOnXOffSet"
      .Add "SystemCreationClassName"
      .Add "SystemName"
      .Add "TimeOfLastReset"
   End With
   Set SerialPort = c
End Function
Private Function SerialPortConfiguration() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "AbortReadWriteOnError"
      .Add "BaudRate"
      .Add "BinaryModeEnabled"
      .Add "BitsPerByte"
      .Add "Caption"
      .Add "ContinueXMitOnXOff"
      .Add "CTSOutflowControl"
      .Add "Description"
      .Add "DiscardNULLBytes"
      .Add "DSROutflowControl"
      .Add "DSRSensitivity"
      .Add "DTRFlowControlType"
      .Add "EOFCharacter"
      .Add "ErrorReplaceCharacter"
      .Add "ErrorReplacementEnabled"
      .Add "EventCharacter"
      .Add "IsBusy"
      .Add "Name"
      .Add "Parity"
      .Add "ParityCheckEnabled"
      .Add "RTSFlowControlType"
      .Add "SettingID"
      .Add "StopBits"
      .Add "XOffCharacter"
      .Add "XOffXMitThreshold"
      .Add "XOnCharacter"
      .Add "XOnXMitThreshold"
      .Add "XOnXOffInFlowControl"
      .Add "XOnXOffOutFlowControl"
   End With
   Set SerialPortConfiguration = c
End Function


Private Function ServerConnection() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "ActiveTime"
      .Add "Caption"
      .Add "ComputerName"
      .Add "ConnectionID"
      .Add "Description"
      .Add "InstallDate"
      .Add "Name"
      .Add "NumberOfFiles"
      .Add "NumberOfUsers"
      .Add "ShareName"
      .Add "Status"
      .Add "UserName"
   End With
   Set ServerConnection = c
End Function
Private Function ServerSession() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "ActiveTime"
      .Add "Caption"
      .Add "ClientType"
      .Add "ComputerName"
      .Add "Description"
      .Add "IdleTime"
      .Add "InstallDate"
      .Add "Name"
      .Add "ResourcesOpened"
      .Add "SessionType"
      .Add "Status"
      .Add "TransportName"
      .Add "UserName"
   End With
   Set ServerSession = c
End Function


Private Function ServiceControl() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Arguments"
      .Add "Caption"
      .Add "Description"
      .Add "Event"
      .Add "ID"
      .Add "Name"
      .Add "ProductCode"
      .Add "SettingID"
      .Add "Wait"
   End With
   Set ServiceControl = c
End Function
Private Function ServiceSpecification() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Caption"
      .Add "CheckID"
      .Add "CheckMode"
      .Add "Dependencies"
      .Add "Description"
      .Add "DisplayName"
      .Add "ErrorControl"
      .Add "ID"
      .Add "LoadOrderGroup"
      .Add "Name"
      .Add "Password"
      .Add "ServiceType"
      .Add "SoftwareElementID"
      .Add "SoftwareElementState"
      .Add "StartName"
      .Add "StartType"
      .Add "TargetOperatingSystem"
      .Add "Version"
   End With
   Set ServiceSpecification = c
End Function


Private Function Session() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Caption"
      .Add "Description"
      .Add "InstallDate"
      .Add "Name"
      .Add "StartTime"
      .Add "Status"
   End With
   Set Session = c
End Function
Private Function ShadowContext() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "ClientAccessible"
      .Add "Differential"
      .Add "ExposedLocally"
      .Add "ExposedRemotely"
      .Add "HardwareAssisted"
      .Add "Imported"
      .Add "Name"
      .Add "NoAutoRelease"
      .Add "NotSurfaced"
      .Add "NoWriters"
      .Add "Persistent"
      .Add "Plex"
      .Add "Transportable"
   End With
   Set ShadowContext = c
End Function


Private Function ShadowCopy() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "ClientAccessible"
      .Add "Count"
      .Add "DeviceObject"
      .Add "Differential"
      .Add "ExposedLocally"
      .Add "ExposedName"
      .Add "ExposedRemotely"
      .Add "HardwareAssisted"
      .Add "ID"
      .Add "Imported"
      .Add "NoAutoRelease"
      .Add "NotSurfaced"
      .Add "NoWriters"
      .Add "OriginatingMachine"
      .Add "Persistent"
      .Add "Plex"
      .Add "ProviderID"
      .Add "ServiceMachine"
      .Add "SetID"
      .Add "State"
      .Add "Transportable"
      .Add "VolumeName"
   End With
   Set ShadowCopy = c
End Function
Private Function ShadowProvider() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "CLSID"
      .Add "ID"
      .Add "Name"
      .Add "Type"
      .Add "Version"
      .Add "VersionID"
   End With
   Set ShadowProvider = c
End Function


Private Function ShortcutAction() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "ActionID"
      .Add "Arguments"
      .Add "Caption"
      .Add "Description"
      .Add "Direction"
      .Add "HotKey"
      .Add "IconIndex"
      .Add "Name"
      .Add "Shortcut"
      .Add "ShowCmd"
      .Add "SoftwareElementID"
      .Add "SoftwareElementState"
      .Add "Target"
      .Add "TargetOperatingSystem"
      .Add "Version"
      .Add "WkDir"
   End With
   Set ShortcutAction = c
End Function
Private Function ShortcutFile() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "AccessMask[]"
      .Add "Archive"
      .Add "Caption"
      .Add "Compressed"
      .Add "CompressionMethod"
      .Add "CreationClassName"
      .Add "CreationDate"
      .Add "CSCreationClassName"
      .Add "CSName"
      .Add "Description"
      .Add "Drive"
      .Add "EightDotThreeFileName"
      .Add "Encrypted"
      .Add "EncryptionMethod"
      .Add "Extension"
      .Add "FileName"
      .Add "FileSize"
      .Add "FileType"
      .Add "FSCreationClassName"
      .Add "FSName"
      .Add "Hidden"
      .Add "InstallDate"
      .Add "InUseCount"
      .Add "LastAccessed"
      .Add "LastModified"
      .Add "Manufacturer"
      .Add "Name"
      .Add "Path"
      .Add "Readable"
      .Add "Status"
      .Add "System"
      .Add "Target"
      .Add "Version"
      .Add "Writeable"
   End With
   Set ShortcutFile = c
End Function


Private Function SID() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "AccountName"
      .Add "BinaryRepresentation[]"
      .Add "ReferencedDomainName"
      .Add "SID"
      .Add "SidLength"
   End With
   Set SID = c
End Function
Private Function SMBIOSMemory() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Access"
      .Add "AdditionalErrorData[]"
      .Add "Availability"
      .Add "BlockSize"
      .Add "Caption"
      .Add "ConfigManagerErrorCode"
      .Add "ConfigManagerUserConfig"
      .Add "CorrectableError"
      .Add "CreationClassName"
      .Add "Description"
      .Add "DeviceID"
      .Add "EndingAddress"
      .Add "ErrorAccess"
      .Add "ErrorAddress"
      .Add "ErrorCleared"
      .Add "ErrorData[]"
      .Add "ErrorDataOrder"
      .Add "ErrorDescription"
      .Add "ErrorInfo"
      .Add "ErrorMethodology"
      .Add "ErrorResolution"
      .Add "ErrorTime"
      .Add "ErrorTransferSize"
      .Add "InstallDate"
      .Add "LastErrorCode"
      .Add "Name"
      .Add "NumberOfBlocks"
      .Add "OtherErrorDescription"
      .Add "PNPDeviceID"
      .Add "PowerManagementCapabilities[]"
      .Add "PowerManagementSupported"
      .Add "Purpose"
      .Add "StartingAddress"
      .Add "Status"
      .Add "StatusInfo"
      .Add "SystemCreationClassName"
      .Add "SystemLevelAddress"
      .Add "SystemName"
   End With
   Set SMBIOSMemory = c
End Function


Private Function SCSIController() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Availability"
      .Add "Caption"
      .Add "ConfigManagerErrorCode"
      .Add "ConfigManagerUserConfig"
      .Add "ControllerTimeouts"
      .Add "CreationClassName"
      .Add "Description"
      .Add "DeviceID"
      .Add "DeviceMap"
      .Add "DriverName"
      .Add "ErrorCleared"
      .Add "ErrorDescription"
      .Add "HardwareVersion"
      .Add "Index"
      .Add "InstallDate"
      .Add "LastErrorCode"
      .Add "Manufacturer"
      .Add "MaxDataWidth"
      .Add "MaxNumberControlled"
      .Add "MaxTransferRate"
      .Add "Name"
      .Add "PNPDeviceID"
      .Add "PowerManagementCapabilities[]"
      .Add "PowerManagementSupported"
      .Add "ProtectionManagement"
      .Add "ProtocolSupported"
      .Add "Status"
      .Add "StatusInfo"
      .Add "SystemCreationClassName"
      .Add "SystemName"
      .Add "TimeOfLastReset"
   End With
   Set SCSIController = c
End Function
Private Function ReserveCost() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Caption"
      .Add "CheckID"
      .Add "CheckMode"
      .Add "Description"
      .Add "Name"
      .Add "ReserveFolder"
      .Add "ReserveKey"
      .Add "ReserveLocal"
      .Add "ReserveSource"
      .Add "SoftwareElementID"
      .Add "SoftwareElementState"
      .Add "TargetOperatingSystem"
      .Add "Version"
   End With
   Set ReserveCost = c
End Function


Private Function RemoveIniAction() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Action"
      .Add "ActionID"
      .Add "Caption"
      .Add "Description"
      .Add "Direction"
      .Add "Key"
      .Add "Name"
      .Add "Section"
      .Add "SoftwareElementID"
      .Add "SoftwareElementState"
      .Add "TargetOperatingSystem"
      .Add "Value"
      .Add "Version"
   End With
   Set RemoveIniAction = c
End Function
Private Function RemoveFileAction() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "ActionID"
      .Add "Caption"
      .Add "Description"
      .Add "Direction"
      .Add "DirProperty"
      .Add "File"
      .Add "FileKey"
      .Add "FileName"
      .Add "InstallMode"
      .Add "Name"
      .Add "SoftwareElementID"
      .Add "SoftwareElementState"
      .Add "TargetOperatingSystem"
      .Add "Version"
   End With
   Set RemoveFileAction = c
End Function
Private Function Thread() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Caption"
      .Add "CreationClassName"
      .Add "CSCreationClassName"
      .Add "CSName"
      .Add "Description"
      .Add "ElapsedTime"
      .Add "ExecutionState"
      .Add "Handle"
      .Add "InstallDate"
      .Add "KernelModeTime"
      .Add "Name"
      .Add "OSCreationClassName"
      .Add "OSName"
      .Add "Priority"
      .Add "PriorityBase"
      .Add "ProcessCreationClassName"
      .Add "ProcessHandle"
      .Add "StartAddress"
      .Add "Status"
      .Add "ThreadState"
      .Add "ThreadWaitReason"
      .Add "UserModeTime"
   End With
   Set Thread = c
End Function
Private Function ThreadStartTrace() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "ProcessID"
      .Add "SECURITY_DESCRIPTOR[]"
      .Add "StackBase"
      .Add "StackLimit"
      .Add "StartAddr"
      .Add "ThreadID"
      .Add "TIME_CREATED"
      .Add "UserStackBase"
      .Add "UserStackLimit"
      .Add "WaitMode"
      .Add "Win32StartAddr"
   End With
   Set ThreadStartTrace = c
End Function
Private Function ThreadStopTrace() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "ProcessID"
      .Add "SECURITY_DESCRIPTOR[]"
      .Add "ThreadID"
      .Add "TIME_CREATED"
   End With
   Set ThreadStopTrace = c
End Function
Private Function ThreadTrace() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "ProcessID"
      .Add "SECURITY_DESCRIPTOR[]"
      .Add "ThreadID"
      .Add "TIME_CREATED"
   End With
   Set ThreadTrace = c
End Function

Private Function TemperatureProbe() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Accuracy"
      .Add "Availability"
      .Add "Caption"
      .Add "ConfigManagerErrorCode"
      .Add "ConfigManagerUserConfig"
      .Add "CreationClassName"
      .Add "CurrentReading"
      .Add "Description"
      .Add "DeviceID"
      .Add "ErrorCleared"
      .Add "ErrorDescription"
      .Add "InstallDate"
      .Add "IsLinear"
      .Add "LastErrorCode"
      .Add "LowerThresholdCritical"
      .Add "LowerThresholdFatal"
      .Add "LowerThresholdNonCritical"
      .Add "MaxReadable"
      .Add "MinReadable"
      .Add "Name"
      .Add "NominalReading"
      .Add "NormalMax"
      .Add "NormalMin"
      .Add "PNPDeviceID"
      .Add "PowerManagementCapabilities[]"
      .Add "PowerManagementSupported"
      .Add "Resolution"
      .Add "Status"
      .Add "StatusInfo"
      .Add "SystemCreationClassName"
      .Add "SystemName"
      .Add "Tolerance"
      .Add "UpperThresholdCritical"
      .Add "UpperThresholdFatal"
      .Add "UpperThresholdNonCritical"
   End With
   Set TemperatureProbe = c
End Function
Private Function TCPIPPrinterPort() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "ByteCount"
      .Add "Caption"
      .Add "CreationClassName"
      .Add "Description"
      .Add "HostAddress"
      .Add "InstallDate"
      .Add "Name"
      .Add "PortNumber"
      .Add "Protocol"
      .Add "Queue"
      .Add "SNMPCommunity"
      .Add "SNMPDevIndex"
      .Add "SNMPEnabled"
      .Add "Status"
      .Add "SystemCreationClassName"
      .Add "SystemName"
      .Add "Type"
   End With
   Set TCPIPPrinterPort = c
End Function
Private Function TapeDrive() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Availability"
      .Add "Capabilities[]"
      .Add "CapabilityDescriptions[]"
      .Add "Caption"
      .Add "Compression"
      .Add "CompressionMethod"
      .Add "ConfigManagerErrorCode"
      .Add "ConfigManagerUserConfig"
      .Add "CreationClassName"
      .Add "DefaultBlockSize"
      .Add "Description"
      .Add "DeviceID"
      .Add "ECC"
      .Add "EOTWarningZoneSize"
      .Add "ErrorCleared"
      .Add "ErrorDescription"
      .Add "ErrorMethodology"
      .Add "FeaturesHigh"
      .Add "FeaturesLow"
      .Add "Id"
      .Add "InstallDate"
      .Add "LastErrorCode"
      .Add "Manufacturer"
      .Add "MaxBlockSize"
      .Add "MaxMediaSize"
      .Add "MaxPartitionCount"
      .Add "MediaType"
      .Add "MinBlockSize"
      .Add "Name"
      .Add "NeedsCleaning"
      .Add "NumberOfMediaSupported"
      .Add "Padding"
      .Add "PNPDeviceID"
      .Add "PowerManagementCapabilities[]"
      .Add "PowerManagementSupported"
      .Add "ReportSetMarks"
      .Add "Status"
      .Add "StatusInfo"
      .Add "SystemCreationClassName"
      .Add "SystemName"
   End With
   Set TapeDrive = c
End Function
Private Function SystemTrace() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "SECURITY_DESCRIPTOR[]"
      .Add "TIME_CREATED"
   End With
   Set SystemTrace = c
End Function
Private Function SystemSlot() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Caption"
      .Add "ConnectorPinout"
      .Add "ConnectorType[]"
      .Add "CreationClassName"
      .Add "CurrentUsage"
      .Add "Description"
      .Add "HeightAllowed"
      .Add "InstallDate"
      .Add "LengthAllowed"
      .Add "Manufacturer"
      .Add "MaxDataWidth"
      .Add "Model"
      .Add "Name"
      .Add "Number"
      .Add "OtherIdentifyingInfo"
      .Add "PartNumber"
      .Add "PMESignal"
      .Add "PoweredOn"
      .Add "PurposeDescription"
      .Add "SerialNumber"
      .Add "Shared"
      .Add "SKU"
      .Add "SlotDesignation"
      .Add "SpecialPurpose"
      .Add "Status"
      .Add "SupportsHotPlug"
      .Add "Tag"
      .Add "ThermalRating"
      .Add "VccMixedVoltageSupport[]"
      .Add "Version"
      .Add "VppMixedVoltageSupport[]"
   End With
   Set SystemSlot = c
End Function
Private Function SystemAccount() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Caption"
      .Add "Description"
      .Add "Domain"
      .Add "InstallDate"
      .Add "LocalAccount"
      .Add "Name"
      .Add "SID"
      .Add "SIDType"
      .Add "Status"
   End With
   Set SystemAccount = c
End Function
Private Function SystemConfigurationChangeEvent() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "EventType"
      .Add "SECURITY_DESCRIPTOR[]"
      .Add "TIME_CREATED"
   End With
   Set SystemConfigurationChangeEvent = c
End Function
Private Function SystemEnclosure() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "AudibleAlarm"
      .Add "BreachDescription"
      .Add "CableManagementStrategy"
      .Add "Caption"
      .Add "ChassisTypes[]"
      .Add "CreationClassName"
      .Add "CurrentRequiredOrProduced"
      .Add "Depth"
      .Add "Description"
      .Add "HeatGeneration"
      .Add "Height"
      .Add "HotSwappable"
      .Add "InstallDate"
      .Add "LockPresent"
      .Add "Manufacturer"
      .Add "Model"
      .Add "Name"
      .Add "NumberOfPowerCords"
      .Add "OtherIdentifyingInfo"
      .Add "PartNumber"
      .Add "PoweredOn"
      .Add "Removable"
      .Add "Replaceable"
      .Add "SecurityBreach"
      .Add "SecurityStatus"
      .Add "SerialNumber"
      .Add "ServiceDescriptions[]"
      .Add "ServicePhilosophy[]"
      .Add "SKU"
      .Add "SMBIOSAssetTag"
      .Add "Status"
      .Add "Tag"
      .Add "TypeDescriptions[]"
      .Add "Version"
      .Add "VisibleAlarm"
      .Add "Weight"
      .Add "Width"
   End With
   Set SystemEnclosure = c
End Function
Private Function SystemMemoryResource() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Caption"
      .Add "CreationClassName"
      .Add "CSCreationClassName"
      .Add "CSName"
      .Add "Description"
      .Add "EndingAddress"
      .Add "InstallDate"
      .Add "Name"
      .Add "StartingAddress"
      .Add "Status"
   End With
   Set SystemMemoryResource = c
End Function


Private Function TimeZone() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Bias"
      .Add "Caption"
      .Add "DaylightBias"
      .Add "DaylightDay"
      .Add "DaylightDayOfWeek"
      .Add "DaylightHour"
      .Add "DaylightMillisecond"
      .Add "DaylightMinute"
      .Add "DaylightMonth"
      .Add "DaylightName"
      .Add "DaylightSecond"
      .Add "DaylightYear"
      .Add "Description"
      .Add "SettingID"
      .Add "StandardBias"
      .Add "StandardDay"
      .Add "StandardDayOfWeek"
      .Add "StandardHour"
      .Add "StandardMillisecond"
      .Add "StandardMinute"
      .Add "StandardMonth"
      .Add "StandardName"
      .Add "StandardSecond"
      .Add "StandardYear"
   End With
   Set TimeZone = c
End Function
Private Function Trustee() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Domain"
      .Add "Name"
      .Add "SID[]"
      .Add "SidLength"
      .Add "SIDString"
   End With
   Set Trustee = c
End Function
Private Function UTCTime() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Day"
      .Add "DayOfWeek"
      .Add "Hour"
      .Add "Milliseconds"
      .Add "Minute"
      .Add "Month"
      .Add "Quarter"
      .Add "Second"
      .Add "WeekInMonth"
      .Add "Year"
   End With
   Set UTCTime = c
End Function
Private Function VideoController() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "AcceleratorCapabilities[]"
      .Add "AdapterCompatibility"
      .Add "AdapterDACType"
      .Add "AdapterRAM"
      .Add "Availability"
      .Add "CapabilityDescriptions[]"
      .Add "Caption"
      .Add "ColorTableEntries"
      .Add "ConfigManagerErrorCode"
      .Add "ConfigManagerUserConfig"
      .Add "CreationClassName"
      .Add "CurrentBitsPerPixel"
      .Add "CurrentHorizontalResolution"
      .Add "CurrentNumberOfColors"
      .Add "CurrentNumberOfColumns"
      .Add "CurrentNumberOfRows"
      .Add "CurrentRefreshRate"
      .Add "CurrentScanMode"
      .Add "CurrentVerticalResolution"
      .Add "Description"
      .Add "DeviceID"
      .Add "DeviceSpecificPens"
      .Add "DitherType"
      .Add "DriverDate"
      .Add "DriverVersion"
      .Add "ErrorCleared"
      .Add "ErrorDescription"
      .Add "ICMIntent"
      .Add "ICMMethod"
      .Add "InfFilename"
      .Add "InfSection"
      .Add "InstallDate"
      .Add "InstalledDisplayDrivers"
      .Add "LastErrorCode"
      .Add "MaxMemorySupported"
      .Add "MaxNumberControlled"
      .Add "MaxRefreshRate"
      .Add "MinRefreshRate"
      .Add "Monochrome"
      .Add "Name"
      .Add "NumberOfColorPlanes"
      .Add "NumberOfVideoPages"
      .Add "PNPDeviceID"
      .Add "PowerManagementCapabilities[]"
      .Add "PowerManagementSupported"
      .Add "ProtocolSupported"
      .Add "ReservedSystemPaletteEntries"
      .Add "SpecificationVersion"
      .Add "Status"
      .Add "StatusInfo"
      .Add "SystemCreationClassName"
      .Add "SystemName"
      .Add "SystemPaletteEntries"
      .Add "TimeOfLastReset"
      .Add "VideoArchitecture"
      .Add "VideoMemoryType"
      .Add "VideoMode"
      .Add "VideoModeDescription"
      .Add "VideoProcessor"
   End With
   Set VideoController = c
End Function
Private Function VideoConfiguration() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "ActualColorResolution"
      .Add "AdapterChipType"
      .Add "AdapterCompatibility"
      .Add "AdapterDACType"
      .Add "AdapterDescription"
      .Add "AdapterRAM"
      .Add "AdapterType"
      .Add "BitsPerPixel"
      .Add "Caption"
      .Add "ColorPlanes"
      .Add "ColorTableEntries"
      .Add "Description"
      .Add "DeviceSpecificPens"
      .Add "DriverDate"
      .Add "HorizontalResolution"
      .Add "InfFilename"
      .Add "InfSection"
      .Add "InstalledDisplayDrivers"
      .Add "MonitorManufacturer"
      .Add "MonitorType"
      .Add "Name"
      .Add "PixelsPerXLogicalInch"
      .Add "PixelsPerYLogicalInch"
      .Add "RefreshRate"
      .Add "ScanMode"
      .Add "ScreenHeight"
      .Add "ScreenWidth"
      .Add "SettingID"
      .Add "SystemPaletteEntries"
      .Add "VerticalResolution"
   End With
   Set VideoConfiguration = c
End Function
Private Function USBHub() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Availability"
      .Add "Caption"
      .Add "ClassCode"
      .Add "ConfigManagerErrorCode"
      .Add "ConfigManagerUserCode"
      .Add "CreationClassName"
      .Add "CurrentAlternativeSettings"
      .Add "CurrentConfigValue"
      .Add "Description"
      .Add "DeviceID"
      .Add "ErrorCleared"
      .Add "ErrorDescription"
      .Add "GangSwitched"
      .Add "InstallDate"
      .Add "LastErrorCode"
      .Add "Name"
      .Add "NumberOfConfigs"
      .Add "NumberOfPorts"
      .Add "PNPDeviceID"
      .Add "PowerManagementCapabilities[]"
      .Add "PowerManagementSupported"
      .Add "ProtocolCode"
      .Add "Status"
      .Add "StatusInfo"
      .Add "SubclassCode"
      .Add "SystemCreationClassName"
      .Add "SystemName"
      .Add "USBVersion"
   End With
   Set USBHub = c
End Function
Private Function TypeLibraryAction() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "ActionID"
      .Add "Caption"
      .Add "Cost"
      .Add "Description"
      .Add "Direction"
      .Add "Language"
      .Add "LibID"
      .Add "Name"
      .Add "SoftwareElementID"
      .Add "SoftwareElementState"
      .Add "TargetOperatingSystem"
      .Add "Version"
   End With
   Set TypeLibraryAction = c
End Function
Private Function UninterruptiblePowerSupply() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "ActiveInputVoltage"
      .Add "Availability"
      .Add "BatteryInstalled"
      .Add "CanTurnOffRemotely"
      .Add "Caption"
      .Add "CommandFile"
      .Add "ConfigManagerErrorCode"
      .Add "ConfigManagerUserConfig"
      .Add "CreationClassName"
      .Add "Description"
      .Add "DeviceID"
      .Add "ErrorCleared"
      .Add "ErrorDescription"
      .Add "EstimatedChargeRemaining"
      .Add "EstimatedRunTime"
      .Add "FirstMessageDelay"
      .Add "InstallDate"
      .Add "IsSwitchingSupply"
      .Add "LastErrorCode"
      .Add "LowBatterySignal"
      .Add "MessageInterval"
      .Add "Name"
      .Add "PNPDeviceID"
      .Add "PowerFailSignal"
      .Add "PowerManagementCapabilities[]"
      .Add "PowerManagementSupported"
      .Add "Range1InputFrequencyHigh"
      .Add "Range1InputFrequencyLow"
      .Add "Range1InputVoltageHigh"
      .Add "Range1InputVoltageLow"
      .Add "Range2InputFrequencyHigh"
      .Add "Range2InputFrequencyLow"
      .Add "Range2InputVoltageHigh"
      .Add "Range2InputVoltageLow"
      .Add "RemainingCapacityStatus"
      .Add "Status"
      .Add "StatusInfo"
      .Add "SystemCreationClassName"
      .Add "SystemName"
      .Add "TimeOnBackup"
      .Add "TotalOutputPower"
      .Add "TypeOfRangeSwitching"
      .Add "UPSPort"
   End With
   Set UninterruptiblePowerSupply = c
End Function
Private Function USBController() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Availability"
      .Add "Caption"
      .Add "ConfigManagerErrorCode"
      .Add "ConfigManagerUserConfig"
      .Add "CreationClassName"
      .Add "Description"
      .Add "DeviceID"
      .Add "ErrorCleared"
      .Add "ErrorDescription"
      .Add "InstallDate"
      .Add "LastErrorCode"
      .Add "Manufacturer"
      .Add "MaxNumberControlled"
      .Add "Name"
      .Add "PNPDeviceID"
      .Add "PowerManagementCapabilities[]"
      .Add "PowerManagementSupported"
      .Add "ProtocolSupported"
      .Add "Status"
      .Add "StatusInfo"
      .Add "SystemCreationClassName"
      .Add "SystemName"
      .Add "TimeOfLastReset"
   End With
   Set USBController = c
End Function
Private Function Volume() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Automount"
      .Add "Capacity"
      .Add "Compressed"
      .Add "DeviceId"
      .Add "DirtyBitSet"
      .Add "DriveLetter"
      .Add "DriveType"
      .Add "FileSystem"
      .Add "FreeSpace"
      .Add "IndexingEnabled"
      .Add "Label"
      .Add "MaximumFileNameLength"
      .Add "QuotasEnabled"
      .Add "QuotasIncomplete"
      .Add "QuotasRebuilding"
      .Add "SerialNumber"
      .Add "SupportsDiskQuotas"
      .Add "SupportsFileBasedCompression"
   End With
   Set Volume = c
End Function
Private Function VolumeChangeEvent() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "DriveName"
      .Add "EventType"
      .Add "SECURITY_DESCRIPTOR[]"
      .Add "TIME_CREATED"
   End With
   Set VolumeChangeEvent = c
End Function
Private Function WindowsProductActivation() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "ActivationRequired"
      .Add "Caption"
      .Add "Description"
      .Add "IsNotificationOn"
      .Add "ProductID"
      .Add "RemainingEvaluationPeriod"
      .Add "RemainingGracePeriod"
      .Add "ServerName"
      .Add "SettingID"
   End With
   Set WindowsProductActivation = c
End Function
Private Function WMISetting() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "ASPScriptDefaultNamespace"
      .Add "ASPScriptEnabled"
      .Add "AutorecoverMofs"
      .Add "AutoStartWin9X"
      .Add "BackupInterval"
      .Add "BackupLastTime"
      .Add "BuildVersion"
      .Add "Caption"
      .Add "DatabaseDirectory"
      .Add "DatabaseMaxSize"
      .Add "Description"
      .Add "EnableAnonWin9xConnections"
      .Add "EnableEvents"
      .Add "EnableStartupHeapPreallocation"
      .Add "HighThresholdOnClientObjects"
      .Add "HighThresholdOnEvents"
      .Add "InstallationDirectory"
      .Add "LastStartupHeapPreallocation"
      .Add "LoggingDirectory"
      .Add "LoggingLevel"
      .Add "LowThresholdOnClientObjects"
      .Add "LowThresholdOnEvents"
      .Add "MaxLogFileSize"
      .Add "MaxWaitOnClientObjects"
      .Add "MaxWaitOnEvents"
      .Add "MofSelfInstallDirectory"
      .Add "SettingID"
   End With
   Set WMISetting = c
End Function
Private Function VoltageProbe() As Collection
   Dim c As Collection
   Set c = New Collection
   With c
      .Add "Accuracy"
      .Add "Availability"
      .Add "Caption"
      .Add "ConfigManagerErrorCode"
      .Add "ConfigManagerUserConfig"
      .Add "CreationClassName"
      .Add "CurrentReading"
      .Add "Description"
      .Add "DeviceID"
      .Add "ErrorCleared"
      .Add "ErrorDescription"
      .Add "InstallDate"
      .Add "IsLinear"
      .Add "LastErrorCode"
      .Add "LowerThresholdCritical"
      .Add "LowerThresholdFatal"
      .Add "LowerThresholdNonCritical"
      .Add "MaxReadable"
      .Add "MinReadable"
      .Add "Name"
      .Add "NominalReading"
      .Add "NormalMax"
      .Add "NormalMin"
      .Add "PNPDeviceID"
      .Add "PowerManagementCapabilities[]"
      .Add "PowerManagementSupported"
      .Add "Resolution"
      .Add "Status"
      .Add "StatusInfo"
      .Add "SystemCreationClassName"
      .Add "SystemName"
      .Add "Tolerance"
      .Add "UpperThresholdCritical"
      .Add "UpperThresholdFatal"
      .Add "UpperThresholdNonCritical"
   End With
   Set VoltageProbe = c
End Function

