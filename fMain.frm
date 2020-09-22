VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form fMain 
   Caption         =   "WMI Monitor"
   ClientHeight    =   6990
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   13245
   Icon            =   "fMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6990
   ScaleWidth      =   13245
   StartUpPosition =   2  'Centrovat na obrazovce
   WindowState     =   2  'Maximalizované
   Begin VB.PictureBox picSplitter 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'Není
      FillColor       =   &H00FF0000&
      Height          =   4800
      Left            =   6675
      ScaleHeight     =   2090.126
      ScaleMode       =   0  'Uživatelské
      ScaleWidth      =   780
      TabIndex        =   4
      Top             =   1200
      Visible         =   0   'False
      Width           =   72
   End
   Begin VB.PictureBox picTitles 
      Align           =   1  'Zarovnání nahoru
      BorderStyle     =   0  'Není
      Height          =   405
      Left            =   0
      ScaleHeight     =   405
      ScaleWidth      =   13245
      TabIndex        =   0
      Top             =   0
      Width           =   13245
      Begin VB.CommandButton Command3 
         Caption         =   "Help"
         Height          =   375
         Left            =   9480
         TabIndex        =   7
         Top             =   0
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Cancel          =   -1  'True
         Caption         =   "End"
         Height          =   375
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Monitor >"
         Default         =   -1  'True
         Height          =   375
         Left            =   990
         TabIndex        =   2
         Top             =   0
         Width           =   975
      End
      Begin VB.ComboBox cMenu 
         Appearance      =   0  'Plochý
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   30
         Width           =   7335
      End
   End
   Begin MSComctlLib.ListView lstLog 
      Height          =   5265
      Left            =   3600
      TabIndex        =   5
      Top             =   360
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   9287
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin MSComctlLib.ListView lstBrowse 
      Height          =   5265
      Left            =   0
      TabIndex        =   6
      Top             =   360
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   9287
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Image imgSplitter 
      Height          =   4785
      Left            =   3240
      MousePointer    =   9  'Šipka Z V
      Top             =   1200
      Width           =   150
   End
   Begin VB.Menu mColsMenu 
      Caption         =   "mColsMenu"
      Begin VB.Menu mCols 
         Caption         =   "Sort As String"
         Index           =   0
      End
      Begin VB.Menu mCols 
         Caption         =   "Sort As Number"
         Index           =   1
      End
      Begin VB.Menu mCols 
         Caption         =   "Sort As DateTime"
         Index           =   2
      End
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Win32_ As String

Dim objSWbemLocator As WbemScripting.SWbemLocator
Dim objSWbemServices As WbemScripting.SWbemServices

Public sServer As String
Public sAccount As String
Public sPassword As String
Dim Sort() As Boolean
Dim Sort2(1 To 2) As Boolean
Dim SortType As ListDataType
Dim mbMoving As Boolean
Const sglSplitLimit = 500
Private Declare Sub InitCommonControls Lib "comctl32.dll" ()


Private Sub Command3_Click()
fHelp.wb.Navigate "http://msdn.microsoft.com/library/default.asp?url=/library/en-us/wmisdk/wmi/win32_" & cMenu.Text & ".asp"
fHelp.Show vbModal, Me
End Sub

Private Sub Form_Resize()
   On Error Resume Next
   If Me.Width < 3000 Then Me.Width = 3000
   SizeControls imgSplitter.Left
End Sub


Private Sub imgSplitter_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   With imgSplitter
      picSplitter.Move .Left, .Top, .Width \ 2, .Height - 20
   End With
   picSplitter.Visible = True
   mbMoving = True
End Sub


Private Sub imgSplitter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim sglPos As Single
   If mbMoving Then
      sglPos = X + imgSplitter.Left
      If sglPos < sglSplitLimit Then
         picSplitter.Left = sglSplitLimit
      ElseIf sglPos > Me.Width - sglSplitLimit Then
         picSplitter.Left = Me.Width - sglSplitLimit
      Else
         picSplitter.Left = sglPos
      End If
   End If
End Sub


Private Sub imgSplitter_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   SizeControls picSplitter.Left
   picSplitter.Visible = False
   mbMoving = False
End Sub


Private Sub lstBrowse_DragDrop(Source As Control, X As Single, Y As Single)
   If Source = imgSplitter Then
      SizeControls X
   End If
End Sub


Sub SizeControls(X As Single)
   On Error Resume Next


   'set the width
   If X < 1500 Then X = 1500
   If X > (Me.Width - 1500) Then X = Me.Width - 1500
   lstBrowse.Width = X
   imgSplitter.Left = X
   lstLog.Left = X + 40
   lstLog.Width = Me.Width - (lstBrowse.Width + 140)


   'set the top
   lstBrowse.Top = picTitles.Height

   lstLog.Top = lstBrowse.Top
   

   lstBrowse.Height = Me.ScaleHeight - (picTitles.Top + picTitles.Height)
   

   lstLog.Height = lstBrowse.Height
   imgSplitter.Top = lstBrowse.Top
   imgSplitter.Height = lstBrowse.Height
End Sub



Public Sub Command1_Click()
   RunMonitor
End Sub

Private Sub Command2_Click()
   Unload Me
   End
End Sub

Private Sub RunMonitor()
   Screen.MousePointer = vbHourglass
   Dim objSWbemLocator
   Dim objSWbemServices
   Dim colSWbemObjectSet
   Dim obj
   Dim i As Integer
   
   Set objSWbemLocator = New WbemScripting.SWbemLocator
   Set objSWbemServices = objSWbemLocator.ConnectServer(sServer, "root\cimv2", sAccount, sPassword)
   Dim col As Collection
   
   Set col = col_Win32(cMenu.Text)
      
   Me.Caption = "WMI Class: W32_" & cMenu.Text
   
   Set colSWbemObjectSet = objSWbemServices.ExecQuery("Select * From Win32_" & cMenu.Text)
   
   
   
   lstLog.ListItems.Clear
   With lstLog.ColumnHeaders
      .Clear
      .Add 1, , " "
      For i = 2 To col.Count
         .Add i, , col.Item(i - 1)
         ReDim Sort(i)
      Next i
   End With
   
   lstLog.Refresh
   lstLog.Visible = False
   
   Dim lv As MSComctlLib.ListItem
   LockWindowUpdate lstLog.hWnd

   On Error Resume Next
   For Each obj In colSWbemObjectSet
      Set lv = lstLog.ListItems.Add(, , "+")
      For i = 1 To obj.Properties_().Count
         lv.SubItems(i) = obj.Properties_(lstLog.ColumnHeaders.Item(i + 1).Text)
      Next i
      Me.Caption = "WMI Class: W32_" & cMenu.Text & " (items: " & lstLog.ListItems.Count & ")"
   Next
   On Error GoTo 0
   LockWindowUpdate 0
    
   lstLog.Refresh
   lstLog.Visible = True
   
   On Error Resume Next
   Call lstLog_ItemClick(lstLog.ListItems.Item(1))
   
   Set objSWbemServices = Nothing
   Set colSWbemObjectSet = Nothing
   Set objSWbemLocator = Nothing
   Set col = Nothing

   Screen.MousePointer = vbNormal
   Beep
End Sub

Private Sub Form_Initialize()
  Call InitCommonControls
 
   Me.mColsMenu.Visible = False
   Me.Show
   fLogin.Show vbModal, Me
End Sub

Private Sub Form_Load()

   With cMenu
      .AddItem "1349Controller"
      .AddItem "Account"
      .AddItem "AutochkSetting"
      .AddItem "BaseBoard"
      .AddItem "BaseService"
      .AddItem "Battery"
      .AddItem "Binary"
      .AddItem "BindImageAction"
      .AddItem "Bios"
      .AddItem "BootConfiguration"
      .AddItem "Bus"
      .AddItem "CDROMDrive"
      .AddItem "COMApplication"
      .AddItem "COMClass"
      .AddItem "COMSetting"
      .AddItem "CacheMemory"
      .AddItem "ClassInfoAction"
      .AddItem "ClassicCOMClass"
      .AddItem "ClassicCOMClassSetting"
      .AddItem "CodecFile"
      .AddItem "CommandLineAccess"
      .AddItem "ComponentCategory"
      .AddItem "ComputerShutdownEvent"
      .AddItem "ComputerSystem"
      .AddItem "ComputerSystemEvent"
      .AddItem "ComputerSystemProduct"
      .AddItem "Condition"
      .AddItem "CreateFolderAction"
      .AddItem "CurrentProbe"
      .AddItem "CurrentTime"
      .AddItem "DCOMApplication"
      .AddItem "DMAChannel"
      .AddItem "Directory"
      .AddItem "DiskDrive"
      .AddItem "DiskPartition"
      .AddItem "DisplayConfiguration"
      .AddItem "DisplayControllerConfiguration"
      .AddItem "DriverVXD"
      .AddItem "DuplicateFileAction"
      .AddItem "Environment"
      .AddItem "EnvironmentSpecification"
      .AddItem "ExtensionInfoAction"
      .AddItem "Fan"
      .AddItem "FileSpecification"
      .AddItem "FloppyController"
      .AddItem "FloppyDrive"
      .AddItem "FontInfoAction"
      .AddItem "Group"
      .AddItem "HeatPipe"
      .AddItem "IDEController"
      .AddItem "IP4PersistedRouteTable"
      .AddItem "IP4RouteTable"
      .AddItem "IRQResource"
      .AddItem "InfraredDevice"
      .AddItem "IniFileSpecification"
      .AddItem "JobObjectStatus"
      .AddItem "Keyboard"
      .AddItem "LaunchCondition"
      .AddItem "LoadOrderGroup"
      .AddItem "LocalTime"
      .AddItem "LogicalDisk"
      .AddItem "LogicalFileSecuritySetting"
      .AddItem "LogicalMemoryConfiguration"
      .AddItem "LogicalProgramGroup"
      .AddItem "LogicalProgramGroupItem"
      .AddItem "LogicalShareSetting"
      .AddItem "LogonSession"
      .AddItem "MIMEInfoAction"
      .AddItem "MSIResource"
      .AddItem "MappedLogicalDisk"
      .AddItem "MemoryArray"
      .AddItem "MemoryDevice"
      .AddItem "ModuleLoadTrace"
      .AddItem "MotherboardDevice"
      .AddItem "MoveFileAction"
      .AddItem "NTDomain"
      .AddItem "NTEventLogFile"
      .AddItem "NTLogEvent"
      .AddItem "NamedJobObject"
      .AddItem "NamedJobObjectActgInfo"
      .AddItem "NamedJobObjectLimitSetting"
      .AddItem "NetworkAdapter"
      .AddItem "NetworkAdapterConfiguration"
      .AddItem "NetworkClient"
      .AddItem "NetworkConnection"
      .AddItem "NetworkLoginProfile"
      .AddItem "NetworkProtocol"
      .AddItem "ODBCAttribute"
      .AddItem "ODBCDataSourceSpecification"
      .AddItem "ODBCDriverSpecification"
      .AddItem "ODBCSourceAttribute"
      .AddItem "ODBCTranslatorSpecification"
      .AddItem "OSRecoveryConfiguration"
      .AddItem "OnBoardDevice"
      .AddItem "OperatingSystem"
      .AddItem "PCMCIAController"
      .AddItem "POTSModem"
      .AddItem "PageFile"
      .AddItem "PageFileSetting"
      .AddItem "PageFileUsage"
      .AddItem "ParallelPort"
      .AddItem "Patch"
      .AddItem "PatchPackage"
      .AddItem "Perf"
      .AddItem "PerfFormattedData"
      .AddItem "PerfFormattedData_ASP_ActiveServerPages"
      .AddItem "PerfFormattedData_ContentFilter_IndexingServiceFilter"
      .AddItem "PerfFormattedData_ContentIndex_IndexingService"
      .AddItem "PerfFormattedData_ISAPISearch_HttpIndexingService"
      .AddItem "PerfFormattedData_InetInfo_InternetInformationServicesGlobal"
      .AddItem "PerfFormattedData_MSDTC_DistributedTransactionCoordinator"
      .AddItem "PerfFormattedData_NTFSDRV_SMTPNTFSStoreDriver"
      .AddItem "PerfFormattedData_PSched_PSchedFlow"
      .AddItem "PerfFormattedData_PSched_PSchedPipe"
      .AddItem "PerfFormattedData_PerfDisk_LogicalDisk"
      .AddItem "PerfFormattedData_PerfDisk_PhysicalDisk"
      .AddItem "PerfFormattedData_PerfNet_Browser"
      .AddItem "PerfFormattedData_PerfNet_Redirector"
      .AddItem "PerfFormattedData_PerfNet_Server"
      .AddItem "PerfFormattedData_PerfNet_ServerWorkQueues"
      .AddItem "PerfFormattedData_PerfOS_Cache"
      .AddItem "PerfFormattedData_PerfOS_Memory"
      .AddItem "PerfFormattedData_PerfOS_Objects"
      .AddItem "PerfFormattedData_PerfOS_PagingFile"
      .AddItem "PerfFormattedData_PerfOS_Processor"
      .AddItem "PerfFormattedData_PerfOS_System"
      .AddItem "PerfFormattedData_PerfProc_FullImage_Costly"
      .AddItem "PerfFormattedData_PerfProc_Image_Costly"
      .AddItem "PerfFormattedData_PerfProc_JobObject"
      .AddItem "PerfFormattedData_PerfProc_JobObjectDetails"
      .AddItem "PerfFormattedData_PerfProc_Process"
      .AddItem "PerfFormattedData_PerfProc_ProcessAddressSpace_Costly"
      .AddItem "PerfFormattedData_PerfProc_Thread"
      .AddItem "PerfFormattedData_PerfProc_ThreadDetails_Costly"
      .AddItem "PerfFormattedData_RSVP_ACSRSVPInterfaces"
      .AddItem "PerfFormattedData_RSVP_ACSRSVPService"
      .AddItem "PerfFormattedData_RemoteAccess_RASPort"
      .AddItem "PerfFormattedData_RemoteAccess_RASTotal"
      .AddItem "PerfFormattedData_SMTPSVC_SMTPServer"
      .AddItem "PerfFormattedData_Spooler_PrintQueue"
      .AddItem "PerfFormattedData_TapiSrv_Telephony"
      .AddItem "PerfFormattedData_Tcpip_ICMP"
      .AddItem "PerfFormattedData_Tcpip_IP"
      .AddItem "PerfFormattedData_Tcpip_NBTConnection"
      .AddItem "PerfFormattedData_Tcpip_NetworkInterface"
      .AddItem "PerfFormattedData_Tcpip_TCP"
      .AddItem "PerfFormattedData_Tcpip_UDP"
      .AddItem "PerfFormattedData_TermService_TerminalServices"
      .AddItem "PerfFormattedData_TermService_TerminalServicesSession"
      .AddItem "PerfFormattedData_W3SVC_WebService"
      .AddItem "PerfRawData"
      .AddItem "PerfRawData_ASP_ActiveServerPages"
      .AddItem "PerfRawData_ContentFilter_IndexingServiceFilter"
      .AddItem "PerfRawData_ContentIndex_IndexingService"
      .AddItem "PerfRawData_ISAPISearch_HttpIndexingService"
      .AddItem "PerfRawData_InetInfo_InternetInformationServicesGlobal"
      .AddItem "PerfRawData_MSDTC_DistributedTransactionCoordinator"
      .AddItem "PerfRawData_NTFSDRV_SMTPNTFSStoreDriver"
      .AddItem "PerfRawData_PSched_PSchedFlow"
      .AddItem "PerfRawData_PSched_PSchedPipe"
      .AddItem "PerfRawData_PerfDisk_LogicalDisk"
      .AddItem "PerfRawData_PerfDisk_PhysicalDisk"
      .AddItem "PerfRawData_PerfNet_Browser"
      .AddItem "PerfRawData_PerfNet_Redirector"
      .AddItem "PerfRawData_PerfNet_Server"
      .AddItem "PerfRawData_PerfNet_ServerWorkQueues"
      .AddItem "PerfRawData_PerfOS_Cache"
      .AddItem "PerfRawData_PerfOS_Memory"
      .AddItem "PerfRawData_PerfOS_Objects"
      .AddItem "PerfRawData_PerfOS_PagingFile"
      .AddItem "PerfRawData_PerfOS_Processor"
      .AddItem "PerfRawData_PerfOS_System"
      .AddItem "PerfRawData_PerfProc_FullImage_Costly"
      .AddItem "PerfRawData_PerfProc_Image_Costly"
      .AddItem "PerfRawData_PerfProc_JobObject"
      .AddItem "PerfRawData_PerfProc_JobObjectDetails"
      .AddItem "PerfRawData_PerfProc_Process"
      .AddItem "PerfRawData_PerfProc_ProcessAddressSpace_Costly"
      .AddItem "PerfRawData_PerfProc_Thread"
      .AddItem "PerfRawData_PerfProc_ThreadDetails_Costly"
      .AddItem "PerfRawData_RSVP_ACSRSVPInterfaces"
      .AddItem "PerfRawData_RSVP_ACSRSVPService"
      .AddItem "PerfRawData_RemoteAccess_RASPort"
      .AddItem "PerfRawData_RemoteAccess_RASTotal"
      .AddItem "PerfRawData_SMTPSVC_SMTPServer"
      .AddItem "PerfRawData_Spooler_PrintQueue"
      .AddItem "PerfRawData_TapiSrv_Telephony"
      .AddItem "PerfRawData_Tcpip_ICMP"
      .AddItem "PerfRawData_Tcpip_IP"
      .AddItem "PerfRawData_Tcpip_NBTConnection"
      .AddItem "PerfRawData_Tcpip_NetworkInterface"
      .AddItem "PerfRawData_Tcpip_TCP"
      .AddItem "PerfRawData_Tcpip_UDP"
      .AddItem "PerfRawData_TermService_TerminalServices"
      .AddItem "PerfRawData_TermService_TerminalServicesSession"
      .AddItem "PerfRawData_W3SVC_WebService"
      .AddItem "PhysicalMedia"
      .AddItem "PhysicalMemory"
      .AddItem "PhysicalMemoryArray"
      .AddItem "PingStatus"
      .AddItem "PnPEntity"
      .AddItem "PnPSignedDriver"
      .AddItem "PointingDevice"
      .AddItem "PortConnector"
      .AddItem "PortResource"
      .AddItem "PortableBattery"
      .AddItem "PowerManagementEvent"
      .AddItem "PrintJob"
      .AddItem "Printer"
      .AddItem "PrinterConfiguration"
      .AddItem "PrinterDriver"
      .AddItem "PrivilegesStatus"
      .AddItem "Process"
      .AddItem "ProcessStartTrace"
      .AddItem "ProcessStartup"
      .AddItem "ProcessStopTrace"
      .AddItem "ProcessTrace"
      .AddItem "Processor"
      .AddItem "Product"
      .AddItem "ProgIDSpecification"
      .AddItem "ProgramGroup"
      .AddItem "ProgramGroupOrItem"
      .AddItem "Property"
      .AddItem "Proxy"
      .AddItem "PublishComponentAction"
      .AddItem "QuickFixEngineering"
      .AddItem "QuotaSetting"
      .AddItem "Refrigeration"
      .AddItem "Registry"
      .AddItem "RegistryAction"
      .AddItem "RemoveFileAction"
      .AddItem "RemoveIniAction"
      .AddItem "ReserveCost"
      .AddItem "SCSIController"
      .AddItem "SID"
      .AddItem "SMBIOSMemory"
      .AddItem "ScheduledJob"
      .AddItem "SecuritySetting"
      .AddItem "SelfRegModuleAction"
      .AddItem "SerialPort"
      .AddItem "SerialPortConfiguration"
      .AddItem "ServerConnection"
      .AddItem "ServerSession"
      .AddItem "Service"
      .AddItem "ServiceControl"
      .AddItem "ServiceSpecification"
      .AddItem "Session"
      .AddItem "ShadowContext"
      .AddItem "ShadowCopy"
      .AddItem "ShadowProvider"
      .AddItem "Share"
      .AddItem "ShortcutAction"
      .AddItem "ShortcutFile"
      .AddItem "SoftwareElement"
      .AddItem "SoftwareElementCondition"
      .AddItem "SoftwareFeature"
      .AddItem "SoundDevice"
      .AddItem "StartupCommand"
      .AddItem "SystemAccount"
      .AddItem "SystemConfigurationChangeEvent"
      .AddItem "SystemEnclosure"
      .AddItem "SystemMemoryResource"
      .AddItem "SystemSlot"
      .AddItem "SystemTrace"
      .AddItem "TCPIPPrinterPort"
      .AddItem "TapeDrive"
      .AddItem "TemperatureProbe"
      .AddItem "Thread"
      .AddItem "ThreadStartTrace"
      .AddItem "ThreadStopTrace"
      .AddItem "ThreadTrace"
      .AddItem "TimeZone"
      .AddItem "Trustee"
      .AddItem "TypeLibraryAction"
      .AddItem "USBController"
      .AddItem "USBHub"
      .AddItem "UTCTime"
      .AddItem "UninterruptiblePowerSupply"
      .AddItem "UserAccount"
      .AddItem "VideoConfiguration"
      .AddItem "VideoController"
      .AddItem "VoltageProbe"
      .AddItem "Volume"
      .AddItem "VolumeChangeEvent"
      .AddItem "WMISetting"
      .AddItem "WindowsProductActivation"
   End With

End Sub

Private Sub lstBrowse_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Sort2(ColumnHeader.Index) = Not Sort2(ColumnHeader.Index)
   SortListView lstBrowse, ColumnHeader.Index, ldtString, Sort2(ColumnHeader.Index)
End Sub

Private Sub lstLog_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   SortType = ldtString
   PopupMenu mColsMenu
   Sort(ColumnHeader.Index) = Not Sort(ColumnHeader.Index)
   SortListView lstLog, ColumnHeader.Index, SortType, Sort(ColumnHeader.Index)
End Sub

Private Sub lstLog_ItemClick(ByVal Item As MSComctlLib.ListItem)

   With lstBrowse
  
      .ListItems.Clear
      With .ColumnHeaders
         .Clear
         .Add 1, , "Properties", "2000"
         .Add 2, , "Value", "3000"
      End With
      .Refresh
      'On Error Resume Next
      Dim lv As MSComctlLib.ListItem
      Dim i As Integer
      For i = 1 To lstLog.ColumnHeaders.Count
         Set lv = .ListItems.Add(, , lstLog.ColumnHeaders.Item(i).Text)
         If i = 1 Then
            lv.SubItems(1) = lstLog.SelectedItem.Text
         Else
            lv.SubItems(1) = lstLog.SelectedItem.SubItems(i - 1)
         End If
      Next i
   End With
End Sub

Private Sub mCols_Click(Index As Integer)
   Select Case Index
      Case 0: SortType = ldtString
      Case 1: SortType = ldtNumber
      Case 2: SortType = ldtDateTime
   End Select
End Sub
