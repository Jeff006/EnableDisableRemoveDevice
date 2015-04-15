Imports System.Runtime.InteropServices
Imports System.Security
Imports System.Runtime.ConstrainedExecution
Imports System.Text

Friend Structure ClassGuid
    Const BatteryDevices As String = "{72631e54-78a4-11d0-bcf7-00aa00b7b32a}"
    Const BiometricDevice As String = "{53D29EF7-377C-4D14-864B-EB3A85769359}"
    Const BluetoothDevices As String = "{e0cbf06c-cd8b-4647-bb8a-263b43f0f974}"
    Const CDROMDrives As String = "{4d36e965-e325-11ce-bfc1-08002be10318}"
    Const DiskDrives As String = "{4d36e967-e325-11ce-bfc1-08002be10318}"
    Const DisplayAdapters As String = "{4d36e968-e325-11ce-bfc1-08002be10318}"
    Const FloppyDiskControllers As String = "{4d36e969-e325-11ce-bfc1-08002be10318}"
    Const FloppyDiskDrives As String = "{4d36e980-e325-11ce-bfc1-08002be10318}"
    Const HardDiskControllers As String = "{4d36e96a-e325-11ce-bfc1-08002be10318}"
    Const HumanInterfaceDevices As String = "{745a17a0-74d3-11d0-b6fe-00a0c90f57da}"
    Const IEEE12844Devices As String = "{48721b56-6795-11d2-b1a8-0080c72e74a2}"
    Const IEEE12844PrintFunctions As String = "{49ce6ac8-6f86-11d2-b1e5-0080c72e74a2}"
    Const IEEE1394Devices61883Protocol As String = "{7ebefbc0-3200-11d2-b4c2-00a0C9697d07}"
    Const IEEE1394DevicesAVCProtocol As String = "{c06ff265-ae09-48f0-812c-16753d7cba83}"
    Const IEEE1394DevicesSBP2Protocol As String = "{d48179be-ec20-11d1-b6b8-00c04fa372a7}"
    Const IEEE1394HostBusController As String = "{6bdd1fc1-810f-11d0-bec7-08002be2092f}"
    Const ImagingDevice As String = "{6bdd1fc6-810f-11d0-bec7-08002be2092f}"
    Const IrDADevices As String = "{6bdd1fc5-810f-11d0-bec7-08002be2092f}"
    Const Keyboard As String = "{4d36e96b-e325-11ce-bfc1-08002be10318}"
    Const MediaChangers As String = "{ce5939ae-ebde-11d0-b181-0000f8753ec4}"
    Const MemoryTechnologyDriver As String = "{4d36e970-e325-11ce-bfc1-08002be10318}"
    Const Modem As String = "{4d36e96d-e325-11ce-bfc1-08002be10318}"
    Const Monitor As String = "{4d36e96e-e325-11ce-bfc1-08002be10318}"
    Const Mouse As String = "{4d36e96f-e325-11ce-bfc1-08002be10318}"
    Const MultifunctionDevices As String = "{4d36e971-e325-11ce-bfc1-08002be10318}"
    Const Multimedia As String = "{4d36e96c-e325-11ce-bfc1-08002be10318}"
    Const MultiportSerialAdapters As String = "{50906cb8-ba12-11d1-bf5d-0000f805f530}"
    Const NetworkAdapter As String = "{4d36e972-e325-11ce-bfc1-08002be10318}"
    Const NetworkClient As String = "{4d36e973-e325-11ce-bfc1-08002be10318}"
    Const NetworkService As String = "{4d36e974-e325-11ce-bfc1-08002be10318}"
    Const NetworkTransport As String = "{4d36e975-e325-11ce-bfc1-08002be10318}"
    Const PCISSLAccelerator As String = "{268c95a1-edfe-11d3-95c3-0010dc4050a5}"
    Const PCMCIAAdapters As String = "{4d36e977-e325-11ce-bfc1-08002be10318}"
    Const PortsCOMandLPTports As String = "{4d36e978-e325-11ce-bfc1-08002be10318}"
    Const Printers As String = "{4d36e979-e325-11ce-bfc1-08002be10318}"
    Const PrintersBusSpecific As String = "{4658ee7e-f050-11d1-b6bd-00c04fa372a7}"
    Const Processors As String = "{50127dc3-0f36-415e-a6cc-4cb3be910b65}"
    Const SCSIandRAIDControllers As String = "{4d36e97b-e325-11ce-bfc1-08002be10318}"
    Const Sensors As String = "{5175d334-c371-4806-b3ba-71fd53c9258d}"
    Const SmartCardReaders As String = "{50dd5230-ba8a-11d1-bf5d-0000f805f530}"
    Const StorageVolumes As String = "{71a27cdd-812a-11d0-bec7-08002be2092f}"
    Const SystemDevices As String = "{4d36e97d-e325-11ce-bfc1-08002be10318}"
    Const TapeDrives As String = "{6d807884-7d21-11cf-801c-08002be10318}"
    Const USBDevice As String = "{88BAE032-5A81-49f0-BC3D-A4FF138216D6}"
    Const WindowsCEUSBActiveSyncDevices As String = "{25dbce51-6c8f-4a72-8a6d-b54c2b4fc835}"
    Const WindowsPortableDevices As String = "{eec5ad98-8080-425f-922a-dabf3de3f69a}"
    Const WindowsSideShow As String = "{997b5d8d-c442-4f2e-baf3-9c8e671e9e21}"


End Structure

Friend Enum SetupDiGetDeviceRegistryProperties
    SPDRP_DEVICEDESC = &H0 'DeviceDesc (R/W)
    SPDRP_HARDWAREID = &H1 'HardwareID (R/W)
    SPDRP_COMPATIBLEIDS = &H2 'CompatibleIDs (R/W)
    SPDRP_UNUSED0 = &H3 'unused
    SPDRP_SERVICE = &H4 'Service (R/W)
    SPDRP_UNUSED1 = &H5 'unused
    SPDRP_UNUSED2 = &H6 'unused
    SPDRP_CLASS = &H7 'Class (R--tied to ClassGUID)
    SPDRP_CLASSGUID = &H8 'ClassGUID (R/W)
    SPDRP_DRIVER = &H9 'Driver (R/W)
    SPDRP_CONFIGFLAGS = &HA 'ConfigFlags (R/W)
    SPDRP_MFG = &HB 'Mfg (R/W)
    SPDRP_FRIENDLYNAME = &HC 'FriendlyName (R/W)
    SPDRP_LOCATION_INFORMATION = &HD 'LocationInformation (R/W)
    SPDRP_PHYSICAL_DEVICE_OBJECT_NAME = &HE 'PhysicalDeviceObjectName (R)
    SPDRP_CAPABILITIES = &HF 'Capabilities (R)
    SPDRP_UI_NUMBER = &H10 'UiNumber (R)
    SPDRP_UPPERFILTERS = &H11 'UpperFilters (R/W)
    SPDRP_LOWERFILTERS = &H12 'LowerFilters (R/W)
    SPDRP_BUSTYPEGUID = &H13 'BusTypeGUID (R)
    SPDRP_LEGACYBUSTYPE = &H14 'LegacyBusType (R)
    SPDRP_BUSNUMBER = &H15 'BusNumber (R)
    SPDRP_ENUMERATOR_NAME = &H16 'Enumerator Name (R)
    SPDRP_SECURITY = &H17 'Security (R/W, binary form)
    SPDRP_SECURITY_SDS = &H18 'Security (W, SDS form)
    SPDRP_DEVTYPE = &H19 'Device Type (R/W)
    SPDRP_EXCLUSIVE = &H1A 'Device is exclusive-access (R/W)
    SPDRP_CHARACTERISTICS = &H1B 'Device Characteristics (R/W)
    SPDRP_ADDRESS = &H1C 'Device Address (R)
    SPDRP_UI_NUMBER_DESC_FORMAT = &H1D 'UiNumberDescFormat (R/W)
    SPDRP_DEVICE_POWER_DATA = &H1E 'Device Power Data (R)
    SPDRP_REMOVAL_POLICY = &H1F 'Removal Policy (R)
    SPDRP_REMOVAL_POLICY_HW_DEFAULT = &H20 'Hardware Removal Policy (R)
    SPDRP_REMOVAL_POLICY_OVERRIDE = &H21 'Removal Policy Override (RW)
    SPDRP_INSTALL_STATE = &H22 'Device Install State (R)
    SPDRP_LOCATION_PATHS = &H23 'Device Location Paths (R)
    SPDRP_BASE_CONTAINERID = &H24  ' Base ContainerID (R)
End Enum

<Flags()> _
Friend Enum SetupDiGetClassDevsFlags
    [Default] = 1
    Present = 2
    AllClasses = 4
    Profile = 8
    DeviceInterface = &H10
End Enum

Friend Enum DiFunction
    SelectDevice = 1
    InstallDevice = 2
    AssignResources = 3
    Properties = 4
    Remove = 5
    FirstTimeSetup = 6
    FoundDevice = 7
    SelectClassDrivers = 8
    ValidateClassDrivers = 9
    InstallClassDrivers = &HA
    CalcDiskSpace = &HB
    DestroyPrivateData = &HC
    ValidateDriver = &HD
    Detect = &HF
    InstallWizard = &H10
    DestroyWizardData = &H11
    PropertyChange = &H12
    EnableClass = &H13
    DetectVerify = &H14
    InstallDeviceFiles = &H15
    UnRemove = &H16
    SelectBestCompatDrv = &H17
    AllowInstall = &H18
    RegisterDevice = &H19
    NewDeviceWizardPreSelect = &H1A
    NewDeviceWizardSelect = &H1B
    NewDeviceWizardPreAnalyze = &H1C
    NewDeviceWizardPostAnalyze = &H1D
    NewDeviceWizardFinishInstall = &H1E
    Unused1 = &H1F
    InstallInterfaces = &H20
    DetectCancel = &H21
    RegisterCoInstallers = &H22
    AddPropertyPageAdvanced = &H23
    AddPropertyPageBasic = &H24
    Reserved1 = &H25
    Troubleshooter = &H26
    PowerMessageWake = &H27
    AddRemotePropertyPageAdvanced = &H28
    UpdateDriverUI = &H29
    Reserved2 = &H30
End Enum

Friend Enum StateChangeAction
    Enable = 1
    Disable = 2
    PropChange = 3
    Start = 4
    [Stop] = 5
End Enum

<Flags()> _
Friend Enum Scopes
    [Global] = 1
    ConfigSpecific = 2
    ConfigGeneral = 4
End Enum

Friend Enum SetupApiError
    NoAssociatedClass = &HE0000200
    ClassMismatch = &HE0000201
    DuplicateFound = &HE0000202
    NoDriverSelected = &HE0000203
    KeyDoesNotExist = &HE0000204
    InvalidDevinstName = &HE0000205
    InvalidClass = &HE0000206
    DevinstAlreadyExists = &HE0000207
    DevinfoNotRegistered = &HE0000208
    InvalidRegProperty = &HE0000209
    NoInf = &HE000020A
    NoSuchHDevinst = &HE000020B
    CantLoadClassIcon = &HE000020C
    InvalidClassInstaller = &HE000020D
    DiDoDefault = &HE000020E
    DiNoFileCopy = &HE000020F
    InvalidHwProfile = &HE0000210
    NoDeviceSelected = &HE0000211
    DevinfolistLocked = &HE0000212
    DevinfodataLocked = &HE0000213
    DiBadPath = &HE0000214
    NoClassInstallParams = &HE0000215
    FileQueueLocked = &HE0000216
    BadServiceInstallSect = &HE0000217
    NoClassDriverList = &HE0000218
    NoAssociatedService = &HE0000219
    NoDefaultDeviceInterface = &HE000021A
    DeviceInterfaceActive = &HE000021B
    DeviceInterfaceRemoved = &HE000021C
    BadInterfaceInstallSect = &HE000021D
    NoSuchInterfaceClass = &HE000021E
    InvalidReferenceString = &HE000021F
    InvalidMachineName = &HE0000220
    RemoteCommFailure = &HE0000221
    MachineUnavailable = &HE0000222
    NoConfigMgrServices = &HE0000223
    InvalidPropPageProvider = &HE0000224
    NoSuchDeviceInterface = &HE0000225
    DiPostProcessingRequired = &HE0000226
    InvalidCOInstaller = &HE0000227
    NoCompatDrivers = &HE0000228
    NoDeviceIcon = &HE0000229
    InvalidInfLogConfig = &HE000022A
    DiDontInstall = &HE000022B
    InvalidFilterDriver = &HE000022C
    NonWindowsNTDriver = &HE000022D
    NonWindowsDriver = &HE000022E
    NoCatalogForOemInf = &HE000022F
    DevInstallQueueNonNative = &HE0000230
    NotDisableable = &HE0000231
    CantRemoveDevinst = &HE0000232
    InvalidTarget = &HE0000233
    DriverNonNative = &HE0000234
    InWow64 = &HE0000235
    SetSystemRestorePoint = &HE0000236
    IncorrectlyCopiedInf = &HE0000237
    SceDisabled = &HE0000238
    UnknownException = &HE0000239
    PnpRegistryError = &HE000023A
    RemoteRequestUnsupported = &HE000023B
    NotAnInstalledOemInf = &HE000023C
    InfInUseByDevices = &HE000023D
    DiFunctionObsolete = &HE000023E
    NoAuthenticodeCatalog = &HE000023F
    AuthenticodeDisallowed = &HE0000240
    AuthenticodeTrustedPublisher = &HE0000241
    AuthenticodeTrustNotEstablished = &HE0000242
    AuthenticodePublisherNotTrusted = &HE0000243
    SignatureOSAttributeMismatch = &HE0000244
    OnlyValidateViaAuthenticode = &HE0000245
End Enum

<StructLayout(LayoutKind.Sequential)> _
Friend Structure DeviceInfoData
    Public Size As Integer
    Public ClassGuid As Guid
    Public DevInst As Integer
    Public Reserved As IntPtr
End Structure

<StructLayout(LayoutKind.Sequential)> _
Friend Structure PropertyChangeParameters
    Public Size As Integer  ' part of header. It's flattened out into 1 structure.
    Public DiFunction As DiFunction
    Public StateChange As StateChangeAction
    Public Scope As Scopes
    Public HwProfile As Integer
End Structure

<StructLayout(LayoutKind.Sequential)> _
Friend Structure RemoveParameters
    Public Size As Integer  ' part of header. It's flattened out into 1 structure.
    Public DiFunction As DiFunction
    Public Scope As Scopes
    Public HwProfile As Integer
End Structure

Friend Class NativeMethods

    Private Const setupapi As String = "setupapi.dll"

    Private Sub New()
    End Sub

    <DllImport(setupapi, CallingConvention:=CallingConvention.Winapi, SetLastError:=True)> _
    Public Shared Function SetupDiCallClassInstaller( _
    ByVal installFunction As DiFunction, _
    ByVal deviceInfoSet As SafeDeviceInfoSetHandle, _
    <[In]()> ByRef deviceInfoData As DeviceInfoData) As <MarshalAs(UnmanagedType.Bool)> Boolean
    End Function

    <DllImport(setupapi, CallingConvention:=CallingConvention.Winapi, SetLastError:=True)> _
    Public Shared Function SetupDiEnumDeviceInfo( _
    ByVal deviceInfoSet As SafeDeviceInfoSetHandle, _
    ByVal memberIndex As Integer, _
    ByRef deviceInfoData As DeviceInfoData) As <MarshalAs(UnmanagedType.Bool)> Boolean
    End Function

    <DllImport(setupapi, CallingConvention:=CallingConvention.Winapi, CharSet:=CharSet.Unicode, SetLastError:=True)> _
    Public Shared Function SetupDiGetClassDevs( _
    <[In]()> ByRef classGuid As Guid, _
    <MarshalAs(UnmanagedType.LPWStr)> ByVal enumerator As String, _
    ByVal hwndParent As IntPtr, _
    ByVal flags As SetupDiGetClassDevsFlags) As SafeDeviceInfoSetHandle
    End Function

    <DllImport(setupapi, CallingConvention:=CallingConvention.Winapi, CharSet:=CharSet.Unicode, SetLastError:=True)> _
    Public Shared Function SetupDiGetDeviceInstanceId( _
    ByVal deviceInfoSet As SafeDeviceInfoSetHandle, _
    <[In]()> ByRef did As DeviceInfoData, _
    <MarshalAs(UnmanagedType.LPTStr)> ByVal _
    deviceInstanceId As StringBuilder, _
    ByVal deviceInstanceIdSize As Integer, _
    <Out()> ByRef requiredSize As Integer) As <MarshalAs(UnmanagedType.Bool)> Boolean
    End Function

    <SuppressUnmanagedCodeSecurity()> _
    <ReliabilityContract(Consistency.WillNotCorruptState, Cer.Success)> _
    <DllImport(setupapi, CallingConvention:=CallingConvention.Winapi, SetLastError:=True)> _
    Public Shared Function SetupDiDestroyDeviceInfoList( _
    ByVal deviceInfoSet As IntPtr) As <MarshalAs(UnmanagedType.Bool)> Boolean
    End Function

    <DllImport(setupapi, CallingConvention:=CallingConvention.Winapi, SetLastError:=True)> _
    Public Shared Function SetupDiSetClassInstallParams( _
    ByVal deviceInfoSet As SafeDeviceInfoSetHandle, _
    <[In]()> ByRef deviceInfoData As DeviceInfoData, _
    <[In]()> ByRef classInstallParams As PropertyChangeParameters, _
    ByVal classInstallParamsSize As Integer) As <MarshalAs(UnmanagedType.Bool)> Boolean
    End Function

    <DllImport(setupapi, CallingConvention:=CallingConvention.Winapi, SetLastError:=True)> _
    Public Shared Function SetupDiSetClassInstallParams( _
    ByVal deviceInfoSet As SafeDeviceInfoSetHandle, _
    <[In]()> ByRef deviceInfoData As DeviceInfoData, _
    <[In]()> ByRef classRemoveParams As RemoveParameters, _
    ByVal classInstallParamsSize As Integer) As <MarshalAs(UnmanagedType.Bool)> Boolean
    End Function

    <DllImport(setupapi, CallingConvention:=CallingConvention.Winapi, SetLastError:=True)> _
    Public Shared Function SetupDiGetDeviceRegistryProperty( _
    ByVal deviceInfoSet As SafeDeviceInfoSetHandle, _
    <[In]()> ByRef deviceInfoData As DeviceInfoData, _
    ByVal propertyNumber As Integer, _
    ByVal PropertyRegDataType As IntPtr, _
    ByVal PropertyBuffer() As Byte, _
    ByVal PropertyBufferSize As Integer, _
    <Out()> ByRef requiredSize As Integer) As <MarshalAs(UnmanagedType.Bool)> Boolean
    End Function

End Class

