Option Strict Off : Option Explicit On

#Region " Imports "

Imports System.Runtime.InteropServices _
      , System.Security

#End Region

Namespace Emicrox

    <SuppressUnmanagedCodeSecurity()>
    Public NotInheritable Class Wlanapi_H

#Region " .E N U M ` S "

        Public Enum DOT11_AUTH_ALGORITHM
            DOT11_AUTH_ALGO_80211_OPEN = 1
            DOT11_AUTH_ALGO_80211_SHARED_KEY = 2
            DOT11_AUTH_ALGO_WPA = 3
            DOT11_AUTH_ALGO_WPA_PSK = 4
            DOT11_AUTH_ALGO_WPA_NONE = 5
            DOT11_AUTH_ALGO_RSNA = 6
            DOT11_AUTH_ALGO_RSNA_PSK = 7
            DOT11_AUTH_ALGO_IHV_START = -2147483648
            DOT11_AUTH_ALGO_IHV_END = -1
        End Enum

        Public Enum DOT11_BSS_TYPE
            dot11_BSS_type_infrastructure = 1
            dot11_BSS_type_independent = 2
            dot11_BSS_type_any = 3
        End Enum

        Public Enum DOT11_CIPHER_ALGORITHM
            DOT11_CIPHER_ALGO_NONE = 0
            DOT11_CIPHER_ALGO_WEP40 = 1
            DOT11_CIPHER_ALGO_TKIP = 2
            DOT11_CIPHER_ALGO_CCMP = 4
            DOT11_CIPHER_ALGO_WEP104 = 5
            DOT11_CIPHER_ALGO_WPA_USE_GROUP = 256
            DOT11_CIPHER_ALGO_RSN_USE_GROUP = 256
            DOT11_CIPHER_ALGO_WEP = 257
            DOT11_CIPHER_ALGO_IHV_START = -2147483648
            DOT11_CIPHER_ALGO_IHV_END = -1
        End Enum

        Public Enum DOT11_PHY_TYPE
            dot11_phy_type_unknown
            dot11_phy_type_any
            dot11_phy_type_fhss
            dot11_phy_type_dsss
            dot11_phy_type_irbaseband
            dot11_phy_type_ofdm
            dot11_phy_type_hrdsss
            dot11_phy_type_erp
            dot11_phy_type_ht
            dot11_phy_type_IHV_start
            dot11_phy_type_IHV_end
        End Enum

        Public Enum Dot11AuthAlgorithm
            IEEE80211_Open = 1
            IEEE80211_SharedKey = 2
            IHV_End = -1
            IHV_Start = -2147483648
            RSNA = 6
            RSNA_PSK = 7
            WPA = 3
            WPA_None = 5
            WPA_PSK = 4
        End Enum

        Public Enum Dot11CipherAlgorithm
            CCMP = 4
            IHV_End = -1
            IHV_Start = -2147483648
            None = 0
            RSN_UseGroup = 256
            TKIP = 2
            WEP = 257
            WEP104 = 5
            WEP40 = 1
            WPA_UseGroup = 256
        End Enum

        Public Enum Dot11OperationMode
            AP = 2
            ExtensibleStation = 4
            NetworkMonitor = -2147483648
            Station = 1
            Unknown = 0
        End Enum

        Public Enum WLAN_INTERFACE_STATE
            wlan_interface_state_not_ready
            wlan_interface_state_connected
            wlan_interface_state_ad_hoc_network_formed
            wlan_interface_state_disconnecting
            wlan_interface_state_disconnected
            wlan_interface_state_associating
            wlan_interface_state_discovering
            wlan_interface_state_authenticating
        End Enum

        Public Enum WLAN_NOTIFICATION_SOURCE As UInteger
            None = 0UI
            All = 65535UI
            ACM = 8UI
            MSM = 16UI
            Security = 32UI
            IHV = 64UI
        End Enum

        Public Enum WlanAccess
            ExecuteAccess = 131105
            ReadAccess = 131073
            WriteAccess = 458787
        End Enum

        Public Enum WlanAdhocNetworkState
            Connected = 1
            Formed = 0
        End Enum

        Public Enum WlanAvailableNetworkFlags
            Connected = 1
            HasProfile = 2
        End Enum

        Public Enum WlanConnectionFlags
            AdhocJoinOnly = 2
            EapolPassthrough = 8
            HiddenNetwork = 1
            IgnorePrivacyBit = 4
        End Enum

        Public Enum WlanConnectionMode
            Auto = 4
            DiscoverySecure = 2
            DiscoveryUnsecure = 3
            Invalid = 5
            Profile = 0
            TemporaryProfile = 1
        End Enum

        Public Enum WlanGetAvailableNetworkFlags
            IncludeAllAdhocProfiles = 1
            IncludeAllManualHiddenProfiles = 2
        End Enum

        Public Enum WlanInterfaceState
            AdHocNetworkFormed = 2
            Associating = 5
            Authenticating = 7
            Connected = 1
            Disconnected = 4
            Disconnecting = 3
            Discovering = 6
            NotReady = 0
        End Enum

        Public Enum WlanNotificationCodeAcm
            AdhocNetworkStateChange = 22
            AutoconfDisabled = 2
            AutoconfEnabled = 1
            BackgroundScanDisabled = 4
            BackgroundScanEnabled = 3
            BssTypeChange = 5
            ConnectionAttemptFail = 11
            ConnectionComplete = 10
            ConnectionStart = 9
            Disconnected = 21
            Disconnecting = 20
            FilterListChange = 12
            InterfaceArrival = 13
            InterfaceRemoval = 14
            NetworkAvailable = 19
            NetworkNotAvailable = 18
            PowerSettingChange = 6
            ProfileChange = 15
            ProfileNameChange = 16
            ProfilesExhausted = 17
            ScanComplete = 7
            ScanFail = 8
        End Enum

        Public Enum WlanNotificationCodeMsm
            AdapterOperationModeChange = 14
            AdapterRemoval = 13
            Associated = 2
            Associating = 1
            Authenticating = 3
            Connected = 4
            Disassociating = 9
            Disconnected = 10
            PeerJoin = 11
            PeerLeave = 12
            RadioStateChange = 7
            RoamingEnd = 6
            RoamingStart = 5
            SignalQualityChange = 8
        End Enum

        <Flags>
        Public Enum WlanNotificationSource
            ACM = 8
            All = 65535
            IHV = 64
            MSM = 16
            None = 0
            Security = 32
        End Enum

        <Flags>
        Public Enum WlanProfileFlags
            AllUser
            GroupPolicy
            User
        End Enum

        Public Enum WlanReasonCode
            AC_BASE = 131072
            AC_CONNECT_BASE = 163840
            AC_END = 196607
            ADHOC_SECURITY_FAILURE = 229386
            ASSOCIATION_FAILURE = 229378
            ASSOCIATION_TIMEOUT = 229379
            AUTO_SWITCH_SET_FOR_ADHOC = 524304
            AUTO_SWITCH_SET_FOR_MANUAL_CONNECTION = 524305
            BASE = 131072
            BSS_TYPE_NOT_ALLOWED = 163845
            BSS_TYPE_UNMATCH = 196611
            CONFLICT_SECURITY = 524299
            CONNECT_CALL_FAIL = 163849
            DATARATE_UNMATCH = 196613
            DISCONNECT_TIMEOUT = 229391
            DRIVER_DISCONNECTED = 229387
            DRIVER_OPERATION_FAILURE = 229388
            GP_DENIED = 163843
            IHV_NOT_AVAILABLE = 229389
            IHV_NOT_RESPONDING = 229390
            IHV_OUI_MISMATCH = 524296
            IHV_OUI_MISSING = 524297
            IHV_SECURITY_NOT_SUPPORTED = 524295
            IHV_SECURITY_ONEX_MISSING = 524306
            IHV_SETTINGS_MISSING = 524298
            IN_BLOCKED_LIST = 163847
            IN_FAILED_LIST = 163846
            INTERNAL_FAILURE = 229392
            INVALID_ADHOC_CONNECTION_MODE = 524302
            INVALID_BSS_TYPE = 524301
            INVALID_PHY_TYPE = 524293
            INVALID_PROFILE_NAME = 524291
            INVALID_PROFILE_SCHEMA = 524289
            INVALID_PROFILE_TYPE = 524292
            KEY_MISMATCH = 163853
            MSM_BASE = 196608
            MSM_CONNECT_BASE = 229376
            MSM_END = 262143
            MSM_SECURITY_MISSING = 524294
            MSMSEC_AUTH_START_TIMEOUT = 294914
            MSMSEC_AUTH_SUCCESS_TIMEOUT = 294915
            MSMSEC_BASE = 262144
            MSMSEC_CANCELLED = 294929
            MSMSEC_CAPABILITY_DISCOVERY = 262165
            MSMSEC_CAPABILITY_NETWORK = 262162
            MSMSEC_CAPABILITY_NIC = 262163
            MSMSEC_CAPABILITY_PROFILE = 262164
            MSMSEC_CAPABILITY_PROFILE_AUTH = 262174
            MSMSEC_CAPABILITY_PROFILE_CIPHER = 262175
            MSMSEC_CONNECT_BASE = 294912
            MSMSEC_DOWNGRADE_DETECTED = 294931
            MSMSEC_END = 327679
            MSMSEC_FORCED_FAILURE = 294933
            MSMSEC_G1_MISSING_GRP_KEY = 294925
            MSMSEC_G1_MISSING_KEY_DATA = 294924
            MSMSEC_KEY_FORMAT = 294930
            MSMSEC_KEY_START_TIMEOUT = 294916
            MSMSEC_KEY_SUCCESS_TIMEOUT = 294917
            MSMSEC_M3_MISSING_GRP_KEY = 294920
            MSMSEC_M3_MISSING_IE = 294919
            MSMSEC_M3_MISSING_KEY_DATA = 294918
            MSMSEC_MAX = 327679
            MSMSEC_MIN = 262144
            MSMSEC_MIXED_CELL = 262169
            MSMSEC_NIC_FAILURE = 294928
            MSMSEC_NO_AUTHENTICATOR = 294927
            MSMSEC_NO_PAIRWISE_KEY = 294923
            MSMSEC_PEER_INDICATED_INSECURE = 294926
            MSMSEC_PR_IE_MATCHING = 294921
            MSMSEC_PROFILE_AUTH_TIMERS_INVALID = 262170
            MSMSEC_PROFILE_DUPLICATE_AUTH_CIPHER = 262151
            MSMSEC_PROFILE_INVALID_AUTH_CIPHER = 262153
            MSMSEC_PROFILE_INVALID_GKEY_INTV = 262171
            MSMSEC_PROFILE_INVALID_KEY_INDEX = 262145
            MSMSEC_PROFILE_INVALID_PMKCACHE_MODE = 262156
            MSMSEC_PROFILE_INVALID_PMKCACHE_SIZE = 262157
            MSMSEC_PROFILE_INVALID_PMKCACHE_TTL = 262158
            MSMSEC_PROFILE_INVALID_PREAUTH_MODE = 262159
            MSMSEC_PROFILE_INVALID_PREAUTH_THROTTLE = 262160
            MSMSEC_PROFILE_KEY_LENGTH = 262147
            MSMSEC_PROFILE_KEY_UNMAPPED_CHAR = 262173
            MSMSEC_PROFILE_KEYMATERIAL_CHAR = 262167
            MSMSEC_PROFILE_NO_AUTH_CIPHER_SPECIFIED = 262149
            MSMSEC_PROFILE_ONEX_DISABLED = 262154
            MSMSEC_PROFILE_ONEX_ENABLED = 262155
            MSMSEC_PROFILE_PASSPHRASE_CHAR = 262166
            MSMSEC_PROFILE_PREAUTH_ONLY_ENABLED = 262161
            MSMSEC_PROFILE_PSK_LENGTH = 262148
            MSMSEC_PROFILE_PSK_PRESENT = 262146
            MSMSEC_PROFILE_RAWDATA_INVALID = 262152
            MSMSEC_PROFILE_TOO_MANY_AUTH_CIPHER_SPECIFIED = 262150
            MSMSEC_PROFILE_WRONG_KEYTYPE = 262168
            MSMSEC_PSK_MISMATCH_SUSPECTED = 294932
            MSMSEC_SEC_IE_MATCHING = 294922
            MSMSEC_SECURITY_UI_FAILURE = 294934
            MSMSEC_TRANSITION_NETWORK = 262172
            MSMSEC_UI_REQUEST_FAILURE = 294913
            NETWORK_NOT_AVAILABLE = 163851
            NETWORK_NOT_COMPATIBLE = 131073
            NO_AUTO_CONNECTION = 163841
            NON_BROADCAST_SET_FOR_ADHOC = 524303
            NOT_VISIBLE = 163842
            PHY_TYPE_UNMATCH = 196612
            PRE_SECURITY_FAILURE = 229380
            PROFILE_BASE = 524288
            PROFILE_CHANGED_OR_DELETED = 163852
            PROFILE_CONNECT_BASE = 557056
            PROFILE_END = 589823
            PROFILE_MISSING = 524290
            PROFILE_NOT_COMPATIBLE = 131074
            PROFILE_SSID_INVALID = 524307
            RANGE_SIZE = 65536
            ROAMING_FAILURE = 229384
            ROAMING_SECURITY_FAILURE = 229385
            SCAN_CALL_FAIL = 163850
            SECURITY_FAILURE = 229382
            SECURITY_MISSING = 524300
            SECURITY_TIMEOUT = 229383
            SSID_LIST_TOO_LONG = 163848
            START_SECURITY_FAILURE = 229381
            Success = 0
            TOO_MANY_SECURITY_ATTEMPTS = 229394
            TOO_MANY_SSID = 524308
            UI_REQUEST_TIMEOUT = 229393
            UNKNOWN = 65537
            UNSUPPORTED_SECURITY_SET = 196610
            UNSUPPORTED_SECURITY_SET_BY_OS = 196609
            USER_CANCELLED = 229377
            USER_DENIED = 163844
            USER_NOT_RESPOND = 163854
        End Enum

#End Region

#Region " .Structure`s "

        Public Structure DOT11_SSID
            Public uSSIDLength As UInteger
            Public ucSSID As Char()
            Public Sub New(ByVal p As IntPtr)
                Me.uSSIDLength = Convert.ToUInt32(Marshal.ReadInt32(p, 0))
                Me.ucSSID = New Char(31) {}
                For index As Integer = 0 To 32 - 1
                    Dim ch As Char = Convert.ToChar(Marshal.ReadByte(p, 4 + index))
                    Me.ucSSID(index) = ch
                Next
            End Sub
        End Structure

        Public Structure WLAN_INTERFACE_INFO_LIST
            Public dwNumberofItems As Integer
            Public dwIndex As Integer
            Public InterfaceInfo As WLAN_INTERFACE_INFO()
            Public Sub New(ByVal pList As IntPtr)
                Me.dwNumberofItems = Marshal.ReadInt32(pList, 0)
                Me.dwIndex = Marshal.ReadInt32(pList, 4)
                Me.InterfaceInfo = New WLAN_INTERFACE_INFO(Me.dwNumberofItems - 1) {}
                For index As Integer = 0 To Me.dwNumberofItems - 1
                    Dim ptr As IntPtr = New IntPtr(pList.ToInt32() + index * 532 + 8)
                    Dim wlanInterfaceInfo As WLAN_INTERFACE_INFO = New WLAN_INTERFACE_INFO()
                    Dim [structure] As WLAN_INTERFACE_INFO = CType(Marshal.PtrToStructure(ptr, GetType(WLAN_INTERFACE_INFO)), WLAN_INTERFACE_INFO)
                    Me.InterfaceInfo(index) = [structure]
                Next
            End Sub
        End Structure

        <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)>
        Public Structure WLAN_INTERFACE_INFO
            Public InterfaceGuid As Guid
            <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=256)>
            Public strInterfaceDescription As String
            Public isState As WLAN_INTERFACE_STATE
        End Structure

        <StructLayout(LayoutKind.Sequential)>
        Public Structure WLAN_CONNECTION_PARAMETERS
            Public wlanConnectionMode As WLAN_CONNECTION_MODE
            Public strProfile As String
            Public pDot11Ssid As IntPtr
            Public pDesiredBssidList As IntPtr
            Public dot11BssType As DOT11_BSS_TYPE
            Public dwFlags As Int32
        End Structure
        <StructLayout(LayoutKind.Sequential)>
        Public Structure DOT11_BSSID_LIST
            Public Header As NDIS_OBJECT_HEADER
            Public uNumOfEntries As UInt32
            Public uTotalNumOfEntries As UInt32
            Public BSSIDs As IntPtr
        End Structure
        <StructLayout(LayoutKind.Sequential)>
        Public Structure DOT11_MAC_ADDRESS
            <MarshalAs(UnmanagedType.ByValArray, ArraySubType:=UnmanagedType.U8, SizeConst:=6)> Public ucDot11MacAddress() As SByte
        End Structure
        <StructLayout(LayoutKind.Sequential)>
        Public Structure NDIS_OBJECT_HEADER
            Public Type As SByte
            Public Revision As SByte
            Public Size As SByte
        End Structure
        Public Enum WLAN_CONNECTION_MODE
            wlan_connection_mode_profile
            wlan_connection_mode_temporary_profile
            wlan_connection_mode_discovery_secure
            wlan_connection_mode_discovery_unsecure
            wlan_connection_mode_auto
            wlan_connection_mode_invalid
        End Enum

        '<StructLayout(LayoutKind.Sequential)>
        'Friend Structure WlanBssListHeader
        '    Friend totalSize As UInt32
        '    Friend numberOfItems As UInt32
        'End Structure

        'Public Structure WlanRateSet
        '    Private rateSetLength As UInteger
        '    <MarshalAs(UnmanagedType.ByValArray, SizeConst:=126)>
        '    Private rateSet As UShort()

        '    Public ReadOnly Property Rates As UShort()
        '        Get
        '            Dim numArray As UShort() = New UShort(CType((Me.rateSetLength / 2UI), IntPtr) - 1) {}
        '            Array.Copy(CType(Me.rateSet, Array), CType(numArray, Array), numArray.Length)
        '            Return numArray
        '        End Get
        '    End Property

        '    Public Function GetRateInMbps(ByVal rate As Integer) As Double
        '        Return CDbl((CInt(Me.rateSet(rate)) And CInt(Short.MaxValue))) * 0.5
        '    End Function
        'End Structure
        'Public Structure WlanBssEntry
        '    Public dot11Ssid As DOT11_SSID
        '    Public phyId As UInteger
        '    <MarshalAs(UnmanagedType.ByValArray, SizeConst:=6)>
        '    Public dot11Bssid As Byte()
        '    Public dot11BssType As Dot11BssType
        '    Public dot11BssPhyType As Dot11PhyType
        '    Public rssi As Integer
        '    Public linkQuality As UInteger
        '    Public inRegDomain As Boolean
        '    Public beaconPeriod As UShort
        '    Public timestamp As ULong
        '    Public hostTimestamp As ULong
        '    Public capabilityInformation As UShort
        '    Public chCenterFrequency As UInteger
        '    Public wlanRateSet As WlanRateSet
        '    Public ieOffset As UInteger
        '    Public ieSize As UInteger
        'End Structure

#End Region


        <DllImport("wlanapi.Dll", CharSet:=CharSet.Ansi)>
        Public Overloads Shared Function WlanOpenHandle(
               <[In]()> ByVal clientVersion As UInteger,
               <[In]()> ByVal pReserved As IntPtr,
               <[In]()> ByRef negotiatedVersion As UInteger,
               <[In]()> ByRef clientHandle As IntPtr
                        ) As UInteger
        End Function

        <DllImport("wlanapi.Dll", CharSet:=CharSet.Ansi)>
        Public Overloads Shared Function WlanCloseHandle(
               <[In]()> ByVal clientHandle As IntPtr,
               <[In]()> ByVal pReserved As IntPtr
                        ) As UInteger
        End Function

        <DllImport("wlanapi.dll", SetLastError:=True)>
        Public Shared Function WlanConnect(
                <[In]()> ByVal hClientHandle As IntPtr,
                <[In]()> ByRef pInterfaceGuid As Guid,
                <[In]()> ByRef pConnectionParameters As WLAN_CONNECTION_PARAMETERS,
                <[In]()> ByVal pReserved As IntPtr) As UInteger
        End Function

        <DllImport("wlanapi.Dll", CharSet:=CharSet.Ansi)>
        Public Overloads Shared Function WlanDisconnect(
               <[In](), MarshalAs(UnmanagedType.LPStruct)> ByVal clientHandle As IntPtr,
               <[In](), MarshalAs(UnmanagedType.LPStruct)> ByVal interfaceGuid As Guid,
               <[In](), MarshalAs(UnmanagedType.LPStruct)> ByRef pReserved As IntPtr
                        ) As UInteger
        End Function

        <DllImport("wlanapi.Dll", CharSet:=CharSet.Ansi)>
        Public Overloads Shared Function WlanEnumInterfaces(
               <[In]()> ByVal clientHandle As UInteger,
               <[In]()> ByVal pReserved As IntPtr,
               <[In]()> ByRef ppInterfaceList As IntPtr
                        ) As UInteger
        End Function

        <DllImport("wlanapi.Dll", CharSet:=CharSet.Ansi)>
        Public Overloads Shared Function WlanScan(
               <[In]()> ByVal clientHandle As IntPtr,
               <[In]()> ByVal interfaceGuid As Guid,
               <[In]()> ByVal pDot11Ssid As IntPtr,
               <[In]()> ByVal pIeData As IntPtr,
               <[In]()> ByVal pReserved As IntPtr
                        ) As UInteger
        End Function

        <DllImport("wlanapi.Dll", CharSet:=CharSet.Ansi)>
        Public Overloads Shared Function WlanFreeMemory(
               <[In]()> ByVal pMemory As IntPtr
                        ) As UInteger
        End Function

        <DllImport("wlanapi.dll", SetLastError:=True)>
        Private Shared Function WlanSetProfile(
                <[In]> ByVal hClientHandle As IntPtr,
                <[In]> ByVal pInterfaceGuid As IntPtr,
                <[In]> ByVal dwFlags As UInteger,
                <[In]> ByVal strProfileXml As String,
                <[In]> ByVal strAllUserProfileSecurity As String,
                <[In]> ByVal bOverwrite As Boolean,
                <[In], Out> ByVal pReserved As IntPtr,
                <[In], Out> ByRef pdwReasonCode As UInteger) As Integer
        End Function
        <DllImport("wlanapi.dll", SetLastError:=True)>
        Private Shared Function WlanGetProfile(
            <[In]> ByVal hClientHandle As IntPtr,
            <[In]> ByVal pInterfaceGuid As IntPtr,
            <[In]> ByVal strProfileName As String,
            <[In]> ByVal pReserved As IntPtr,
            <[In], Out> ByRef pstrProfileXml As IntPtr,
            <[In], Out> ByRef pdwFlags As IntPtr,
            <[In], Out> ByRef pdwGrantedAccess As IntPtr) As Integer
        End Function

        Public Structure WLAN_AVAILABLE_NETWORK_LIST
            Public dwNumberOfItems As UInteger
            Public dwIndex As UInteger
            Public Networks As WLAN_AVAILABLE_NETWORK()
            Public Sub New(ByVal list As IntPtr)
                Me.dwNumberOfItems = CUInt(Marshal.ReadInt32(list, 0))
                Me.dwIndex = CUInt(Marshal.ReadInt32(list, 4))
                Me.Networks = New WLAN_AVAILABLE_NETWORK(CInt(Me.dwNumberOfItems) - 1) {}

                Try
                    Dim index As Integer = 0

                    While CLng(index) < CLng(Me.dwNumberOfItems)
                        Me.Networks(index).strProfileName = Marshal.PtrToStringUni(New IntPtr(list.ToInt32() + 8))
                        Me.Networks(index).dot11Ssid = New DOT11_SSID(New IntPtr(list.ToInt32() + 520))
                        Me.Networks(index).dot11BssType = CType(Marshal.ReadInt32(list, 556), DOT11_BSS_TYPE)
                        Me.Networks(index).uNumberOfBssids = CUInt(Marshal.ReadInt32(list, 560))
                        Me.Networks(index).bNetworkConnectable = Marshal.ReadByte(list, 564) = CByte(1)
                        Me.Networks(index).wlanNotConnectableReason = CUInt(Marshal.ReadInt32(list, 568))
                        Me.Networks(index).uNumberOfPhyTypes = CUInt(Marshal.ReadInt32(list, 572))
                        Me.Networks(index).bMorePhyTypes = Marshal.ReadByte(list, 608) = CByte(1)
                        Me.Networks(index).wlanSignalQuality = CUInt(Marshal.ReadInt32(list, 612))
                        Me.Networks(index).bSecurityEnabled = Marshal.ReadByte(list, 616) = CByte(1)
                        Me.Networks(index).dot11DefaultAuthAlgorithm = CType(Marshal.ReadInt32(list, 620), DOT11_AUTH_ALGORITHM)
                        Me.Networks(index).dot11DefaultCipherAlgorithm = CType(Marshal.ReadInt32(list, 624), DOT11_CIPHER_ALGORITHM)
                        Me.Networks(index).dwFlags = CUInt(Marshal.ReadInt32(list, 628))
                        Me.Networks(index).dwReserved = CUInt(Marshal.ReadInt32(list, 632))
                        Dim num As UInteger = CUInt(Marshal.ReadInt32(list, 636))
                        list = New IntPtr(list.ToInt32() + 632 - 4)
                        index += 1
                    End While

                Catch ex As Exception
                    Console.WriteLine(ex.ToString())
                End Try
            End Sub
        End Structure

        Public Structure WLAN_AVAILABLE_NETWORK
            <MarshalAs(UnmanagedType.LPWStr)>
            Public strProfileName As String
            Public dot11Ssid As DOT11_SSID
            Public dot11BssType As DOT11_BSS_TYPE
            Public uNumberOfBssids As UInteger
            Public bNetworkConnectable As Boolean
            Public wlanNotConnectableReason As UInteger
            Public uNumberOfPhyTypes As UInteger
            Public dot11PhyTypes As DOT11_PHY_TYPE()
            Public bMorePhyTypes As Boolean
            Public wlanSignalQuality As UInteger
            Public bSecurityEnabled As Boolean
            Public dot11DefaultAuthAlgorithm As DOT11_AUTH_ALGORITHM
            Public dot11DefaultCipherAlgorithm As DOT11_CIPHER_ALGORITHM
            Public dwFlags As UInteger
            Public dwReserved As UInteger
        End Structure

        Friend Structure WlanBssListHeader
            Friend totalSize As UInteger
            Friend numberOfItems As UInteger
        End Structure

        Public Structure WlanBssEntry
            Public dot11Ssid As Dot11Ssid
            Public phyId As UInteger
            <MarshalAs(UnmanagedType.ByValArray, SizeConst:=6)>
            Public dot11Bssid As Byte()
            Public dot11BssType As Dot11BssType
            Public dot11BssPhyType As Dot11PhyType
            Public rssi As Integer
            Public linkQuality As UInteger
            Public inRegDomain As Boolean
            Public beaconPeriod As UShort
            Public timestamp As ULong
            Public hostTimestamp As ULong
            Public capabilityInformation As UShort
            Public chCenterFrequency As UInteger
            Public wlanRateSet As WlanRateSet
            Public ieOffset As UInteger
            Public ieSize As UInteger
        End Structure

        Public Structure WlanRateSet
            Private rateSetLength As UInteger
            <MarshalAs(UnmanagedType.ByValArray, SizeConst:=126)>
            Private rateSet As UShort()
            'Public ReadOnly Property Rates As UShort()
            '    Get
            '        Dim numArray As UShort() = New UShort(CType((Me.rateSetLength / 2UI), IntPtr) - 1) {}
            '        Array.Copy(CType(Me.rateSet, Array), CType(numArray, Array), numArray.Length)
            '        Return numArray
            '    End Get
            'End Property
            'Public Function GetRateInMbps(ByVal rate As Integer) As Double
            '    Return CDbl((CInt(Me.rateSet(rate)) And CInt(Short.MaxValue))) * 0.5
            'End Function
        End Structure


        Public Structure Dot11Ssid
            Public SSIDLength As UInteger
            <MarshalAs(UnmanagedType.ByValArray, SizeConst:=32)>
            Public SSID As Byte()
        End Structure

        Public Enum Dot11BssType As UInteger
            Infrastructure = 1
            Independent = 2
            Any = 3
        End Enum

        Public Enum Dot11PhyType As UInteger
            Any = 0
            DSSS = 2
            ERP = 6
            FHSS = 1
            HRDSSS = 5
            IHV_End = 2147483648UI
            IHV_Start = 4294967295UI
            IrBaseband = 3
            OFDM = 4
            Unknown = 0
        End Enum

        <DllImport("wlanapi.dll")>
        Public Shared Function WlanGetAvailableNetworkList(
            <[In]> ByVal clientHandle As IntPtr,
            <MarshalAs(UnmanagedType.LPStruct), [In]> ByVal interfaceGuid As Guid,
            <[In]> ByVal flags As WlanGetAvailableNetworkFlags,
            <[In], Out> ByVal reservedPtr As IntPtr, <Out> ByRef availableNetworkListPtr As IntPtr
                        ) As Integer
        End Function

        <DllImport("wlanapi.dll")>
        Public Shared Function WlanGetNetworkBssList(
            <[In]> ByVal clientHandle As IntPtr,
            <MarshalAs(UnmanagedType.LPStruct), [In]> ByVal interfaceGuid As Guid,
            <[In]> ByVal dot11SsidInt As IntPtr,
            <[In]> ByVal dot11BssType As Dot11BssType,
            <[In]> ByVal securityEnabled As Boolean,
            ByVal reservedPtr As IntPtr,
            ByRef wlanBssList As IntPtr
                    ) As Integer
        End Function

        Private Shared Function ConvertBssListPtr(bssListPtr As IntPtr) As WlanBssEntry()
            Dim bssEntries As WlanBssEntry()
            Try
                Dim bssListHeader As WlanBssListHeader = DirectCast(Marshal.PtrToStructure(bssListPtr, GetType(WlanBssListHeader)), WlanBssListHeader)
                Dim bssListIt As Long = bssListPtr.ToInt64() + Marshal.SizeOf(GetType(WlanBssListHeader))
                bssEntries = New WlanBssEntry(bssListHeader.numberOfItems - 1) {}
                For i As Integer = 0 To bssListHeader.numberOfItems - 1
                    bssEntries(i) = DirectCast(Marshal.PtrToStructure(New IntPtr(bssListIt), GetType(WlanBssEntry)), WlanBssEntry)
                    bssListIt += Marshal.SizeOf(GetType(WlanBssEntry))
                Next
            Catch ex As Exception
            End Try
            Return bssEntries
        End Function

        Public Shared Function GetChannel(ByVal chCenterFrequency As UInteger) As Integer
            Select Case chCenterFrequency
                Case 2412000
                    Return 1
                Case 2417000
                    Return 2
                Case 2422000
                    Return 3
                Case 2427000
                    Return 4
                Case 2432000
                    Return 5
                Case 2437000
                    Return 6
                Case 2442000
                    Return 7
                Case 2447000
                    Return 8
                Case 2452000
                    Return 9
                Case 2457000
                    Return 10
                Case 2462000
                    Return 11
                Case 2467000
                    Return 12
                Case 2472000
                    Return 13
            End Select
            Return -1
        End Function

        Public Shared Function GetStringForSSID(ByVal ssid As DOT11_SSID) As String
            Dim str As String = ""
            Try
                Dim index As Integer = 0
                While CLng(index) < CLng(ssid.uSSIDLength)
                    str += CStr(CObj(ssid.ucSSID(index)))
                    index += 1
                End While
            Catch ex As Exception
            End Try
            Return str
        End Function

        <DllImport("Wlanapi", EntryPoint:="WlanQueryInterface")>
        Public Shared Function WlanQueryInterface(
                <[In]> ByVal clientHandle As IntPtr,
                <[In]> ByRef interfaceGuid As Guid,
                ByVal OpCode As WlanIntfOpcode,
                ByVal pReserved As IntPtr,
                <Out> ByRef pdwDataSize As UInteger,
                ByRef ppData As IntPtr,
                ByVal pWlanOpcodeValueType As IntPtr
                        ) As UInteger
        End Function

        <DllImport("wlanapi.Dll", CharSet:=CharSet.Ansi)>
        Public Overloads Shared Function WlanSetInterface(
                <[In]> ByVal clientHandle As IntPtr,
                <[In]> ByRef interfaceGuid As Guid,
                ByVal OpCode As WlanIntfOpcode,
                <[In]> ByVal pdwDataSize As UInteger,
                <[In]> ByVal ppData As IntPtr,
                ByVal pReserved As IntPtr
                        ) As UInteger
        End Function

        Public Enum WlanIntfOpcode
            AutoconfStart = 0
            AutoconfEnabled = 1
            BackgroundScanEnabled = 2
            BssType = 5
            ChannelNumber = 8
            CurrentConnection = 7
            CurrentOperationMode = 12
            IhvEnd = 1073741823
            IhvStart = 805306368
            InterfaceState = 6
            MediaStreamingMode = 3
            MsmStart = 268435712
            MsmEnd = 536870911
            RadioState = 4
            RSSI = 268435714
            AutoconfEnd = 268435455
            SecurityEnd = 805306367
            SecurityStart = 536936448
            Statistics = 268435713
            SupportedAdhocAuthCipherPairs = 10
            SupportedCountryOrRegionStringList = 11
            SupportedInfrastructureAuthCipherPairs = 9
        End Enum

        Public Enum WlanAutoConfOpcode
            _start = 0
            ShowDeniedNetworks = 1
            PowerSetting = 2
            OnlyUseGroupProfilesForAllowedNetworks = 3
            AllowExplicitCredentials = 4
            BlockPeriod = 5
            AllowVirtualStationExtensibility = 6
            _End = 7
        End Enum

        Public Enum WlanOpcodeValueType
            Invalid = 3
            QueryOnly = 0
            SetByGroupPolicy = 1
            SetByUser = 2
        End Enum

        <DllImport("wlanapi.Dll", CharSet:=CharSet.Ansi)>
        Public Shared Function WlanQueryAutoConfigParameter(
                <[In]> ByVal clientHandle As IntPtr,
                ByVal OpCode As WlanAutoConfOpcode,
                ByVal pReserved As IntPtr,
                <Out> ByRef pdwDataSize As UInteger,
                ByRef ppData As IntPtr,
                ByVal pWlanOpcodeValueType As IntPtr
                        ) As Integer
        End Function

        <DllImport("wlanapi.Dll", CharSet:=CharSet.Ansi)>
        Public Shared Function WlanSetAutoConfigParameter(
                <[In]> ByVal clientHandle As IntPtr,
                <[In]> ByVal OpCode As WlanAutoConfOpcode,
                <[In]> ByVal pdwDataSize As UInteger,
                <[In]> ByVal ppData As IntPtr,
                ByVal pReserved As IntPtr
                        ) As Integer
        End Function


        Public Shared Function GetValue(ByVal OpCode As WlanIntfOpcode) As Int32
            '
            Dim ret As Integer = 0
            Try
                Dim [clientHandle] As IntPtr = Nothing
                If WlanOpenHandle(1, IntPtr.Zero, IntPtr.Zero, [clientHandle]) = 0 Then
                    '
                    Dim [InterfacesPtr] As IntPtr = IntPtr.Zero
                    If WlanEnumInterfaces([clientHandle], IntPtr.Zero, [InterfacesPtr]) = 0 Then
                        For Each interfaces As WLAN_INTERFACE_INFO In New WLAN_INTERFACE_INFO_LIST([InterfacesPtr]).InterfaceInfo
                            '
                            Dim valuePtr As IntPtr
                            Dim valueSize As Integer
                            Dim opcodeValueType As WlanOpcodeValueType
                            WlanQueryInterface(clientHandle, interfaces.InterfaceGuid, OpCode, IntPtr.Zero, valueSize, valuePtr, opcodeValueType)
                            ret = Marshal.ReadInt32(valuePtr)
                            'MessageBox.Show(String.Concat(OpCode, " | value:", ret, " size:", valueSize))
                            '
                        Next
                        '
                        If [InterfacesPtr] <> IntPtr.Zero Then WlanFreeMemory([InterfacesPtr])
                    End If
                    WlanCloseHandle([clientHandle], IntPtr.Zero)
                End If
            Catch ex As Exception
                ret = 0
            End Try
            '
            Return ret
            '
        End Function

        Public Shared Function GetAutoConfValue(ByVal OpCode As WlanAutoConfOpcode) As Int32
            '
            Dim ret As Integer = 0
            Try
                Dim [clientHandle] As IntPtr = Nothing
                If WlanOpenHandle(1, IntPtr.Zero, IntPtr.Zero, [clientHandle]) = 0 Then
                    '
                    Dim [InterfacesPtr] As IntPtr = IntPtr.Zero
                    If WlanEnumInterfaces([clientHandle], IntPtr.Zero, [InterfacesPtr]) = 0 Then
                        For Each interfaces As WLAN_INTERFACE_INFO In New WLAN_INTERFACE_INFO_LIST([InterfacesPtr]).InterfaceInfo
                            '
                            Dim valueSize As Integer = 4
                            Dim valuePtr As IntPtr = Marshal.AllocHGlobal(valueSize)
                            Dim opcodeValueType As WlanOpcodeValueType
                            WlanQueryAutoConfigParameter(clientHandle _
                                    , OpCode, IntPtr.Zero, valueSize, valuePtr, opcodeValueType)
                            ret = Marshal.ReadInt32(valuePtr)
                            Marshal.FreeHGlobal(valuePtr)
                            '
                        Next
                        '
                        If [InterfacesPtr] <> IntPtr.Zero Then WlanFreeMemory([InterfacesPtr])
                    End If
                    WlanCloseHandle([clientHandle], IntPtr.Zero)
                End If
            Catch ex As Exception
                ret = 0
            End Try
            '
            Return ret
            '
        End Function


        Public Shared Sub SetValue(ByVal OpCode As WlanIntfOpcode, value As Int32)
            '
            Try
                Dim [clientHandle] As IntPtr = Nothing
                If WlanOpenHandle(1, IntPtr.Zero, IntPtr.Zero, [clientHandle]) = 0 Then
                    '
                    Dim [InterfacesPtr] As IntPtr = IntPtr.Zero
                    If WlanEnumInterfaces([clientHandle], IntPtr.Zero, [InterfacesPtr]) = 0 Then
                        For Each interfaces As WLAN_INTERFACE_INFO In New WLAN_INTERFACE_INFO_LIST([InterfacesPtr]).InterfaceInfo
                            '
                            Dim valueSize As Integer = 4
                            Dim valuePtr As IntPtr = Marshal.AllocHGlobal(valueSize)
                            Marshal.WriteInt32(valuePtr, 0, value)
                            WlanSetInterface(clientHandle, interfaces.InterfaceGuid _
                                               , OpCode, valueSize, valuePtr, IntPtr.Zero)
                            Marshal.FreeHGlobal(valuePtr)
                            '
                        Next
                        '
                        If [InterfacesPtr] <> IntPtr.Zero Then WlanFreeMemory([InterfacesPtr])
                    End If

                    WlanCloseHandle([clientHandle], IntPtr.Zero)
                End If
            Catch ex As Exception
            End Try
            '
        End Sub

        Public Shared Sub SetAutoConfValue(ByVal OpCode As WlanAutoConfOpcode, value As Int32)
            '
            Try
                Dim [clientHandle] As IntPtr = Nothing
                If WlanOpenHandle(1, IntPtr.Zero, IntPtr.Zero, [clientHandle]) = 0 Then
                    '
                    Dim [InterfacesPtr] As IntPtr = IntPtr.Zero
                    If WlanEnumInterfaces([clientHandle], IntPtr.Zero, [InterfacesPtr]) = 0 Then
                        For Each interfaces As WLAN_INTERFACE_INFO In New WLAN_INTERFACE_INFO_LIST([InterfacesPtr]).InterfaceInfo
                            '
                            Dim valueSize As Integer = 4
                            Dim valuePtr As IntPtr = Marshal.AllocHGlobal(valueSize)
                            Marshal.WriteInt32(valuePtr, 0, value)
                            Dim ret As Integer = WlanSetAutoConfigParameter(clientHandle _
                                    , OpCode, valueSize, valuePtr, IntPtr.Zero)
                            Marshal.FreeHGlobal(valuePtr)
                            '
                        Next
                        '
                        If [InterfacesPtr] <> IntPtr.Zero Then WlanFreeMemory([InterfacesPtr])
                    End If

                    WlanCloseHandle([clientHandle], IntPtr.Zero)
                End If
            Catch ex As Exception
            End Try
            '
        End Sub

        Public Shared Is_WiFi_Enable As Boolean = False

        Public Shared Sub TestWiFi()
            '
            Try
                '
                Is_WiFi_Enable = False
                Dim [clientHandle] As IntPtr = Nothing
                If WlanOpenHandle(1, IntPtr.Zero, IntPtr.Zero, [clientHandle]) = 0 Then
                    '
                    Dim [InterfacesPtr] As IntPtr = IntPtr.Zero
                    If WlanEnumInterfaces([clientHandle], IntPtr.Zero, [InterfacesPtr]) = 0 Then
                        For Each interfaces As WLAN_INTERFACE_INFO In New WLAN_INTERFACE_INFO_LIST([InterfacesPtr]).InterfaceInfo
                            '
                            Is_WiFi_Enable = True
                            '
                        Next
                        '
                        If [InterfacesPtr] <> IntPtr.Zero Then WlanFreeMemory([InterfacesPtr])
                    End If
                    WlanCloseHandle([clientHandle], IntPtr.Zero)
                End If
                '
            Catch ex As Exception
            End Try
            '
        End Sub


    End Class

End Namespace


