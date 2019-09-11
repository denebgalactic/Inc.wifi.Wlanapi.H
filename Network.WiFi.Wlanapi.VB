Option Strict Off : Option Explicit On

#Region " Imports "

Imports System.Runtime.InteropServices _
      , System.Security

#End Region

Namespace Network

    <SuppressUnmanagedCodeSecurity()>
    Public Class Wlanapi

#Region " .E N U M ` S "

        Public Enum DOT11_BSS_TYPE
            dot11_BSS_type_infrastructure = 1
            dot11_BSS_type_independent = 2
            dot11_BSS_type_any = 3
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
            ShowDeniedNetworks = 1
            AllowExplicitCredentials = 4
            BlockPeriod = 5
        End Enum

        Public Enum WlanOpcodeValueType
            Invalid = 3
            QueryOnly = 0
            SetByGroupPolicy = 1
            SetByUser = 2
        End Enum

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

        Public Enum WLAN_CONNECTION_MODE
            wlan_connection_mode_profile
            wlan_connection_mode_temporary_profile
            wlan_connection_mode_discovery_secure
            wlan_connection_mode_discovery_unsecure
            wlan_connection_mode_auto
            wlan_connection_mode_invalid
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
        Public Structure DOT11_BSSID_LIST
            Public Header As NDIS_OBJECT_HEADER
            Public uNumOfEntries As UInt32
            Public uTotalNumOfEntries As UInt32
            Public BSSIDs As IntPtr
        End Structure

        <StructLayout(LayoutKind.Sequential)>
        Public Structure NDIS_OBJECT_HEADER
            Public Type As SByte
            Public Revision As SByte
            Public Size As SByte
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
        End Structure

        Public Structure Dot11Ssid
            Public SSIDLength As UInteger
            <MarshalAs(UnmanagedType.ByValArray, SizeConst:=32)>
            Public SSID As Byte()
        End Structure

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

        <DllImport("wlanapi.Dll", CharSet:=CharSet.Ansi)>
        Public Overloads Shared Function WlanEnumInterfaces(
               <[In]()> ByVal clientHandle As UInteger,
               <[In]()> ByVal pReserved As IntPtr,
               <[In]()> ByRef ppInterfaceList As IntPtr
                        ) As UInteger
        End Function

        <DllImport("wlanapi.Dll", CharSet:=CharSet.Ansi)>
        Public Overloads Shared Function WlanFreeMemory(
               <[In]()> ByVal pMemory As IntPtr
                        ) As UInteger
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

    End Class

End Namespace

