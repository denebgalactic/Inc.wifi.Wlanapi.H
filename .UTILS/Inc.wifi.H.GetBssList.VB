
#Region " Imports "

Imports System.Runtime.InteropServices _
      , System.Text _
      , System.IO

#End Region

Public Class wifi

#Region " .Structure`s "

    Friend Structure WlanBssListHeader
        Friend [T1] As UInteger _
             , [T2] As UInteger
    End Structure
    Public Structure Dot11Ssid
        Public [T1] As UInteger
        <MarshalAs(UnmanagedType.ByValArray, SizeConst:=32)>
        Public [T2] As Byte()
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

    Sub EnumRouters(T1 As Integer, T2 As Integer)
        '
        Dim [clientHandle] As IntPtr = IntPtr.Zero
        If WlanOpenHandle(1UI, IntPtr.Zero, Nothing, [clientHandle]) = 0 Then
            '
            Dim zero As IntPtr = IntPtr.Zero
            If WlanEnumInterfaces([clientHandle], IntPtr.Zero, zero) = 0 Then
                '
                Dim interfaceInfoList As WLAN_INTERFACE_INFO_LIST = New WLAN_INTERFACE_INFO_LIST(zero)
                Dim bssListPtr As IntPtr = 0
                WlanGetNetworkBssList([clientHandle], interfaceInfoList.InterfaceInfo(0).InterfaceGuid _
                                            , Nothing, DOT11_BSS_TYPE.dot11_BSS_type_any _
                                            , False, IntPtr.Zero, bssListPtr)
                For Each Bss As WlanBssEntry In ConvertBssListPtr(bssListPtr)
                    '

                    '
                Next
                WlanFreeMemory(bssListPtr)
            End If
            '
            WlanCloseHandle([clientHandle], IntPtr.Zero)
        End If
        '
    End Sub

End Class

