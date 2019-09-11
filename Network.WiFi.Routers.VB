Option Strict Off : Option Explicit On

#Region " Imports "

Imports System.Runtime.InteropServices _
      , System.Security

#End Region

Namespace Network

    <SuppressUnmanagedCodeSecurity()>
    Partial Public NotInheritable Class WiFi
        Inherits Network.Wlanapi

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
                        Dim item As ListViewItem = New ListViewItem(New String() {
                                                            Encoding.ASCII.GetString(Bss.dot11Ssid.SSID, 0, CInt(Bss.dot11Ssid.SSIDLength)),
                                                            String.Concat("SignalQuality: ", Bss.linkQuality.ToString()),
                                                            GetChannel(Bss.chCenterFrequency).ToString()})
                        ListView.Items.Add(item)
                        '
                    Next
                    WlanFreeMemory(bssListPtr)
                End If
                '
                WlanCloseHandle([clientHandle], IntPtr.Zero)
            End If
            '
        End Sub

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

    End Class

End Namespace

