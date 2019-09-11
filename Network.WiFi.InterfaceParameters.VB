Option Strict Off : Option Explicit On

#Region " Imports "

Imports System.Runtime.InteropServices _
      , System.Security

#End Region

Namespace Network

    <SuppressUnmanagedCodeSecurity()>
    Partial Public NotInheritable Class WiFi
        Inherits Network.Wlanapi

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
                            Dim valuePtr As [IntPtr] _
                              , valueSize As Integer _
                              , opcodeValueType As WlanOpcodeValueType
                            WlanQueryInterface(clientHandle, interfaces.InterfaceGuid, OpCode, IntPtr.Zero, valueSize, valuePtr, opcodeValueType)
                            ret = Marshal.ReadInt32(valuePtr)
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
                            Dim valueSize As Integer = 4 _
                              , valuePtr As IntPtr = Marshal.AllocHGlobal(valueSize)
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

    End Class

End Namespace

