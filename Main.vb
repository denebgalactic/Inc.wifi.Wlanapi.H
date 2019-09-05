#Region " Imports "

Imports System.ComponentModel _
      , System.Net.NetworkInformation _
      , System.Security.Principal _
      , System.Drawing.Design _
      , System.Reflection _
      , System.Threading _
      , System.Management _
      , Microsoft.Win32 _
      , System.Text _
      , System.IO

#End Region

Public Class Main

    Sub Main()
        '
        Emicrox.CIMWin32.Disabled(NetworkInterface.GetAllNetworkInterfaces()(0).Id)
        '
        'Win32_NetworkAdapterConfiguration:
        Emicrox.CIMWin32.SetValue("SetMTU", "MTU", 1500)
        Emicrox.CIMWin32.SetValue("SetKeepAliveInterval", "KeepAliveInterval", 1000)
        Emicrox.CIMWin32.SetValue("SetTcpipNetbios", "TcpipNetbiosOptions", 0)
        Emicrox.CIMWin32.SetValue("SetIGMPLevel", "IGMPLevel", 2)
        Emicrox.CIMWin32.SetValue("SetTcpWindowSize", "TcpWindowSize", 16384)
        Emicrox.CIMWin32.SetValue("SetNumForwardPackets", "NumForwardPackets", 128)
        Emicrox.CIMWin32.SetValue("EnableDHCP", "DHCPEnabled", True)
        Emicrox.CIMWin32.SetValue("SetPMTUBHDetect", "PMTUBHDetectEnabled", True)
        '......
        '...
        '
        'registry
        Registry.SetValue(String.Concat("HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\Interfaces\" _
                                        , NetworkInterface.GetAllNetworkInterfaces()(0).Id, "\"), "TcpInitialRtt", 3, RegistryValueKind.DWord)
        Registry.SetValue("HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\", "DontPingGateway" _
                                        , True, RegistryValueKind.DWord)
        '......
        '...
        '
        'WiFi:
        Emicrox.Wlanapi_H.SetValue(Emicrox.Wlanapi_H.WlanIntfOpcode.AutoconfEnabled, True)
        Emicrox.Wlanapi_H.SetAutoConfValue(Emicrox.Wlanapi_H.WlanAutoConfOpcode.ShowDeniedNetworks, True)
        Emicrox.Wlanapi_H.SetAutoConfValue(Emicrox.Wlanapi_H.WlanAutoConfOpcode.AllowExplicitCredentials, True)
        Emicrox.Wlanapi_H.SetAutoConfValue(Emicrox.Wlanapi_H.WlanAutoConfOpcode.AllowVirtualStationExtensibility, True)
        Emicrox.Wlanapi_H.SetAutoConfValue(Emicrox.Wlanapi_H.WlanAutoConfOpcode.BlockPeriod, 5000)
        '......
        '...
        '
        Emicrox.CIMWin32.Enabled(NetworkInterface.GetAllNetworkInterfaces()(0).Id)
        '
        Console.ReadLine()
        '
    End Sub

End Class
