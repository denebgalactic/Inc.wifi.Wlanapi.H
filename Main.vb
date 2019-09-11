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

Module Main

    Sub Main()
        '
        Network.CIMWin32.EnumInterfaces()
        '
        Network.CIMWin32.Disabled(NetworkInterface.GetAllNetworkInterfaces()(0).Id)
        '
        Dim [T1] As Integer = Network.CIMWin32.GetValue("TcpWindowSize") _
          , [T2] As Integer = Network.CIMWin32.GetValue("MTU")
        '
        Network.CIMWin32.SetValue("SetMTU", "MTU", [T2])
        Network.CIMWin32.SetValue("SetTcpWindowSize", "TcpWindowSize", [T1])
        Network.CIMWin32.SetValue("SetTcpipNetbios", "TcpipNetbiosOptions", 0)
        Network.CIMWin32.SetValue("SetIGMPLevel", "IGMPLevel", 2)
        Network.CIMWin32.SetValue("SetPMTUBHDetect", "PMTUBHDetectEnabled", True)
        '
        Registry.SetValue(String.Concat("HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\Interfaces\" _
                                        , NetworkInterface.GetAllNetworkInterfaces()(0).Id, "\"), "TcpInitialRtt", 3, RegistryValueKind.DWord)
        Registry.SetValue("HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\", "DontPingGateway" _
                                        , True, RegistryValueKind.DWord)
        '
        Network.WiFi.SetValue(Network.WiFi.WlanIntfOpcode.AutoconfEnabled, True)
        Network.WiFi.SetAutoConfValue(Network.WiFi.WlanAutoConfOpcode.ShowDeniedNetworks, True)
        Network.WiFi.SetAutoConfValue(Network.WiFi.WlanAutoConfOpcode.AllowExplicitCredentials, True)
        Network.WiFi.SetAutoConfValue(Network.WiFi.WlanAutoConfOpcode.AllowVirtualStationExtensibility, True)
        Network.WiFi.SetAutoConfValue(Network.WiFi.WlanAutoConfOpcode.BlockPeriod, 5000)
        '
        Network.CIMWin32.Enabled(NetworkInterface.GetAllNetworkInterfaces()(0).Id)
        '
        Console.ReadLine()
        '
    End Sub

End Module
