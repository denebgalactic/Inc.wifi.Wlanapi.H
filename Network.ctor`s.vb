Option Strict Off : Option Explicit On

#Region " Imports "

Imports System.Runtime.InteropServices _
      , System.Net.NetworkInformation _
      , System.Security.Principal _
      , System.Management _
      , System.Security

#End Region

Namespace Network

    <SuppressUnmanagedCodeSecurity()>
    Partial Public NotInheritable Class CIMWin32

#Region " .ctor "

        Public Shared INTERFACES As Dictionary(Of String, ADAPTERPROPERTY) _
                              = New Dictionary(Of String, ADAPTERPROPERTY)

        <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Ansi, Pack:=1)>
        Public Structure ADAPTERPROPERTY
            Public [T1] As NetworkInterface _
                 , [T2] As ManagementObject _
                 , [T3] As ManagementObject
        End Structure

#End Region


    End Class

End Namespace
