Option Strict Off : Option Explicit On

#Region " Imports "

Imports System.Runtime.InteropServices _
      , System.Net.NetworkInformation _
      , System.Security.Principal _
      , System.Management _
      , System.Security

#End Region

Namespace Emicrox

    <SuppressUnmanagedCodeSecurity()>
    Public NotInheritable Class CIMWin32

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

        Shared Sub Processing()
            '
            INTERFACES = New Dictionary(Of String, ADAPTERPROPERTY)
            For Each adapter As NetworkInterface In NetworkInterface.GetAllNetworkInterfaces()
                Dim T As ADAPTERPROPERTY = New ADAPTERPROPERTY
                If Not (INTERFACES.TryGetValue(adapter.Id, T)) Then
                    T.T1 = adapter
                    INTERFACES.Add(adapter.Description, T)
                End If
            Next
            '
            For Each [T] As ManagementObject In New ManagementClass("Win32_NetworkAdapterConfiguration").GetInstances()
                With [T]
                    Dim IROPERTY As ADAPTERPROPERTY _
                                  = New ADAPTERPROPERTY
                    If Not (INTERFACES.TryGetValue(.GetPropertyValue("Description"), IROPERTY)) Then
                    Else
                        IROPERTY.T2 = [T]
                        INTERFACES(.GetPropertyValue("Description")) = IROPERTY
                    End If
                End With
            Next
            '
            For Each [T] As ManagementObject In New ManagementClass("Win32_NetworkAdapter").GetInstances()
                With [T]
                    Dim IROPERTY As ADAPTERPROPERTY _
                                  = New ADAPTERPROPERTY
                    If Not (INTERFACES.TryGetValue(.GetPropertyValue("Description"), IROPERTY)) Then
                    Else
                        IROPERTY.T3 = [T]
                        INTERFACES(.GetPropertyValue("Description")) = IROPERTY
                    End If
                End With
            Next
            '
        End Sub

        Public Shared Sub SetValue(MethodName As String, ParameterName As String, value As Integer)
            '
            With New ManagementClass("Win32_NetworkAdapterConfiguration")
                Dim [T1] As ManagementBaseObject = .GetMethodParameters(MethodName) _
                                  , [T2] As ManagementBaseObject
                [T1](ParameterName) = UInt32.Parse(value)
                [T2] = .InvokeMethod(MethodName, [T1], Nothing)
            End With
            '
        End Sub

        Public Shared Function Disabled(ByVal ID As String) As Boolean
            '
            Dim IROPERTY As ADAPTERPROPERTY _
                      = New ADAPTERPROPERTY
            If INTERFACES.TryGetValue(ID, IROPERTY) Then
                If Not (IsNothing(IROPERTY.T2)) Then
                    Dim [T1] As ManagementBaseObject = IROPERTY.T3.GetMethodParameters("Disable")
                    Dim [T2] As ManagementBaseObject _
                            = IROPERTY.T3.InvokeMethod("Disable", [T1], Nothing)
                End If
            End If
            Return True
            '
        End Function

        Public Shared Function Enabled(ByVal ID As String) As Boolean
            '
            Dim IROPERTY As ADAPTERPROPERTY _
                      = New ADAPTERPROPERTY
            If INTERFACES.TryGetValue(ID, IROPERTY) Then
                If Not (IsNothing(IROPERTY.T2)) Then
                    Dim [T1] As ManagementBaseObject = IROPERTY.T3.GetMethodParameters("Enable")
                    Dim [T2] As ManagementBaseObject _
                                = IROPERTY.T3.InvokeMethod("Enable", [T1], Nothing)
                End If
            End If
            Return True
            '
        End Function

    End Class

End Namespace
