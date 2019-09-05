Option Strict Off : Option Explicit On

#Region " Imports "

Imports System.Runtime.InteropServices _
      , System.Security

#End Region

Namespace Emicrox

    <SuppressUnmanagedCodeSecurity()>
    Public NotInheritable Class propertygrid_dynamic

        Private Sub SetPropertyGridFieldVisableState(ByVal PropName As String, ByVal state As Boolean)
            Dim theDescriptor As PropertyDescriptor _
                    = TypeDescriptor.GetProperties(Me.PropertyGrid1.SelectedObject.[GetType]())(PropName)
            Dim theDescriptorBrowsableAttribute As BrowsableAttribute = CType(theDescriptor.Attributes(GetType(BrowsableAttribute)), BrowsableAttribute)
            Dim isBrowsable As FieldInfo = theDescriptorBrowsableAttribute.[GetType]().GetField("Browsable", BindingFlags.IgnoreCase Or BindingFlags.NonPublic Or BindingFlags.Instance)
            isBrowsable.SetValue(theDescriptorBrowsableAttribute, state)
        End Sub

        Public Sub SelectPropertyGridItem(ByVal property_grid As PropertyGrid, ByVal name As String)
            '
            property_grid.SelectedGridItem = [IFinder].FindPropertyGridItem(property_grid, name)
            '
        End Sub

    End Class

End Namespace


