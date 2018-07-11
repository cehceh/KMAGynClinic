Imports System
Imports System.ComponentModel
Imports Microsoft.Win32

Public Class RegistryLicenseProvider
    Inherits LicenseProvider

    Public Overloads Overrides Function GetLicense(
    ByVal context As LicenseContext,
    ByVal type As Type,
    ByVal instance As Object,
    ByVal allowExceptions As Boolean) As License

        Dim licenseKey As RegistryKey = Registry.CurrentUser.OpenSubKey("Software\\Acme\\HostKeys")

        If context.UsageMode = LicenseUsageMode.Runtime Then
            If Not licenseKey Is Nothing Then
                Dim strLic As String = CType(licenseKey.GetValue(type.GUID.ToString()), String)
                If Not strLic Is Nothing Then
                    If String.Compare("Installed", strLic, False) = 0 Then
                        GetLicense = New RuntimeRegistryLicense(type)
                        Exit Function
                    End If
                End If

                If allowExceptions = True Then
                    Throw New LicenseException(type, instance, "Your license is invalid")
                End If

                GetLicense = Nothing
            End If
        Else
            GetLicense = New DesigntimeRegistryLicense(type)
        End If

#Disable Warning BC42105 ' Function doesn't return a value on all code paths
    End Function
#Enable Warning BC42105 ' Function doesn't return a value on all code paths
End Class

