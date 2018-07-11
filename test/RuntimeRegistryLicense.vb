Imports System
Imports System.ComponentModel

Public Class RuntimeRegistryLicense
    Inherits License

    Private type As Type = Nothing

    Public Sub New(ByVal type As Type)
        Me.type = type
    End Sub

    Public Overrides ReadOnly Property LicenseKey() As String
        Get
            Return type.GUID.ToString()
        End Get
    End Property

    Public Overrides Sub Dispose()

    End Sub

End Class


