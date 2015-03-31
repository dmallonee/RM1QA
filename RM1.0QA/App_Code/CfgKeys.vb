Imports System
Imports System.Configuration

Public Class CfgKeys
    Public Shared ReadOnly Property ConnString() As String
        Get
            Return ConfigurationManager.ConnectionStrings.Item("ProductionConnectionString").ConnectionString
        End Get
    End Property
End Class
