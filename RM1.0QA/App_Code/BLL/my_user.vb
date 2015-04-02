Imports Microsoft.ApplicationBlocks.Data
Imports System
Imports System.Data

Namespace BLL
    Public Class my_user
        Public Shared Function get_org_id(ByVal user_id As Object) As Long
            Dim value As Object = SqlHelper.ExecuteDatacol( _
                CfgKeys.ConnString, "user_select", "org_id", _
                user_id, Nothing)

            Return IIf(value Is Nothing, 0, Convert.ToInt32(value))
        End Function
    End Class
End Namespace

