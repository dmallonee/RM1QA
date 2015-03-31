Imports Microsoft.ApplicationBlocks.Data
Imports System
Imports System.Data

Namespace BLL
    Public Class user_city
        Public Shared Function GetRows(ByVal user_id As Object) As DataRowCollection
            Return SqlHelper.ExecuteDataset(CfgKeys.ConnString, "user_city_select", _
                user_id).Tables(0).Rows
        End Function
    End Class
End Namespace

