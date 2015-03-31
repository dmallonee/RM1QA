Imports Microsoft.ApplicationBlocks.Data
Imports System
Imports System.Data

Namespace BLL
    Public Class car_rate_rule_schedule_header
        Public Shared Function GetList(ByVal org_id As Object, ByVal schedule_type_id As Object, ByVal user_id As Object) As DataView
            Return SqlHelper.ExecuteDataset(CfgKeys.ConnString, "car_rate_rule_schedule_select", _
                New Object() {org_id, schedule_type_id, user_id}).Tables.Item(0).DefaultView
        End Function

        Public Shared Function Save(ByVal schedule_id As Object, ByVal schedule_desc As Object, ByVal org_id As Object, ByVal schedule_type_id As Object) As Integer
            Return Convert.ToInt32(SqlHelper.ExecuteScalar(CfgKeys.ConnString, "car_rate_rule_schedule_header_save", _
                New Object() {schedule_id, schedule_desc, org_id, schedule_type_id}))
        End Function
    End Class
End Namespace

