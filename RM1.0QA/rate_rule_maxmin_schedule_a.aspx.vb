Imports Schedules

Partial Class rate_rule_maxmin_schedule_a
    Inherits System.Web.UI.Page


    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack Then
            BindSchedules()
        End If

        btnDelete.Attributes.Add("onclick", "javascript:return confirm('Are you sure you want to delete this schedule? You cannot undo this operation.');")
    End Sub

    Private Sub BindSchedules()
        ddlSchedules.DataSource = BLL.car_rate_rule_schedule_header.GetList(Nothing, 4, Request.QueryString("user_id"))
        ddlSchedules.DataBind()
    End Sub

    Protected Sub btnEdit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnEdit.Click
        ' need to also pass the Request.QueryString("user_id") value because it is needed on the next page
        Response.Redirect(String.Format("rate_rule_maxmin_schedule_b.aspx?schedule_id={0}&user_id={1}&schedule_name={2}", ddlSchedules.SelectedValue, Request.QueryString("user_id"), ddlSchedules.SelectedItem.Text))
    End Sub



    Protected Sub btnNew_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnNew.Click
        Dim org_id As Integer = BLL.my_user.get_org_id(Request.QueryString("user_id"))
        Dim schedule_id As Integer = BLL.car_rate_rule_schedule_header.Save(Nothing, txtNewName.Text, org_id, 4)

        ddlSchedules.Items.Insert(0, New ListItem(txtNewName.Text, schedule_id.ToString()))
        ddlSchedules.SelectedIndex = 0

        btnEdit_Click(sender, e)
    End Sub

    Protected Sub btnDelete_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim id As Integer = ddlSchedules.SelectedValue

        Dim da As New SchedulesTableAdapters.QueriesTableAdapter
        Call da.car_rate_rule_schedule_delete(id)
        ddlSchedules.Items.Clear()
        BindSchedules()
    End Sub
End Class
