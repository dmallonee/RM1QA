
Partial Class ProfileSearchScheduleA
    Inherits System.Web.UI.Page


    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack Then

            If IsNothing(Session("user_id")) Then
                Response.Write("error loading session.")
                Response.End()
                Exit Sub
                'Else
                '    Response.Write(Session("user_id"))
                '    Response.End()
                '    Exit Sub
            End If
            'Session("user_id") = 110
            'Session("org_id") = 35
            LoadSchedules()
        End If
    End Sub

    Private Sub LoadSchedules()
        Dim da As New DataSetProfileSearchScheduleTableAdapters.ScheduleTableAdapter
        Dim dt As New DataSetProfileSearchSchedule.ScheduleDataTable
        Dim r As DataSetProfileSearchSchedule.ScheduleRow
        Dim li As ListItem

        da.Fill(dt, CInt(Session("user_id")))
        lblMessage.Text = "there are " & dt.Rows.Count & " schedules"

        For Each r In dt.Rows
            li = New ListItem(r.schedule_desc, r.schedule_id)
            ddlSchedule.Items.Add(li)
        Next
    End Sub


    Protected Sub btnEdit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnEdit.Click
        Response.Redirect("ProfileSearchScheduleB.aspx?id=" & ddlSchedule.SelectedValue, False)
    End Sub

    Protected Sub btnNew_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnNew.Click
        Response.Redirect("ProfileSearchScheduleB.aspx?name=" & Server.UrlEncode(txtNewName.Text), False)
    End Sub
End Class
