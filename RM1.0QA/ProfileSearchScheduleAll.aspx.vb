
Partial Class ProfileSearchScheduleAll
    Inherits System.Web.UI.Page


    Private iCount As Integer = 0

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack Then
            REM Make sure to comment out for production
            'Session("user_id") = 40
            'Session("org_id") = 35

            'lblMessage.Text = Session("user_id")

            btnUpdate.Enabled = False
            LoadGroups()
            LoadSchedules()
        End If
    End Sub

    Private Sub LoadGroups()
        'Dim da As New datasetprofilesearchscheduletableadapters.schedulewithgroupincludetableadapter
        'Dim dt As New datasetprofilesearchschedule.schedulewithgroupincludedatatable

        Dim da As New datasetprofilesearchscheduletableadapters.schedulegrouptableadapter
        Dim dt As New datasetprofilesearchschedule.schedulegroupdatatable

        da.Fill(dt, CInt(Session("user_id")))

        ddlScheduleGroup.DataSource = dt
        ddlScheduleGroup.DataBind()

        Dim li As New ListItem("choose group", -1)

        ddlScheduleGroup.Items.Insert(0, li)

    End Sub

    Private Sub LoadSchedules()
        'Dim da As New DataSetProfileSearchScheduleTableAdapters.ScheduleTableAdapter
        'Dim dt As New DataSetProfileSearchSchedule.ScheduleDataTable

        'da.Fill(dt, Session("user_id"))

        'dlSchedule.DataSource = dt
        'iCount = 0
        'dlSchedule.DataBind()


    End Sub



    'Protected Sub dlSchedule_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataListItemEventArgs) Handles dlSchedule.ItemDataBound
    '    Dim lblA As Label
    '    Dim lblB As Label
    '    Dim lblC As Label
    '    Dim lblD As Label
    '    Dim lblE As Label
    '    Dim btn As Button

    '    lblA = e.Item.FindControl("lblCount")
    '    iCount = iCount + 1
    '    lblA.Text = iCount


    '    lblB = e.Item.FindControl("lblDesc")
    '    lblB.Text = e.Item.DataItem("schedule_desc")

    '    lblC = e.Item.FindControl("lblDow")
    '    Dim sDow As String = e.Item.DataItem("schedule_dow_list")
    '    sDow = Replace(sDow, "1", "Sun")
    '    sDow = Replace(sDow, "2", "Mon")
    '    sDow = Replace(sDow, "3", "Tue")
    '    sDow = Replace(sDow, "4", "Wed")
    '    sDow = Replace(sDow, "5", "Thu")
    '    sDow = Replace(sDow, "6", "Fri")
    '    sDow = Replace(sDow, "7", "Sat")
    '    lblC.Text = sDow

    '    lblD = e.Item.FindControl("lblDttm")
    '    lblD.Text = ""

    '    lblE = e.Item.FindControl("lblUpdated")
    '    lblE.Text = e.Item.DataItem("updated")

    '    btn = e.Item.FindControl("btnDelete")
    '    btn.CommandArgument = e.Item.DataItem("schedule_id")
    'End Sub

    Protected Sub gvSchedule_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles gvSchedule.RowDataBound
        If e.Row.RowType = DataControlRowType.DataRow Then
            Dim lblDow As Label
            lblDow = e.Row.FindControl("lblDow")

            Dim sDow As String = e.Row.DataItem("schedule_dow_list")
            sDow = Replace(sDow, "1", "Su")
            sDow = Replace(sDow, "2", "M")
            sDow = Replace(sDow, "3", "T")
            sDow = Replace(sDow, "4", "W")
            sDow = Replace(sDow, "5", "R")
            sDow = Replace(sDow, "6", "F")
            sDow = Replace(sDow, "7", "S")
            lblDow.Text = sDow

            Dim lblType As Label
            lblType = e.Row.FindControl("lblType")
            Select Case CInt(e.Row.DataItem("schedule_type"))
                Case 1
                    lblType.Text = "fixed time - selected days of the week"

                Case 2
                    lblType.Text = "fixed time - specific date"

                Case 3
                    lblType.Text = "fixed time - certain day of the month"

                Case 4
                    lblType.Text = "random time - selected days of the week"

                Case 5
                    lblType.Text = "random time - specific date"

                Case 6
                    lblType.Text = "random time - certain day of the month"

            End Select

            'Dim lblST As Label
            ' Dim lblUpdated As Label
            'Dim dScheduled As Date
            'Dim dUpdated As Date

            'lblST = e.Row.FindControl("lblScheduledTime")

            'dScheduled = CDate(e.Row.DataItem("scheduled_time"))

            'lblST.Text = Hour(dScheduled) & ":" & Minute(dScheduled)

            'lblUpdated = e.Row.FindControl("lblUpdated")

            'dUpdated = CDate(e.Row.DataItem("updated"))

            ' lblUpdated.Text = Month(dUpdated) & "/" & Day(dUpdated) & "/" & Year(dUpdated)

            Dim chk As checkbox
            Dim hdn As HiddenField
            chk = e.Row.FindControl("chkIncluded")
            hdn = e.Row.FindControl("schedule_id")
            hdn.Value = CInt(e.Row.DataItem("schedule_id"))
            If CInt(e.row.dataItem("included")) > 0 Then
                chk.checked = True
            End If
        End If



    End Sub


    Protected Sub ddlScheduleGroup_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddlScheduleGroup.SelectedIndexChanged
        If ddlScheduleGroup.SelectedValue = "-1" Then
            btnUpdate.Enabled = False
        Else
            btnUpdate.Enabled = True
        End If

        oDSSchedule.DataBind()
        gvSchedule.DataBind()
    End Sub

    Protected Sub btnUpdate_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnUpdate.Click
        Dim item As GridViewRow
        Dim iGroup As Integer
        Dim da As New DataSetProfileSearchScheduleTableAdapters.QueriesTableAdapter
        iGroup = CInt(ddlScheduleGroup.SelectedValue)
        For Each item In gvSchedule.Rows
            If item.RowType = DataControlRowType.DataRow Then
                Dim chk As CheckBox
                Dim hdn As HiddenField

                chk = item.FindControl("chkIncluded")
                hdn = item.FindControl("schedule_id")
                Dim iSchedule As Integer
                Dim bInclude As Integer
                If chk.Checked Then
                    bInclude = 1
                Else
                    bInclude = 0
                End If
                bInclude = IIf(chk.Checked = True, 1, 0)
                iSchedule = CInt(hdn.Value)

                da.ScheduleGroupScheduleInclude(iGroup, iSchedule, bInclude)


            End If

        Next

        lblMessage.Text = "Schedule Group updated."
    End Sub

    Protected Sub btnNewGroup_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnNewGroup.Click
        Dim dq As New DataSetProfileSearchScheduleTableAdapters.QueriesTableAdapter

        Dim iOrg As Integer
        dq.GetOrgFromUser(CInt(Session("user_id")), iOrg)

        dq.ScheduleGroupInsert(iOrg, CStr(txtNewGroup.Text))


        LoadGroups()
        LoadSchedules()
        oDSSchedule.DataBind()
        gvSchedule.DataBind()

        lblMessage.Text = "Group created: " & txtNewGroup.Text
        txtNewGroup.Text = ""

    End Sub
End Class
