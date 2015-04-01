Imports Utility


Partial Class ProfileSearchScheduleB
    Inherits System.Web.UI.Page

    Private iScheduleID As Integer

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack Then

            FillTimeSelector(ddlScheduledTime)
            FillTimeSelector(ddlScheduledTime0)
            FillTimeSelector(ddlScheduledTime1)
            FillNumericalDaySelector(ddlDayOfMonth)

            If Len(Request.QueryString("id")) > 0 Then
                iScheduleID = CInt(Request.QueryString("id"))
                LoadSchedule()
                btnSubmit.Text = "Update"
            ElseIf Len(Request.Form("ddlSchedule")) > 0 Then
                iScheduleID = CInt(Request.Form("ddlSchedule"))
                LoadSchedule()
                btnSubmit.Text = "Update"
            Else

                If Len(Request.QueryString("name")) > 0 Then
                    txtName.Text = Server.UrlDecode(Request.QueryString("name"))
                End If

                SetDefaults()
                iScheduleID = -1
                btnSubmit.Text = "Create"
            End If
        End If

    End Sub

    Private Sub SetDefaults()
        rdoWeek.Checked = True
        rdoFixed.Checked = True
        chkSaveCopy.Visible = False
    End Sub

    Private Sub LoadSchedule()
        Dim da As New DataSetProfileSearchScheduleTableAdapters.ScheduleTableAdapter
        Dim dt As New DataSetProfileSearchSchedule.ScheduleDataTable
        Dim r As DataSetProfileSearchSchedule.ScheduleRow
        chkSaveCopy.Visible = True
        chkSaveCopy.Checked = False
        da.FillByID(dt, CInt(iScheduleID))

        If dt.Rows.Count > 0 Then
            Dim dow As String
            Dim dCheckTime As Date
            Dim sCheckTime As String



            '1	fixed time - selected days of the week
            '2	fixed time - specific date
            '3	fixed time - certain day of the month
            '4	random time - selected days of the week
            '5	random time - specific date
            '6	random time - certain day of the month

            r = dt.Rows(0)

            txtName.Text = r.schedule_desc
            schedule_id.Value = iScheduleID
            schedule_type.Value = r.schedule_type
            Dim iType As Integer = r.schedule_type

            Select Case iType
                Case 1
                    rdoFixed.Checked = True
                    rdoWeek.Checked = True
                Case 2
                    rdoFixed.Checked = True
                    rdoDayFixed.Checked = True
                Case 3
                    rdoFixed.Checked = True
                    rdoMonthly.Checked = True
                Case 4
                    rdoRandom.Checked = True
                    rdoWeek.Checked = True
                Case 5
                    rdoRandom.Checked = True
                    rdoDayFixed.Checked = True
                Case 6
                    rdoRandom.Checked = True
                    rdoMonthly.Checked = True

            End Select

            ' Select days of the week
            If iType = 1 Or iType = 4 Then
                dow = IIFNotNull(r.schedule_dow_list)
                If InList(1, dow) Then
                    chkDowSunday.Checked = True
                End If
                If InList(2, dow) Then
                    chkDowMonday.Checked = True
                End If
                If InList(3, dow) Then
                    chkDowTuesday.Checked = True
                End If
                If InList(4, dow) Then
                    chkDowWednesday.Checked = True
                End If
                If InList(5, dow) Then
                    chkDowThursday.Checked = True
                End If
                If InList(6, dow) Then
                    chkDowFriday.Checked = True
                End If
                If InList(7, dow) Then
                    chkDowSaturday.Checked = True
                End If
            End If

            ' fixed time
            If iType = 1 Or iType = 2 Or iType = 3 Then
                dCheckTime = r.scheduled_time

                'If Hour(dCheckTime) < 10 Then
                'sCheckTime = "0" & Hour(dCheckTime) & ":" & Minute(dCheckTime)
                'Else
                'sCheckTime = Hour(dCheckTime) & ":" & Minute(dCheckTime)
                'End If
                sCheckTime = dCheckTime.ToString("HH:mm")
                'lblMessage.Text = sCheckTime
                ddlScheduledTime.SelectedValue = sCheckTime
            
            End If

            schedule_id.Value = r.schedule_id
            schedule_type.Value = r.schedule_type

        Else
            lblMessage.Text = "Could not locate data."
        End If




    End Sub

    Private Sub FillTimeSelector(ByRef ddl As DropDownList)

        Dim li As ListItem
        Dim intIndex As Integer
        Dim intMinuteIndex As Integer
        Dim strTime As String
        Dim strTimeValue As String

        li = New ListItem("Midnight", "00:00")
        ddl.Items.Add(li)

        li = New ListItem("12:15 am", "00:15")
        ddl.Items.Add(li)

        li = New ListItem("12:30 am", "00:30")
        ddl.Items.Add(li)

        li = New ListItem("12:45 am", "00:45")
        ddl.Items.Add(li)

        For intIndex = 1 To 11
            For intMinuteIndex = 0 To 3
                Select Case intMinuteIndex
                    Case 0
                        strTime = intIndex & ":00 am"
                    Case 1
                        strTime = intIndex & ":15 am"
                    Case 2
                        strTime = intIndex & ":30 am"
                    Case 3
                        strTime = intIndex & ":45 am"
                    Case Else
                        strTime = ""
                End Select
                strTimeValue = Trim(FormatDateTime(strTime, 4))
                li = New ListItem(strTime, strTimeValue)
                ddl.Items.Add(li)
            Next
        Next

        li = New ListItem("Noon", "12:00")
        ddl.Items.Add(li)

        li = New ListItem("12:15 pm", "12:15")
        ddl.Items.Add(li)

        li = New ListItem("12:30 pm", "12:30")
        ddl.Items.Add(li)

        li = New ListItem("12:45 pm", "12:45")
        ddl.Items.Add(li)

        For intIndex = 1 To 11
            For intMinuteIndex = 0 To 3
                Select Case intMinuteIndex
                    Case 0
                        strTime = intIndex & ":00 pm"
                    Case 1
                        strTime = intIndex & ":15 pm"
                    Case 2
                        strTime = intIndex & ":30 pm"
                    Case 3
                        strTime = intIndex & ":45 pm"
                    Case Else
                        strTime = ""
                End Select
                strTimeValue = Trim(FormatDateTime(strTime, 4))
                li = New ListItem(strTime, strTimeValue)
                ddl.Items.Add(li)
            Next
        Next


    End Sub

    Private Sub FillNumericalDaySelector(ByRef ddl As DropDownList)
        Dim i As Integer
        Dim li As ListItem


        For i = 1 To 31
            li = New ListItem(i, i)
            ddl.Items.Add(li)
        Next
    End Sub

    Protected Sub btnSubmit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSubmit.Click
        Dim da As New DataSetProfileSearchScheduleTableAdapters.ScheduleTableAdapter
        Dim dt As New DataSetProfileSearchSchedule.ScheduleDataTable
        Dim dq As New DataSetProfileSearchScheduleTableAdapters.QueriesTableAdapter
        Dim r As DataSetProfileSearchSchedule.ScheduleRow
        Dim blnNew As Boolean
        Dim OrgID As Integer

        If IsInt(schedule_id.Value) Then
            ' saving an existing
            
            da.FillByID(dt, schedule_id.Value)
            blnNew = False
            r = dt.Rows(0)
        Else
            ' saving a new
            da.Fill(dt, CInt(Session("user_id")))
            blnNew = True
            r = dt.NewScheduleRow
            dq.GetOrgFromUser(CInt(Session("user_id")), OrgID)

            r.org_id = OrgID
        End If
        r.schedule_desc = txtName.Text

        If rdoFixed.Checked Then

            If rdoWeek.Checked Then
                r.schedule_type = 1
            ElseIf rdoDayFixed.Checked Then
                r.schedule_type = 2
            Else
                r.schedule_type = 3
            End If
        End If
        If rdoRandom.Checked Then
            If rdoWeek.Checked Then
                r.schedule_type = 4
            ElseIf rdoDayFixed.Checked Then
                r.schedule_type = 5
            Else
                r.schedule_type = 6
            End If
        End If

        Dim strDow As String = ""
        If chkDowSunday.Checked = True Then
            strDow = strDow & "1,"
        End If
        If chkDowMonday.Checked = True Then
            strDow = strDow & "2,"
        End If
        If chkDowTuesday.Checked = True Then
            strDow = strDow & "3,"
        End If
        If chkDowWednesday.Checked = True Then
            strDow = strDow & "4,"
        End If
        If chkDowThursday.Checked = True Then
            strDow = strDow & "5,"
        End If
        If chkDowFriday.Checked = True Then
            strDow = strDow & "6,"
        End If
        If chkDowSaturday.Checked = True Then
            strDow = strDow & "7,"
        End If

        If Len(strDow) > 0 Then
            strDow = Left(strDow, Len(strDow) - 1)
        End If
        r.schedule_dow_list = strDow

        If rdoFixed.Checked Then
            r.scheduled_time = FormatDateTime("1/1/1900 " & ddlScheduledTime.SelectedValue, DateFormat.GeneralDate)
        End If

        r.start_hr = 0
        r.end_hr = 0
        r.schedule_status = "E"
        r.updated = Now()

        If rdoRandom.Checked Then
            r.random = True
        Else
            r.random = False
        End If

        If blnNew Then
            dt.Rows.Add(r)
            da.Update(dt)
            lblMessage.Text = "Schedule has been added."
            schedule_id.Value = r.schedule_id
            btnSubmit.Text = "Update"

        Else
            da.Update(r)
            da.Update(dt)
            lblMessage.Text = "Schedule has been updated."
        End If

    End Sub
End Class
