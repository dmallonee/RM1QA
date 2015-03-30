Imports ComponentArt.Web.UI
Imports System.Data

Partial Class rate_rule_maxmin_schedule_b
    Inherits System.Web.UI.Page

    Private arrLocations As ArrayList

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack Then
            'BindLocations(0)
            'BindMonths(0)
            'BindGrid()
            LoadTabStrip(0)
            txtName.Text = Request.QueryString("schedule_name")
            chkDebug.Checked = False
            linkBack.NavigateUrl = "~/rate_rule_maxmin_schedule_a.aspx?user_id=" & Request.QueryString("user_id")
        End If
    End Sub



    Private Sub LoadTabStrip(ByVal index As Integer)
        Dim newTab As ComponentArt.Web.UI.TabStripTab
        Dim da As New SchedulesTableAdapters.user_cityTableAdapter
        Dim dt As New Schedules.user_cityDataTable
        'arrLocations = New ArrayList
        da.Fill(dt, Request.QueryString("user_id"))

        'Dim r As Schedules.user_cityRow
        'Dim items As ListItemCollection = New ListItemCollection()
        'items.Add(New ListItem("Default"))
        'For Each r In dt
        '    items.Add(New ListItem(Convert.ToString(r.city_cd)))

        'Next


        'dlLocations.DataSource = dt

        'dlLocations.SelectedIndex = index
        'dlLocations.DataBind()

        Dim r As Schedules.user_cityRow
        For Each r In dt
            newTab = New ComponentArt.Web.UI.TabStripTab()
            newTab.Text = r.city_cd

            If Trim(r.city_cd.ToString) <> "Default" Then
                newTab.Value = r.city_cd
                Dim subTab As ComponentArt.Web.UI.TabStripTab
                Dim i As Integer
                For i = 1 To 12
                    subTab = New ComponentArt.Web.UI.TabStripTab()
                    subTab.Text = MonthName(i, True)
                    subTab.Value = i
                    newTab.Tabs.Add(subTab)
                Next
            Else
                newTab.Value = "***"
                Dim subTab As New ComponentArt.Web.UI.TabStripTab
                subTab.Text = "Default"
                subTab.Value = 0
                newTab.Tabs.Add(subTab)
            End If

            TabStrip1.Tabs.Add(newTab)


        Next

        TabStrip1.SelectedTab = TabStrip1.Tabs(0).Tabs(0)
        'currentLocation.Value = "***"
        'currentMonth.Value = "0"
        'currentState.Value = "0"
        BindGrid("***", 0)
    End Sub



    Private Sub TabStrip1_ItemSelected(ByVal sender As System.Object, ByVal e As ComponentArt.Web.UI.TabStripTabEventArgs) Handles TabStrip1.ItemSelected
        Dim sLocation As String = "***"
        Dim iMonth As Integer = 0

        'If currentState.Value = "1" Then
        'DoUpdate()

        'End If


        If Not IsNothing(e.Tab.ParentTab) Then
            lblCB.Text = e.Tab.ParentTab.Text & ": " & e.Tab.Text
            sLocation = e.Tab.ParentTab.Value
            iMonth = e.Tab.Value
        Else

            TabStrip1.SelectedTab = e.Tab.Tabs(0)
            lblCB.Text = e.Tab.Text & ": " & e.Tab.Tabs(0).Text
            sLocation = e.Tab.Value
            iMonth = e.Tab.Tabs(0).Value

        End If

        'currentLocation.Value = sLocation
        'currentMonth.Value = iMonth.ToString
        'currentState.Value = "0"
        BindGrid(sLocation, iMonth)

    End Sub


    'Private Sub BindLocations(ByVal index As Integer)
    '    Dim rows As DataRowCollection = BLL.user_city.GetRows(Request.QueryString("user_id"))

    '    Dim items As ListItemCollection = New ListItemCollection()
    '    items.Add(New ListItem("Default"))

    '    For Each row As DataRow In rows
    '        items.Add(New ListItem(Convert.ToString(row("city_cd"))))
    '    Next

    '    dlLocations.DataSource = items
    '    dlLocations.SelectedIndex = index
    '    dlLocations.DataBind()
    'End Sub

    'Private Sub BindLocations(ByVal index As Integer)
    '    Dim da As New SchedulesTableAdapters.user_cityTableAdapter
    '    Dim dt As New Schedules.user_cityDataTable
    '    arrLocations = New ArrayList
    '    da.Fill(dt, Request.QueryString("user_id"))

    '    'Dim r As Schedules.user_cityRow
    '    'Dim items As ListItemCollection = New ListItemCollection()
    '    'items.Add(New ListItem("Default"))
    '    'For Each r In dt
    '    '    items.Add(New ListItem(Convert.ToString(r.city_cd)))

    '    'Next


    '    dlLocations.DataSource = dt

    '    dlLocations.SelectedIndex = index
    '    dlLocations.DataBind()

    '    Dim r As Schedules.user_cityRow
    '    For Each r In dt
    '        arrLocations.Add(r.city_cd)
    '    Next


    'End Sub

    'Private Sub BindMonths(ByVal index As Integer)
    '    Dim cols As ListItemCollection = New ListItemCollection()
    '    cols.Add(New ListItem("Default"))

    '    Dim i As Integer
    '    For i = 1 To 12
    '        cols.Add(New ListItem(MonthName(i, True)))
    '    Next

    '    dlMonths.DataSource = cols
    '    dlMonths.SelectedIndex = index
    '    dlMonths.DataBind()
    'End Sub

    Private Sub BindGrid(ByVal SelLocation As String, ByVal SelMonth As Integer)
        Dim da As New SchedulesTableAdapters.car_rate_rule_schedule_detail_4TableAdapter
        Dim dt As New Schedules.car_rate_rule_schedule_detail_4DataTable

        da.Fill(dt, Request.QueryString("schedule_id"), SelLocation, SelMonth)


        grdCarTypes.DataSource = dt
        grdCarTypes.DataBind()
    End Sub


    'Protected Sub dlMonths_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dlMonths.SelectedIndexChanged
    '    BindLocations(dlLocations.SelectedIndex)
    '    BindMonths(dlMonths.SelectedIndex)
    '    BindGrid()
    'End Sub

    'Protected Sub dlLocations_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dlLocations.SelectedIndexChanged
    '    BindLocations(dlLocations.SelectedIndex)
    '    BindMonths(0)
    '    BindGrid()
    'End Sub

    Protected Sub btnUpdate_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnUpdate.Click
        
        DoUpdate()
    End Sub


    Private Sub DoUpdate()
        Dim sLocation As String = "***"
        Dim iMonth As Integer = 0
        Dim scheduleID As Integer = Request.QueryString("schedule_id")
        Dim qa As New SchedulesTableAdapters.QueriesTableAdapter

        'If currentLocation.Value <> "-1" Then
        '    sLocation = currentLocation.Value
        'Else
        sLocation = TabStrip1.SelectedTab.ParentTab.Value
        'End If

        'If currentMonth.Value <> "-1" Then
        '    iMonth = CInt(currentMonth.Value)
        'Else
        iMonth = CInt(TabStrip1.SelectedTab.Value)
        'End If



        'Dim da2 As New SchedulesTableAdapters.car_rate_rule_schedule_detail_4TableAdapter
        'Dim dt2 As New Schedules.car_rate_rule_schedule_detail_4DataTable
        'da2.Fill(dt2, scheduleID, arrLocations(dlLocations.SelectedIndex).ToString, dlMonths.SelectedIndex)
        lblStatus.Text = "Updating for " & sLocation & "," & iMonth
        qa.car_rate_rule_schedule_detail_4_delete(scheduleID, sLocation, iMonth)
        Dim gvr As GridViewRow
        For Each gvr In grdCarTypes.Rows


            Dim cartype As Label = gvr.Cells(0).Controls(1)
            Dim min As TextBox = gvr.Cells(1).Controls(1)
            Dim max As TextBox = gvr.Cells(2).Controls(1)


            If min.Text.Length > 0 And max.Text.Length > 0 Then
                lblStatus.Text = lblStatus.Text & "<br/>car_rate_rule_schedule_detail_4_insert " & scheduleID & "," & cartype.Text & "," & min.Text & "," & max.Text & "," & iMonth & "," & sLocation

                qa.car_rate_rule_schedule_detail_4_insert(scheduleID, cartype.Text, CDbl(min.Text), CDbl(max.Text), iMonth, sLocation)
            Else
                lblStatus.Text = lblStatus.Text & "<br/>skipping because min and max not provided " & scheduleID & "," & cartype.Text & "," & min.Text & "," & max.Text & "," & iMonth & "," & sLocation


            End If

        Next

    End Sub

    Protected Sub chkDebug_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkDebug.CheckedChanged

        lblCB.Visible = chkDebug.Checked
        lblStatus.Visible = chkDebug.Checked
       
    End Sub

    Protected Sub grdCarTypes_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdCarTypes.SelectedIndexChanged

    End Sub
End Class
