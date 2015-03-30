Imports Utility

Partial Class RezCentralHeader
    Inherits System.Web.UI.Page

    'Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
    '    If e.Row.RowType = DataControlRowType.DataRow Then
    '        Dim ddl As DropDownList
    '        ddl = e.Row.FindControl("ddlSystem")

    '        If Not IsNothing(ddl) Then
    '            Dim da As New DataSetLookupsTableAdapters.rezcentral_systemsTableAdapter
    '            Dim dt As New DataSetLookups.rezcentral_systemsDataTable

    '            da.Fill(dt)

    '            ddl.DataSource = dt
    '            ddl.DataValueField = "tsd_system"
    '            ddl.DataTextField = "tsd_system"
    '            ddl.DataBind()
    '        End If



    '    End If
    'End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not Page.IsPostBack Then

            If IsInt(Request.QueryString("uid")) Then
                Session("user_id") = CInt(Request.QueryString("uid"))
            End If

            If IsNothing(Session("user_id")) Then
                Response.Clear()
                Response.Write("error accessing this file.")
                Response.End()
                Exit Sub
            End If


            'If Len(Request.ServerVariables("HTTP_REFERER")) > 0 Then
                'lnkBack.NavigateUrl = Request.ServerVariables("HTTP_REFERER") ' We cant use this because after an add or delete it refers to itself
                lnkBack.NavigateUrl = "rezcentral_tethering.asp"
                lnkBack.Visible = True
            'Else
            '    lnkBack.Visible = False
            'End If

            Dim dq As New DataSetProfileSearchScheduleTableAdapters.QueriesTableAdapter
            Dim iOrg As Integer
            'Session("user_id") = 33

            dq.GetOrgFromUser(CInt(Session("user_id")), iOrg)

            Session("org_id") = iOrg

            GridView1.DataBind()



        End If

        If Len(Request.QueryString("message")) > 0 Then
            lblMessage.Text = Server.UrlDecode(Request.QueryString("messsage"))
        End If

    End Sub

    Protected Sub GridView1_RowCommand(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewCommandEventArgs) Handles GridView1.RowCommand
        If e.CommandName = "Insert" Then
            Dim txtRateCode As TextBox = GridView1.FooterRow.FindControl("txtRateCode")
            Dim txtBranch As TextBox = GridView1.FooterRow.FindControl("txtBranch")
            Dim ddlSystem As DropDownList = GridView1.FooterRow.FindControl("ddlSystem")

            Dim txtDailyDiff As TextBox = GridView1.FooterRow.FindControl("txtDailyDiff")
            Dim chkDailyIsDollar As CheckBox = GridView1.FooterRow.FindControl("chkDailyIsDollar")
            Dim txtWeekendDiff As TextBox = GridView1.FooterRow.FindControl("txtWeekendDiff")
            Dim chkWeekendIsDollar As CheckBox = GridView1.FooterRow.FindControl("chkWeekendIsDollar")
            Dim txtWeeklyDiff As TextBox = GridView1.FooterRow.FindControl("txtWeeklyDiff")
            Dim chkWeeklyIsDollar As CheckBox = GridView1.FooterRow.FindControl("chkWeeklyIsDollar")
            Dim txtMonthlyDiff As TextBox = GridView1.FooterRow.FindControl("txtMonthlyDiff")
            Dim chkMonthlyIsDollar As CheckBox = GridView1.FooterRow.FindControl("chkMonthlyIsDollar")

            Dim txtExtraDayDiff As TextBox = GridView1.FooterRow.FindControl("txtExtraDayDiff")
            Dim chkExtraDayIsDollar As CheckBox = GridView1.FooterRow.FindControl("chkExtraDayIsDollar")
            Dim txtWkndExtraDayDiff As TextBox = GridView1.FooterRow.FindControl("txtWkndExtraDayDiff")
            Dim chkWkndExtraDayIsDollar As CheckBox = GridView1.FooterRow.FindControl("chkWkndExtraDayIsDollar")

            Dim txtWeeklyExtraDayFactor As TextBox = GridView1.FooterRow.FindControl("txtWeeklyExtraDayFactor")
            Dim txtMonthlyExtraDayFactor As TextBox = GridView1.FooterRow.FindControl("txtMonthlyExtraDayFactor")
            Dim txtWeekendDOW As TextBox = GridView1.FooterRow.FindControl("txtWeekendDOW")
            Dim ddlGovType As DropDownList = GridView1.FooterRow.FindControl("ddlGovType")

            Dim da As New DataSetRezTableAdapters.newrez_headerTableAdapter
            Dim dt As New DataSetRez.newrez_headerDataTable
            Dim r As DataSetRez.newrez_headerRow


            da.FillByOrg(dt, CInt(Session("org_id")))

            r = dt.Newnewrez_headerRow
            r.org_id = CInt(Session("org_id"))

            If Len(txtBranch.Text) > 1 Then
                r.Branch = txtBranch.Text
            End If

            r.RateCode = txtRateCode.Text

            r.DailyDiff = CDbl(txtDailyDiff.Text)
            r.DailyIsDollar = chkDailyIsDollar.Checked
            r.WeekendDiff = CDbl(txtWeekendDiff.Text)
            r.WeekendIsDollar = chkWeekendIsDollar.Checked
            r.WeeklyDiff = CDbl(txtWeeklyDiff.Text)
            r.WeeklyIsDollar = chkWeeklyIsDollar.Checked
            r.MonthlyDiff = CDbl(txtMonthlyDiff.Text)
            r.MonthlyIsDollar = chkMonthlyIsDollar.Checked

            r.ExtraDayDiff = CDbl(txtExtraDayDiff.Text)
            r.ExtraDayIsDollar = chkExtraDayIsDollar.Checked
            r.WkndExtraDayDiff = CDbl(txtWkndExtraDayDiff.Text)
            r.WkndExtraDayIsDollar = chkExtraDayIsDollar.Checked
            r.WeeklyExtraDayFactor = CDbl(txtWeeklyExtraDayFactor.Text)
            r.MonthlyExtraDayFactor = CDbl(txtMonthlyExtraDayFactor.Text)

            r.WeekendDOW = IsEmpty(txtWeekendDOW.Text, "")
            r.TsdSystem = ddlSystem.SelectedValue
            r.GovType = ddlGovType.SelectedValue
            r.SenderID = "RHI01"
            r.RecipientID = "TRN"
            r.TradingPartnerCode = "WEB01"
            r.MessageID = "ADDRAT"

            dt.Addnewrez_headerRow(r)

            da.Update(dt)



            Response.Redirect("RezCentralHeader.aspx?message=" & Server.UrlEncode("Successfully added row: " & r.RezCentralHeaderID))


        End If
    End Sub

    Protected Sub GridView1_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowCreated
        'Dim del As LinkButton = DirectCast(e.Row.Cells(22).Controls(1), LinkButton)
        'del.OnClientClick = "javascript: if (!myConfirm()) {return false;};"
    End Sub

    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        If e.Row.RowState = DataControlRowState.Normal Or e.Row.RowState = DataControlRowState.Alternate Then
            If e.Row.RowType = DataControlRowType.DataRow Or e.Row.RowType = DataControlRowType.EmptyDataRow Then
                Dim del As LinkButton = e.Row.FindControl("btnDelete")
                del.OnClientClick = "javascript: return (confirm('Are you sure you want to delete?'));"
            End If
        End If
        
        If (e.Row.RowState = DataControlRowState.Alternate Or e.Row.RowState = DataControlRowState.Normal) And e.Row.RowType = DataControlRowType.DataRow Then
            Dim lbl As Label = e.Row.FindControl("lblGovType")
            Dim lbl2 As Label = e.Row.FindControl("lblTsdSystem")
            Dim ddl As DropDownList = e.Row.FindControl("ddlGovType")
            Dim ddl2 As DropDownList = e.Row.FindControl("ddlSystem")

            If Not IsNothing(lbl) And Not IsNothing(ddl) Then
                lbl.Text = ddl.SelectedItem.Text
            End If

            If Not IsNothing(lbl2) And Not IsNothing(ddl2) Then
                lbl2.Text = ddl2.SelectedItem.Text
            End If
        End If

    End Sub

    Protected Sub GridView1_RowDeleted(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewDeletedEventArgs) Handles GridView1.RowDeleted
        Response.Redirect("RezCentralHeader.aspx")
    End Sub

    Protected Sub GridView1_RowUpdating(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewUpdateEventArgs) Handles GridView1.RowUpdating

    End Sub
End Class
