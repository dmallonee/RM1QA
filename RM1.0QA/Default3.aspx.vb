Imports System.Data.SqlClient
Imports System.Text
Partial Class Default3
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not Me.IsPostBack Then

            btnShow.Attributes.Add("OnClick", " return CheckItems(); ")
        End If
    End Sub

    Protected Sub btnLoad_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnLoad.Click
        txt.Text = ""
        BindGrid()
        btnShow.Visible = True
    End Sub
    Public Sub BindGrid()
        Dim da As New SqlDataAdapter(txtSQL.Text, txtConn.Text)
        Dim dt As New System.Data.DataTable()
        da.Fill(dt)

        Grid1.DataSource = dt
        Grid1.DataBind()



    End Sub

End Class
