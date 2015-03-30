
Partial Class admin_Transfer
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim i As Integer


        For i = 0 To Request.Form.Count - 1
            Session(Request.Form.GetKey(i)) = Request.Form(i).ToString
        Next

        If Len(Request.QueryString("goto")) > 0 Then
            Dim sURL As String = "http://www.rate-monitor.com/"
            sURL = sURL & Request.QueryString("goto")
            'Response.Write("going to " & sURL)
            'Response.Redirect("/" & Request.QueryString("goto"), True)
            Response.Redirect(sURL)
        End If

        For i = 0 To Session.Contents.Count - 1
            'lblResult.Text = lblResult.Text & "assigned to " & Session.Keys(i).ToString() & " value: " & Session(i).ToString() & "<br />"
        Next
  
    End Sub
End Class
