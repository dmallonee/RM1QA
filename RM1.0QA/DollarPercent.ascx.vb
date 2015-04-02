
Partial Class DollarPercent
    Inherits System.Web.UI.UserControl


    Private m_State As String

    
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        m_State = "$"

        SetCaption()
    End Sub

    Sub SetCaption()
        btnDollarPercent.Text = m_State
    End Sub

    Protected Sub btnDollarPercent_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDollarPercent.Click
        If m_State = "$" Then
            m_State = "%"
        Else
            m_State = "$"
        End If

        SetCaption()
    End Sub
End Class
