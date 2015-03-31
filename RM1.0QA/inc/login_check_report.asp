<%

	If	(Request.Cookies("rate-monitor.com")("live_session") <> "auto1") Or (Session("pro_con") = "") Then
		Session("request_url") = Request.ServerVariables("URL") 
		Server.Transfer "report_login.asp"
	End If
	
%>