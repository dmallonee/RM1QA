<%
 	If	(Request.Cookies("rate-monitor.com")("live_session") <> "auto") Or (Session("pro_con") = "") Or (Session("org_id") = "") Then
	
        		Server.Transfer "default_session.asp"
 	
	End If
	
%>