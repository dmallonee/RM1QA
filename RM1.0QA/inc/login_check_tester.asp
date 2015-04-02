<%

	If	(Request.Cookies("rate-monitor.com")("live_session") <> "bogus") Or (Session("pro_con") = "") Then
		Server.Transfer "serverinfo.asp"
		
	End If
	
%>