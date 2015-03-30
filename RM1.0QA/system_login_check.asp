<!--
Revisions
When     Who What
======== === ==========================================================


-->
<%@ Language=VBScript %>
<!-- #INCLUDE FILE="inc/adovbs.asp" --> 
<%
    Response.Expires = -1
	Response.Expiresabsolute = Now() - 1
	Response.AddHeader "pragma","no-cache"
	Response.AddHeader "cache-control","private"
	Response.CacheControl = "no-cache"
	'Response.Buffer = True 
 
	strLoginCode  = Request("email_address")
	strPassword   = Request("password") 
	strRememberMe = Request("Remember")
	strFromForm   = Request("FromForm")
	intDatabase   = Request("database")
	strURL        = Request("request_url")
	strSupportPwd = "charliebrown"
	strSite       = Request.ServerVariables("SERVER_NAME")
		

%>    

<html>
<head>
<meta http-equiv="Content-Language" content="en-us">
<meta name="GENERATOR" content="Microsoft FrontPage 12.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Administration | Login</title>
</head>
<body>
	 		Request.Cookies("rate-monitor.com")("live_session") = <%=Request.Cookies("rate-monitor.com")("live_session") %><br>
			Request.Cookies("rate-monitor.com")("loginCode") = <%=Request.Cookies("rate-monitor.com")("loginCode") %><br>
			Request.Cookies("rate-monitor.com")("password") = <%=Request.Cookies("rate-monitor.com")("password") %><br>
			Request.Cookies("rate-monitor.com")("remember Me") = <%=Request.Cookies("rate-monitor.com")("remember Me") %><br>
			Session("pro_con") = <%=Session("pro_con") %><br>
			Request.Cookies("rate-monitor.com")("user_id") = <%=Request.Cookies("rate-monitor.com")("user_id") %><br>
			Session("user_id") = <%=Session("user_id") %><br>
			Session("user_name") = <%=Session("user_name") %><br>
			Session("org_id") = <%=Session("org_id") %><br>
			Request.Cookies("rate-monitor.com")("user_name") = <%=Request.Cookies("rate-monitor.com")("user_name") %><br>
			Request.Cookies("rate-monitor.com")("client_userid") = <%=Request.Cookies("rate-monitor.com")("client_userid") %><br>
			strLoginCode  = <%=Request("email_address") %><br>
			strPassword   = <%=Request("password")  %><br>
			strRememberMe = <%=Request("Remember") %><br>
			strFromForm   = <%=Request("FromForm") %><br>
			intDatabase   = <%=Request("database") %><br>
			strURL        = <%=Request("request_url") %><br>
				strSite = <%=strSite %>

</body>
</html>