<%@ Language=VBScript %>
<!--
Revisions
When     Who What
======== === ==========================================================
-->
<!-- #INCLUDE FILE="inc/adovbs.asp" --> 
<%

	On Error Resume Next
	
    Response.Expires = -1
	Response.Expiresabsolute = Now() - 1
	Response.AddHeader "pragma","no-cache"
	Response.AddHeader "cache-control","private"
	Response.CacheControl = "no-cache"
	'Response.Buffer = True 
 
	Dim strLoginCode
	Dim strPassword 
	Dim strRememberMe 
	Dim strFromForm
	Dim strConn
	Dim strSQL
	Dim adoCmd
	Dim ReturnValue
	Dim Database
	Dim objItem 
	Dim intLOB_id
	Dim strURL
	Dim intParentID
	Dim strSupportPwd
	Dim strSite
	
	strLoginCode  = Request.QueryString("email_address") & ""
	strPassword   = Request.QueryString("password") & ""
	strRememberMe = Request.QueryString("Remember") & ""
	strFromForm   = Request.QueryString("FromForm") & ""
	intDatabase   = Request.QueryString("database") & ""
	strURL        = Request.QueryString("request_url") & ""
	strSupportPwd = "charliebrown"
	strSite       = Request.ServerVariables("SERVER_NAME")
	

	If err.number <> 0 Then
	   pad="&nbsp;&nbsp;&nbsp;&nbsp;"
	   response.write "<b>VBScript Errors Occured!<br>"
	   response.write parm_msg & "</b><br>"
	   response.write pad & "Error Number= #<b>" & err.number & "</b><br>"
	   response.write pad & "Error Desc.= <b>" & err.description & "</b><br>"
	   response.write pad & "Help Context= <b>" & err.HelpContext & "</b><br>"
	   response.write pad & "Help File Path=<b>" & err.helpfile & "</b><br>"
	   response.write pad & "Error Source= <b>" & err.source & "</b><br><hr>"
	End If


%>    

<html>
<head>
<meta http-equiv="Content-Language" content="en-us">
<meta name="GENERATOR" content="Microsoft FrontPage 12.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Administration | Login</title>
<script type="text/javascript" >
function submitform()
{
  document.login.submit();
}
</script>
</head>
<body onload="submitform();" > 
<form action="system_login.asp" method="post" name="login">
	<input name="password" type="hidden" value="<%=strPassword %>">
	<input name="email_address" type="hidden" value="<%=strLoginCode %>">
	<input name="Remember" type="hidden" value="<%=strRememberMe %>">
	<input name="FromForm" type="hidden" value="<%=strFromForm %>">
	<input name="database" type="hidden" value="<%=intDatabase %>">
	<input name="request_url" type="hidden" value="<%=strURL %>">
</form>
</body>
</html>