<%@ Language=VBScript %>
<% Response.Expires = -1
   Response.cachecontrol="private" 
   Response.AddHeader "pragma", "no-cache"

   strUser     = Request.Cookies("iPhone-User")
   strPassword = Request.Cookies("iPhone-Password")

   Server.ScriptTimeout = 180

	On Error Resume Next

	Dim strConn	
	Dim adoCmd	
	Dim adoRS
	Dim adoCmd1	
	Dim adoRS1
	Dim adoCmd2	
	Dim adoRS2
	Dim adoCmd3	
	Dim adoRS3
	Dim adoCmd4
	Dim adoRS4


	Dim adoPrices
	Dim strUserId
	Dim intRuleId
	Dim strAlertDesc
	Dim datBeginDate


	'strUserId = Request.Cookies("rate-monitor.com")("user_id")
	
	strConn = "Provider=SQLOLEDB.1; Network Library=dbmssocn;Password=@ppleWEB@ccess;User ID=iPhone;Initial Catalog=production;Data Source=thor.rate-monitor.com;"

	Set adoCmd = CreateObject("ADODB.Command")

	adoCmd.ActiveConnection =  strConn
	adoCmd.CommandText = "machine_restriction_select"

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

	
	Set adoRS = adoCmd.Execute

   
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta content="text/html; charset=windows-1252" http-equiv="Content-Type" />
<title>Search Machines</title>
<meta content="yes" name="apple-mobile-web-app-capable" />
<meta content="index,follow" name="robots" />
<meta content="text/html; charset=iso-8859-1" http-equiv="Content-Type" />
<link href="pics/homescreen.png" rel="apple-touch-icon" />
<meta content="minimum-scale=1.0, width=device-width, maximum-scale=0.6667, user-scalable=no" name="viewport" />
<link href="css/style.css" rel="stylesheet" type="text/css" />
<script src="javascript/functions.js" type="text/javascript"></script>
</head>
<body class="musiclist">
<div id="topbar">
	<div id="leftnav">
		<a href="default.htm"><img alt="home" src="images/home.png" /></a><a href="default.htm">Home</a></div>
	<div id="rightnav">
		<a href="site_list.asp">Site List</a></div>
</div>
<div id="content">
	<ul class="autolist">
	    <% Dim intCount
	       intCount = 1
	    %>
		<% While adoRS.EOF = False %>
		<li>
		<a href="assignment_list.asp?machine_name=<%=adoRS.Fields("machine_name").Value %>">
		<span class="number"><%=intCount %></span><span class="name"><%=adoRS.Fields("machine_name").Value %></span><span class="time">(org id = <%=adoRS.Fields("org_id").Value %>)</span><span class="arrow"></span></a>
		</li>
		<% adoRS.MoveNext %>
		<% intCount = intCount + 1%>
		<% Wend %>
	</ul>
</div>
<div id="footer">
	<a href="http://iwebkit.net">Powered by iWebKit</a></div>
</body>
</html>
