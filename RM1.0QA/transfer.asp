<html>
<head>
<title>transfer</title>
</head>
<body>
<%
' We graf all the session variable names/values and stick them in a form
' and then we submit the form to our receiving ASP.NET page (ASPNETPage1.aspx)...
Response.Write("<form name='t' id='t' action='Transfer.aspx?goto=" & Request.QueryString("goto") & "' method='post' >")
'For each Item in Session.Contents
For each Item in Request.Cookies("rate-monitor.com")
Response.Write("<input type='hidden' name='" & Item)
Response.Write( "' value='" & Request.Cookies("rate-monitor.com")(item) & "'>")
next
Response.Write("</form>")
Response.Write("<script language='javascript'>document.t.submit();</script>")
%>
</body>
</html>