<%@ Language=VBScript %>
<% Response.Expires = -1
   Response.cachecontrol="private" 
   Response.AddHeader "pragma", "no-cache"


	Response.Cookies("iPhone-User") = Request.Form("user")
	Response.Cookies("iPhone-User").Expires = #January 01, 2099#
	Response.Cookies("iPhone-User").Domain = "www.rate-monitor.com"
		
	Response.Cookies("iPhone-Password") = Request.Form("password")
	Response.Cookies("iPhone-Password").Expires = #January 01, 2099#
	Response.Cookies("iPhone-Password").Domain = "www.rate-monitor.com"


	If Request.Form("user") = "michael" Then
		Response.Redirect("home.asp")
	Else
		Response.Redirect("default.asp")
	End If

   
%>
<head>
</head>
<html>
<body>
</body>
</html>		
