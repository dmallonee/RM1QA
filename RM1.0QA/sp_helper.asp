<%@ Language=VBScript %>
<!-- #INCLUDE FILE="inc/adovbs.asp" --> 

<% Response.Expires = -1
   Response.cachecontrol="private" 
   Response.AddHeader "pragma", "no-cache"

   Server.ScriptTimeout = 180
 
	Dim strConn	
	Dim adoCmd	
	Dim adoRS
	Dim adoPrices
	Dim intIndex
	Dim blnExit
	Dim intReportRequestID
	Dim intErrorCount
	Dim intTimeoutCount
	Dim strCarType 
	Dim intResults
	Dim intPrice
	Dim strName
	
	strName = Request.Form("sp_name")
	
	
	If (strName = "") Then
	
		rem Msg box?
	Else
	
		strConn = Session("pro_con")
	
  		Set adoCmd = CreateObject("ADODB.Command")

		adoCmd.ActiveConnection =  strConn
		adoCmd.CommandText = strName
		adoCmd.CommandType = 4

		adoCmd.Parameters.Refresh 
		
		For intIndex = 1 to adoCmd.Parameters.Count - 1
			Response.Write "adoCmd.Parameters.Append adoCmd.CreateParameter(|" & adoCmd.Parameters.Item(intIndex).Name & "|, " & adoCmd.Parameters.Item(intIndex).Type & ", " & adoCmd.Parameters.Item(intIndex).Direction & ", " & adoCmd.Parameters.Item(intIndex).Size & ")" &  "<br>"
		
		Next



	
	
	End If

	%>
	

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Language" content="en-us">

<title>Rate-Monitor by Rate-Highway, Inc. | Search Queue</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<link rel="stylesheet" type="text/css" href="inc/rh_standard.css">
<link rel="stylesheet" type="text/css" href="inc/rh_report.css">
<link rel="stylesheet" type="text/css" href="inc/sitewide.css">
<script src="inc/sitewide.js" language="javascript"></script>
</head>

<html>
<body >
<form method="POST" name="sp" class="search" action="sp_helper.asp">
  <div style="position: absolute; width: 100px; height: 100px; z-index: 1" id="layer1">
&nbsp;</div>
  <p>&nbsp;</p>
  <p align="center">Stored Procedure <input type="text" name="sp_name" size="20"></p>
  <p align="center"><input type="submit" value="Submit" name="B1"><input type="reset" value="Reset" name="B2"></p>
</form>
<p align="center"><a href="serverinfo.asp">Server Info</a></p>
<!--#INCLUDE FILE="footer.asp"-->
</body> 
</html>