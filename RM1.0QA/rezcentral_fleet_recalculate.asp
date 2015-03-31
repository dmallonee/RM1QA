<%@ Language=VBScript %>
<%
Response.Buffer = true

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta content="text/html; charset=windows-1252" http-equiv="Content-Type" />
<title>Untitled 1</title>
<style type="text/css">
.auto-style1 {
	vertical-align: middle;
}
.auto-style2 {
	text-align: center;
}
</style>
</head>
<body>
<!-- this bit goes at the very top of the page: -->
<div id='interstitial' class="auto-style2">
   <img src="images/animation_processing.gif" width="200" height="200" alt="Processing ... Please wait... " class="auto-style1" />
</div>
<!-- This is for those times where you have to [asp: response.end] [php: exit(); ] for some reason-->
<script type="text/javascript" language="javascript">
    setTimeout('interstitial.style.display="none";',10000);
</script>
<!-- if using Classic ASP: -->
<%
response.flush
%>

<%
   	on error resume next

   	Server.ScriptTimeout = 30

	strConn = Session("pro_con")
	
	Rem Get the data sources
	Set adoCmd = CreateObject("ADODB.Command")

	adoCmd.ActiveConnection =  strConn
	adoCmd.CommandText = "car_fleet_fox_recalculate"
	adoCmd.CommandType = adCmdStoredProc

	adoCmd.Execute

%>



<!-- this bit goes at the very bottom of the page: -->
<script type="text/javascript" language="javascript">
     interstitial.style.display="none";
</script>
</body>

</html>
