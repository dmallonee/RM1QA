<%@ Language=VBScript %>
<!-- #INCLUDE FILE="inc/login_check.asp" -->
<!-- #INCLUDE FILE="inc/adovbs.asp" -->
<% Response.Expires = -1
   Response.cachecontrol="private" 
   Response.AddHeader "pragma", "no-cache"
   Response.Buffer = True
   
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
	Dim intSearchId
	Dim strAlertDesc
	Dim datBeginDate
	Dim strVerboseChecked




	intSearchId = Request("txt_report")
	If CStr(intSearchId) = "" Then
		intSearchId = Request("recent_searches")
	End If
	intRuleId = Request("rule")	
	intRuleDepth = Request("rule_depth")
	intUtilization = Request("txt_utilization")

	strUserId = Request.Cookies("rate-monitor.com")("user_id")
	
	strConn = Session("pro_con")
	
	Rem Get the data sources
	Set adoCmd1 = CreateObject("ADODB.Command")

	adoCmd1.ActiveConnection =  strConn
	adoCmd1.CommandText = "car_shop_request_recent"
	adoCmd1.CommandType = 4

	adoCmd1.Parameters.Append adoCmd1.CreateParameter("@user_id", 3, 1,  0, strUserId)
		
	Set adoRS1 = adoCmd1.Execute

	Set adoCmd4 = CreateObject("ADODB.Command")
	
	adoCmd4.CommandTimeout = 0

	adoCmd4.ActiveConnection =  strConn
	adoCmd4.CommandText = "car_rate_rule_select_names"
	adoCmd4.CommandType = 4
	adoCmd4.Parameters.Append adoCmd4.CreateParameter("@rate_rule_id",      3, 1, 0, Null)
	adoCmd4.Parameters.Append adoCmd4.CreateParameter("@user_id",           3, 1, 0, strUserId)
	adoCmd4.Parameters.Append adoCmd4.CreateParameter("@rate_rule_type_cd", 3, 1, 0, 1)
	adoCmd4.Parameters.Append adoCmd4.CreateParameter("@rule_status",     200, 1, 1, "E")
				
	Set adoRS4 = adoCmd4.Execute

	If err.number <> 0 Then
	   pad="&nbsp;&nbsp;&nbsp;&nbsp;"
	   response.write "<b>VBScript Errors occured selecting the rules<br>"
	   response.write "</b><br>"
	   response.write pad & "Error Number= #<b>" & err.number & "</b><br>"
	   response.write pad & "Error Desc.= <b>" & err.description & "</b><br>"
	   response.write pad & "Help Context= <b>" & err.HelpContext & "</b><br>"
	   response.write pad & "Help File Path=<b>" & err.helpfile & "</b><br>"
	   response.write pad & "Error Source= <b>" & err.source & "</b><br><hr>"

	End If


	If IsNumeric(intSearchId) Then
	
		If intSearchId  > 0 Then

			Set adoCmd3 = CreateObject("ADODB.Command")

			adoCmd3.ActiveConnection =  strConn
			adoCmd3.CommandText = "car_rule_pre_evaluation_test"
			adoCmd3.CommandType = 4

			adoCmd3.CommandTimeout = 0

			adoCmd3.Parameters.Append adoCmd3.CreateParameter("@shop_request_id2", 3, 1, 0, intSearchId)
		
			If intRuleId > 0 Then
				adoCmd3.Parameters.Append adoCmd3.CreateParameter("@rate_rule_id2",    3, 1, 0, intRuleId)
			Else
				adoCmd3.Parameters.Append adoCmd3.CreateParameter("@rate_rule_id2",    3, 1, 0, Null)
			End If
	
			adoCmd3.Parameters.Append adoCmd3.CreateParameter("@debug",           11, 1, 0, 0)
			adoCmd3.Parameters.Append adoCmd3.CreateParameter("@rule_depth",       3, 1, 0, intRuleDepth)	

			
			If Request.Form("verbose") = "true" Then
				adoCmd3.Parameters.Append adoCmd3.CreateParameter("@verbose",         11, 1, 0, 1)
				strVerboseChecked = "checked=""checked"" "
			Else
				adoCmd3.Parameters.Append adoCmd3.CreateParameter("@verbose",         11, 1, 0, 0)
				strVerboseChecked = ""
			End If
			
			adoCmd3.Parameters.Append adoCmd3.CreateParameter("@test_date",        135, 1, 0, Null)
			adoCmd3.Parameters.Append adoCmd3.CreateParameter("@test_car_type_cd", 200, 1, 4, Null)
			
			If IsNumeric(intUtilization) Then
				adoCmd3.Parameters.Append adoCmd3.CreateParameter("@utilization",    6, 1, 0, intUtilization)			
			Else
				adoCmd3.Parameters.Append adoCmd3.CreateParameter("@utilization",    6, 1, 0, Null)
			End If

			
				
			Set adoRS3 = adoCmd3.Execute(,,adExecuteNoRecords)

			'While adoRS3.State = adStateExecuting 
			
			
			If err.number <> 0 Then
			   pad="&nbsp;&nbsp;&nbsp;&nbsp;"
			   response.write "<b>VBScript Errors during pre-evaluation<br>"
			   response.write "</b><br>"
response.write intSearchId & "<br>"
response.write intRuleId & "<br>"
response.write intRuleDepth & "<br>"
response.write intUtilization & "<br>"
			   response.write pad & "Error Number= #<b>" & err.number & "</b><br>"
			   response.write pad & "Error Desc.= <b>" & err.description & "</b><br>"
			   response.write pad & "Help Context= <b>" & err.HelpContext & "</b><br>"
			   response.write pad & "Help File Path=<b>" & err.helpfile & "</b><br>"
			   response.write pad & "Error Source= <b>" & err.source & "</b><br><hr>"
			End If
			

			strConn = Session("pro_con")
	
			Set adoRS = CreateObject("ADODB.Recordset")
			Set adoCmd = CreateObject("ADODB.Command")
	
			adoCmd.ActiveConnection = strConn
			adoCmd.CommandText = "car_rate_rule_change_select_test"
			adoCmd.CommandType = 4

			adoCmd.CommandTimeout = 0
	
			adoCmd.Parameters.Append adoCmd.CreateParameter("@shop_request_id", 3, 1, 0, intSearchId) 'Request("reportrequestid"))

			Set adoRS = adoCmd.Execute

			If err.number <> 0 Then
			   pad="&nbsp;&nbsp;&nbsp;&nbsp;"
			   response.write "<b>VBScript Errors during rule change select<br>"
			   response.write "</b><br>"
			   response.write pad & "Error Number= #<b>" & err.number & "</b><br>"
			   response.write pad & "Error Desc.= <b>" & err.description & "</b><br>"
			   response.write pad & "Help Context= <b>" & err.HelpContext & "</b><br>"
			   response.write pad & "Help File Path=<b>" & err.helpfile & "</b><br>"
			   response.write pad & "Error Source= <b>" & err.source & "</b><br><hr>"

			End If


		Else
	
			Set adoRS = CreateObject("ADODB.Recordset")

		End If

		
	End If

	
	If err.number <> 0 Then
	   pad="&nbsp;&nbsp;&nbsp;&nbsp;"
	   response.write "<b>VBScript Errors Occured!<br>"
	   response.write "</b><br>"
	   response.write pad & "Error Number= #<b>" & err.number & "</b><br>"
	   response.write pad & "Error Desc.= <b>" & err.description & "</b><br>"
	   response.write pad & "Help Context= <b>" & err.HelpContext & "</b><br>"
	   response.write pad & "Help File Path=<b>" & err.helpfile & "</b><br>"
	   response.write pad & "Error Source= <b>" & err.source & "</b><br><hr>"

	End If
	

%>



<!doctype HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>

<head>
<meta name="GENERATOR" content="Microsoft FrontPage 12.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Language" content="en-us">
<title>Rate-Monitor by Rate-Highway, Inc. | Alerts! | Rate Rule Tester</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<style type="text/css" >
<!--
.profile_header {height="48" text-align:left; padding-left:3; padding-right:3; padding-top:0; background-color:#879AA2; font-family:Verdana, Arial, Helvetica, sans-serif; font-size:10pt; vertical-align:bottom; text-align:left }
.profile_light { height="68" text-align:left; padding-left:3; padding-right:3; padding-top:3; background-color:#CFD7DB; font-family:Verdana, Arial, Helvetica, sans-serif; font-size:10pt; vertical-align:bottom; text-align:left }
.profile_dark { height="48" text-align:left; padding-left:3; padding-right:3; padding-top:3; background-color:#B2BEC4; font-family:Verdana, Arial, Helvetica, sans-serif; font-size:10pt; vertical-align:bottom; text-align:left }
.style1 {
	font-size: x-small;
}
.style2 {
	height: "48" text-align:left;
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-size: 10pt;
	vertical-align: bottom;
	text-align: left;
	border-width: 0;
	padding-left: 3;
	padding-right: 3;
	padding-top: 0;
	background-color: #879AA2;
}
-->
</style>
<link rel="stylesheet" type="text/css" href="inc/rh_standard.css">
<link rel="stylesheet" type="text/css" href="inc/rh_report.css">
<link rel="stylesheet" type="text/css" href="inc/sitewide.css">
<script language="Javascript" type="text/javascript" src="inc/sitewide.js"></script>
<script language="Javascript" type="text/javascript" src="inc/header_menu_support.js" ></script>
</head>
<script language="javascript" type="text/javascript" >
    // This script is intended for use with a minimum of Netscape 4 or IE 4.
    if (document.getElementById) {
        var upLevel = true;
    }
    else if (document.layers) {
        var ns4 = true;
    }
    else if (document.all) {
        var ie4 = true;
    }

    function showObject(obj) {
        if (ns4) obj.visibility = "show";
        else if (ie4 || upLevel) obj.style.visibility = "visible";
    }
    function hideObject(obj) {
        if (ns4) {
            obj.visibility = "hide";
        }
        if (ie4 || upLevel) {
            obj.style.visibility = "hidden";
        }
    }

    function clearbox() {
        document.test_rules.txt_report.value = '';
    }

</script>
<body style="font-family:Verdana, Arial, Helvetica, sans-serif; font-size:x-small; margin-top:0; margin-left:0; background-color:white; " >
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td background="images/top_tile.gif">
    <img src="images/top.jpg" width="770" height="91" alt=""></td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td background="images/b_tile.gif">
    <!-- #INCLUDE FILE="inc/page_header_buttons.htm" -->
    </td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td background="images/med_bar_tile.gif">
    <img src="images/med_bar.gif" width="12" height="8"></td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td background="images/user_tile.gif">
    <table width="100" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td><img src="images/user_left.gif" width="580" height="31"></td>
        <td background="images/user_tile.gif">
        <table width="100" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td valign="bottom">
            <table width="100" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td align="right">
                <div align="right">
                  <font face="Vendana, Arial, Helvetica, sans-serif" size="2" color="#FFFFFF">
                  User: <%=Request.Cookies("rate-monitor.com")("user_name")%></font></div>
                </td>
              </tr>
              <tr>
                <td><img src="images/separator.gif" width="183" height="6"></td>
              </tr>
            </table>
            </td>
            <td><img src="images/user_tile.gif" width="7" height="31"></td>
          </tr>
        </table>
        </td>
      </tr>
    </table>
    </td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td background="images/h_tile.gif">
    <table width="100" border="0" cellspacing="0" cellpadding="0" id="table1">
      <tr>
        <td><img src="images/h_alerts.gif" width="368" height="31"></td>
        <td><img src="images/h_right.gif" width="402" height="31"></td>
      </tr>
    </table>
    </td>
  </tr>
</table>

<p align="right">&nbsp;</p>

<DIV ID="splashScreen" STYLE="position:absolute;z-index:1;top:30%;left:35%">
    <TABLE BGCOLOR="#000000" BORDER=1 BORDERCOLOR="#000000" 
	CELLPADDING=0 CELLSPACING=0 HEIGHT=200 WIDTH=300>
      <TR>
        <TD WIDTH="100%" HEIGHT="100%" BGCOLOR="#FFFFFF" ALIGN="CENTER">
          <BR><BR>
          <b><font face="Helvetica,Verdana,Arial" color="#000066">Results</font></b><FONT FACE="Helvetica,Verdana,Arial" SIZE=3 COLOR="#000066"><B> 
          loading.  Please wait...</B></FONT> <br>
          <BR> 
          <IMG SRC="images/waiting.gif" BORDER=0 width="28" height="28" ><BR><BR> 
        </TD>
      </TR>
    </TABLE>
</DIV>
<%
Response.Flush
%>                
<div align="center">
<table cellpadding="0" cellspacing="0" border="0" width="1310" bgcolor="#FFFFFF">
<tr height="1">
<td colspan="1" width="1">&nbsp;</td>
<td rowspan="2" width="169"><img src="images/ratemanagementalerts2_a.JPG" width="169" height="25" hspace="0" vspace="0" border="0" alt="Rate Management" description=""></td>
<td colspan="1" >&nbsp;</td>
</tr>
<tr height="1">
<td bgcolor="#000000" colspan="1" height="1"><img src="images/pixel.gif" width="1" height="1"></td>
<td bgcolor="#000000" colspan="1" height="1"><img src="images/pixel.gif" width="1" height="1"></td>
<td bgcolor="#000000" colspan="1" height="1"><img src="images/pixel.gif" width="1" height="1"></td>
</tr>
</table>
</div>
<table cellpadding="0" cellspacing="0" border="0" ALIGN="CENTER" width="100%" bgcolor="#FFFFFF">
<tr >
<td  width="1" bgcolor="#000000"><img src="images/pixel.gif" width="1" height="1"></td>
<td colspan=3 bgcolor="#D9DEE1">
<table border="0" cellspacing="5" cellpadding="5">
<tr><td>
<font color="#080000">
<P>
<!-- JUSTTABS TOP OPEN-END -->
&nbsp;</p>
<form method="post" name="test_rules" class="search">
<table border="0" style="border-collapse: collapse" bordercolor="#FFFFFF" width="100%" cellspacing="0" height="4">
    <tr>
      <td background="images/ruler.gif"></td>
    </tr>
  </table>
 
  <table width="745" border="0" cellspacing="0" cellpadding="0" background="images/alt_color.gif">
    <tr>
      <td>
      <table width="1108" border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" background="images/alt_color.gif">
        <tr valign="bottom">
          <td width="10" height="51">&nbsp;</td>
          <td width="179" height="51">
          <img border="0" src="images/test_rule.GIF" width="162" height="25" alt=""></td>
          <td width="583" colspan="3" height="51">
          <span style="font-size:x-small">To test 
          a single rule or all rules assigned to a report, select a recent 
          report, and either &quot;Execute all assigned rules&quot; or a single rule, then 
          click the test rule button.</span></td>
          <td width="336" height="51">&nbsp;</td>
        </tr>
        <tr valign="bottom">
          <td width="10" height="26">&nbsp;</td>
          <td width="179" height="26">&nbsp;</td>
          <td width="177" height="26">
          <font size="2">Recent report to use:</font> </td>
          <td width="80" height="26">
          <select size="1" name="recent_searches"  style="border:1px solid #000000; width:300; background-color:#FF9933" onclick="clearbox();">
<%
		While adoRS1.EOF = False
		
			If CLng(intSearchId) = adoRS1.Fields("shop_request_id").Value Then %>
				<option selected value="<%=adoRS1.Fields("shop_request_id").Value %>"><%=adoRS1.Fields("recent_searches").Value %></option>
			<% Else %>
				<option value="<%=adoRS1.Fields("shop_request_id").Value %>"><%=adoRS1.Fields("recent_searches").Value %></option>
			<% End If
			adoRS1.MoveNext
		
		Wend
					
%>


          </select></td>
          <td width="662" colspan="2" height="26">
          <font size="2" face="Verdana, Arial, Helvetica, sans-serif">
          <input name="search" type="submit" id="Open2224" value="    Test Rule   " class="rh_button" ></font></td>
        </tr>
        <tr valign="bottom">
          <td width="10" height="22">&nbsp;</td>
          <td width="179" height="22">&nbsp;</td>
          <td width="177" height="22">
<font color="#080000">
			<span class="style1">&nbsp;&nbsp;&nbsp; or rpt. number:</span>&nbsp;</font></td>
          <td  height="22">
		  <input name="txt_report" id="txt_report" type="text" value="<%=intSearchId %>" style="text-align:right"></td>
          <td width="662" colspan="2" height="22">&nbsp;</td>
        </tr>
        <tr valign="bottom">
          <td width="10" height="22">&nbsp;</td>
          <td width="179" height="22">&nbsp;</td>
          <td width="177" height="22">
          <font size="2">Rule to test with</font><font size="2" face="Verdana, Arial, Helvetica, sans-serif">:</font></td>
          <td width="80" height="22">
          <select size="1" name="rule" style="border:1px solid #000000; width:300; background-color:#FF9933">
          		<option selected value="0">(Execute all assigned rules)</option>
                <% While adoRS4.EOF = False %> 
                	<% If CLng(intRuleId) = adoRS4.Fields("rate_rule_id").Value Then %>
		                	<option selected value="<%=adoRS4.Fields("rate_rule_id").Value %>"><%=adoRS4.Fields("alert_desc").Value %></option>
		            <% Else %>
		                	<option value="<%=adoRS4.Fields("rate_rule_id").Value %>"><%=adoRS4.Fields("alert_desc").Value %></option>
		            <% End If %>
                <%	adoRS4.MoveNext
				   Wend
				   Set adoRS4 = Nothing
				%>
					

          </select></td>
          <td width="662" colspan="2" height="22"><span style="font-size:x-small" >(optional)&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
		  *Type 0 = Valid change</span></td>
        </tr>
        <tr valign="bottom">
          <td width="10" height="22">&nbsp;</td>
          <td width="179" height="22">&nbsp;</td>
          <td width="177" height="22">
<font color="#080000">
          <font size="2">Rule depth</font><font size="2" face="Verdana, Arial, Helvetica, sans-serif">:</font></font></td>
          <td width="80" height="22">
<font color="#080000">
          <select size="1" name="rule_depth" style="border:1px solid #000000; width:300; background-color:#FF9933">
          	<% If intRuleDepth = 100 Then %>
          		<option selected value="100">(Execute all assigned rules)</option>
          	<% Else   %>
          		<option value="100">(Execute all assigned rules)</option>
          	<% End If %>
          	
          	<% If intRuleDepth = 0 Then %>
          		<option selected value="0">1 rule deep</option>
          	<% Else   %>
          		<option value="0">1 rule deep</option>
          	<% End If %>
          	
           	<% If intRuleDepth = 1 Then %>
          		<option selected value="1">2 rules deep</option>
          	<% Else   %>
          		<option value="1">2 rules deep</option>
          	<% End If %>
          		
           	<% If intRuleDepth = 2 Then %>
          		<option selected value="2">3 rules deep</option>
          	<% Else   %>
          		<option value="2">3 rules deep</option>
          	<% End If %>
          	
           	<% If intRuleDepth = 3 Then %>
          		<option selected value="3">4 rules deep</option>
          	<% Else   %>
          		<option value="3">4 rules deep</option>
          	<% End If %>
          	
           	<% If intRuleDepth = 4 Then %>
          		<option selected  value="4">5 rules deep</option>
          	<% Else   %>
          		<option value="4">5 rules deep</option>
          	<% End If %>
          	
           	<% If intRuleDepth = 5 Then %>
          		<option selected value="5">6 rules deep</option>
          	<% Else   %>
          		<option value="5">6 rules deep</option>
          	<% End If %>
          	
           	<% If intRuleDepth = 6 Then %>
          		<option selected value="6">7 rules deep</option>
          	<% Else   %>
          		<option value="6">7 rules deep</option>
          	<% End If %>
          	
           	<% If intRuleDepth = 7 Then %>
          		<option selected value="7">8 rules deep</option>
          	<% Else   %>
          		<option value="7">8 rules deep</option>
          	<% End If %>
          	
           	<% If intRuleDepth = 8 Then %>
          		<option selected value="8">9 rules deep</option>
          	<% Else   %>
          		<option value="8">9 rules deep</option>
          	<% End If %>
          		
           	<% If intRuleDepth = 9 Then %>
          		<option selected value="9">10 rules deep</option>
          	<% Else   %>
          		<option value="9">10 rules deep</option>
          	<% End If %>

          </select></font></td>
          <td width="662" colspan="2" height="22">
<font color="#080000">
		  <span style="font-size:x-small" >(optional)&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
		  1 = No change detected</span></font></td>
        </tr>
        <tr valign="bottom">
          <td width="10" style="height: 22px"></td>
          <td width="179" style="height: 22px"></td>
          <td width="177" style="height: 22px"><span style="font-size:x-small" >Display deleted:</span></td>
          <td width="80" style="height: 22px">
          	<% If Request.Form("verbose") = "true" Then %>
			<input name="verbose" type="checkbox" style="width: 20px" value="true" checked="checked"  >
          	<% Else %>
			<input name="verbose" type="checkbox" style="width: 20px" value="true" >
			<% End If %>			
		  </td>
          <td width="662" colspan="2" style="height: 22px"><span style="font-size:x-small" >(<font color="#080000"><span style="font-size:x-small" >optional)&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
		  2 = Duplicate date</span></font></span></td>
        </tr>
        <tr valign="bottom">
          <td width="10" height="22">&nbsp;</td>
          <td width="179" height="22">&nbsp;</td>
          <td width="177" height="22"><span style="font-size:x-small" >Forced Utilization:</span></td>
          <td width="80" height="22">
			<font color="#080000">
		  <input name="txt_utilization" type="text" value="<%=intUtilization %>" style="text-align:right" size="4"></font></td>
          <td width="662" colspan="2" height="22">
<font color="#080000">
		  <span style="font-size:x-small" >(<span style="font-size:x-small" >optional)&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
		  3 = Out of util. range</span></span></font></td>
        </tr>
        <tr valign="bottom">
          <td width="10" height="22">&nbsp;</td>
          <td width="179" height="22">&nbsp;</td>
          <td width="177" height="22">&nbsp;</td>
          <td width="80" height="22">
			&nbsp;</td>
          <td width="662" colspan="2" height="22">&nbsp;</td>
        </tr>
        <tr valign="bottom">
          <td width="10" height="22">&nbsp;</td>
          <td width="179" height="22">&nbsp;</td>
          <td width="177" height="22">&nbsp;</td>
          <td width="80" height="22">
		  &nbsp;</td>
          <td width="662" colspan="2" height="22">&nbsp;</td>
        </tr>
      </table>
      </td>
    </tr>
  </table>
  </form>
 <table width="1300" border="0" cellpadding="2" background="images/alt_color.gif" style="border-collapse: collapse" bordercolor="#111111" id="table1">
    <tr valign="bottom">
      <td >&nbsp;Rate Rule Test Results</td>
    </tr>
  </table>
  <table border="0" style="border-collapse: collapse" bordercolor="#FFFFFF" width="1300" height="4" id="table2">
    <tr>
      <td background="images/ruler.gif"></td>
    </tr>
  </table>
  <form name="rate_change" method="POST" action="rate_change_request.asp">
  <input type="hidden" name="action" value="1">
  <input type="hidden" name="refresh_from" value="search">
  <table border="1" bordercolor="#FFFFFF" id="table3" width="1300" cellspacing="0" cellpadding="0" >
    <tr>
      <th align="left" valign="bottom" bgcolor="#879AA2"><font size="2">Id</font></th>
      <th class="style2" style="background-color: #E07D1A"><font size="2">Alert Description</font></th>
      <th class="style2"><font size="2">Rate Code</font></th>
      <th class="profile_header"><font size="2">Proposed Rate</font></th>
      <th class="profile_header" width="75"><font size="2">Diff.</font></th>
      <th class="profile_header"><font size="2">City</font></th>
      <th class="profile_header"><font size="2">Car Type</font></th>
      <th class="profile_header"><font size="2">Pick-up</font></th>
      <th class="profile_header"><font size="2">LOR</font> </th>
      <th class="profile_header"><font size="2">Current Rate</font> </th>
      <th class="profile_header"><font size="2">Comp. Set Max</font> </th>
      <th class="profile_header"><font size="2">Comp. Set Min</font> </th>
      <th class="profile_header"><font size="2">Min Rate Vendor</font> </th>
      <th class="profile_header"><font size="2">Util. level</font> </th>      
      <th class="profile_header"><font size="2">True Rule</font> </th> 
      <th class="profile_header"><font size="2">False Rule</font></th> 
      <th class="profile_header"><font size="2">Type*</font></th> 
   </tr>
    
 <%
        
        Dim strClass
        Dim strOrange
        Dim intCount
        
        If adoRS Is Nothing Then
		%>
		
		Nothing
		
		<%

		ElseIf (adoRS.State <> adStateOpen) Then
		%>
		
		Closed State = <%=adoRS.State %>
		
		<%
		
		Else

		While adoRS.EOF = False
		
			If strClass = "profile_light" Then
				strClass = "profile_dark"
				strOrange = "bgcolor='#E07D1A'"
			Else
				strClass = "profile_light"
				strOrange = "bgcolor='#FDC677'"
			End If
			
			intCount = intCount + 1
			
		%>
    
    <tr>
    <td class="<%=strClass %>" width="29">
	<%=adoRS.Fields("car_rate_rule_change_id").Value%></td>
    <td bgcolor="#FDC677" width="400" title="<%=adoRS.Fields("alert_desc").Value %>" ><font size="2">
    <% If Len(adoRS.Fields("alert_desc").Value) > 50 Then %>
    <%=Left(adoRS.Fields("alert_desc").Value, 50) & "..." %>
    <% Else %>
    <%=adoRS.Fields("alert_desc").Value %>
    <% End If %>
    
    </font></td>
     <td class="<%=strClass %>" width="70" >
	<%=adoRS.Fields("rate_cd").Value & "" %>
	</td>
    <td class="<%=strClass %>_right" align="right" width="74">
    <% If adoRS.Fields("new_rt_amt").Value = -1000 Then %>
	<font color='red'>&lt;too low&lt;</font>
    <% ElseIf adoRS.Fields("new_rt_amt").Value = 10000 Then %>
	<font color='blue'>close</font>
    <% ElseIf adoRS.Fields("new_rt_amt").Value = 10001 Then %>
	<font color='green'>open</font>
    <% ElseIf adoRS.Fields("new_rt_amt").Value = 20000 Then %>
	<font color='black'>drop chrg</font>
    <% Else %>
	<%=FormatCurrency(adoRS.Fields("new_rt_amt").Value) %>
    <% End If %>
	</td>
    <td class="<%=strClass %>_right" >
	<font size="-1">
    <% If adoRS.Fields("new_rt_amt").Value >= 10000 Then %>
		n/a 
	<% ElseIf adoRS.Fields("new_rt_amt").Value > adoRS.Fields("rt_amt").Value Then %>
	<%=FormatCurrency( adoRS.Fields("new_rt_amt").Value - adoRS.Fields("rt_amt").Value)	%>
	<% Else	%>
	<font color="#FF0000">
	<%=FormatCurrency(adoRS.Fields("new_rt_amt").Value - adoRS.Fields("rt_amt").Value)	%>
	</font>
	<% End If %>
	</font></td>
    <td class="<%=strClass %>" width="50" >
	<%=adoRS.Fields("city_cd").Value %>
	</td>
    <td class="<%=strClass %>" width="73"><font size="-1">
	<%=adoRS.Fields("shop_car_type_cd").Value %> </font></td>
    <td class="<%=strClass %>_ctr" width="74">
    <%=FormatDateTime(adoRS.Fields("arv_dt").Value, 2) %></td>
   
    <td class="<%=strClass %>_ctr" width="45">
    <font size="-1">
    <%=adoRS.Fields("lor").Value%></font></td>
   
    <td class="<%=strClass %>_right" align="right" width="75">
    <font size="-1" >
	<%=FormatCurrency(adoRS.Fields("rt_amt").Value) %></font></td>
   
    <td class="<%=strClass %>_right" align="right" width="74">
    <font size="-1">
	<%=FormatCurrency(adoRS.Fields("max_rt_amt").Value) %></font></td>
   
    <td class="<%=strClass %>_right" align="right" width="74">
    <font size="-1">
	<%=FormatCurrency(adoRS.Fields("min_rt_amt").Value) %></font></td>
   
    <td class="<%=strClass %>" height="24" width="74">
    <font size="-1">
	<%=adoRS.Fields("min_vend_cd").Value%></font></td>

    <td class="<%=strClass %>" height="24" width="74">
    <font size="-1">
	<%=FormatNumber(adoRS.Fields("current_util").Value) & "%" %></font></td>

    <td class="<%=strClass %>_right" height="24" width="74">
    <font size="-1">
	<%=adoRS.Fields("on_success_id").Value %></font></td>

    <td class="<%=strClass %>_right" height="24" width="74">
    <font size="-1">
	<%=adoRS.Fields("on_failure_id").Value %></font></td>

    <td class="<%=strClass %>_right" height="24" width="24">
    <font size="-1">
	<%=adoRS.Fields("change_type_id").Value %></font></td>
   
    </tr>

<%	
	adoRS.MoveNext
	Wend
	
	End If
	
%>    
    
</table>
 
  <table border="0" style="border-collapse: collapse" bordercolor="#FFFFFF" width="1300" height="4" id="table4">
    <tr>
      <td background="images/ruler.gif"></td>
    </tr>
  </table>
  <table border="0" style="border-collapse: collapse" bordercolor="#FFFFFF" width="1300" height="4">
    <tr>
       <td background="images/ruler.gif"></td>
    </tr>
  </table>
  <input type="hidden" name="refresh_from" value="create">
  <input type="hidden" name="rule_status" value="E">
</form>
<!-- Content goes before this comment -->
<!-- JUSTTABS BOTTOM OPEN -->
</font></td></tr></table>
</td>
<td  width="1" bgcolor="#000000"><img src="images/pixel.gif" width="1" height="1"></td>
</tr>
<tr bgcolor="#000000" height="1">
<td colspan="5"><img src="images/pixel.gif" width="1" height="1"></td>
</tr>
</table>
<!-- JUSTTABS BOTTOM CLOSE -->
<p><%=intSearchId %>/<%=intRuleId %>/<%=intRuleDepth %></p>
<p align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">© 
2002 - 2013 - All rights reserved<br>
<b>Rate-Highway, Inc.</b><br>
</font><font size="1" face="Verdana, Arial, Helvetica, sans-serif">18001 Cowan, 
    Suite F<br />
    Irvine, CA 92614<br />
    (949) 614-0751&nbsp; </font>

    </p>
<font size="2" face="Verdana, Arial, Helvetica, sans-serif">
<p align="center">&nbsp;</p>
</font>
            
<SCRIPT type="text/javascript"  LANGUAGE="JavaScript">
    if (upLevel) {
        var splash = document.getElementById("splashScreen");
    }
    else if (ns4) {
        var splash = document.splashScreen;
    }
    else if (ie4) {
        var splash = document.all.splashScreen;
    }
    hideObject(splash);
</SCRIPT>
<!--#INCLUDE FILE="footer.asp"-->
</body>
</html>
