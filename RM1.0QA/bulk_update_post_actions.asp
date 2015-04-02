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


	strUserId = Request.Cookies("rate-monitor.com")("user_id")
	strRuleId = Request.Cookies("rate-monitor.com")("repeat_rule_ids")
	
	strSelectedRules = Request.Form("id_selected")
		
	strConn = Session("pro_con")
	
	Rem Get the data sources
	Set adoCmd = CreateObject("ADODB.Command")

	adoCmd.ActiveConnection =  strConn
	adoCmd.CommandText = "car_rate_rule_select_names"
	adoCmd.CommandType = adCmdStoredProc
	adoCmd.Parameters.Append adoCmd.CreateParameter("@rate_rule_id",      3, 1, 0, Null)
	adoCmd.Parameters.Append adoCmd.CreateParameter("@user_id",           3, 1, 0, strUserId)
	adoCmd.Parameters.Append adoCmd.CreateParameter("@rate_rule_type_cd", 3, 1, 0, 1)
	adoCmd.Parameters.Append adoCmd.CreateParameter("@rule_status",     200, 1, 1, "E")

	Set adoRS18a = adoCmd.Execute
	Set adoRS18b = adoCmd.Execute
	
	If Len(strSelectedRules) > 0 Then

		Set adoCmd = CreateObject("ADODB.Command")
	
		adoCmd.ActiveConnection =  strConn
		adoCmd.CommandText = "car_rate_rule_bulk_update_utilization"
		adoCmd.CommandType = adCmdStoredProc

		adoCmd.Parameters.Append adoCmd.CreateParameter("@rule_id_string", 200, 1, 4028, strSelectedRules)
		

		Dim intUtilCount
		Dim strUtil
		Dim strUtilParam
			
		
		For intUtilCount = 0 To 7
		
			strUtil = "util_min_" & CStr(intUtilCount)
			strUtilParam = "@util_min_" & CStr(intUtilCount)
			
			If Trim(Request.Form(strUtil)) = "" Then
				adoCmd.Parameters.Append adoCmd.CreateParameter(strUtilParam, 2, 1, 0, Null)
			Else
				adoCmd.Parameters.Append adoCmd.CreateParameter(strUtilParam, 2, 1, 0, Request.Form(strUtil))
			
			End If
	
		Next

		For intUtilCount = 0 To 7
		
			strUtil = "util_max_" & CStr(intUtilCount)
			strUtilParam = "@util_max_" & CStr(intUtilCount)
			
			If Trim(Request.Form(strUtil)) = "" Then
				adoCmd.Parameters.Append adoCmd.CreateParameter(strUtilParam, 2, 1, 0, Null)
			Else
				adoCmd.Parameters.Append adoCmd.CreateParameter(strUtilParam, 2, 1, 0, Request.Form(strUtil))
			
			End If
	
		Next


'		adoCmd.Parameters.Append adoCmd.CreateParameter("@util_min_0", 2, 1, 0, Request.Form("util_min_0"))
'		adoCmd.Parameters.Append adoCmd.CreateParameter("@util_min_1", 2, 1, 0, Request.Form("util_min_1"))
'		adoCmd.Parameters.Append adoCmd.CreateParameter("@util_min_2", 2, 1, 0, Request.Form("util_min_2"))
'		adoCmd.Parameters.Append adoCmd.CreateParameter("@util_min_3", 2, 1, 0, Request.Form("util_min_3"))
'		adoCmd.Parameters.Append adoCmd.CreateParameter("@util_min_4", 2, 1, 0, Request.Form("util_min_4"))
'		adoCmd.Parameters.Append adoCmd.CreateParameter("@util_min_5", 2, 1, 0, Request.Form("util_min_5"))
'		adoCmd.Parameters.Append adoCmd.CreateParameter("@util_min_6", 2, 1, 0, Request.Form("util_min_6"))
'		adoCmd.Parameters.Append adoCmd.CreateParameter("@util_min_7", 2, 1, 0, Request.Form("util_min_7"))
'		adoCmd.Parameters.Append adoCmd.CreateParameter("@util_max_0", 2, 1, 0, Request.Form("util_max_0"))
'		adoCmd.Parameters.Append adoCmd.CreateParameter("@util_max_1", 2, 1, 0, Request.Form("util_max_1"))
'		adoCmd.Parameters.Append adoCmd.CreateParameter("@util_max_2", 2, 1, 0, Request.Form("util_max_2"))
'		adoCmd.Parameters.Append adoCmd.CreateParameter("@util_max_3", 2, 1, 0, Request.Form("util_max_3"))
'		adoCmd.Parameters.Append adoCmd.CreateParameter("@util_max_4", 2, 1, 0, Request.Form("util_max_4"))
'		adoCmd.Parameters.Append adoCmd.CreateParameter("@util_max_5", 2, 1, 0, Request.Form("util_max_5"))
'		adoCmd.Parameters.Append adoCmd.CreateParameter("@util_max_6", 2, 1, 0, Request.Form("util_max_6"))
'		adoCmd.Parameters.Append adoCmd.CreateParameter("@util_max_7", 2, 1, 0, Request.Form("util_max_7"))
		
		adoCmd.Execute
		
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
<title>Rate-Monitor by Rate-Highway, Inc. | Alerts! | Bulk Update</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<script language="Javascript" type="text/javascript" src="inc/sitewide.js"></script>
<script language="Javascript" type="text/javascript" src="inc/header_menu_support.js" ></script>
<script language="Javascript" type="text/javascript" src="inc/js_calendar_v2.js"></script>
<script language="Javascript" type="text/javascript" src="inc/validate2.js"></script>
<script language="JavaScript" type="text/javascript" src="inc/multiple_select_support.js"></script>
<script language="JavaScript" type="text/javascript" src="inc/multiple_select_support2.js"></script>
<style type="text/css" >
<!--
.profile_header { height="48" text-align:left; padding-left:3; padding-right:3; padding-top:0; background-color:#879AA2; font-family:Verdana, Arial, Helvetica, sans-serif; font-size:10pt; vertical-align:bottom; text-align:left }
.profile_light { height="68" text-align:left; padding-left:3; padding-right:3; padding-top:3; background-color:#CFD7DB; font-family:Verdana, Arial, Helvetica, sans-serif; font-size:10pt; vertical-align:bottom; text-align:left }
.profile_dark { height="48" text-align:left; padding-left:3; padding-right:3; padding-top:3; background-color:#B2BEC4; font-family:Verdana, Arial, Helvetica, sans-serif; font-size:10pt; vertical-align:bottom; text-align:left }
.style1 {
	text-align: center;
}
.style2 {
	font-family: Verdana;
	font-size: x-small;
	color: #080000;
}
.style3 {
	background-color: #CFD7DB;
}
-->
</style>
<link rel="stylesheet" type="text/css" href="inc/rh_standard.css">
<link rel="stylesheet" type="text/css" href="inc/rh_report.css">
<link rel="stylesheet" type="text/css" href="inc/sitewide.css">
<script language="Javascript" type="text/javascript" src="inc/sitewide.js"></script>
<script language="Javascript" type="text/javascript" src="inc/header_menu_support.js" ></script>
</head>

<SCRIPT LANGUAGE="JavaScript" type="text/javascript">
function CreateAlert()
{ 
	var valid_form = true;
	var numSelected = 0;
	var i;
	
	var listBox = document.update_utilization.id_selected;
	var len = listBox.length;
	for(var x=0;x<len;x++){
		listBox.options[x].selected= true;
	}


	if (valid_form) {
		//// change debug to true for debug messages
		//alert("1about to transfer to " + window.document.create_alert.action.value);
		////window.document.create_alert.action = "car_rate_rule_insert1.asp?debug=false";
		//window.document.create_alert.txtaction.value = "car_rate_rule_insert1.asp?debug=true";
		//window.document.create_alert.action.value = "car_rate_rule_insert1.asp?debug=true";
		//alert("2about to transfer to " + window.document.create_alert.action.value);
		////window.document.create_alert.txtaction.value = window.document.create_alert.action.value ;
		window.document.update_utilization.submit();
		return true;
		}
	else {
		return false;
		}
}
</SCRIPT><body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" >
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td background="images/top_tile.gif">
    <img src="images/top.jpg" width="770" height="91"></td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td background="images/b_tile.gif">
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


<br>                
<div align="center">
<table cellpadding="0" cellspacing="0" border="0" bgcolor="#FFFFFF" style="width: 830px">
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
<form name="update_utilization" method="post" OnSubmit="return BulkUpdate();">
<table cellpadding="0" cellspacing="0" border="0" ALIGN="CENTER" class="style3" style="width: 830px">
    <tr>
      <td width="8" height="22">&nbsp;</td>
      <td width="217" height="22">&nbsp;</td>
      <td width="210" height="22">&nbsp;</td>
      <td width="510" height="22" colspan="8">
		&nbsp;</td>
      <td width="262" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="22">&nbsp;</td>
      <td width="217" height="22">&nbsp;</td>
      <td width="210" height="22"><span class="style2">Rules</span><font face="Verdana" size="2" color="#080000">:<br>
		<br>
		&nbsp;&nbsp;&nbsp;&nbsp;Unselected:<br>
		<br>
		</font></td>
      <td width="510" height="22" colspan="9" >
		<select name="id_unselected" style="width:600px; font-family:Verdana; font-size:10pt;" size="45" multiple="multiple">
				<% intLoopCount = 0 %>
                <% While (adoRS18b.EOF = False) And (intLoopCount < 800)  %> 
    	               <% If adoRS18b.Fields("rule_status").Value = "E" Then %>
		                	<option value="<%=adoRS18b.Fields("rate_rule_id").Value %>"><%=adoRS18b.Fields("alert_desc").Value %></option>
		                <% End If %> 
	                <%	adoRS18b.MoveNext
                    intLoopCount = intLoopCount + 1
				   Wend
				%>

				<%					   
				   Set adoRS18b = Nothing
				%>
				</select></td>
      
    </tr>
    <tr>
      <td width="8" height="22">&nbsp;</td>
      <td width="217" height="22">&nbsp;</td>
      <td width="210" height="22">&nbsp;</td>
      <td width="510" height="22" colspan="9" class="style1">
<font color="#080000">
                      <img border="0" src="images/down_button.GIF" width="24" height="22"  onclick="moveDualList( document.update_utilization.id_unselected, document.update_utilization.id_selected, false );return false" alt="Select true follow-on rule" >
                      <img border="0" src="images/up_button.GIF"   width="24" height="22"  onclick="moveDualList( document.update_utilization.id_selected, document.update_utilization.id_unselected, false );return false" alt="Un-select true follow-on rule" class="style10" ></font></td>
      
    </tr>
    <tr>
      <td width="8" height="22">&nbsp;</td>
      <td width="217" height="22">&nbsp;</td>
      <td width="210" height="22"><font face="Verdana" size="2" color="#080000">
		&nbsp; &nbsp; Selected:</font></td>
      <td width="510" height="22" colspan="9">
		<font color="#080000">
		<select name="id_selected" style="width:600px; font-family:Verdana; font-size:10pt;" size="15" multiple="multiple">
		</select>
		</font>
	  </td>
     
    </tr>
    
        <tr>
      <td width="8" height="22">&nbsp;</td>
      <td width="217" height="22">&nbsp;</td>
      <td width="210" height="25">&nbsp;</td>
      <td width="510" height="25" colspan="8">&nbsp;</td>
      <td width="262" height="22">&nbsp;</td>
    </tr>
    
        <tr>
      <td width="8" height="22">&nbsp;</td>
      <td width="217" height="22">&nbsp;</td>
      <td width="210" height="25"><font face="Verdana" size="2">
		Utilization Range:</font> </td>
      <td width="510" height="25" colspan="8"><font face="Verdana" size="2">
      (<font face="Verdana" size="2" color="#080000">leave blank to indicate 
		no limit</font>)</font></td>
      <td width="262" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="22">&nbsp;</td>
      <td width="217" height="22">&nbsp;</td>
      <td width="210" height="22"><font face="Verdana" size="2">&nbsp;&nbsp;&nbsp;&nbsp; Days out:</font></td>
      <td width="64" height="22" align="center" class="style5"><font face="Verdana" size="2">Same</font></td>
      <td width="64" height="22" align="center" class="style5"><font face="Verdana" size="2">Next</font></td>
      <td width="64" height="22" align="center" class="style5"><font face="Verdana" size="2">2 - 4</font></td>
      <td height="22" align="center" style="width: 64px" class="style5"><font face="Verdana" size="2">5 - 7</font></td>
      <td width="64" height="22" align="center" class="style5"><font face="Verdana" size="2">8 - 14</font></td>
      <td width="64" height="22" align="center" class="style5"><font face="Verdana" size="2">15 - 30</font></td>
      <td width="64" height="22" align="center" class="style5"><font face="Verdana" size="2">31 - 50</font></td>
      <td height="22" align="center" style="width: 64px" class="style5"><font face="Verdana" size="2">51 +</font></td>
	  <td width="262" height="22" >&nbsp;</td>
      
    </tr>
    <tr>
      <td width="8" style="height: 7px"></td>
      <td width="217" style="height: 7px"></td>
      <td width="210" style="height: 7px"><font face="Verdana" size="2">&nbsp;&nbsp;&nbsp;&nbsp; Maximum:</font></td>
      <td width="64" style="height: 7px" class="style5"><font face="Verdana" size="2"><input type="text"  name="util_max_0" size="20" style="width:59; font-family:Verdana; font-size:10pt; text-align:right; height:22" value=""></font></td>
      <td width="64" style="height: 7px" class="style5"><font face="Verdana" size="2"><input type="text"  name="util_max_1" size="20" style="width:59; font-family:Verdana; font-size:10pt; text-align:right; height:22" value=""></font></td>
      <td width="64" style="height: 7px" class="style5"><font face="Verdana" size="2"><input type="text"  name="util_max_2" size="20" style="width:59; font-family:Verdana; font-size:10pt; text-align:right; height:22" value=""></font></td>
      <td style="height: 7px; width: 64px;" class="style5"><font face="Verdana" size="2"><input type="text"  name="util_max_3" size="20" style="width:59; font-family:Verdana; font-size:10pt; text-align:right; height:22" value=""></font></td>
      <td width="64" style="height: 7px" class="style5"><font face="Verdana" size="2"><input type="text"  name="util_max_4" size="20" style="width:59; font-family:Verdana; font-size:10pt; text-align:right; height:22" value=""></font></td>
      <td width="64" style="height: 7px" class="style5"><font face="Verdana" size="2"><input type="text"  name="util_max_5" size="20" style="width:59; font-family:Verdana; font-size:10pt; text-align:right; height:22" value=""></font></td>
      <td width="64" style="height: 7px" class="style5"><font face="Verdana" size="2"><input type="text"  name="util_max_6" size="20" style="width:59; font-family:Verdana; font-size:10pt; text-align:right; height:22" value=""></font></td>
      <td style="height: 7px; width: 64px;" class="style5"><font face="Verdana" size="2"><input type="text"  name="util_max_7" size="20" style="width:59; font-family:Verdana; font-size:10pt; text-align:right; height:22" value=""></font></td>
	  <td width="262" style="height: 7px" >&nbsp;</td>
	  
    </tr>
    <tr>
      <td width="8" style="height: 24px"></td>
      <td width="217" style="height: 24px"></td>
      <td width="210" style="height: 24px"><font face="Verdana" size="2">&nbsp;&nbsp;&nbsp;&nbsp; Minimum:</font> </td>
      <td width="64" style="height: 24px" class="style5"><font face="Verdana" size="2"><input type="text" name="util_min_0" size="20" style="width:59; font-family:Verdana; font-size:10pt; text-align:right; height:22" value=""></font></td>
      <td width="64" style="height: 24px" class="style5"><font face="Verdana" size="2"><input type="text" name="util_min_1" size="20" style="width:59; font-family:Verdana; font-size:10pt; text-align:right; height:22" value=""></font></td>
      <td width="64" style="height: 24px" class="style5"><font face="Verdana" size="2"><input type="text" name="util_min_2" size="20" style="width:59; font-family:Verdana; font-size:10pt; text-align:right; height:22" value=""></font></td>
      <td style="height: 24px; width: 64px;" class="style5"><font face="Verdana" size="2"><input type="text" name="util_min_3" size="20" style="width:59; font-family:Verdana; font-size:10pt; text-align:right; height:22" value=""></font></td>
      <td width="64" style="height: 24px" class="style5"><font face="Verdana" size="2"><input type="text" name="util_min_4" size="20" style="width:59; font-family:Verdana; font-size:10pt; text-align:right; height:22" value=""></font></td>
      <td width="64" style="height: 24px" class="style5"><font face="Verdana" size="2"><input type="text" name="util_min_5" size="20" style="width:59; font-family:Verdana; font-size:10pt; text-align:right; height:22" value=""></font></td>
      <td width="64" style="height: 24px" class="style5"><font face="Verdana" size="2"><input type="text" name="util_min_6" size="20" style="width:59; font-family:Verdana; font-size:10pt; text-align:right; height:22" value=""></font></td>
      <td style="height: 24px; width: 64px;" class="style5"><font face="Verdana" size="2"><input type="text" name="util_min_7" size="20" style="width:59; font-family:Verdana; font-size:10pt; text-align:right; height:22" value=""></font></td>
	  <td width="262" style="height: 24px">&nbsp;</td>
      
    </tr>
    <tr>
      <td width="8" height="22">&nbsp;</td>
      <td width="217" height="22">&nbsp;</td>
      <td width="210" height="22">
      &nbsp;</td>
      <td width="510" height="22" colspan="8">
      <input type="checkbox" name="util_in_percent" id="util_in_percent" value="True" checked disabled ><label for="util_in_percent"><font size="2">values 
      are listed as percentages (please do not include percent signs)</font></label></td>
      <td width="262" height="22">&nbsp;</td>
    </tr>

    <tr>
      <td width="8" height="22">&nbsp;</td>
      <td width="217" height="22">&nbsp;</td>
      <td width="210" height="22">
      &nbsp;</td>
      <td width="510" height="22" colspan="8">
      &nbsp;</td>
      <td width="262" height="22">&nbsp;</td>
    </tr>

    <tr>
      <td width="8" height="22">&nbsp;</td>
      <td width="217" height="22">&nbsp;</td>
      <td width="210" height="22">
      &nbsp;</td>
      <td width="510" height="22" colspan="8" class="style1">
      <input name="submit" type="submit" id="submit" value="  Update  " class="rh_button">
</td>
      <td width="262" height="22">&nbsp;</td>
    </tr>

    <tr>
      <td width="8" height="22">&nbsp;</td>
      <td width="217" height="22">&nbsp;</td>
      <td width="210" height="22">
      &nbsp;</td>
      <td width="510" height="22" colspan="8">
      <input name="remember" type="checkbox" value="true"></td>
      <td width="262" height="22">&nbsp;</td>
    </tr>

<tr bgcolor="#000000" height="1" >
<td><img src="images/pixel.gif" width="1" height="1" alt=""></td>
<td><img src="images/pixel.gif" width="1" height="1" alt=""></td>
<td><img src="images/pixel.gif" width="1" height="1" alt=""></td>
<td><img src="images/pixel.gif" width="1" height="1" alt=""></td>
<td><img src="images/pixel.gif" width="1" height="1" alt=""></td>
<td><img src="images/pixel.gif" width="1" height="1" alt=""></td>
<td><img src="images/pixel.gif" width="1" height="1" alt=""></td>
<td><img src="images/pixel.gif" width="1" height="1" alt=""></td>
<td><img src="images/pixel.gif" width="1" height="1" alt=""></td>
<td><img src="images/pixel.gif" width="1" height="1" alt=""></td>
<td><img src="images/pixel.gif" width="1" height="1" alt=""></td>
<td><img src="images/pixel.gif" width="1" height="1" alt=""></td>
</tr>
</table>
</form>
<!--#INCLUDE FILE="footer.asp"-->
</body>
</html>
