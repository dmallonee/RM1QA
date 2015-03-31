<%@ Language=VBScript %>
<!-- #INCLUDE FILE="inc/adovbs.asp" --> 
<!-- #INCLUDE FILE="inc/login_check.asp" -->
<% Response.Expires = -1  
   Response.cachecontrol="private" 
   Response.AddHeader "pragma", "no-cache" 
   
   Dim strCarClasses
   Dim strDataValues
   
   	'on error resume next

   	Server.ScriptTimeout = 180

	strUserId = Request.Cookies("rate-monitor.com")("user_id")
	
	strConn = Session("pro_con")
	
	Rem Get the data sources
	Set adoCmd = CreateObject("ADODB.Command")

	adoCmd.ActiveConnection =  strConn
	adoCmd.CommandText = "org_select"
	adoCmd.CommandType = adCmdStoredProc

	adoCmd.Parameters.Append adoCmd.CreateParameter("@org_id",   3, 1,  0, Null)
	adoCmd.Parameters.Append adoCmd.CreateParameter("@user_id",  3, 1,  0, strUserId)
		
	Set adoRS = adoCmd.Execute

	Rem Get the data sources
	Set adoCmd = CreateObject("ADODB.Command")

	adoCmd.ActiveConnection =  strConn
	adoCmd.CommandText = "user_city_select"
	adoCmd.CommandType = adCmdStoredProc

	adoCmd.Parameters.Append adoCmd.CreateParameter("@user_id",  3, 1,  0, strUserId)
		
	Set adoRS1 = adoCmd.Execute
	
	strCityCd = Request("city_cd")
	If IsDate( Request("util_date")) Then
		datUtilDate = Request("util_date")
	Else
		datUtilDate = Now
	End If


	If strCityCd <> "" And IsDate(datUtilDate) Then

		Rem Get the data sources
		Set adoCmd = CreateObject("ADODB.Command")

		adoCmd.ActiveConnection =  strConn
		adoCmd.CommandText = "car_utilization_select"
		adoCmd.CommandType = adCmdStoredProc

		adoCmd.Parameters.Append adoCmd.CreateParameter("@org_id",        3, 1, 0, adoRS.Fields("org_id").Value)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@city_cd",     200, 1, 6, strCityCd)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@car_type_cd", 200, 1, 4, Null)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@fleet_date",  135, 1, 0, datUtilDate)
		
		Set adoRS2 = adoCmd.Execute


		Rem Get the data sources
		Set adoCmd = CreateObject("ADODB.Command")

		adoCmd.ActiveConnection =  strConn
		adoCmd.CommandText = "car_rental_transaction_date_check"
		adoCmd.CommandType = adCmdStoredProc

		adoCmd.Parameters.Append adoCmd.CreateParameter("@org_id",        3, 1, 0, adoRS.Fields("org_id").Value)

		Set adoRS3 = adoCmd.Execute
        	if isnull(adoRS3.Fields("most_recent")) then
		    	Set adoRS2 = CreateObject("ADODB.Recordset")
    	    		datCurrentDate = ""
		    	datCurrentTime = ""
        	else
		    	datCurrentDate = FormatDateTime(adoRS3.Fields("most_recent").Value, 2) 
		
		    	datCurrentTime = FormatDateTime(adoRS3.Fields("most_recent").Value, 3)
		    	If datCurrentTime = "12:00:00 AM" Then
			    datCurrentTime = ""
		    	End If
		End if
	Else
		Set adoRS2 = CreateObject("ADODB.Recordset")
		datCurrentDate = ""
		datCurrentTime = ""

	End If
	
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
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html xmlns:m="http://schemas.microsoft.com/office/2004/12/omml" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office">
<head>
<meta http-equiv="Content-Language" content="en-us">

<title>Rate-Monitor by Rate-Highway, Inc. |&nbsp; Utilization Settings</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<link rel="stylesheet" type="text/css" href="inc/sitewide.css">
<script type="text/javascript" language="javascript" src="inc/sitewide.js" ></script>
<script type="text/javascript" language="javascript" src="inc/pupdate.js"></script>
<script type="text/JavaScript" language="JavaScript" >
<!--
    function MM_preloadImages() { //v3.0
        var d = document; if (d.images) {
            if (!d.MM_p) d.MM_p = new Array();
            var i, j = d.MM_p.length, a = MM_preloadImages.arguments; for (i = 0; i < a.length; i++)
                if (a[i].indexOf("#") != 0) { d.MM_p[j] = new Image; d.MM_p[j++].src = a[i]; } 
        }
    }

    function MM_swapImgRestore() { //v3.0
        var i, x, a = document.MM_sr; for (i = 0; a && i < a.length && (x = a[i]) && x.oSrc; i++) x.src = x.oSrc;
    }

    function MM_findObj(n, d) { //v4.01
        var p, i, x; if (!d) d = document; if ((p = n.indexOf("?")) > 0 && parent.frames.length) {
            d = parent.frames[n.substring(p + 1)].document; n = n.substring(0, p);
        }
        if (!(x = d[n]) && d.all) x = d.all[n]; for (i = 0; !x && i < d.forms.length; i++) x = d.forms[i][n];
        for (i = 0; !x && d.layers && i < d.layers.length; i++) x = MM_findObj(n, d.layers[i].document);
        if (!x && d.getElementById) x = d.getElementById(n); return x;
    }

    function MM_swapImage() { //v3.0
        var i, j = 0, x, a = MM_swapImage.arguments; document.MM_sr = new Array; for (i = 0; i < (a.length - 2); i += 3)
            if ((x = MM_findObj(a[i])) != null) { document.MM_sr[j++] = x; if (!x.oSrc) x.oSrc = x.src; x.src = a[i + 2]; }
    }


    function openWindow(theURL, winName, features) { //v2.0
        window.open(theURL, winName, features);
    }

    function DisableTSDinfo() {

        var s = document.getElementById("update_rezcentral");
        var v = s.options[s.selectedIndex].text;


        if (v == 'Yes') {
            document.utilization_update.tsd_customer_number.disabled = false;
            document.utilization_update.tsd_passcode.disabled = false;
        }
        else {
            document.utilization_update.tsd_customer_number.disabled = true;
            document.utilization_update.tsd_passcode.disabled = true;
        }

    }

//-->
</script>



<link rel="stylesheet" type="text/css" href="inc/rh_standard.css">
<style type="text/css" >
<!--
.profile_header { height="48" text-align:left; padding-left:3; padding-right:3; padding-top:0; background-color:#879AA2; font-family:Verdana, Arial, Helvetica, sans-serif; font-size:10pt; vertical-align:bottom; text-align:left }
.profile_light { height="68" text-align:left; padding-left:3; padding-right:3; padding-top:3; background-color:#CFD7DB; font-family:Verdana, Arial, Helvetica, sans-serif; font-size:10pt; vertical-align:bottom; text-align:left }
.profile_dark { height="48" text-align:left; padding-left:3; padding-right:3; padding-top:3; background-color:#B2BEC4; font-family:Verdana, Arial, Helvetica, sans-serif; font-size:10pt; vertical-align:bottom; text-align:left }
.style2 {
	font-size: x-small;
}
.style3 {
	font-size: x-small;
	text-align: right;
}
.style5 {
	text-align: center;
	font-size: medium;
}
.style6 {
	text-align: center;
	color: #FF0000;
}
.style7 {
	border-collapse: collapse;
}
.style10 {
	text-align: right;
	border-style: solid;
	border-width: 0;
}
numeric_input {
	text-align: right;
}
.style12 {
	border-style: solid;
	border-width: 0;
}
}
.UtilGridValue {
	border: 0 solid #FFFFFF;
	text-align: center;
}


-->
</style>
<base target="_self">
</head>

<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="MM_preloadImages('images/b_search_pro_on.gif','images/b_search_que_on.gif','images/b_search_cri_on.gif','images/b_rate_on.gif','images/b_alert_on.gif','images/b_user_on.gif','images/b_system_on.gif')" >
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td background="images/top_tile.gif"><img src="images/top.jpg" width="770" height="91"></td>
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
    <td background="images/med_bar_tile.gif"><img src="images/med_bar.gif" width="12" height="8"></td>
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
                      <td><div align="right">
                  <font face="Vendana, Arial, Helvetica, sans-serif" size="2" color="#FFFFFF">
                  User: <%=Request.Cookies("rate-monitor.com")("user_name")%></font></div></td>
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
        <td><img src="images/h_system.gif" width="368" height="31"></td>
        <td><map name="logout_map">
        <area alt="Click to log out of Rate-Monitor" href="http://www.rate-monitor.com/" shape="rect" coords="337, 10, 394, 25">
        </map>
        <img src="images/bottom_right.gif" width="402" height="31" usemap="#logout_map" border="0"></td>
      </tr>
    </table>
    </td>
  </tr>
</table>
<p>&nbsp;&nbsp;&nbsp;<br>&nbsp;
<font size="2" face="Vendana, Arial, Helvetica, sans-serif">
<% Select Case strUserId %>
<% Case 333 %>
&nbsp;<a href="javascript:not_enabled()" >[custom city codes]</a>
&nbsp;<a href="javascript:not_enabled()">[system performance]</a>
&nbsp;<a href="javascript:not_enabled()">[RezCentral settings]</a>
&nbsp;<b>[utilization settings]</b>
&nbsp;<a href="javascript:not_enabled()">[utilization car groups]</a>
&nbsp;<a href="javascript:not_enabled()">[one-way settings]</a>
<% Case 15  %>
&nbsp;<a href="javascript:not_enabled()" >[custom city codes]</a>
&nbsp;<a href="system_performance.asp">[system performance]</a>
&nbsp;<a href="early_late_charge_maint.asp">[Early/Late Charges]</a>
&nbsp;<b>[utilization settings]</b>
&nbsp;<a title="click to manage the utilization car groups" href="system_utilization_car_groups.asp">[utilization car groups]</a>
&nbsp;<a title="click to manage the utilization car groups" href="system_drop_charge_factors.asp">[one-way settings]</a>
<% Case 40 %>
&nbsp;<a href="javascript:not_enabled()" >[custom city codes]</a>
&nbsp;<a href="system_performance.asp">[system performance]</a>
&nbsp;<a href="early_late_charge_maint.asp">[Early/Late Charges]</a>
&nbsp;<a href="rezcentral_tethering.asp">[RezCentral settings]</a>
&nbsp;<b>[utilization settings]</b>
&nbsp;<a title="click to manage the utilization car groups" href="system_utilization_car_groups.asp">[utilization car groups]</a>
&nbsp;<a title="click to manage the utilization car groups" href="system_drop_charge_factors.asp">[one-way settings]</a>
<% Case Else %>
&nbsp;<a href="javascript:not_enabled()" >[custom city codes]</a>
&nbsp;<a href="system_performance.asp">[system performance]</a>
&nbsp;<a href="rezcentral_tethering.asp">[RezCentral settings]</a>
&nbsp;<b>[utilization settings]</b>
&nbsp;<a title="click to manage the utilization car groups" href="system_utilization_car_groups.asp">[utilization car groups]</a>
&nbsp;<a title="click to manage the utilization car groups" href="system_drop_charge_factors.asp">[one-way settings]</a>
<% End Select %>

</font><br>&nbsp;</p>
<table border="0" style="border-collapse: collapse" bordercolor="#FFFFFF" width="1114" cellspacing="0" height="4">
    <tr>
      <td background="images/ruler.gif"></td>
      <td background="images/ruler.gif"></td>
      <td background="images/ruler.gif"></td>
      <td background="images/ruler.gif"></td>
      <td background="images/ruler.gif"></td>
      <td background="images/ruler.gif"></td>
      <td background="images/ruler.gif"></td>
      <td background="images/ruler.gif"></td>
      <td background="images/ruler.gif"></td>
      <td background="images/ruler.gif"></td>
      <td background="images/ruler.gif"></td>
      <td background="images/ruler.gif"></td>
      <td background="images/ruler.gif"></td>
    </tr>
  </table>
<p>&nbsp;
     <p>&nbsp;<form method="post" name="utilization_update" action="system_update.asp">
  <table style="width: 100%">
	<tr>
		<td class="style3">My system reads utilization from:&nbsp;</td>
		<td colspan="2">
  <p >
  <select size="1" name="utilization_input_id" class="style2" atomicselection="true">
  <% Select Case adoRS.Fields("utilization_input_id").Value %>
  	<% Case 0 %>
	  <option selected value="0">Manual Input - values displayed below</option>
	  <option value="1">TSD counter system</option>
	  <option value="4">Custom feed</option>
	  <option value="2">TSD ASP system</option>
	  <option value="3">Bluebird counter system</option>
  	<% Case 1 %>
	  <option value="0">Manual Input - values displayed below</option>
	  <option selected value="1">TSD counter system</option>
	  <option value="4">Custom feed</option>
	  <option value="2">TSD ASP system</option>
	  <option value="3">Bluebird counter system</option>
  	<% Case 2 %>
	  <option value="0">Manual Input - values displayed below</option>
	  <option value="1">TSD counter system</option>
	  <option value="4">Custom feed</option>
	  <option selected value="2">TSD ASP system</option>
	  <option value="3">Bluebird counter system</option>
  	<% Case 3 %>
	  <option value="0">Manual Input - values displayed below</option>
	  <option value="1">TSD counter system</option>
	  <option value="4">Custom feed</option>
	  <option value="2">TSD ASP system</option>
	  <option selected value="3">Bluebird counter system</option>
  	<% Case 4 %>
	  <option value="0">Manual Input - values displayed below</option>
	  <option value="1">TSD counter system</option>
	  <option selected value="4">Custom feed</option>
	  <option value="2">TSD ASP system</option>
	  <option value="3">Bluebird counter system</option>
  <% End Select %>



  </select></p>
		</td>
	</tr>
	<tr>
		<td class="style3">No-show percentage:&nbsp; </td>
		<td class="style2" colspan="2"><input name="no_show_percentage" type="text" style="width: 42px" value="<%=adoRS.Fields("no_show_percentage").Value %>"> 
		(please do not include the % sign)</td>
	</tr>
	<tr>
		<td class="style3">Update RezCentral:&nbsp;  </td>
		<td class="style12" colspan="2">
		<select name="update_rezcentral" id="update_rezcental" onchange="DisableTSDinfo();">
		<% If adoRS.Fields("update_rezcentral").Value = True Then %>
		<option selected value="1">Yes</option>
		<option value="0">No</option>
		<% Else %>
		<option value="1">Yes</option>
		<option selected value="0">No</option>
		<% End If %>
		</select>
		</td>
	</tr>
	<tr>
		<td class="style3">
		<span style="font-size:11.0pt;line-height:115%;
font-family:&quot;Wingdings 3&quot;;mso-ascii-font-family:Calibri;mso-ascii-theme-font:
minor-latin;mso-fareast-font-family:Calibri;mso-fareast-theme-font:minor-latin;
mso-hansi-font-family:Calibri;mso-hansi-theme-font:minor-latin;mso-bidi-font-family:
&quot;Times New Roman&quot;;mso-bidi-theme-font:minor-bidi;mso-ansi-language:EN-US;
mso-fareast-language:EN-US;mso-bidi-language:AR-SA;mso-char-type:symbol;
mso-symbol-font-family:&quot;Wingdings 3&quot;">
		<span style="mso-char-type:symbol;
mso-symbol-font-family:&quot;Wingdings 3&quot;">Ê</span></span></td>
		<td class="style3" width="110"><label id="lbl_tsd_customer_number" >TSD Customer No:</label></td>
		<td class="style12">
		<input name="tsd_customer_number" type="text" value="<%=adoRS.Fields("tsd_customer_number").Value %>"></td>
	</tr>
	<tr>
		<td class="style3">
		<span style="font-size:11.0pt;line-height:115%;
font-family:&quot;Wingdings 3&quot;;mso-ascii-font-family:Calibri;mso-ascii-theme-font:
minor-latin;mso-fareast-font-family:Calibri;mso-fareast-theme-font:minor-latin;
mso-hansi-font-family:Calibri;mso-hansi-theme-font:minor-latin;mso-bidi-font-family:
&quot;Times New Roman&quot;;mso-bidi-theme-font:minor-bidi;mso-ansi-language:EN-US;
mso-fareast-language:EN-US;mso-bidi-language:AR-SA;mso-char-type:symbol;
mso-symbol-font-family:&quot;Wingdings 3&quot;">
		<span style="mso-char-type:symbol;
mso-symbol-font-family:&quot;Wingdings 3&quot;">Ê</span></span></td>
		<td class="style3"><label id="lbl_tsd_passcode" >TSD Passcode:</label></td>
		<td class="style12">
		<input name="tsd_passcode" type="text" value="<%=adoRS.Fields("tsd_passcode").Value %>"></td>
	</tr>
	<tr>
		<td class="style3">Update Bluebird:&nbsp;  </td>
		<td class="style2" colspan="2">
		<select name="update_bluebird">
		<% If adoRS.Fields("update_bluebird").Value = True Then %>
		<option selected value="1">Yes</option>
		<option value="0">No</option>
		<% Else %>
		<option value="1">Yes</option>
		<option selected value="0">No</option>
		<% End If %>
		</select>
		</td>
	</tr>
	<tr>
		<td class="style3">Update TSD:&nbsp;  </td>
		<td class="style2" colspan="2">
		<select name="update_tsd">
		<% If adoRS.Fields("update_tsd").Value = True Then %>
		<option selected value="1">Yes</option>
		<option value="0">No</option>
		<% Else %>
		<option value="1">Yes</option>
		<option selected value="0">No</option>
		<% End If %>
		</select>
		</td>
	</tr>
	<tr>
		<td class="style3">FTP Client Id:&nbsp; </td>
		<td class="style2" colspan="2">
		<input name="ftp_client_id" type="text" value="<%=adoRS.Fields("ftp_client_id").Value %>">
		</td>
	</tr>
	<tr>
		<td class="style3">Time zone Offset:&nbsp; </td>
		<td class="style2" colspan="2"><select name="time_zone_offset">
		<% Select Case adoRS.Fields("time_zone_offset").Value %>
			<% Case -2 %>
				<option selected="" value="-2">+2 Hawaii</option>
				<option value="-1">+1 Alaska</option>
				<option value="0">No offset (PST)</option>
				<option value="1">-1 Mountain</option>
				<option value="2">-2 Central</option>
				<option value="3">-3 Eastern</option>
			<% Case -1 %>
				<option value="-2">+2 Hawaii</option>
				<option selected="" value="-1">+1 Alaska</option>
				<option value="0">No offset (PST)</option>
				<option value="1">-1 Mountain</option>
				<option value="2">-2 Central</option>
				<option value="3">-3 Eastern</option>
			<% Case 0 %>
				<option value="-2">+2 Hawaii</option>
				<option value="-1">+1 Alaska</option>
				<option selected="" value="0">No offset (PST)</option>
				<option value="1">-1 Mountain</option>
				<option value="2">-2 Central</option>
				<option value="3">-3 Eastern</option>
			<% Case 1 %>
				<option value="-2">+2 Hawaii</option>
				<option value="-1">+1 Alaska</option>
				<option value="0">No offset (PST)</option>
				<option selected="" value="1">-1 Mountain</option>
				<option value="2">-2 Central</option>
				<option value="3">-3 Eastern</option>
			<% Case 2 %>
				<option value="-2">+2 Hawaii</option>
				<option value="-1">+1 Alaska</option>
				<option value="0">No offset (PST)</option>
				<option value="1">-1 Mountain</option>
				<option selected="" value="2">-2 Central</option>
				<option value="3">-3 Eastern</option>
			<% Case 3 %>
				<option value="-2">+2 Hawaii</option>
				<option value="-1">+1 Alaska</option>
				<option value="0">No offset (PST)</option>
				<option value="1">-1 Mountain</option>
				<option value="2">-2 Central</option>
				<option selected="" value="3">-3 Eastern</option>

		<% End Select%>
		</select></td>
	</tr>
	<tr>
		<td class="style3">Weekly LOR:&nbsp;&nbsp; </td>
		<td class="style2" colspan="2">
		<input name="weekly_lor" type="text" style="width: 42px" value="<%=adoRS.Fields("weekly_lor").Value %>" size="2"></td>
	</tr>
	<tr>
		<td class="style3">Enable Rule Processing:&nbsp;&nbsp; </td>
		<td class="style2" colspan="2">
        <input name="enable_rule_processing" type="checkbox" value="1"<%if adoRS.Fields("enable_rule_processing").Value = "True" Then Response.Write " checked"%> /></td>
	</tr>
	</table>
	<p align="center">
	<input type="submit" value="Update" name="update_submit" class="rh_button"></p>
	 </FORM>
	<!-- 
	<p align="center">
		Peter - use this report for right now please =&gt;
	<a href="system_utilization_report.asp">utilization report</a></p>
	-->
<form method="post" name="display_utilization" >
<p class="style5">&nbsp;
        Current Utilization by Location<div align="center">
        <table border="0" cellpadding="0" style="width: 600px;" bordercolor="#111111" class="style7">
          <tr>
           <td width="100%" class="boxtitle" colspan="7" style="height: 15px"><font size="2"><b>
           Directions:</b> To manage the utilization settings manually please 
           use this page. Select a city code, then click the view button to 
           display the selected city. Once the system is displaying the city you 
           would like to modify, you may edit the current utilization levels or 
           the total number or cars for that car type at that location. Once you 
           are satisfied with your changes simple press the update button. If 
           you want to discard your changes and not save them, either navigate 
           away from this page or click the view button.</font><p>
           <font size="2">&nbsp; </font>&nbsp;</p>
           </td>
           
          </tr>
          <tr>
           	<td class="boxtitle"  style="height: 15px">&nbsp;</td>
           	<td class="boxtitle"  style="height: 15px">
           &nbsp;</td>
			<td class="style10"  style="height: 15px" colspan="2"> 
           <font size="2">Currently viewing:&nbsp;&nbsp; </font></td>
            <td class="boxtitle" colspan="2"> 
           <select size="1" name="city_cd">
                   <%   While (adoRS1.EOF = False) 
 		                  If adoRS1.Fields("city_cd").Value = strCityCd Then %>
		                    <option selected ><%=adoRS1.Fields("city_cd").Value %></option>		           
		           <%     Else %>	 
		                    <option ><%=adoRS1.Fields("city_cd").Value %></option>
		           <%     End If %>
		 		   <%     adoRS1.MoveNext %>
		           <%   Wend %>
		   </select></td>
            <td class="boxtitle">&nbsp;</td>

			
		  </tr>
          <tr>
           	<td class="boxtitle" >&nbsp;</td>
           	<td class="boxtitle" >&nbsp;</td>
			<td class="style10" colspan="2" >
			<font size="2">Date to view:&nbsp;&nbsp; </font></td>
            <td class="boxtitle" colspan="2">
			<input name="util_date" id="util_date" type="text" value="<%=FormatDateTime(datUtilDate, 2) %>" size="8"><img src="images/cal_button.gif" class="DatePicker" alt="Pick a date to display utilization for" height="20" width="32" onClick="getCalendarFor(document.display_utilization.util_date);return false" >
			</td>
            <td class="boxtitle">&nbsp;</td>

		  </tr>
          <tr>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="style10"  style="height: 15px"><font size="2">Collected on:&nbsp;&nbsp;</font></td>
           <td  class="style10"  style="height: 15px" align="left" ><font size="2"><%=datCurrentDate %></font></td>
           <td  class="style10"  style="height: 15px" align="left" ><font size="2"><%=datCurrentTime %></font></td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           
          </tr>
          <tr>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="style10"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           
          </tr>
          <tr>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px"> 
          <input type=submit value='Display' name=submit caption="Display Utilization" class="rh_button" ></td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           
          </tr>
          <tr>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           
          </tr>
          <tr class="profile_header">
            <td  class="boxtitle">Car Util. Group</td>
            <td  class="boxtitle">Current Util.</td>
            <td  class="boxtitle">Reserved<sup>1</sup></td>
            <td  class="boxtitle">On Rent<sup>2</sup></td>
            <td  class="boxtitle">Canceled<sup>3</sup></td>
            <td  class="boxtitle">Returns<sup>4</sup></td>
            <td  class="boxtitle">Fleet Size</td>
          </tr>
          
          <% 	Dim intCount             	%>
          <% 	Dim intGroups             	%>
          <% 	Dim curUtilization         	%>
          <% 	Dim curLocationUtil        	%>                    
          <% 	intCount = 0	            %>
          <% 	strClass = "profile_dark"	%>
          
          <% If adoRS2.State = adStateOpen Then %>
          <% While (adoRS2.EOF = False) %>
          <%   If strClass = "profile_dark" Then
          	     strClass = "profile_light"
          	   Else
          	     strClass = "profile_dark"
          	   End If
          %>
          <tr  class="<%=strClass %>" >
          	<% strCarClasses = strCarClasses & adoRS2.Fields("car_type_cd").Value & "|" %>
            <td class="boxtitle" style="width: 14%"><font size="2"><%=adoRS2.Fields("car_type_cd").Value %></font></td>
            <td class="boxtitle" style="width: 14%">
            <% strDataValues = strDataValues & FormatNumber(((adoRS2.Fields("fleet_rsvd").Value + adoRS2.Fields("fleet_out").Value - adoRS2.Fields("fleet_invoiced").Value) / adoRS2.Fields("fleet_count").Value) * 100, 2) & "," %>
            <%=FormatPercent((adoRS2.Fields("fleet_rsvd").Value + adoRS2.Fields("fleet_out").Value - adoRS2.Fields("fleet_invoiced").Value) / adoRS2.Fields("fleet_count").Value) %></td>
            <td class="UtilGridValue" style="width: 14%"><font size="2">
       		<a href="system_utilization_reservations.asp?date=<%=Server.URLEncode(datUtilDate) %>&car_type_cd=<%=adoRS2.Fields("car_type_cd").Value %>&org_id=<%=adoRS.Fields("org_id").Value %>&city_cd=<%=strCityCd %>&no_show_percentage=<%=adoRS.Fields("no_show_percentage").Value %> "><%=adoRS2.Fields("fleet_rsvd").Value %></a></font></td>
            <td class="UtilGridValue" style="width: 14%"><font size="2">			
            <a href="system_utilization_on_rent.asp?date=<%=Server.URLEncode(datUtilDate) %>&car_type_cd=<%=adoRS2.Fields("car_type_cd").Value %>&org_id=<%=adoRS.Fields("org_id").Value %>&city_cd=<%=strCityCd %>&no_show_percentage=<%=adoRS.Fields("no_show_percentage").Value %> "><%=adoRS2.Fields("fleet_out").Value %></a>&nbsp;&nbsp;&nbsp;</font>
            </td>
            <td class="UtilGridValue" style="width: 14%"><font size="2"><%=adoRS2.Fields("fleet_canceled").Value %></font>
            <!--  
            <input type="text" name="canceled" size="20" readonly value="<%=adoRS2.Fields("fleet_canceled").Value %>" style="text-align: right; width: 80px;">
            -->
            </td>
            <td class="UtilGridValue" style="width: 14%"><font size="2">
			<%=adoRS2.Fields("fleet_invoiced").Value %>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</font>
            <!-- 
            <td class="boxtitle" style="width: 14%">
            <input type="text" name="returned" size="20" readonly value="<%=adoRS2.Fields("fleet_invoiced").Value %>" style="text-align: right; width: 80px;">
            -->
            </td>
            <!-- 
            <td class="boxtitle" style="width: 14%">
            <input type="text" name="total" size="20" readonly value="<%=adoRS2.Fields("fleet_count").Value %>" style="text-align: right; width: 80px;">
            -->
            <td class="UtilGridValue" style="width: 14%"><font size="2">
			<%=adoRS2.Fields("fleet_count").Value %>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</font>
			</td>
          </tr>
          <%    totreserve = totreserve + adoRS2.Fields("fleet_rsvd").Value %>
		  <%    totrent = totrent + adoRS2.Fields("fleet_out").Value %>
		  <%    totcancel = totcancel + adoRS2.Fields("fleet_canceled").Value %>
		  <%    totreturn = totreturn + adoRS2.Fields("fleet_invoiced").Value %>
		  <%    totfleet = totfleet + adoRS2.Fields("fleet_count").Value %>
          <% 	intCount = intCount + 1	%>
          <%   adoRS2.MoveNext         	%>
          <% Wend                     	%>
          <% End If 					%>
 		  <% if(totfleet<>0) Then
                totpct = FormatNumber(((totreserve + totrent - totreturn) / totfleet) * 100, 2)
             end if%>
		  <tr class="profile_header">
           <td  class="boxtitle">TOTAL</td>
            <td  class="boxtitle"><%=totpct%>%</td>
            <td  class="boxtitle"><%=totreserve%></td>
            <td  class="boxtitle"><%=totrent%></td>
            <td  class="boxtitle"><%=totcancel%></td>
            <td  class="boxtitle"><%=totreturn%></td>
            <td  class="boxtitle"><%=totfleet%></td>
		  </tr>
         <tr>
          <td colspan="7" class="style2" ><sup>1</sup> Reserved count does not include reservations that have converted to open rentals, cancelations or returned rentals. 
			Reservations that are more than two hours past pick-up are not 
			counted. Reservations are for the day displayed ONLY.<br>
			<sup>2</sup> On Rent is comprised of the open contracts and reservations that 
			will be open on the date displayed. It does not contain the 
			reservations that are expected to convert to open contracts. For 
			displayed date's reservations that will be picked up, please view 
			the reserved count.<br><sup>3</sup> Cancelations have already been removed 
		  from the Reserved count, but are provided here for informational 
		  purposes only.<br>
			<sup>4</sup> Returns are all cars due in that have not yet been returned, but 
			are due on the date listed.</td>
          </tr>
          <tr>
          <td colspan="7" class="style2" >
          <% If adoRS2.State = adStateOpen Then %>
          <%
          
          strCarClasses = Left(strCarClasses, (Len(strCarClasses) - 1))
          strDataValues = Left(strDataValues, (Len(strDataValues) - 1))
          %>
          <!-- 
          <img alt="Utilization Chart" src="http://chart.apis.google.com/chart?cht=bvs&chbh=30,10&chd=t:<%=strDataValues %>&chds=0,100&chs=600x300&chl=<%=strCarClasses %>&chf=c,lg,90,879AA2,0.5,ffffff,0|bg,s,EFEFEF&chco=FDC677" > 
          -->
		  <% End If %>
          </td>
          </tr>
          </table>
        </div>
        <p class="style6">
        <% If strCityCd = "" Then %>
        <strong>Please select a city code and click the display button below to display utilization</strong>
        
        <% End If %>
        <p align="center">&nbsp;
        &nbsp;&nbsp; 
          </p>
        <p align="center">&nbsp;</p>
  </FORM>

  <table border="0" style="border-collapse: collapse" bordercolor="#FFFFFF" width="1114" cellspacing="0" height="4" id="table1">
    <tr>
      <td background="images/ruler.gif"></td>
      <td background="images/ruler.gif"></td>
      <td background="images/ruler.gif"></td>
      <td background="images/ruler.gif"></td>
      <td background="images/ruler.gif"></td>
      <td background="images/ruler.gif"></td>
      <td background="images/ruler.gif"></td>
      <td background="images/ruler.gif"></td>
      <td background="images/ruler.gif"></td>
      <td background="images/ruler.gif"></td>
      <td background="images/ruler.gif"></td>
      <td background="images/ruler.gif"></td>
      <td background="images/ruler.gif"></td>
    </tr>
  </table>
<p align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">© 
2002 - 2013 - All rights reserved<br>
<b>Rate-Highway, Inc.</b><br>
</font>18001 Cowan, 
    Suite F<br />
    Irvine, CA 92614<br />
    (949) 614-0751&nbsp; </font>
<p align="center">&nbsp;</p>
<p>&nbsp;</p>
<script language="JavaScript"type="text/JavaScript">
<!--
    if (document.all) {
        document.writeln("<div id=\"PopUpCalendar\" style=\"position:absolute; left:0px; top:0px; z-index:7; width:200px; height:77px; overflow: visible; visibility: hidden; background-color: #FFFFFF; border: 1px none #000000\" onMouseOver=\"if(ppcTI){clearTimeout(ppcTI);ppcTI=false;}\" onMouseOut=\"ppcTI=setTimeout(\'hideCalendar()\',500)\">");
        document.writeln("<div id=\"monthSelector\" style=\"position:absolute; left:0px; top:0px; z-index:9; width:181px; height:27px; overflow: visible; visibility:inherit\">");
    }
    else if (document.layers) {
        document.writeln("<layer id=\"PopUpCalendar\" pagex=\"0\" pagey=\"0\" width=\"200\" height=\"200\" z-index=\"100\" visibility=\"hide\" bgcolor=\"#FFFFFF\" onMouseOver=\"if(ppcTI){clearTimeout(ppcTI);ppcTI=false;}\" onMouseOut=\"ppcTI=setTimeout('hideCalendar()',500)\">");
        document.writeln("<layer id=\"monthSelector\" left=\"0\" top=\"0\" width=\"181\" height=\"27\" z-index=\"9\" visibility=\"inherit\">");
    }
    else {
        document.writeln("<p><font color=\"#FF0000\"><b>Error ! The current browser is either too old or too modern (usind DOM document structure).</b></font></p>");
    }
 -->
</script>
<noscript><p><font color="#FF0000"><b>JavaScript is not activated !</b></font></p></noscript>
<table border="1" cellspacing="1" cellpadding="2" width="200" bordercolorlight="#000000" bordercolordark="#000000" vspace="0" hspace="0"><form name="ppcMonthList"><tr><td align="center" bgcolor="#CCCCCC"><a href="javascript:moveMonth('Back')" onMouseOver="window.status=' ';return true;"><font face="Arial, Helvetica, sans-serif" size="2" color="#000000"><b> </b></font></a><font face="MS Sans Serif, sans-serif" size="1"> 
<select name="sItem" onMouseOut="if(ppcIE){window.event.cancelBubble = true;}" onChange="switchMonth(this.options[this.selectedIndex].value)" style="font-family: 'MS Sans Serif', sans-serif; font-size: 9pt"><option value="0" selected>2000 • January</option><option value="1">2000 • February</option><option value="2">2000 • March</option><option value="3">2000 • April</option><option value="4">2000 • May</option><option value="5">2000 • June</option><option value="6">2000 • July</option><option value="7">2000 • August</option><option value="8">2000 • September</option><option value="9">2000 • October</option><option value="10">2000 • November</option><option value="11">2000 • December</option><option value="0">2001 • January</option></select></font><a href="javascript:moveMonth('Forward')" onMouseOver="window.status=' ';return true;"><font face="Arial, Helvetica, sans-serif" size="2" color="#000000"><b></b></font></a></td></tr></form></table>
<table border="1" cellspacing="1" cellpadding="2" bordercolorlight="#000000" bordercolordark="#000000" width="200" vspace="0" hspace="0"><tr align="center" bgcolor="#CCCCCC"><td width="20" bgcolor="#FFFFCC"><b><font face="MS Sans Serif, sans-serif" size="1">Su</font></b></td><td width="20"><b><font face="MS Sans Serif, sans-serif" size="1">Mo</font></b></td><td width="20"><b><font face="MS Sans Serif, sans-serif" size="1">Tu</font></b></td><td width="20"><b><font face="MS Sans Serif, sans-serif" size="1">We</font></b></td><td width="20"><b><font face="MS Sans Serif, sans-serif" size="1">Th</font></b></td><td width="20"><b><font face="MS Sans Serif, sans-serif" size="1">Fr</font></b></td><td width="20" bgcolor="#FFFFCC"><b><font face="MS Sans Serif, sans-serif" size="1">Sa</font></b></td></tr></table>
<script language="JavaScript" type="text/JavaScript">
<!--
    if (document.all) {
        document.writeln("</div>");
        document.writeln("<div id=\"monthDays\" style=\"position:absolute; left:0px; top:52px; z-index:8; width:200px; height:17px; overflow: visible; visibility:inherit; background-color: #FFFFFF; border: 1px none #000000\"> </div></div>");
    }
    else if (document.layers) {
        document.writeln("</layer>");
        document.writeln("<layer id=\"monthDays\" left=\"0\" top=\"52\" width=\"200\" height=\"17\" z-index=\"8\" bgcolor=\"#FFFFFF\" visibility=\"inherit\"> </layer></layer>");
    }
    else { /*NOP*/ }
-->
</script>
<!--#INCLUDE FILE="footer.asp"-->
</body>
</html>
<% Set adoRS = Nothing
   Set adoCmd = Nothing 
%>