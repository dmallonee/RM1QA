<%@ Language=VBScript %>
<!-- #INCLUDE FILE="inc/adovbs.asp" --> 
<!-- #INCLUDE FILE="inc/login_check.asp" -->
<% Response.Expires = -1  
   Response.cachecontrol="private" 
   Response.AddHeader "pragma", "no-cache" 
   
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
	datUtilDate = Request("util_date")


	If strCityCd <> "" And IsDate(datUtilDate) Then

		Rem Get the data sources
		Set adoCmd = CreateObject("ADODB.Command")

		adoCmd.ActiveConnection =  strConn
		adoCmd.CommandText = "car_utilization_select"
		adoCmd.CommandType = adCmdStoredProc

		adoCmd.Parameters.Append adoCmd.CreateParameter("@org_id",        3, 1, 0, adoRS.Fields("org_id").Value)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@city_cd",     200, 1, 5, strCityCd)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@car_type_cd", 200, 1, 4, Null)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@fleet_date",  135, 1, 0, datUtilDate)
		
		Set adoRS2 = adoCmd.Execute
	Else
		Set adoRS2 = CreateObject("ADODB.Recordset")

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
<html>
<head>
<meta http-equiv="Content-Language" content="en-us">

<title>Rate-Monitor by Rate-Highway, Inc. |&nbsp; User Configuration</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<link rel="stylesheet" type="text/css" href="inc/sitewide.css">
<script src="inc/sitewide.js" language="javascript"></script>
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}


function openWindow(theURL,winName,features) 
{ //v2.0
  window.open(theURL,winName,features);
}
//-->
</script>

<link rel="stylesheet" type="text/css" href="inc/rh_standard.css">
<style>
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
.style11 {
	border: 0 solid #FFFFFF;
	text-align: right;
}
numeric_input {
	text-align: right;
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
        <td><img src="images/h_user_configuration.gif" width="368" height="31"></td>
        <td><map name="logout_map">
        <area alt="Click to log out of Rate-Monitor" href="http://www.rate-monitor.com/" shape="rect" coords="337, 10, 394, 25">
        </map>
        <img src="images/bottom_right.gif" width="402" height="31" usemap="#logout_map" border="0"></td>
      </tr>
    </table>
    </td>
  </tr>
</table>
<p>&nbsp;&nbsp;&nbsp; <br>
&nbsp;<font size="2" face="Vendana, Arial, Helvetica, sans-serif">&nbsp;<a href="javascript:not_enabled()">[company 
settings]</a>&nbsp;<b>[user settings]</b></font><br>
&nbsp;</p>
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
     <p>&nbsp;<form method="post" name="utilization_update" action="utilization_update.asp">
  <table style="width: 100%">
	<tr>
		<td class="style3">My system reads utilization from:&nbsp;</td>
		<td>
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
		<td class="style2"><input name="no_show_percentage" type="text" style="width: 42px" value="<%=adoRS.Fields("no_show_percentage").Value %>"> 
		(please do not include the % sign)</td>
	</tr>
	<tr>
		<td class="style3">Update RezCentral:&nbsp;  </td>
		<td class="style2">
		<select name="update_rezcentral">
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
		<td class="style3">Update Bluebird:&nbsp;  </td>
		<td class="style2">
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
		<td class="style2">
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
		<td class="style2">
		<input name="ftp_client_id" type="text" value="<%=adoRS.Fields("ftp_client_id").Value %>">
		</td>
	</tr>
	<tr>
		<td class="style3">Time zone Offset:&nbsp; </td>
		<td class="style2"><select name="time_zone_offset">
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
		<td class="style2">
		<input name="weekly_lor" type="text" style="width: 42px" value="<%=adoRS.Fields("weekly_lor").Value %>" size="2"></td>
	</tr>
	<tr>
		<td class="style3">&nbsp;</td>
		<td class="style2">&nbsp;</td>
	</tr>
	</table>
	<p align="center">
	<input type="submit" value="Update" name="B2" class="rh_button"></p>
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
			<% If IsDate(datUtilDate) Then %>
			<input name="util_date" type="text" value="<%=FormatDateTime(datUtilDate, 2) %>" size="10">
			<% Else %>
			<input name="util_date" type="text" value="<%=FormatDateTime(now, 2) %>" size="10">
			<% End If %>
			</td>
            <td class="boxtitle">&nbsp;</td>

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
          <tr>
            <td  class="boxtitle"><u><font size="2">Car Type</font></u></td>
            <td  class="boxtitle"><u><font size="2">Current Util.</font></u></td>
            <td  class="boxtitle"><u><font size="2">Reserved*</font></u></td>
            <td  class="boxtitle"><u><font size="2">On Rent**</font></u></td>
            <td  class="boxtitle"><u><font size="2">Canceled</font></u></td>
            <td  class="boxtitle"><u><font size="2">Returns***</font></u></td>
            <td  class="boxtitle"><u><font size="2">Fleet Size</font></u></td>
          </tr>
          
          <% If adoRS2.State = adStateOpen Then %>
          <% While (adoRS2.EOF = False) %>
          <tr>
            <td class="boxtitle" style="width: 14%"><font size="2"><%=adoRS2.Fields("car_type_cd").Value %></font></td>
            <td class="boxtitle" style="width: 14%">
            <input type="text" name="utilization" size="20" readonly value="<%=FormatPercent((adoRS2.Fields("fleet_rsvd").Value + adoRS2.Fields("fleet_out").Value - adoRS2.Fields("fleet_invoiced").Value) / adoRS2.Fields("fleet_count").Value) %>" style="text-align: right; width: 80px;"></td>
            <td class="style11" style="width: 14%"><font size="2">
			<a href="system_utilization_res_detail.asp?date=<%=Server.URLEncode(datUtilDate) %>&car_type_cd=<%=adoRS2.Fields("car_type_cd").Value %>&org_id=<%=adoRS.Fields("org_id").Value %>&city_cd=<%=strCityCd %>"><%=adoRS2.Fields("fleet_rsvd").Value %></a>&nbsp;&nbsp;&nbsp;</font></td>
            <td class="style11" style="width: 14%"><font size="2">
			<a href="system_utilization_out_detail.asp?date=<%=Server.URLEncode(datUtilDate) %>&car_type_cd=<%=adoRS2.Fields("car_type_cd").Value %>&org_id=<%=adoRS.Fields("org_id").Value %>&city_cd=<%=strCityCd %>"><%=adoRS2.Fields("fleet_out").Value %></a>&nbsp;&nbsp;&nbsp;</font>
            <!-- 
            <td class="boxtitle" style="width: 14%">
            <input type="text" name="on_rent" size="20" readonly value="<%=adoRS2.Fields("fleet_out").Value %>" style="text-align: right; width: 80px;">
            -->
            </td>
            <td class="boxtitle" style="width: 14%">
            <!-- 
            <input type="text" name="canceled" size="20" readonly value="<%=adoRS2.Fields("fleet_canceled").Value %>" style="text-align: right; width: 80px;">
            -->
            </td>
            <td class="style11" style="width: 14%"><font size="2">
			<a href="system_utilization_return_detail.asp?date=<%=Server.URLEncode(datUtilDate) %>&car_type_cd=<%=adoRS2.Fields("car_type_cd").Value %>&org_id=<%=adoRS.Fields("org_id").Value %>&city_cd=<%=strCityCd %>"><%=adoRS2.Fields("fleet_invoiced").Value %></a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</font>
            <!-- 
            <td class="boxtitle" style="width: 14%">
            <input type="text" name="returned" size="20" readonly value="<%=adoRS2.Fields("fleet_invoiced").Value %>" style="text-align: right; width: 80px;">
            -->
            </td>
            <!-- 
            <td class="boxtitle" style="width: 14%">
            <input type="text" name="total" size="20" readonly value="<%=adoRS2.Fields("fleet_count").Value %>" style="text-align: right; width: 80px;">
            -->
            <td class="style11" style="width: 14%"><font size="2">
			<a href="system_utilization_fleet_detail.asp?date=<%=Server.URLEncode(datUtilDate) %>&car_type_cd=<%=adoRS2.Fields("car_type_cd").Value %>&org_id=<%=adoRS.Fields("org_id").Value %>&city_cd=<%=strCityCd %>"><%=adoRS2.Fields("fleet_count").Value %></a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</font>
			</td>
          </tr>
          <%   adoRS2.MoveNext        %>
          <% Wend                     %>
          <% End If %>
          <tr>
          <td colspan="7" class="style2" >* Reserved count does not include reservations that have converted to open rentals, cancelations or returned rentals. 
			Reservations that are more than two hours past pick-up are not 
			counted. Reservations are for the day displayed ONLY.<br>
			** On Rent is comprised of the open contracts and reservations that 
			will be open on the date displayed. It does not contain the 
			reservations that are expected to convert to open contracts. For 
			displayed date's reservations that will be picked up, please view 
			the reserved count.<br>
			*** Returns are all cars due in that have not yet been returned, but 
			are due on the date listed.</td>
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
<!--#INCLUDE FILE="footer.asp"-->
</body>
</html>
<% Set adoRS = Nothing
   Set adoCmd = Nothing 
%>