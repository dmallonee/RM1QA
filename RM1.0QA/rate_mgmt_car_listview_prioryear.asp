<%@ Language=VBScript %>
<!-- #INCLUDE FILE="inc/login_check.asp" -->
<!-- #INCLUDE FILE="inc/adovbs.asp" -->
<% Response.Expires = -1
   Response.cachecontrol="private" 
   Response.AddHeader "pragma", "no-cache"

   Server.ScriptTimeout = 180

	Dim strConn	
	Dim adoCmd	
	Dim adoRS
	Dim adoCmd1	
	Dim adoRS1
	Dim adoCmd2	
	Dim adoRS2

	Dim adoPrices
	Dim strUserId
	Dim strCityCd
	Dim strVendor
	Dim strCarType
	Dim strDataSource
	
	
	strCityCd =      Request.QueryString("city_cd")
	strVendor =      Request.QueryString("vend_cd")
	strLOR =         Request.QueryString("lor")
	strCompVendor =  Request.QueryString("comp_vend_cd")
	strCarType =     Request.QueryString("car_type_cd")
	strDataSource =  Request.QueryString("data_source")
	If Request.QueryString("include_non_closed") = "true" Then
		strNonClosedOnly = 1
	Else
		strNonClosedOnly = 0
	End If
	
	
	strUserId = Request.Cookies("rate-monitor.com")("user_id")

	On Error Resume Next
	
	strConn = Session("pro_con")
	
	Rem Get the data sources

	Set adoCmd = CreateObject("ADODB.Command")

	adoCmd.ActiveConnection =  strConn
	adoCmd.CommandText = "data_source_select"
	adoCmd.CommandType = 4

	adoCmd.Parameters.Append adoCmd.CreateParameter("@lob_id",  3, 1, 0, 2)
	adoCmd.Parameters.Append adoCmd.CreateParameter("@user_id", 3, 1, 0, strUserId)
		
	Set adoRS1 = adoCmd.Execute
	
	If Err.Number <> 0 Then
	   pad="&nbsp;&nbsp;&nbsp;&nbsp;"
	   response.write "<b>Error Occured - selecting data sources<br>"
	   response.write "</b><br>"
	   response.write pad & "Error Number= #<b>" & Err.Number & "</b><br>"
	   response.write pad & "Error Desc.= <b>" & err.description & "</b><br>"
	   response.write pad & "Help Context= <b>" & err.HelpContext & "</b><br>"
	   response.write pad & "Help File Path=<b>" & err.helpfile & "</b><br>"
	   response.write pad & "Error Source= <b>" & err.source & "</b><br><hr>"

	End If
	
	Rem Get the vendors
	Set adoCmd = CreateObject("ADODB.Command")

	adoCmd.ActiveConnection =  strConn
	adoCmd.CommandText = "vendor_select"
	adoCmd.CommandType = 4
		
	Set adoRS2  = adoCmd.Execute
	Set adoRS2a = adoCmd.Execute

	If Err.Number <> 0 Then
	   pad="&nbsp;&nbsp;&nbsp;&nbsp;"
	   response.write "<b>Error Occured - selecting vendors<br>"
	   response.write "</b><br>"
	   response.write pad & "Error Number= #<b>" & Err.Number & "</b><br>"
	   response.write pad & "Error Desc.= <b>" & err.description & "</b><br>"
	   response.write pad & "Help Context= <b>" & err.HelpContext & "</b><br>"
	   response.write pad & "Help File Path=<b>" & err.helpfile & "</b><br>"
	   response.write pad & "Error Source= <b>" & err.source & "</b><br><hr>"

	End If

		
	Rem Get the city
	Set adoCmd = CreateObject("ADODB.Command")

	adoCmd.ActiveConnection =  strConn
	adoCmd.CommandText = "user_city_select"
	adoCmd.CommandType = 4
	
	adoCmd.Parameters.Append adoCmd.CreateParameter("@user_id", 3, 1, 0, strUserId)
		
	Set adoRS3 = adoCmd.Execute

	If Err.Number <> 0 Then
	   pad="&nbsp;&nbsp;&nbsp;&nbsp;"
	   response.write "<b>Error Occured - selecting cities<br>"
	   response.write "</b><br>"
	   response.write pad & "Error Number= #<b>" & Err.Number & "</b><br>"
	   response.write pad & "Error Desc.= <b>" & err.description & "</b><br>"
	   response.write pad & "Help Context= <b>" & err.HelpContext & "</b><br>"
	   response.write pad & "Help File Path=<b>" & err.helpfile & "</b><br>"
	   response.write pad & "Error Source= <b>" & err.source & "</b><br><hr>"

	End If

	Rem Get the car types
	Set adoCmd5 = CreateObject("ADODB.Command")

	adoCmd5.ActiveConnection =  strConn
	adoCmd5.CommandText = "car_type_select"
	adoCmd5.CommandType = 4
		
	Set adoRS5 = adoCmd5.Execute

	If Err.Number <> 0 Then
	   pad="&nbsp;&nbsp;&nbsp;&nbsp;"
	   response.write "<b>Error Occured - selecting car types<br>"
	   response.write "</b><br>"
	   response.write pad & "Error Number= #<b>" & Err.Number & "</b><br>"
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
<title>Rate-Monitor by Rate-Highway, Inc. | Reports | Rate Management</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<link rel="stylesheet" type="text/css" href="inc/rh_standard.css">
<link rel="stylesheet" type="text/css" href="inc/rh_report.css">
<link rel="stylesheet" type="text/css" href="inc/sitewide.css">
<script language="javascript" src="inc/sitewide.js" ></script>
<script language="javascript" src="inc/header_menu_support.js" ></script>
<script language="javascript">
	function NewWindow(mypage, myname, w, h, scroll) {
		var winl = (screen.width - w) / 2;
		var wint = (screen.height - h) / 2;
		winprops = 'height='+h+',width='+w+',top='+wint+',left='+winl+',scrollbars='+scroll+',resizable=0'
		win = window.open(mypage, myname, winprops);
		win.window.focus();
	}
</script>
<style>
<!--
.off_day     { background-color: #879AA2 }
.style1 {
	border-collapse: collapse;
	border-top-width: 0px;
}
.style3 {
	font-size: 11pt;
}
.style4 {
	font-size: x-small;
}
.style5 {
	height= "48" text-align:left;
	padding-left: 3;
	padding-right: 3;
	padding-top: 0;
	background-color: #879AA2;
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-size: x-small;
	vertical-align: bottom;
	text-align: left;
}
.style6 {
	border-collapse: collapse;
}
.style7 {
	border-style: solid;
	border-width: 1px;
}
.style8 {
	border-style: solid;
	border-width: 1px;
	height= "68" text-align:left;
	padding-left: 3;
	padding-right: 3;
	padding-top: 3;
	background-color: #CFD7DB;
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-size: 10pt;
	vertical-align: bottom;
	text-align: right;
}
.style9 {
	border-style: solid;
	border-width: 1px;
	height= "68" text-align:left;
	padding-left: 3;
	padding-right: 3;
	padding-top: 3;
	background-color: #CFD7DB;
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-size: 10pt;
	vertical-align: bottom;
	text-align: center;
}
-->
</style>
</head>

<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onload="MM_preloadImages('images/b_search_pro_on.gif','images/b_search_que_on.gif','images/b_search_cri_on.gif','images/b_rate_on.gif','images/b_alert_on.gif','images/b_user_on.gif','images/b_system_on.gif')" style="font-family: Tahoma; font-size: 10pt">

<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td background="images/top_tile.gif">
    <img src="images/top.jpg" width="770" height="91"></td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td background="images/b_tile.gif">
    <table width="400" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td><img src="images/b_left.jpg" width="62" height="32"></td>
        <td>
        <a href="search_profiles_car.asp" onmouseover="MM_swapImage('s1','','images/b_search_pro_on.gif',1)" onmouseout="MM_swapImgRestore()">
        <img src="images/b_search_pro_of.gif" name="s1" width="183" height="32" border="0" id="s1"></a></td>
        <td>
        <a href="search_queue_car.asp" onmouseover="MM_swapImage('s2','','images/b_search_que_on.gif',1)" onmouseout="MM_swapImgRestore()">
        <img src="images/b_search_que_of.gif" name="s2" width="97" height="32" border="0" id="s2"></a></td>
        <td>
        <a href="search_criteria_car.asp" onmouseover="MM_swapImage('s3','','images/b_search_cri_on.gif',1)" onmouseout="MM_swapImgRestore()">
        <img src="images/b_search_cri_of.gif" name="s3" width="103" height="32" border="0" id="s3"></a></td>
        <td>
        <a href="javascript:not_enabled()" onmouseover="MM_swapImage('ra','','images/b_rate_on.gif',1)" onmouseout="MM_swapImgRestore()">
        <img src="images/b_rate_of.gif" name="ra" width="88" height="32" border="0" id="ra"></a></td>
        <td>
        <a href="javascript:not_enabled()" onmouseover="MM_swapImage('al','','images/b_alert_on.gif',1)" onmouseout="MM_swapImgRestore()">
        <img src="images/b_alert_of.gif" name="al" width="53" height="32" border="0" id="al"></a></td>
        <td>
        <a href="javascript:not_enabled()" onmouseover="MM_swapImage('us','','images/b_user_on.gif',1)" onmouseout="MM_swapImgRestore()">
        <img src="images/b_user_of.gif" name="us" width="126" height="32" border="0" id="us"></a></td>
        <td>
        <a href="javascript:not_enabled()" onmouseover="MM_swapImage('sy','','images/b_system_on.gif',1)" onmouseout="MM_swapImgRestore()">
        <img src="images/b_system_of.gif" name="sy" width="58" height="32" border="0" id="sy"></a></td>
      </tr>
    </table>
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
                  User: <%=Request.Cookies("rate-monitor.com")("user_name")%> </font></div>
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
    <table width="100" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td><img src="images/h_rates_reports.gif" width="368" height="31"></td>
		<td><map name="logout_map">
        <area alt="Click to log out of Rate-Monitor" href="http://www.rate-monitor.com/" shape="rect" coords="337, 10, 394, 25">
        </map>
        <img src="images/bottom_right.gif" width="402" height="31" usemap="#logout_map" border="0"></td>
      </tr>
    </table>
    </td>
  </tr>
</table>
<!-- JUSTTABS TOP OPEN -->
<p>&nbsp;</p>
<table cellpadding="0" cellspacing="0" border="0" ALIGN="CENTER" width="100%" bgcolor="#FFFFFF">
<tr height="1">
<td colspan="1" width="10">&nbsp;</td>
<td rowspan="2" width="335"><a href="rate_calendar_car.asp">
<img src="images/ratecalendar0_ia.GIF" width="100" height="25" hspace="0" vspace="0" border="0" alt="Rate Calendar" description="Rate Calendar"></a><a href="rate_wizard_car.asp"><img src="images/ratewizards1_ia.GIF" width="97" height="25" hspace="0" vspace="0" border="0" alt="Rate Wizards" description="Rate Calendar"></a><img src="images/ratemananagement2_a.GIF" width="138" height="25" hspace="0" vspace="0" border="0" alt="New ALT Text" description="Rate Calendar"></td>
<td colspan="1" >&nbsp;</td>
</tr>
<tr height="1">
<td bgcolor="#000000" colspan="1" height="1">
<img src=pixel.gif width="1" height="1"></td>
<td bgcolor="#000000" colspan="1" height="1">
<img src=pixel.gif width="1" height="1"></td>
</tr>
</table>
<table cellpadding="0" cellspacing="0" border="0" ALIGN="CENTER" width="100%" bgcolor="#FFFFFF">
<tr >
<td  width="1" bgcolor="#000000"><img src=pixel.gif width="1" height="1"></td>
<td colspan=3 bgcolor="#CED7DB">
<table border="0" cellspacing="5" cellpadding="5">
<tr><td>
<font color="#080000">
<!-- JUSTTABS TOP CLOSE -->
    <form method="get" action="rate_mgmt_car.asp" name="calendar_request">
      <table border="1" style="width: 400px;" id="table3" bordercolor="#384F5B" class="style1">
        <tr>
          <td colspan="2" bgcolor="#384F5B" style="border-left:1px solid #384F5B; border-right:1px solid #384F5B; border-top:1px solid #384F5B; "><font color="#FFFFFF">
          &nbsp;<span class="style3">Monthly Rental Activity (list view) </span></font></td>
        </tr>
        <tr>
          <td width="96" style="border-left-style: solid; border-left-width: 1px; border-right-style:none; border-right-width:medium; border-top-style:none; border-top-width:medium; border-bottom-style:none; border-bottom-width:medium" bgcolor="#FFFFFF">
          &nbsp;</td>
          <td style="border-right:1px solid #384F5B; border-left-style:none; border-left-width:medium; border-top-style:none; border-top-width:medium; border-bottom-style:none; border-bottom-width:medium; width: 238px;" bgcolor="#FFFFFF">
          &nbsp;</td>
        </tr>
        <tr>
          <td width="96" style="border-left-style: solid; border-left-width: 1px; border-right-style:none; border-right-width:medium; border-top-style:none; border-top-width:medium; border-bottom-style:none; border-bottom-width:medium" bgcolor="#FFFFFF">
          <font size="2">&nbsp;Location:</font></td>
          <td style="border-right:1px solid #384F5B; border-left-style:none; border-left-width:medium; border-top-style:none; border-top-width:medium; border-bottom-style:none; border-bottom-width:medium; width: 238px;" bgcolor="#FFFFFF">
          <select size="1" name="city_cd" style="width: 70px">
          <% While adoRS3.EOF = False %>
            <% If adoRS3.Fields("city_cd").Value = strCityCd Then %>
			<option selected><%=adoRS3.Fields("city_cd").Value %></option>          
			<% Else %>
			<option><%=adoRS3.Fields("city_cd").Value %></option>          
			<% End If %>
			<% adoRS3.MoveNext %>
		  <% Wend %> 
          </select></td>
        </tr>
        <tr>
          <td width="96" style="border-left-style: solid; border-left-width: 1px; border-right-style:none; border-right-width:medium; border-top-style:none; border-top-width:medium; border-bottom-style:none; border-bottom-width:medium" bgcolor="#FFFFFF"><font size="2">
          &nbsp;Car Type:</font></td>
          <td style="border-right:1px solid #384F5B; border-top-style: none; border-top-width: medium; border-left-style:none; border-left-width:medium; border-bottom-style:none; border-bottom-width:medium; width: 238px;" bgcolor="#FFFFFF">
          <p><select size="1" name="car_type_cd" style="width: 70px">
          <% While adoRS5.EOF = False %>
            <% If adoRS5.Fields("car_type_cd").Value = strCarType Then %>
			<option selected value="<%=adoRS5.Fields("car_type_cd").Value %>"><%=adoRS5.Fields("car_type_cd").Value %></option>
			<% Else %>
			<option value="<%=adoRS5.Fields("car_type_cd").Value %>"><%=adoRS5.Fields("car_type_cd").Value %></option>
			<% End If %>
			<% adoRS5.MoveNext %>
		  <% Wend %>
		  <option value="XXXX">All</option>
          </select></p>
          </td>
        </tr>
        <tr>
<font color="#080000">
          <td width="96" style="border-left-style: solid; border-left-width: 1px; border-right-style:none; border-right-width:medium; border-top-style:none; border-top-width:medium; border-bottom-style:none; border-bottom-width:medium" bgcolor="#FFFFFF">
			<font size="2" color="#080000">&nbsp;Length of Rent:</font></td>
          <td style="border-right:1px solid #384F5B; border-left-style:none; border-left-width:medium; border-top-style:none; border-top-width:medium; border-bottom-style:none; border-bottom-width:medium; width: 238px;" bgcolor="#FFFFFF">
			<select name="lor" style="width: 68px">
			<% For intIndex = 1 To 31 %>
			<% If CInt(intIndex) = CInt(strLOR) Then %>
				<option selected="selected" value="<%=intIndex %>"><%=intIndex %></option>
			<% Else                      %>
				<option value="<%=intIndex %>"><%=intIndex %></option>
			<% End If                    %>
			<% Next %>			
			<option value="0">All</option>
			</select></td>
    </font>
    	</tr>
        <tr>
          <td width="96" style="border-left-style: solid; border-left-width: 1px; border-right-style:none; border-right-width:medium; border-top-style:none; border-top-width:medium; border-bottom-style:none; border-bottom-width:medium" bgcolor="#FFFFFF">
			&nbsp;</td>
          <td style="border-right:1px solid #384F5B; border-top-style: none; border-top-width: medium; border-left-style:none; border-left-width:medium; border-bottom-style:none; border-bottom-width:medium; width: 238px;" bgcolor="#FFFFFF">
          <input name="Checkbox1" type="checkbox"><font size="2">Color code alert levels</font></td>
        </tr>
        <tr>
          <td width="96" style="border-left-style: solid; border-left-width: 1px; border-right-style:none; border-right-width:medium; border-top-style:none; border-top-width:medium; border-bottom-style:none; border-bottom-width:medium" bgcolor="#FFFFFF">
			&nbsp;</td>
          <td style="border-right:1px solid #384F5B; border-top-style: none; border-top-width: medium; border-left-style:none; border-left-width:medium; border-bottom-style:none; border-bottom-width:medium; width: 238px;" bgcolor="#FFFFFF">
          <select name="Select1">
			<option selected="">Week 1</option>
			<option>Week 2</option>
			<option>Week 3</option>
			<option>Week 4</option>
			<option>Week 5</option>
			</select></td>
        </tr>
        <tr>
          <td width="96" style="border-bottom:1px solid #384F5B; border-left-style: solid; border-left-width: 1px; border-right-style:none; border-right-width:medium; border-top-style:none; border-top-width:medium" bgcolor="#FFFFFF">&nbsp;</td>
          <td style="border-right:1px solid #384F5B; border-bottom-style: solid; border-bottom-width: 1px; border-left-style:none; border-left-width:medium; border-top-style:none; border-top-width:medium; width: 238px;" bgcolor="#FFFFFF" class="style4">
			<input name="Checkbox2" type="checkbox" checked="checked" value="true">Display 
			goal indicators</td>
        </tr>
        <tr>
          <td width="96" style="border-bottom:1px solid #384F5B; border-left-style: solid; border-left-width: 1px; border-right-style:none; border-right-width:medium; border-top-style:none; border-top-width:medium" bgcolor="#FFFFFF">
			&nbsp;</td>
          <td style="border-right:1px solid #384F5B; border-bottom-style: solid; border-bottom-width: 1px; border-left-style:none; border-left-width:medium; border-top-style:none; border-top-width:medium; width: 238px;" bgcolor="#FFFFFF" class="style4">
			<input name="Checkbox3" type="checkbox">Show prior year comparison</td>
        </tr>
        <tr>
          <td width="96" style="border-bottom:1px solid #384F5B; border-left-style: solid; border-left-width: 1px; border-right-style:none; border-right-width:medium; border-top-style:none; border-top-width:medium" bgcolor="#FFFFFF">
			&nbsp;</td>
          <td style="border-right:1px solid #384F5B; border-bottom-style: solid; border-bottom-width: 1px; border-left-style:none; border-left-width:medium; border-top-style:none; border-top-width:medium; width: 238px;" bgcolor="#FFFFFF" class="style4">
<font color="#080000">
			<button name="display" type="submit">Display</button>
    </font>
    		</td>
        </tr>
        <tr>
          <td width="96" style="border-bottom:1px solid #384F5B; border-left-style: solid; border-left-width: 1px; border-right-style:none; border-right-width:medium; border-top-style:none; border-top-width:medium" bgcolor="#FFFFFF">
			&nbsp;</td>
          <td style="border-right:1px solid #384F5B; border-bottom-style: solid; border-bottom-width: 1px; border-left-style:none; border-left-width:medium; border-top-style:none; border-top-width:medium; width: 238px;" bgcolor="#FFFFFF" class="style4">
			&nbsp;</td>
        </tr>
      </table>
      <p>&nbsp;</p>
      <input type="hidden" name="arv_dt" value="<%=Month(datNewDate) & "/1/" & Year(datNewDate)%>">
      <input type="hidden" name="display_request" value="true">
      <input type="hidden" name="rtrn_dt" value="<%=DateAdd("m", 1, Month(datNewDate) & "/1/" & Year(datNewDate))%>">
      <input type="hidden" name="new_date" value="<%=datNewDate %>">
    </form>
    <form action="rate_mgmt_car.asp" name="calendar" method="POST">
      <p></p>
      <p align="center">|&nbsp;<a href='rate_mgmt_car.asp?new_date=<%=Month(DateAdd("m", -1, datNewDate)) & "/1/" & Year(datNewDate)%>'>&lt;</a>&nbsp;|&nbsp;March 2007 (week 1) |
      <a href='rate_mgmt_car.asp?new_date=<%=Month(DateAdd("m", 1, datNewDate)) & "/1/" & Year(datNewDate) %>'>&gt;</a> |</p>
     
        
      <input type="hidden" name="current_date" value="<%=datPreviousDate %>">		
  
        
    </form>
<font size="-1">
		<br>	

		


		<table width="1000" border="0" cellpadding="2" background="images/alt_color.gif" style="border-collapse: collapse" bordercolor="#111111" id="desc<%=intTableCount %>">
		    <tr valign="bottom">
		      <td >&nbsp;</td>
		    </tr>
		</table>
		<table border="0" style="width: 800px;" bordercolor="#FFFFFF" height="4" id="headerbar<%=intTableCount %>" class="style6">
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
 
		<table bordercolor="#FFFFFF" id="rate_changes<%=intTableCount %>" cellspacing="0" cellpadding="0" id="rate_changes<%=intTableCount %>" style="width: 800px" class="style7">
    		<tr>
      			<th bgcolor="#879AA2">&nbsp;</th>
      			<th class="style5" style="width: 102px">D<font size="2">ate</font></th>
      			<th class="profile_header" style="width: 44px"><font size="2">Rsvd.</font></th>
      			<th class="style5" style="width: 36px">Cncld</th>
      			<th class="style5" style="width: 34px">N<font size="2">S</font></th>
      			<th class="style5" style="width: 42px">O<font size="2">n Rent</font></th>
      			<th class="style5">R<font size="2">trnd</font></th>
      			<th class="profile_header"><font size="2">Avg. Rate</font></th>
      			<th class="profile_header" height="45"><font size="2">Util. level</th>      
  
 			</tr>
		


	    
	
    
  		
    
	    <tr>
	    <td class="style9" width="29" align="center" >
<font size="-1" color="#080000">
		<img border="0" src="images/remaining-sm.gif" width="14" height="14" alt=""></font></td>
	    <td class="style9" style="width: 102px" >
	    Thu - 4/1/07</td>
	    <td class="style8" style="width: 44px" >
		45</td>
	    <td class="style9" style="width: 36px" >
		4</td>
	    <td class="style9" style="width: 34px">0</td>
	    <td class="style9" style="width: 42px">
	    134</td>
   
	    <td class="style9" width="45">
	    36</td>
   
	    <td class="style8" align="right" width="75">
	    $45.25</td>
   
	    <td class="style8" height="24" width="74">
	    78.0%</td>
   
	    </tr>


    
	    <tr>
	    <td class="style9" width="29" align="center" >
<font size="-1" color="#080000">
		<img border="0" src="images/success-sm.gif" width="14" height="14" alt=""></font></td>
	    <td class="style9" style="width: 102px" >
	    Fri - 4/2/07</td>
	    <td class="style8" style="width: 44px" >
		35</td>
	    <td class="style9" style="width: 36px" >
		5</td>
	    <td class="style9" style="width: 34px">1</td>
	    <td class="style9" style="width: 42px">
	    216</td>
   
	    <td class="style9" width="45">
	    67</td>
   
	    <td class="style8" align="right" width="75">
<font size="-1" color="#080000">
		$45.25</font></td>
   
	    <td class="style8" height="24" width="74">
	    8<font size="-1" color="#080000">8.4%</font></td>
   
	    </tr>


    
	    <tr>
	    <td class="style9" width="29" align="center" >
<font size="-1" color="#080000">
		<img border="0" src="images/remaining-sm.gif" width="14" height="14" alt=""></font></td>
	    <td class="style9" style="width: 102px" >
	    Sat - 4/3/07</td>
	    <td class="style8" style="width: 44px" >
		46</td>
	    <td class="style9" style="width: 36px" >
		6</td>
	    <td class="style9" style="width: 34px">3</td>
	    <td class="style9" style="width: 42px">
	    176</td>
   
	    <td class="style9" width="45">
	    34</td>
   
	    <td class="style8" align="right" width="75">
<font size="-1" color="#080000">
		$45.25</font></td>
   
	    <td class="style8" height="24" width="74">
	    6<font size="-1" color="#080000">8.2%</font></td>
   
	    </tr>


    
	    <tr>
	    <td class="style9" width="29" align="center" >
<font size="-1" color="#080000">
		<img border="0" src="images/remaining-sm.gif" width="14" height="14" alt=""></font></td>
	    <td class="style9" style="width: 102px" >
	    Sun - 4/4/07</td>
	    <td class="style8" style="width: 44px" >
		67</td>
	    <td class="style9" style="width: 36px" >
		6</td>
	    <td class="style9" style="width: 34px">2</td>
	    <td class="style9" style="width: 42px">
	    154</td>
   
	    <td class="style9" width="45">
	    24</td>
   
	    <td class="style8" align="right" width="75">
<font size="-1" color="#080000">
		$45.25</font></td>
   
	    <td class="style8" height="24" width="74">
	    65.4<font size="-1" color="#080000">%</font></td>
   
	    </tr>


    
	    <tr>
	    <td class="style9" width="29" align="center" >
<font size="-1" color="#080000">
		<img border="0" src="images/success-sm.gif" width="14" height="14" alt=""></font></td>
	    <td class="style9" style="width: 102px" >
	    Mon - 4/5/07</td>
	    <td class="style8" style="width: 44px" >
		45</td>
	    <td class="style9" style="width: 36px" >
		0</td>
	    <td class="style9" style="width: 34px">2</td>
	    <td class="style9" style="width: 42px">
	    137</td>
   
	    <td class="style9" width="45">
	    34</td>
   
	    <td class="style8" align="right" width="75">
<font size="-1" color="#080000">
		$45.25</font></td>
   
	    <td class="style8" height="24" width="74">
	    8<font size="-1" color="#080000">8.1%</font></td>
   
	    </tr>


    
	    <tr>
	    <td class="style9" width="29" align="center" >
<font size="-1" color="#080000">
		<img border="0" src="images/success-sm.gif" width="14" height="14" alt=""></font></td>
	    <td class="style9" style="width: 102px" >
	    Tue - 4/6/07</td>
	    <td class="style8" style="width: 44px" >
		33</td>
	    <td class="style9" style="width: 36px" >
		3</td>
	    <td class="style9" style="width: 34px">4</td>
	    <td class="style9" style="width: 42px">
	    289</td>
   
	    <td class="style9" width="45">
	    42</td>
   
	    <td class="style8" align="right" width="75">
<font size="-1" color="#080000">
		$45.25</font></td>
   
	    <td class="style8" height="24" width="74">
	    98.5%</td>
   
	    </tr>


    
	    <tr>
	    <td class="style9" width="29" align="center" >
<font size="-1" color="#080000">
		<img border="0" src="images/success-sm.gif" width="14" height="14" alt=""></font></td>
	    <td class="style9" style="width: 102px" >
	    Wed - 4/7/07</td>
	    <td class="style8" style="width: 44px" >
		56</td>
	    <td class="style9" style="width: 36px" >
		4</td>
	    <td class="style9" style="width: 34px">4</td>
	    <td class="style9" style="width: 42px">
	    157</td>
   
	    <td class="style9" width="45">
	    12</td>
   
	    <td class="style8" align="right" width="75">
<font size="-1" color="#080000">
		$45.25</font></td>
   
	    <td class="style8" height="24" width="74">
<font size="-1" color="#080000">
		98.3%</font></td>
   
	    </tr>


  </table>
 
  <table border="0" style="width: 800px;" bordercolor="#FFFFFF" height="4" id="footerbar<%=intTableCount %>" class="style6">
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
    

	


	


<font size="2">

      &nbsp;</font><font size="3" color="#000000"> </font>
</font>
    <table width="745" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td><img src="images/ruler.gif" width="745" height="2"></td>
      </tr>
    </table>
    <p>&nbsp;</p>
    <p class="style4">
<font color="#080000">
	<a href="rate_mgmt_car.asp">[Monthly Rental Activity calendar view]</a></font></p>
    </font>
    </td>
  </tr>
  </table>
<p>&nbsp;</p>
<p>
&nbsp;</p>
</td></tr></table>
&nbsp;<!-- JUSTTABS BOTTOM CLOSE -->
<!--#INCLUDE FILE="footer.asp"-->
</body>
</html>
<%
Rem Clean-up

 Set adoCmd = Nothing
 Set adoRS = Nothing
 Set adoRS1 = Nothing
 Set adoRS2 = Nothing
 Set adoRS2a = Nothing


%>