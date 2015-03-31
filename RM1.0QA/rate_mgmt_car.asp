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

	'On Error Resume Next
	
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



	datPreviousDate = Request.Form("previous_date")

	If IsDate(Request.QueryString("new_date")) Then
		datNewDate = Request.QueryString("new_date")
	Else
		datNewDate = now
	End If

	If Request("display_request") = "true" Then
		Rem Get the rates
		Set adoCmd = CreateObject("ADODB.Command")

		adoCmd.ActiveConnection =  strConn
		adoCmd.CommandText = "car_rental_transaction_summary_select"
		adoCmd.CommandType = 4
			
'		adoCmd.Parameters.Append adoCmd.CreateParameter("@city_cd",      200, 1, 6, strCityCd)
'		adoCmd.Parameters.Append adoCmd.CreateParameter("@data_source",  200, 1, 3, strDataSource)
'		adoCmd.Parameters.Append adoCmd.CreateParameter("@lor",            3, 1, 0, strLOR)
'		adoCmd.Parameters.Append adoCmd.CreateParameter("@car_type_cd",  200, 1, 4, strCarType)
'		adoCmd.Parameters.Append adoCmd.CreateParameter("@user_id",        3, 1, 0, strUserId)
'		adoCmd.Parameters.Append adoCmd.CreateParameter("@arv_dt",       135, 1, 0, Month(datNewDate) & "/1/" & Year(datNewDate))
'		adoCmd.Parameters.Append adoCmd.CreateParameter("@rtrn_dt",      135, 1, 0, DateAdd("m", 1, Month(datNewDate) & "/1/" & Year(datNewDate)))
'		adoCmd.Parameters.Append adoCmd.CreateParameter("@vend_cd",      200, 1, 2, strVendor)
'		adoCmd.Parameters.Append adoCmd.CreateParameter("@comp_vend_cd", 200, 1, 2, Null)

		adoCmd.Parameters.Append adoCmd.CreateParameter("@city_cd",      200, 1, 6, "LAS")
		adoCmd.Parameters.Append adoCmd.CreateParameter("@car_type_cd",  200, 1, 4, "CCAR")
		adoCmd.Parameters.Append adoCmd.CreateParameter("@lor",            3, 1, 0, 2)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@user_id",        3, 1, 0, 15)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@start_dt",     135, 1, 0, "3/1/2007")
		adoCmd.Parameters.Append adoCmd.CreateParameter("@end_dt",       135, 1, 0, "3/30/2007")
		
		If Err.Number <> 0 Then
			pad="&nbsp;&nbsp;&nbsp;&nbsp;"
			response.write "<b>Error Occured - preparing to select rates<br>"
	   		response.write "</b><br>"
	   		response.write pad & "Error Number   = <b>" & Err.Number & "</b><br>"
	   		response.write pad & "Error Desc.    = <b>" & err.description & "</b><br>"
	   		response.write pad & "Help Context   = <b>" & err.HelpContext & "</b><br>"
	   		response.write pad & "Help File Path = <b>" & err.helpfile & "</b><br>"
	   		response.write pad & "Error Source   = <b>" & err.source & "</b><br><hr>"

		End If
	
		Set adoRS6 = adoCmd.Execute
		
		'http://www.rate-monitor.com/rate_mgmt_car.asp?city_cd=LAS&car_type_cd=CCAR&lor=2&display=Display&arv_dt=3%2F1%2F2007&display_request=true&rtrn_dt=4%2F1%2F2007&new_date=3%2F13%2F2007+9%3A40%3A21+PM
		
		If Err.Number <> 0 Then
			pad="&nbsp;&nbsp;&nbsp;&nbsp;"
			response.write "<b>Error Occured - selecting rates<br>"
	   		response.write "</b><br>"
	   		response.write pad & "Error Number   = <b>" & Err.Number & "</b><br>"
	   		response.write pad & "Error Desc.    = <b>" & err.description & "</b><br>"
	   		response.write pad & "Help Context   = <b>" & err.HelpContext & "</b><br>"
	   		response.write pad & "Help File Path = <b>" & err.helpfile & "</b><br>"
	   		response.write pad & "Error Source   = <b>" & err.source & "</b><br><hr>"

		End If
		
	Else

		Rem Get the rates
		Set adoCmd = CreateObject("ADODB.Command")

		adoCmd.ActiveConnection =  strConn
		adoCmd.CommandText = "car_shopped_rate_calendar"
		adoCmd.CommandType = 4
			
		adoCmd.Parameters.Append adoCmd.CreateParameter("@city_cd", 200, 1, 5, "")
		adoCmd.Parameters.Append adoCmd.CreateParameter("@data_source", 200, 1, 3, "")
		adoCmd.Parameters.Append adoCmd.CreateParameter("@lor", 3, 1, 0, 2)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@car_type_cd", 200, 1, 4, "")
		adoCmd.Parameters.Append adoCmd.CreateParameter("@org_id", 3, 1, 0, 5)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@arv_dt", 135, 1, 0, Null)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@rtrn_dt", 135, 1, 0, Null)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@vend_cd", 200, 1, 2, Null)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@comp_vend_cd", 200, 1, 2, Null)

		Set adoRS6 = adoCmd.Execute

		If Err.Number <> 0 Then
			pad="&nbsp;&nbsp;&nbsp;&nbsp;"
			response.write "<b>Error Occured - selecting default rates<br>"
	   		response.write "</b><br>"
	   		response.write pad & "Error Number= #<b>" & Err.Number & "</b><br>"
	   		response.write pad & "Error Desc.= <b>" & err.description & "</b><br>"
	   		response.write pad & "Help Context= <b>" & err.HelpContext & "</b><br>"
	   		response.write pad & "Help File Path=<b>" & err.helpfile & "</b><br>"
	   		response.write pad & "Error Source= <b>" & err.source & "</b><br><hr>"

		End If


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
.style2 {
	background-color: #879AA2;
	font-size: x-small;
}
.style3 {
	font-size: 11pt;
}
.style4 {
	font-size: x-small;
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
<td bgcolor="#000000" colspan="1" height="1"><img src=pixel.gif width="1" height="1"></td>
<td bgcolor="#000000" colspan="1" height="1"><img src=pixel.gif width="1" height="1"></td>
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
          &nbsp;<span class="style3">Monthly Rental Activity (calendar view) </span></font></td>
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
          &nbsp;</td>
        </tr>
        <tr>
          <td width="96" style="border-left-style: solid; border-left-width: 1px; border-right-style:none; border-right-width:medium; border-top-style:none; border-top-width:medium; border-bottom-style:none; border-bottom-width:medium" bgcolor="#FFFFFF">
			&nbsp;</td>
          <td style="border-right:1px solid #384F5B; border-left-style:none; border-left-width:medium; border-top-style:none; border-top-width:medium; border-bottom-style:none; border-bottom-width:medium; width: 238px;" bgcolor="#FFFFFF">
          &nbsp;<button name="display" type="submit">Display</button></td>
        </tr>
        <tr>
          <td width="96" style="border-bottom:1px solid #384F5B; border-left-style: solid; border-left-width: 1px; border-right-style:none; border-right-width:medium; border-top-style:none; border-top-width:medium" bgcolor="#FFFFFF">&nbsp;</td>
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
      <p align="center">|&nbsp;<a href="rate_mgmt_car.asp?new_date=<%=Month(DateAdd("m", -1, datNewDate)) & "/1/" & Year(datNewDate)%>">&lt;</a>&nbsp;|&nbsp;<%=MonthName(Month(datNewDate)) & " " & Year(datNewDate)%> |
      <a href="rate_mgmt_car.asp?new_date=<%=Month(DateAdd("m", 1, datNewDate)) & "/1/" & Year(datNewDate) %>">&gt;</a> |</p>
      <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="100%" id="table1">
          <tr>
          <td bgcolor="#384F5B" align="center" width="14%"><font color="#FFFFFF">
          Sunday</font></td>
          <td bgcolor="#384F5B" align="center" width="14%"><font color="#FFFFFF">
          Monday</font></td>
          <td bgcolor="#384F5B" align="center" width="14%"><font color="#FFFFFF">
          Tuesday</font></td>
          <td bgcolor="#384F5B" align="center" width="14%"><font color="#FFFFFF">
          Wednesday</font></td>
          <td bgcolor="#384F5B" align="center" width="14%"><font color="#FFFFFF">
          Thursday</font></td>
          <td bgcolor="#384F5B" align="center" width="14%"><font color="#FFFFFF">
          Friday</font></td>
          <td bgcolor="#384F5B" align="center" width="14%"><font color="#FFFFFF">
          Saturday</font></td>
          </tr>
        <% Dim curRate  
           Dim curTotalRate
           Dim curTotalPrice
           Dim Rates(31,6)
           
           Const arv_dt = 0
           Const rt_amt = 1
           Const total_rt_amt = 2
           Const est_rental_chrg_amt = 3
           Const shop_dttm = 4
           Const shop_request_id = 5


   		   Do While (adoRS6.EOF = False)  
             Rates(Day(adoRS6.Fields("arv_dt").Value), arv_dt)              = adoRS6.Fields("arv_dt").Value
             Rates(Day(adoRS6.Fields("arv_dt").Value), rt_amt)              = adoRS6.Fields("rt_amt").Value
             Rates(Day(adoRS6.Fields("arv_dt").Value), total_rt_amt)        = adoRS6.Fields("total_rt_amt").Value
             Rates(Day(adoRS6.Fields("arv_dt").Value), est_rental_chrg_amt) = adoRS6.Fields("est_rental_chrg_amt").Value
             Rates(Day(adoRS6.Fields("arv_dt").Value), shop_dttm)           = adoRS6.Fields("shop_dttm").Value
             Rates(Day(adoRS6.Fields("arv_dt").Value), shop_request_id)     = adoRS6.Fields("shop_request_id").Value

             'If IsNumeric(adoRS6.Fields("rt_amt").Value) Then
             '   Rates(Day(adoRS6.Fields("arv_dt").Value), rate_amt) = FormatNumber(adoRS6.Fields("rt_amt").Value,0	)
	 	     'Else
             '   Rates(Day(adoRS6.Fields("arv_dt").Value), rate_amt) = "N/A"
		     'End If
 		     adoRS6.MoveNext       
		
		   Loop
		


           
		   datFirstDate = Month(datNewDate) & "/1/" & Year(datNewDate) 
  		   datIndex = DateAdd("d", -1, datFirstDate ) 

		%>		
		<% For intRowIndex = 1 To 5  %>
        <tr>
  
 		<%   For intColIndex = 1 To 7    %>  

        <%       If (intRowIndex = 1) And (intColIndex < Weekday(datFirstDate, 1)) Then %> 	
			          <td class="off_day" >&nbsp;&nbsp;</td>

        <%       ElseIf DateDiff("m", datFirstDate, DateAdd("d", 1, datIndex)) > 0 Then %> 	
			          <td class="off_day" >&nbsp;</td>

		<%		 ElseIf (DateAdd("d", 1, datIndex) < Now) Then %>
	 	<%         datIndex = DateAdd("d", 1, datIndex) %>
			          <td>
				      <table border="0" cellpadding="0" style="border-collapse: collapse" width="100%" id="table6">
			            <tr>
			              <td bgcolor="#CFD7DB" >&nbsp;</td>
			              <td align="right" bgcolor="#CFD7DB">
							<a title="<%="Data feed date:: " & FormatDateTime(Rates(Day(datIndex), shop_dttm), 2) & " " & FormatDateTime(Rates(Day(datIndex), shop_dttm), 4) %>" href="rate_mgmt_car.asp"><%=Day(datIndex) %></a>&nbsp;&nbsp; </td>
			            </tr>
			            <tr>
			              <td align="right" class="style2" >Rsvd:</td>
						  <td align="right" class="off_day" ><font size="2">
			              <% If IsNumeric(Rates(Day(datIndex), rt_amt)) Then %>
			              	<%=FormatNumber(Rates(Day(datIndex), rt_amt),0)%>
			              <% Else %>
			              	<%="N/A" %>
			              <% End If %></font></td>
				        </tr>
			            <tr>
			              <td align="right" class="off_day" ><font size="2">Cncl:</font></td>
			              <td align="right" class="off_day" ><font size="2">
			              <% If IsNumeric(Rates(Day(datIndex), total_rt_amt)) Then %>
			              	<%=FormatNumber(Rates(Day(datIndex), total_rt_amt),0)%>
			              <% Else %>
			              	<%="N/A" %>
			              <% End If %></font></td>
			            </tr>
			            <tr>
			              <td align="right" class="off_day" ><font size="2">N-S:</font></td>
			              <td align="right" class="off_day" ><font size="2">
			              <% If IsNumeric(Rates(Day(datIndex), total_rt_amt)) Then %>
			              	<%=FormatNumber(Rates(Day(datIndex), total_rt_amt),0)%>
			              <% Else %>
			              	<%="N/A" %>
			              <% End If %></font></td>
			            </tr>

			            <tr>
			              <td align="right" class="off_day" ><font size="2">Rtrn:</font></td>
			              <td align="right" class="off_day" ><font size="2">
			              &nbsp;&nbsp;&nbsp; </font></td>
			            </tr>
			          </table>
			          </td>
			          

			          
		<%		 Else %>
	 	<%         datIndex = DateAdd("d", 1, datIndex) %>
			          <td>
				      <table border="0" cellpadding="0" style="border-collapse: collapse" width="100%" id="table6">
			            <tr>
			              <td bgcolor="#CFD7DB">&nbsp;</td>
			              <td align="right" bgcolor="#CFD7DB"><a title="<%="Data feed date:: " & FormatDateTime(Rates(Day(datIndex), shop_dttm), 2) & " " & FormatDateTime(Rates(Day(datIndex), shop_dttm), 4) %>" href="rate_mgmt_car.asp"><%=Day(datIndex) %></a>&nbsp;&nbsp; </td>
			            </tr>
			            <tr>
			              <td align="right" bgcolor="#FFFFFF"><font size="2">
							<a title="click to display the reservation detail for this day" >Rsvd:</a></font></td>
			              <td align="right" bgcolor="#FFFFFF"><font size="2">
			              <% If IsNumeric(Rates(Day(datIndex), rt_amt)) Then %>
			              	<%=FormatNumber(Rates(Day(datIndex), rt_amt),0)%>
			              <% Else %>
			              	<%="N/A" %>
			              <% End If %></font></td>
				          
				        </tr>
			            <tr>
			              <td align="right" bgcolor="#FFFFFF"><font size="2">
							<a title="click to display the reservation cancelation detail for this day" >Cncl:</a></font></td>
			              <td align="right" bgcolor="#FFFFFF"><font size="2">
			              <% If IsNumeric(Rates(Day(datIndex), total_rt_amt)) Then %>
			              	<%=FormatNumber(Rates(Day(datIndex), total_rt_amt),0)%>
			              <% Else %>
			              	<%="N/A" %>
			              <% End If %></font></td>
			            </tr>
			            <tr>
			              <td align="right" bgcolor="#FFFFFF"><font size="2">
							<a title="click to display the open rental detail for this day" >Open:</a></font></td>
			              <td align="right" bgcolor="#FFFFFF"><font size="2">
			              <% If IsNumeric(Rates(Day(datIndex), est_rental_chrg_amt)) Then %>
			              	<%=FormatNumber(Rates(Day(datIndex), est_rental_chrg_amt),0)%>
			              <% Else %>
			              	<%="N/A" %>
			              <% End If %></font></td>
			            </tr>
			            <tr>
			              <td align="right" bgcolor="#FFFFFF"><font size="2"><a title="click to display the return/invoice detail for this day" >Rtrn:</a></font></td>
			              <td align="right" bgcolor="#FFFFFF"><font size="2">
			              <% If IsNumeric(Rates(Day(datIndex), est_rental_chrg_amt)) Then %>
			              	<%=FormatNumber(Rates(Day(datIndex), est_rental_chrg_amt),0)%>
			              <% Else %>
			              	<%="N/A" %>
			              <% End If %></font></td>
			            </tr>

			          </table>
			          </td>
			          
		<%      End If %>
		<%   Next      %>	

		</tr>
		<% Next        %>			
		
</table>		
  
        
      <input type="hidden" name="current_date" value="<%=datPreviousDate %>">		
  
        
    </form>
    <p>&nbsp;&nbsp;</p>
    <table width="745" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td><img src="images/ruler.gif" width="745" height="2"></td>
      </tr>
    </table>
    <p>&nbsp;</p>
    <p class="style4">
<font color="#080000">
	<a href="rate_mgmt_car_listview.asp">[Monthly Rental Activity list view]</a></font></p>
    </font>
    </td>
  </tr>
  </table>
<p>&nbsp;</p>
<p>
&nbsp;</p>
</td></tr></table>
<!--#INCLUDE FILE="footer.asp"-->
&nbsp;<!-- JUSTTABS BOTTOM CLOSE -->
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