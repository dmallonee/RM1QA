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
	Dim adoPrices
	Dim intIndex
	Dim blnExit
	Dim intReportRequestID
	Dim intErrorCount
	Dim intTimeoutCount
	Dim strCarType 
	Dim intResults
	Dim intPrice
	Rem we have no clue how many, so cross your fingers
	Dim varCarTypes()
	Dim varDataSources()
	Dim varDates()

	
		strConn = Session("pro_con")
	
  		Set adoCmd = CreateObject("ADODB.Command")

		adoCmd.ActiveConnection =  strConn
		adoCmd.CommandText = "car_shop_request_select_payless"
		adoCmd.CommandType = 4

		Set adoRS = adoCmd.Execute
	

	'If adoRS Is Nothing Then
	'	Set adoRS = CreateObject("ADODB.Recordset")
	'End If

	%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<title>Rate-Monitor by Rate-Highway, Inc. | Search Queue</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta http-equiv="refresh" content="120;url=search_queue_car_payless.asp">
<link rel="stylesheet" type="text/css" href="inc/rh_standard.css">
<link rel="stylesheet" type="text/css" href="inc/rh_report.css">
<link rel="stylesheet" type="text/css" href="inc/sitewide.css">
<script src="inc/sitewide.js" language="javascript"></script>
<script src="inc/header_menu_support.js" language="JavaScript" type="text/JavaScript"></script>
</head>

<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onload="MM_preloadImages('images/b_search_pro_on.gif','images/b_search_que_on.gif','images/b_search_cri_on.gif','images/b_rate_on.gif','images/b_alert_on.gif','images/b_user_on.gif','images/b_system_on.gif')">

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
        <img src="images/b_search_pro_of.gif" name="s1" border="0" id="s1" width="183" height="32"></a></td>
        <td>
        <a href="search_queue_car_payless.asp" onmouseover="MM_swapImage('s2','','images/b_search_que_on.gif',1)" onmouseout="MM_swapImgRestore()">
        <img src="images/b_search_que_of.gif" name="s2" border="0" id="s2" width="97" height="32"></a></td>
        <td>
        <a href="search_criteria_car.asp" onmouseover="MM_swapImage('s3','','images/b_search_cri_on.gif',1)" onmouseout="MM_swapImgRestore()">
        <img src="images/b_search_cri_of.gif" name="s3" border="0" id="s3" width="103" height="32"></a></td>
        <td>
        <% If Request.Cookies("rate-monitor.com")("rate_reports") = "True" Then %>
        <a href="rate_calendar_car.asp" onmouseover="MM_swapImage('ra','','images/b_rate_on.gif',1)" onmouseout="MM_swapImgRestore()">
        <% Else %>
        <a href="javascript:not_enabled()" onmouseover="MM_swapImage('ra','','images/b_rate_on.gif',1)" onmouseout="MM_swapImgRestore()">
        <% End If %>
        <img src="images/b_rate_of.gif" name="ra" border="0" id="ra" width="88" height="32"></a></td>
        <td>
        <a href="javascript:not_enabled()" onmouseover="MM_swapImage('al','','images/b_alert_on.gif',1)" onmouseout="MM_swapImgRestore()">
        <img src="images/b_alert_of.gif" name="al" border="0" id="al" width="53" height="32"></a></td>
        <td>
        <a href="javascript:not_enabled()" onmouseover="MM_swapImage('us','','images/b_user_on.gif',1)" onmouseout="MM_swapImgRestore()">
        <img src="images/b_user_of.gif" name="us" border="0" id="us" width="126" height="32"></a></td>
        <td>
        <a href="javascript:not_enabled()" onmouseover="MM_swapImage('sy','','images/b_system_on.gif',1)" onmouseout="MM_swapImgRestore()">
        <img src="images/b_system_of.gif" name="sy" border="0" id="sy" width="58" height="32"></a></td>
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
                <td>
                <div align="right">
                  <a href="default.asp"><font face="Arial, Helvetica, sans-serif" size="2" color="#FFFFFF">User: <%=Request.Cookies("rate-monitor.com")("user_name")%></font></div>
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
        <td><img src="images/h_search_que.gif" width="368" height="31"></td>
        <td><img src="images/h_right.gif" width="402" height="31"></td>
      </tr>
    </table>
    </td>
  </tr>
</table>
&nbsp;

  Payless Searches Displayed Below Only.<table border="0" style="border-collapse: collapse" bordercolor="#FFFFFF" width="2110" height="4">
    <tr>
      <td background="images/ruler.gif">
		<p align="left">&nbsp;</td>
    </tr>
  </table>

<table width="1110" border="0" cellpadding="2" background="images/alt_color.gif" style="border-collapse: collapse" bordercolor="#111111">
  <tr valign="bottom">
    <td width="169"><font size="2">&nbsp;|
	<a href="javascript:document.queue.submit();">Send</a> | </font></td>
  </tr>
</table>
<table border="0" style="border-collapse: collapse" bordercolor="#FFFFFF" width="2110" height="4">
  <tr>
    <td background="images/ruler.gif"></td>
  </tr>
</table>
<form name="queue" method="POST" action="send_car_notification_payless.asp" >
<table border="1" style="border-collapse: collapse" bordercolor="#FFFFFF" id="profiles" width="2110">
<thead >
  <tr>
    <td align="left" valign="bottom" bgcolor="#879AA2" height="45" width="26">&nbsp;</td>
    <td class="profile_header" width="63" style="background-color: #E07D1A" height="45">Selected</td>
    <td class="profile_header" width="45" height="45">Search ID</td>
    <td class="profile_header" width="50" height="45">Search Status</td>
    <td class="profile_header" width="75" height="45">Request</td>
    <td class="profile_header" width="75" height="45">User</td>
    <td class="profile_header" width="300" height="45">Profile</td>
    <td class="profile_header" width="60" height="45">Action</td>
    <td class="profile_header" width="76" height="45">Search Units</td>
    <td class="profile_header" width="73" height="45">Rate Units Expected</td>
    <td class="profile_header" width="79" height="45">Rate Units Complete</td>
    <td class="profile_header" width="70" height="45">Pickup City</td>
    <td class="profile_header" width="72" height="45">First Rental Date</td>
    <td class="profile_header" width="82" height="45">Last Rental Date</td>
    <td class="profile_header" width="82" height="45">Car Types</td>
    <td class="profile_header" width="82" height="45">Companies</td>
  </tr>
  </thead> 
  <%
        
        Dim strClass
        Dim strOrange
        Dim intCount

		

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
    <td class="<%=strClass %>" height="20"><%=intCount  %></td>
    <td   bgcolor="#FDC677" align="center" height="20">
    <input type="radio" value="<%=adoRS.Fields("shop_request_id").Value %>" name="search_id"></td>
    <td class="<%=strClass %>" height="20" width="45">
    <a href="view_report_car2.asp?ReportRequestID=<%=adoRS.Fields("shop_request_id").Value %>" target="_blank">
    <%=adoRS.Fields("shop_request_id").Value %></a></td>
    <td class="<%=strClass %>" height="20" width="50">
    <% Select Case adoRS.Fields("request_status").Value %>
	<%     Case "R" %>
	Running
	<%     Case "C" %>
	Cancelled
	<%     Case "N" %>
	New
	<%     Case "S" %>
	Successful
	<%     Case "F" %>
	Failure
	<%     Case Else %>
	<%=adoRS.Fields("request_status").Value %>
	<% End Select %></td>
    <td class="<%=strClass %>" height="20" width="75">
    <% If DateDiff("d", Now, adoRS.Fields("scheduled_dttm").Value) = 0 Then %>
      <%=FormatDateTime(adoRS.Fields("scheduled_dttm").Value, 4) %>
      <!--
      <%=DatePart("h", adoRS.Fields("scheduled_dttm").Value) & ":" & DatePart("n", adoRS.Fields("scheduled_dttm").Value) & ":" & DatePart("s", adoRS.Fields("scheduled_dttm").Value) %>
      -->
    <% Else %>
      <%=FormatDateTime(adoRS.Fields("scheduled_dttm").Value, 2) %>
    <% End If %>
    </td>
    <td class="<%=strClass %>" height="20" width="75"><%=adoRS.Fields("client_userid").Value %></td>
    <td class="<%=strClass %>" height="20" width="300">
	<% If adoRS.Fields("profile_desc").Value = "" Then %>
	 <i>[none]</i> On-demand report request
	<% Else %>
	  <%=Left(adoRS.Fields("profile_desc").Value, 40) %>
	<% End If %>
	</td>
    <td class="<%=strClass %>" height="20">Display</td>
    <td class="<%=strClass %>" height="20"><%=adoRS.Fields("work_units").Value %></td>
    <td class="<%=strClass %>" height="20"><%=adoRS.Fields("work_units").Value %></td>
    <td class="<%=strClass %>" height="20"><%=adoRS.Fields("work_units_complete").Value %></td>
    <td class="<%=strClass %>" height="20"><%=adoRS.Fields("city_cd").Value %></td>
    <td class="<%=strClass %>" height="20"><%=adoRS.Fields("begin_arv_dt").Value %></td>
    <td class="<%=strClass %>" height="20"><%=adoRS.Fields("end_arv_dt").Value %></td>
    <td class="<%=strClass %>" height="20"><%=adoRS.Fields("shop_car_type_cds").Value %></td>
    <% If adoRS.Fields("vend_cd").Value = "" Then %>
    <td class="<%=strClass %>" height="20">All</td>
    <% Else %>
    <td class="<%=strClass %>" height="20"><%=adoRS.Fields("vend_cd").Value %></td>
    <% End If %>
  </tr>
  <%
        
        	adoRS.MoveNext
        	
        Wend
        
   		adoRS.Close
		Set adoRS1 = Nothing
		Set adoCmd = Nothing

	

		%>
		
</table>
<table border="0" style="border-collapse: collapse" bordercolor="#FFFFFF" width="2110" height="4">
  <tr>
      <td background="images/ruler.gif"></td>
  </tr>
</table>
<p>&nbsp;| <a href="javascript:document.queue.submit();">Send</a> 
| </p>
</form>
<!--#INCLUDE FILE="footer.asp"-->
</body>
</html>
