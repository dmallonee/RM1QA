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
	Dim intRefresh

	strClientUserid = Request.Form("userid")
	strCity = Request.Form("city")
	strCarType = Request.Form("car_type")
	strCompany = Request.Form("company")
	
	strSearched = False
	
	If Request("refresh") = "" Then
		intRefresh = 300
	Else
		intRefresh = Request("refresh")
		If IsNumeric(intRefresh) = False Then
			intRefresh = 300
		End If
	End If
	
	If (strClientUserid = "") And (strCity = "") And (strCarType  = "") And (strCompany = "") Then
	
		strConn = Session("pro_con")
	
  		Set adoCmd = CreateObject("ADODB.Command")

		adoCmd.ActiveConnection =  strConn
		adoCmd.CommandText = "car_shop_request_select1"
		adoCmd.CommandType = 4

		adoCmd.Parameters.Append adoCmd.CreateParameter("@client_userid",      200, 1, 20)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@city_cd",            200, 1, 5)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@shop_car_type_cds",  200, 1, 255)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@vendor_cd",          200, 1, 255)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@user_id",              3, 1, 0)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@days_to_include",      3, 1, 0)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@linked_to_send_dttm", 11, 1, 0)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@day_back",             3, 1, 0)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@user_role",            3, 1, 0)
		
		adoCmd.Parameters("@user_id").Value = Request.Cookies("rate-monitor.com")("user_id")
		adoCmd.Parameters("@client_userid").Value = Null
		adoCmd.Parameters("@city_cd").Value = Null
		adoCmd.Parameters("@shop_car_type_cds").Value = Null
		adoCmd.Parameters("@vendor_cd").Value = Null 
		adoCmd.Parameters("@days_to_include") = 0
		adoCmd.Parameters("@linked_to_send_dttm") = Null
		adoCmd.Parameters("@day_back") = 0 
		adoCmd.Parameters("@user_role").Value = Request.Cookies("rate-monitor.com")("user_role")



				
		Set adoRS = adoCmd.Execute
	
		strSearched = True
 
	Else
	
		strConn = Session("pro_con")
	
  		Set adoCmd = CreateObject("ADODB.Command")

		adoCmd.ActiveConnection =  strConn
		adoCmd.CommandText = "car_shop_request_select1"
		adoCmd.CommandType = 4

		adoCmd.Parameters.Refresh 

		adoCmd.Parameters("@user_id").Value = Request.Cookies("rate-monitor.com")("user_id")

		If Trim(strClientUserid) <> "" Then
			adoCmd.Parameters("@client_userid").Value = strClientUserid 
		Else
			adoCmd.Parameters("@client_userid").Value = Null
		End If


		If Trim(strCity) <> "" Then
			adoCmd.Parameters("@city_cd").Value = strCity 
		Else
			adoCmd.Parameters("@city_cd").Value = Null
		End If


		If Trim(strCarType) <> "" Then
			adoCmd.Parameters("@shop_car_type_cds").Value = strCarType 
		Else
			adoCmd.Parameters("@shop_car_type_cds").Value = Null
		End If
	

		If Trim(strCompany) <> "" Then
			adoCmd.Parameters("@vendor_cd").Value = strCompany 
		Else
			adoCmd.Parameters("@vendor_cd").Value = Null 
		End If

		Set adoRS = adoCmd.Execute
		'Set adoRS1 = adoCmd.Execute
	
		strSearched = True
		

	
	
	End If

	If strClientUserId = "" Then
		strClientUserId = Request.Cookies("rate-monitor.com")("client_userid")
	End If


	'If adoRS Is Nothing Then
	'	Set adoRS = CreateObject("ADODB.Recordset")
	'End If

	%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<title>Rate-Monitor by Rate-Highway, Inc. | Today's completed reports</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta http-equiv="refresh" content="<%=intRefresh %>;url=search_queue_car_today.asp">
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
<form method="POST" action="search_queue_car_today.asp" name="search" class="search">
  <table border="0" style="border-collapse: collapse" bordercolor="#FFFFFF" width="1110" height="4">
    <tr>
      <td background="images/ruler.gif"></td>
     </tr>
  </table>
  <table border="0" style="border-collapse: collapse" bordercolor="#FFFFFF" width="1110" height="4">
    <tr>
      <td background="images/ruler.gif"></td>
    </tr>
  </table>
</form>
<table width="1110" border="0" cellpadding="2" background="images/alt_color.gif" style="border-collapse: collapse" bordercolor="#111111">
  <tr valign="bottom">
    <td width="169">&nbsp;<font size="2" face="Verdana, Arial, Helvetica, sans-serif"><a href="search_queue_car_today.asp">|&lt;</a>
    <a href="search_queue_car_today.asp">&lt;</a> Page 1 of 1
    <a href="search_queue_car_today.asp">&gt;</a> <a href="search_queue_car_today.asp">&gt;|</a></font></td>
  </tr>
</table>
<table border="0" style="border-collapse: collapse" bordercolor="#FFFFFF" width="1110" height="4">
  <tr>
    <td background="images/ruler.gif"></td>
  </tr>
</table>
<form name="queue" method="POST" action="cancel_search_car.asp" >
<table border="1" style="border-collapse: collapse" bordercolor="#FFFFFF" id="profiles" width="1110">
<thead >
  <tr>
    <td align="left" valign="bottom" bgcolor="#879AA2" height="45" width="43">&nbsp;</td>
    <td class="profile_header" width="82" height="45">Report ID</td>
    <td class="profile_header" width="975" height="45">Report or Profile 
    Description <%Session("testing")%> </td>
  </tr>
  </thead> 
  <%
        
        Dim strClass
        Dim strOrange
        Dim intCount
        Dim intRateReport
        
        intRateReport = True

		If strSearched = True Then

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

	<% If intRateReport = True Then %>
  <tr>
    <td class="<%=strClass %>" height="20"><%=intCount  %></td>
    <td class="<%=strClass %>" height="20" width="82">
    <a href="car_report_by_type.asp?ReportRequestID=<%=adoRS.Fields("shop_request_id").Value %>&security_code=<%=Escape(adoRS.Fields("security_code").Value) %>" target="_blank">
    <%=adoRS.Fields("shop_request_id").Value %></a></td>
    <td class="<%=strClass %>" height="20" width="975">
	<% If adoRS.Fields("profile_desc").Value = "" Then %>
	 <i>[none]</i> On-demand report request
	<% Else %>
	  <%=Left(adoRS.Fields("profile_desc").Value, 40) %>
	<% End If %>
	</td>
    <% If adoRS.Fields("vend_cd").Value = "" Then %>
    <% Else %>
    <% End If %>
  </tr>

  <% Else %>
  <tr>
    <td class="<%=strClass %>" height="20"><%=intCount  %></td>
    <td class="<%=strClass %>" height="20" width="82">
    <a alt="<%=adoRS.Fields("profile_desc").Value %>" href="rate_change_report.asp?reportrequestid=<%=adoRS.Fields("shop_request_id").Value %>&security_code=<%=Escape(adoRS.Fields("security_code").Value) %>" target="_blank">
    <%=adoRS.Fields("shop_request_id").Value %></a></td>
    <td class="<%=strClass %>" height="20" width="975">
	<% If adoRS.Fields("profile_desc").Value = "" Then %>
	 <i>[none]</i> On-demand report request
	<% Else %>
	  <%="Alerts! for " & Left(adoRS.Fields("profile_desc").Value, 40) %>
	<% End If %>
	</td>
    <% If adoRS.Fields("vend_cd").Value = "" Then %>
    <% Else %>
    <% End If %>
  </tr>

<% intCount = intCount + 1 %>

  <tr>
    <td class="<%=strClass %>" height="20"><%=intCount  %></td>
    <td class="<%=strClass %>" height="20" width="82">
    <a alt="<%=adoRS.Fields("profile_desc").Value %>" href="rate_change_report_override.asp?reportrequestid=<%=adoRS.Fields("shop_request_id").Value %>&security_code=<%=Escape(adoRS.Fields("security_code").Value) %>" target="_blank">
    <%=adoRS.Fields("shop_request_id").Value %></a></td>
    <td class="<%=strClass %>" height="20" width="975">
	<% If adoRS.Fields("profile_desc").Value = "" Then %>
	 <i>[none]</i> On-demand report request
	<% Else %>
	  <%="Overrideable Alerts! for " & Left(adoRS.Fields("profile_desc").Value, 40) %>
	<% End If %>
	</td>
    <% If adoRS.Fields("vend_cd").Value = "" Then %>
    <% Else %>
    <% End If %>
  </tr>

	<% adoRS.MoveNext %>

  <% End If %>

  <%
        
   		intRateReport = Not intRateReport 
        	
        Wend
        
   		adoRS.Close
		Set adoRS1 = Nothing
		Set adoCmd = Nothing

		Else

		%>
		
		
  <tr>
    <td class="profile_light" height="20"></td>
    <td class="profile_light" height="20" width="82"></td>
    <td class="profile_light" height="20" width="975"></td>
  </tr>
<!--
  <tr>
    <td width="26" class="profile_light" height="20"></td>
    <td width="63" bgcolor="#FDC677" align="center" height="20">
    <input type="radio" value="V1" name="selected"></td>
    <td width="26" class="profile_light" height="20"></td>
    <td width="26" class="profile_light" height="20"></td>
    <td width="26" class="profile_light" height="20"></td>
    <td width="54" class="profile_light" height="20"></td>
    <td width="60" class="profile_light" height="20"></td>
    <td width="26" class="profile_light" height="20"></td>
    <td width="26" class="profile_light" height="20"></td>
    <td width="1" class="profile_light" height="20"></td>
    <td width="1" class="profile_light" height="20"></td>
    <td width="13" class="profile_light" height="20"></td>
    <td width="1" class="profile_light" height="20"></td>
    <td width="1" class="profile_light" height="20"></td>
    <td width="1" class="profile_light" height="20"></td>
  </tr>

-->
  <%

		End If
		        
        %>
</table>
<table border="0" style="border-collapse: collapse" bordercolor="#FFFFFF" width="1110" height="4">
  <tr>
      <td background="images/ruler.gif"></td>
  </tr>
</table>
<p>&nbsp;</p>
</form>
<!--#INCLUDE FILE="footer.asp"-->
</body>
</html>
