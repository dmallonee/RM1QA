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

	strClientUserid = Request.Form("userid")
	strCity = Request.Form("city")
	strCarType = Request.Form("car_type")
	strCompany = Request.Form("company")
	
	strSearched = False
	
	If (strClientUserid = "") And (strCity = "") And (strCarType  = "") And (strCompany = "") Then
	
		strConn = Session("pro_con")
	
  		Set adoCmd = CreateObject("ADODB.Command")

		adoCmd.ActiveConnection =  strConn
		adoCmd.CommandText = "car_shop_request_select1"
		adoCmd.CommandType = 4

		'adoCmd.Parameters.Refresh 
		adoCmd.Parameters.Append adoCmd.CreateParameter("@client_userid", 200, 1, 20)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@city_cd", 200, 1, 5)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@shop_car_type_cds", 200, 1, 255)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@vendor_cd", 200, 1, 255)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@user_id", 3, 1, 0)

		adoCmd.Parameters("@user_id").Value = Request.Cookies("rate-monitor.com")("user_id")
		adoCmd.Parameters("@client_userid").Value = Null
		adoCmd.Parameters("@city_cd").Value = Null
		adoCmd.Parameters("@shop_car_type_cds").Value = Null
		adoCmd.Parameters("@vendor_cd").Value = Null 
		
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
<title>Rate-Monitor by Rate-Highway, Inc. | Search Queue</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta http-equiv="refresh" content="120;url=search_queue_car_ex.asp">
<link rel="stylesheet" type="text/css" href="inc/rh_standard.css">
<link rel="stylesheet" type="text/css" href="inc/rh_report.css">
<link rel="stylesheet" type="text/css" href="inc/sitewide.css">
<script src="inc/sitewide.js" language="javascript"></script>
<script src="inc/header_menu_support.js" language="JavaScript" type="text/JavaScript"></script>
<script language ="javascript" >
<!--

function not_enabled() {
	alert("This section is currently  unavailable."  + "\n" + "Please contact your Rate-Highway rep if you would like to enable it");
	//return true;
}
//-->
</script>
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
        <a href="search_queue_car_ex.asp" onmouseover="MM_swapImage('s2','','images/b_search_que_on.gif',1)" onmouseout="MM_swapImgRestore()">
        <img src="images/b_search_que_of.gif" name="s2" border="0" id="s2" width="97" height="32"></a></td>
        <td>
        <a href="search_criteria_car.asp" onmouseover="MM_swapImage('s3','','images/b_search_cri_on.gif',1)" onmouseout="MM_swapImgRestore()">
        <img src="images/b_search_cri_of.gif" name="s3" border="0" id="s3" width="103" height="32"></a></td>
        <td>
        <a href="javascript:not_enabled()" onmouseover="MM_swapImage('ra','','images/b_rate_on.gif',1)" onmouseout="MM_swapImgRestore()">
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
<form method="POST" action="search_queue_car_ex.asp" name="search" class="search">
  <table border="0" style="border-collapse: collapse" bordercolor="#FFFFFF" width="2110" height="4">
    <tr>
      <td background="images/ruler.gif"></td>
     </tr>
  </table>
  <table border="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#111111" width="1110" id="AutoNumber1" background="images/alt_color.gif">
    <tr>
      <td width="19" height="18">&nbsp;</td>
      <td width="182">&nbsp;</td>
      <td width="278" height="18" colspan="2">
      <font size="2" face="Verdana, Arial, Helvetica, sans-serif">To search enter 
      the last name, or a portion of. You may also optionally enter city, car type 
      and/or the car company.</font></td>
      <td width="699" height="18">&nbsp;</td>
    </tr>
    <tr>
      <td width="19" height="26">&nbsp;</td>
      <td width="182"><img border="0" src="images/search.GIF"></td>
      <td width="87" height="26">
      <font size="2" face="Verdana, Arial, Helvetica, sans-serif">User Name:
      </font></td>
      <td width="191" height="26">
      <input type="text" name="userid" size="20" value="<%=strClientUserid %>" onfocus="this.className='focus';" onblur="this.className='';">
      </td>
      <td width="699" height="26">
      <font size="2" face="Verdana, Arial, Helvetica, sans-serif">
      <input type="submit" value="  Display  " name="submit" class="rh_button"></font></td>
    </tr>
    <tr>
      <td width="19" height="22">&nbsp;</td>
      <td width="182">&nbsp;</td>
      <td width="87" height="22">
      <font size="2" face="Verdana, Arial, Helvetica, sans-serif">City:</font></td>
      <td width="191" height="22">
      <input type="text" name="city" size="20" value="<%=strCity %>" onfocus="this.className='focus';cl(this,'<%=strCity %>');" onblur="this.className='';fl(this,'<%=strstrCity %>');"></td>
      <td width="699" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="19" height="22">&nbsp;</td>
      <td width="182">&nbsp;</td>
      <td width="87" height="22">
      <font size="2" face="Verdana, Arial, Helvetica, sans-serif">Car Type:</font></td>
      <td width="191" height="22">
      <input type="text" name="car_type" size="20" value="<%=strCarType %>" onfocus="this.className='focus';cl(this,'<%=strCarType %>');" onblur="this.className='';fl(this,'<%=strCarType %>');"></td>
      <td width="699" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="19" height="22">&nbsp;</td>
      <td width="182">&nbsp;</td>
      <td width="87" height="22">
      <font size="2" face="Verdana, Arial, Helvetica, sans-serif">Company:</font></td>
      <td width="191" height="22">
      <input type="text" name="company" size="20" value="<%=strCompany %>" onfocus="this.className='focus';cl(this,'<%=strCompany %>');" onblur="this.className='';fl(this,'<%=strCompany %>');"></td>
      <td width="699" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="19" height="18">&nbsp;</td>
      <td width="182">&nbsp;</td>
      <td width="87" height="18">&nbsp;</td>
      <td width="191" height="18">&nbsp;</td>
      <td width="699" height="18">&nbsp;</td>
    </tr>
  </table>
  <table border="0" style="border-collapse: collapse" bordercolor="#FFFFFF" width="2110" height="4">
    <tr>
      <td background="images/ruler.gif"></td>
    </tr>
  </table>
</form>
<table width="1110" border="0" cellpadding="2" background="images/alt_color.gif" style="border-collapse: collapse" bordercolor="#111111">
  <tr valign="bottom">
    <td width="169">&nbsp;<font size="2" face="Verdana, Arial, Helvetica, sans-serif"><a href="search_queue_car_ex.asp">|&lt;</a>
    <a href="search_queue_car_ex.asp">&lt;</a> Page 1 of 1
    <a href="search_queue_car_ex.asp">&gt;</a> <a href="search_queue_car_ex.asp">&gt;|</a></font></td>
  </tr>
</table>
<table border="0" style="border-collapse: collapse" bordercolor="#FFFFFF" width="1110" height="4">
  <tr>
    <td background="images/ruler.gif"></td>
  </tr>
</table>
<form name="queue" method="POST" action="cancel_search_car.asp" >

<table border="1" style="border-collapse: collapse" bordercolor="#FFFFFF" width="1110" id="profiles">
  <!-- MSTableType="layout" -->
  <tr>
    <td class="profile_header" align="left" valign="bottom" bgcolor="#879AA2" height="45" width="26">ID</td>
    <td class="profile_header" width="63" style="background-color: #E07D1A" height="45">
    Selected</td>
    <td class="profile_header" width="102" height="45">User</td>
    <td class="profile_header" width="102" height="45">Search Status</td>
    <td class="profile_header" width="100" height="45">Profile</td>
    <td class="profile_header" width="50" height="45">Action</td>
    <td class="profile_header" width="54" height="45">Search Units</td>
    <td class="profile_header" width="58" height="45">Rate Units Expected</td>
    <td class="profile_header" width="80" height="45">Rate Units Complete</td>
    <td class="profile_header" width="86" height="45">First/Last Rental Date</td>
  </tr>
  <% Dim strClass %> 
  <% If adoRS.State = adStateOpen Then %> 
  <% While adoRS.EOF = False %>
  <%	If strClass <> "profile_dark" Then
        			strClass = "profile_dark" 'background-color:#B2BEC4
        		Else
        			strClass = "profile_light" 'background-color:#CFD7DB
        		End If
        		
			'OnClick="this.cell(,3).background = '#E07D1A'        	
			'onMouseOver="this.style.background = '#E07D1A'" onMouseOut="this.style = 'profile_light'"
        	%>
  <tr class="<%=strClass %>">
    <td width="26" class="<%=strClass %>" height="80" rowspan="5" valign="middle">
    <a href="view_report_car2.asp?ReportRequestID=<%=adoRS.Fields("shop_request_id").Value %>" target="_blank">
    <%=adoRS.Fields("shop_request_id").Value %></a></td>
    <td width="63" bgcolor="#FDC677" align="center" height="80" rowspan="5"  valign="middle">
<!--
    <input type="radio" name="shop_request_id" value="<%=adoRS.Fields("shop_request_id").Value %>" onclick="UpdateProfileForMaint(this.value)" ></td>
  -->
    <% If adoRS.Fields("request_status").Value = "C" Then %>
    <input type="checkbox" value='<%=adoRS.Fields("shop_request_id").Value %>' name="shop_request_id" disabled></td>
    <% Else %>
    <input type="checkbox" value='<%=adoRS.Fields("shop_request_id").Value %>' name="shop_request_id"></td>
    <% End If %>

    
    <td width="97" height="20">
    <%=adoRS.Fields("client_userid").Value %></td>
    <td width="102" height="20" >
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
    <td width="100" class="<%=strClass %>" height="20">
	<% If adoRS.Fields("profile_desc").Value = "" Then %>
	 <i>[none]</i> On-demand report request
	<% Else %>
	  <%=Left(adoRS.Fields("profile_desc").Value, 40) %>
	<% End If %>
	</td>
    <td width="50" class="<%=strClass %>" height="20"><%=adoRS.Fields("rtrn_city_cd").Value %></td>
    <td width="58" class="<%=strClass %>" height="20"><%=adoRS.Fields("work_units").Value %></td>
    <td width="80" class="<%=strClass %>" height="20"><%=adoRS.Fields("work_units").Value %></td>
    <td width="76" class="<%=strClass %>" height="20"><%=adoRS.Fields("work_units_complete").Value %></td>
    <td height="20"><%=adoRS.Fields("begin_arv_dt").Value %>/<%=adoRS.Fields("end_arv_dt").Value %></td>
  </tr>
  <tr class="<%=strClass %>">
    <td width="97" height="20" bgcolor="#E1DFCC" align="right">City Codes:</td>
    <td height="20" colspan="7" ><%=adoRS.Fields("city_cd").Value %></td>
  </tr>
  <tr class="<%=strClass %>">
    <td width="97" height="20" bgcolor="#E1DFCC" align="right">Car Type:</td>
    <td height="20" colspan="7" ><%=adoRS.Fields("shop_car_type_cds").Value %></td>
  </tr>
  <tr class="<%=strClass %>">
    <td width="97" height="20" bgcolor="#E1DFCC" align="right">Companies:</td>
    <td height="20" colspan="7" ><%=adoRS.Fields("vend_cd").Value %></td>
  </tr>
  <tr  class="<%=strClass %>">
    <td width="97" height="20" bgcolor="#E1DFCC" align="right">Full Profile:</td>
    <td height="20" colspan="7" >
    <%=adoRS.Fields("profile_desc").Value %></td>
  </tr>
  <% adoRS.MoveNext %> 
  <% Wend %> 
  <% End If %>
  </table>

<table border="0" style="border-collapse: collapse" bordercolor="#FFFFFF" width="1110" height="4">
  <tr>
      <td background="images/ruler.gif"></td>
  </tr>
</table>
<p>&nbsp;| <a href="javascript:document.queue.submit();">Cancel</a> 
| Extended view | <a href="search_queue_car.asp">Standard view</a> |</p>
</form>
<!--#INCLUDE FILE="footer.asp"-->
</body>
</html>
