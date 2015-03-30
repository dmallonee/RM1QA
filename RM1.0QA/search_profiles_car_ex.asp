<%@ Language=VBScript %>
<!-- #INCLUDE FILE="inc/login_check.asp" -->
<!-- #INCLUDE FILE="inc/adovbs.asp" -->
<% Response.Expires = -1
   Response.cachecontrol="private" 
   Response.AddHeader "pragma", "no-cache"

	On Error Resume Next
	
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
	Dim strUserId 
	Dim blnExpandedTable 

	strProfileDesc = Request.Form("profile_desc")
	strProfileCarType = Request.Form("profile_car_type")
	strProfileCarCo = Request.Form("profile_car_co")
	strUserId = Request.Cookies("rate-monitor.com")("user_id") 
	'Session("user_id")
	
	blnExpandedTable = False
	
	strSearched = False
	
	
	If strUserId <> "" Then
	
		strConn = Session("pro_con")
	
  		Set adoCmd = CreateObject("ADODB.Command")

		adoCmd.ActiveConnection =  strConn
		adoCmd.CommandText = "car_shop_profile_select"
		adoCmd.CommandType = 4

		'adoCmd.Parameters.Refresh 
		adoCmd.Parameters.Append adoCmd.CreateParameter("@desc", 200, 1, 255)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@shop_car_type_cds", 200, 1, 4096)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@vend_cds", 200, 1, 1024)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@user_id", 3, 1, 0)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@profile_id", 3, 1, 0)

		adoCmd.Parameters("@user_id").Value = strUserId 

		If Trim(strProfileDesc) <> "" Then
			adoCmd.Parameters("@desc").Value = strProfileDesc 
		Else
			adoCmd.Parameters("@desc").Value = Null
		End If


		If Trim(strProfileCarType) <> "" Then
			adoCmd.Parameters("@shop_car_type_cds").Value = strProfileCarType 
		Else
			adoCmd.Parameters("@shop_car_type_cds").Value = Null
		End If
	

		If Trim(strProfileCarCo) <> "" Then
			adoCmd.Parameters("@vend_cds").Value = strProfileCarCo 
		Else
			adoCmd.Parameters("@vend_cds").Value = Null 
		End If

		Set adoRS = adoCmd.Execute

'		Set adoRS = CreateObject("ADODB.Recordset")


  		Set adoCmd = CreateObject("ADODB.Command")

		adoCmd.ActiveConnection =  strConn
		adoCmd.CommandText = "user_select"
		adoCmd.CommandType = 4

		adoCmd.Parameters.Append adoCmd.CreateParameter("@user_id", 3, 1, 0)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@user_id_filter", 3, 1, 0, strUserId)

		Set adoUsers = adoCmd.Execute
	
		strSearched = True
		
	Else
		Set adoRS = CreateObject("ADODB.Recordset")
		Set adoUsers = CreateObject("ADODB.Recordset")

	End If



%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<title>Rate-Monitor by Rate-Highway, Inc. | Search Profiles</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<link rel="stylesheet" type="text/css" href="inc/rh_standard.css">
<link rel="stylesheet" type="text/css" href="inc/rh_report.css">
<link rel="stylesheet" type="text/css" href="inc/sitewide.css">
<script src="inc/sitewide.js" language="javascript"></script>
<script src="inc/header_menu_support.js" language="JavaScript" type="text/JavaScript"></script>
<script language="JavaScript" type="text/JavaScript">
function UpdateProfileForMaint(this_radio_value) {

	document.maintenance.profile_id.value = this_radio_value;
	//alert(this_radio) ;

	}
	
	//document.profiles.profile_radio



function not_enabled() {
	alert("This section is currently  unavailable."  + "\n" + "Please contact your Rate-Highway rep if you would like to enable it");
	//return true;
}

function ChangeExpandedView(checked) {
	alert("the checkbox is = " + checked );

}

</script>
</head>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onload="MM_preloadImages('images/b_search_pro_on.gif','images/b_search_que_on.gif','images/b_search_cri_on.gif','images/b_rate_on.gif','images/b_alert_on.gif','images/b_user_on.gif','images/b_system_on.gif')">

<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td background="images/top_middle.jpg">
    <img src="images/top_left.jpg" width="423" height="91"></td>
    <td background="images/top_middle.jpg"></td>
    <td background="images/top_tile.gif">
    <img src="images/top_right.jpg" width="365" height="91"></td>
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
                <td>
                <div align="right">
                  <font face="Arial, Helvetica, sans-serif" size="2" color="#FFFFFF">
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

<table width="1110" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td background="images/h_tile.gif">
    <table width="100" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td><img src="images/h_search_profiles.gif" width="368" height="31"></td>
        <td><img src="images/h_right.gif" width="402" height="31"></td>
      </tr>
    </table>
    </td>
  </tr>
</table>
&nbsp;<form method="POST" action="search_profiles_car_ex.asp" name="search_profiles" class="search">
  <table border="0" style="border-collapse: collapse" bordercolor="#FFFFFF" width="1110" cellspacing="0" height="4" id="table1">
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
  <table border="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#111111" width="1110" id="table2" background="images/alt_color.gif">
    <tr>
      <td width="19" height="18">&nbsp;</td>
      <td width="182">&nbsp;</td>
      <td width="278" height="18" colspan="2">
      <font size="2" face="Verdana, Arial, Helvetica, sans-serif">To search for 
      a profile, enter the last name, or a portion of. You may also enter the car 
      type and/or the car companies.</font></td>
      <td width="699" height="18">&nbsp;</td>
    </tr>
    <tr>
      <td width="19" height="26">&nbsp;</td>
      <td width="182"><img border="0" src="images/search.GIF"></td>
      <td width="87" height="26"><font size="2">Profile</font><font size="2" face="Verdana, Arial, Helvetica, sans-serif">:
      </font></td>
      <td width="191" height="26">
      <input type="text" name="profile_desc" size="20" value="<%=strProfileDesc %>" onfocus="this.className='focus';cl(this,'<%=strClientUserid %>');" onblur="this.className='';fl(this,'<%=strClientUserid %>');"></td>
      <td width="699" height="26">
      <font size="2" face="Verdana, Arial, Helvetica, sans-serif">
      <input type="submit" value="  Display  " name="submit0" class="rh_button"></font></td>
    </tr>
    <tr>
      <td width="19" height="22">&nbsp;</td>
      <td width="182">&nbsp;</td>
      <td width="87" height="22">
      <font size="2" face="Verdana, Arial, Helvetica, sans-serif">Car Type:</font></td>
      <td width="191" height="22">
      <input type="text" name="profile_car_type" size="20" value="<%=strProfileCarType %>" onfocus="this.className='focus';cl(this,'<%=strCarType %>');" onblur="this.className='';fl(this,'<%=strCarType %>');"></td>
      <td width="699" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="19" height="22">&nbsp;</td>
      <td width="182">&nbsp;</td>
      <td width="87" height="22">
      <font size="2" face="Verdana, Arial, Helvetica, sans-serif">Company:</font></td>
      <td width="191" height="22">
      <input type="text" name="profile_car_co" size="20" value="<%=strProfileCarCo  %>" onfocus="this.className='focus';cl(this,'<%=strCompany %>');" onblur="this.className='';fl(this,'<%=strCompany %>');"></td>
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
  <table border="0" style="border-collapse: collapse" bordercolor="#FFFFFF" width="1110" height="4" id="table3">
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
</form>
<form name='profiles' method="POST" action="enable_profiles_car.asp">
<table width="1108" border="0" cellpadding="2" cellspacing="0" background="images/alt_color.gif" style="border-collapse: collapse" bordercolor="#111111">
  <tr valign="bottom">
    <td width="169">&nbsp;<font size="2" face="Verdana, Arial, Helvetica, sans-serif"><a href="search_profiles_car_ex.asp">|&lt;</a>
    <a href="search_profiles_car_ex.asp">
    &lt;</a> Page 1 of 
    1 <a href="search_profiles_car_ex.asp">&gt;</a> 
    <a href="search_profiles_car_ex.asp">
    &gt;|</a></font></td>
  </tr>
</table>
<table border="0" style="border-collapse: collapse" bordercolor="#FFFFFF" width="1110" cellspacing="0" height="4">
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
<table border="1" style="border-collapse: collapse" bordercolor="#FFFFFF" width="1110" id="profiles">
  <tr>
    <td align="left" valign="bottom" bgcolor="#879AA2" height="45" width="26">&nbsp;</td>
    <td class="profile_header" width="63" style="background-color: #E07D1A" height="45">
    Selected</td>
    <td class="profile_header" width="139" height="45">User</td>
    <td class="profile_header" width="341" height="45">Profile</td>
    <td class="profile_header" width="103" height="45">Action</td>
    <td class="profile_header" width="54" height="45">Rtrn<br>
    City</td>
    <td class="profile_header" width="58" height="45">LOR</td>
    <td class="profile_header" width="80" height="45">First<br>
    Rental<br>
    Date</td>
    <td class="profile_header" width="76" height="45">Last<br>
    Rental<br>
    Date</td>
    <td class="profile_header" width="129" height="45">Data Sources</td>
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
    <td width="26" class="<%=strClass %>" height="80" rowspan="5" valign="middle"><%=adoRS.Fields("profile_id").Value %><%=lcase(adoRS.Fields("profile_status").Value) %></td>
    <td width="63" bgcolor="#FDC677" align="center" height="80" rowspan="5"  valign="middle">
    <input type="radio" name="profile_id" value="<%=adoRS.Fields("profile_id").Value %>" onclick="UpdateProfileForMaint(this.value)" alt="<%=adoRS.Fields("desc").Value %>" ></td>
    <td width="139" height="20"><%=adoRS.Fields("last_name").Value %></td>
    <td width="341" height="20" ><font title="<%=adoRS.Fields("desc").Value %>"><a href="search_criteria_car.asp?profile=<%=adoRS.Fields("profile_id").Value %>">
    							 <% If Len(adoRS.Fields("desc").Value) > 35 Then %>
    									<%=Left(adoRS.Fields("desc").Value, 35) %>
								 <%	Else %>   									
										<%=adoRS.Fields("desc").Value %>
								 <%	End If %>
								 </font></a></td>
    <td width="103" class="<%=strClass %>" height="20"><%=Left(adoRS.Fields("action_desc").Value, 7) %></td>
    <td width="54" class="<%=strClass %>" height="20"><%=adoRS.Fields("rtrn_city_cd").Value %></td>
    <td width="58" class="<%=strClass %>" height="20">
    <% Select Case adoRS.Fields("lor").Value %>
    <%	Case 1000	%>
    Daily
    <%	Case 1001	%>
    Wkend Daily
    <%	Case 1002	%>
    Weekly
    <%	Case Else	%>
    <%=adoRS.Fields("lor").Value %>
    <% End Select %>
    </td>
    <td width="80" class="<%=strClass %>" height="20"><%=FormatDateTime(adoRS.Fields("begin_arv_dt").Value, 2 ) %></td>
    <td width="76" class="<%=strClass %>" height="20"><%=FormatDateTime(adoRS.Fields("end_arv_dt").Value, 2 ) %></td>
    <td width="129" class="<%=strClass %>" height="20"><%=adoRS.Fields("data_sources").Value %></td>
  </tr>
  <tr class="<%=strClass %>">
    <td width="139" height="20" bgcolor="#E1DFCC" align="right">City Codes:</td>
    <td width="836" height="20" colspan="7" ><%=adoRS.Fields("city_cd").Value %></td>
  </tr>
  <tr class="<%=strClass %>">
    <td width="139" height="20" bgcolor="#E1DFCC" align="right">Car Type:</td>
    <td width="836" height="20" colspan="7" ><%=adoRS.Fields("shop_car_type_cds").Value %></td>
  </tr>
  <tr class="<%=strClass %>">
    <td width="139" height="20" bgcolor="#E1DFCC" align="right">Companies:</td>
    <td width="671" height="20" colspan="7" ><%=adoRS.Fields("vend_cds").Value %></td>
  </tr>
  <tr  class="<%=strClass %>">
    <td width="139" height="20" bgcolor="#E1DFCC" align="right">Full Profile:</td>
    <td width="671" height="20" colspan="7" ><%=adoRS.Fields("desc").Value %></td>
  </tr>
  <% adoRS.MoveNext %> 
  <% Wend %> 
  <% End If %>
  </table>
<table border="0" style="border-collapse: collapse" bordercolor="#FFFFFF" width="1110" cellspacing="0" height="4">
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
</form>
<p>| <a href="javascript:document.profiles.submit();">Enable</a> | <a href="javascript:document.profiles.submit();">Disable</a> |
Extended view | <a href="search_profiles_car.asp">Standard view</a> |<form method="POST" action="search_profiles_maint_car.asp" name="maintenance">
<table border="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#111111" width="1110" id="AutoNumber1" background="images/alt_color.gif">
  <tr>
    <td width="11">&nbsp;</td>
    <td width="247">&nbsp;</td>
    <td width="715">&nbsp;</td>
    <td width="326">&nbsp;</td>
  </tr>
  <tr>
    <td width="11">&nbsp;</td>
    <td width="247"><img border="0" src="images/maintenance.GIF"></td>
    <td width="715">
    <input type="radio" name="maint_action" value="enable" checked><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Enable 
    Profile</td>
    <td width="326">&nbsp;</td>
  </tr>
  <tr>
    <td width="11">&nbsp;</td>
    <td width="247">&nbsp;</td>
    <td width="715">
    <input type="radio" name="maint_action" value="disable"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Disable 
    Profile</td>
    <td width="326">&nbsp;</td>
  </tr>
  <tr>
    <td width="11">&nbsp;</td>
    <td width="247">&nbsp;</td>
    <td width="715">
    <input type="radio" name="maint_action" value="delete"><font size="2">Delete</font><font size="2" face="Verdana, Arial, Helvetica, sans-serif"> 
    Profile</td>
    <td width="326">&nbsp;</td>
  </tr>
  <tr>
    <td width="11">&nbsp;</td>
    <td width="247">&nbsp;</td>
    <td width="715">
    <input type="radio" name="maint_action" value="copy"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Copy 
    Profile (same user)</td>
    <td width="326">&nbsp;</td>
  </tr>
  <tr>
    <td width="11">&nbsp;</td>
    <td width="247">&nbsp;</td>
    <td width="715">
    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <font size="2">N</font><font size="2" face="Verdana, Arial, Helvetica, sans-serif">ew 
    name: <input type="text" name="new_name" size="33" style="width: 250"></td>
    <td width="326">&nbsp;</td>
  </tr>
  <tr>
    <td width="11">&nbsp;</td>
    <td width="247">&nbsp;</td>
    <td width="715">
    <input type="radio" name="maint_action" value="copy_new_user"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Copy 
    Profile (to another user)</td>
    <td width="326">&nbsp;</td>
  </tr>
  <tr>
    <td width="11">&nbsp;</td>
    <td width="247">&nbsp;</td>
    <td width="715">
    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <font size="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
    User</font><font size="2" face="Verdana, Arial, Helvetica, sans-serif">: 
    <select size="1" name="users" style="width: 250">
    <% While adoUsers.EOF = False			%>
	    
	    <option value='<%=adoUsers.Fields("user_id").Value %>'><%=adoUsers.Fields("client_userid").Value %></option>
	    <% adoUsers.MoveNext %>

	<% Wend									%>
    </select></td>
    <td width="326">&nbsp;</td>
  </tr>
  <tr>
    <td width="11">&nbsp;</td>
    <td width="247">&nbsp;</td>
    <td width="715">
      &nbsp;</td>
    <td width="326">&nbsp;</td>
  </tr>
  <tr>
    <td width="11">&nbsp;</td>
    <td width="247">&nbsp;</td>
    <td width="715"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">
    <input type="submit" value="Submit" name="maint_submit" class="rh_button"></td>
    <td width="326">&nbsp;</td>
  </tr>
  <tr>
    <td width="11">&nbsp;</td>
    <td width="247">&nbsp;</td>
    <td width="715">&nbsp;</td>
    <td width="326">&nbsp;</td>
  </tr>
</table>
<input type="hidden" name="profile_id" value="0">
</form>
<!--#INCLUDE FILE="footer.asp"-->
</body>
</html>
<%		Set adoRS = Nothing %>