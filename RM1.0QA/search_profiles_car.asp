<%@ Language=VBScript %>
<!-- #INCLUDE FILE="inc/login_check_ex.asp" -->
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
	

	'Declare variables
	Dim intCurrentPage
	Dim iPageSize
	Dim sPageURL

	'Retrieve the name of the current ASP document
	sPageURL = Request.ServerVariables("SCRIPT_NAME")

	'Retrieve the current page number from the QueryString
	intCurrentPage = Request.QueryString("page")
	If intCurrentPage = "" Or intCurrentPage = 0 Then intCurrentPage = 1

	'Set the number of records to be displayed on each page
	iPageSize = 200

	strProfileDesc = Request("profile_desc")
	strProfileCarType = Request("profile_car_type")
	strProfileCarCo = Request("profile_car_co")
	strProfileCity = Request("profile_city")
    if Request("profile_status") = "" Then
        intProfileStatus = 1
    else
	    intProfileStatus = Request("profile_status")
    end if

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
		adoCmd.Parameters.Append adoCmd.CreateParameter("@desc",              200, 1, 255)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@shop_car_type_cds", 200, 1, 1024)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@vend_cds",          200, 1, 1024)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@user_id",             3, 1, 0)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@profile_id",          3, 1, 0)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@city_cd",           200, 1, 1024)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@enabled",             3, 1, 0)
		'adoCmd.Parameters.Append adoCmd.CreateParameter("@size",                3, 1, 0)
		'adoCmd.Parameters.Append adoCmd.CreateParameter("@pages",               3, adParamOutput, 0)
		
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

		If Trim(strProfileCity) <> "" Then
			adoCmd.Parameters("@city_cd").Value = strProfileCity 
		Else
			adoCmd.Parameters("@city_cd").Value = Null 
		End If

		If IsNumeric(intProfileStatus) Then
			adoCmd.Parameters("@enabled").Value = intProfileStatus 
		Else
			adoCmd.Parameters("@enabled").Value = 1 
		End If

		If strPage= "" Then
			strPage= 1
		End If

		'If IsNumeric(strPage) Then
		'	adoCmd.Parameters("@page").Value = strPage
		'Else
		'	adoCmd.Parameters("@page").Value = 1
		'	strPage = 1
		'End If


		'If strPageSize = "" Then
		'	strPageSize = 100
		'End If

		'If IsNumeric(strPageSize) Then
		'	If strPageSize > 200 Then
		'		adoCmd.Parameters("@size").Value = 200
		'		strPageSize = 200
		'	Else
		'		adoCmd.Parameters("@size").Value = strPageSize 
		'	End If
		'Else
		'	adoCmd.Parameters("@size").Value = 100
		'	strPageSize = 100
		'End If

		'Set adoRS = adoCmd.Execute


		'Create an ADO RecordSet object
		Set adoRS = Server.CreateObject("ADODB.Recordset")

		'Set the RecordSet PageSize property
		adoRS.PageSize = iPageSize 
		adoRS.CursorLocation = adUseClient 

		'Set the RecordSet CacheSize property to the
		'number of records that are returned on each page of results
		adoRS.CacheSize = iPageSize 

		'Open the RecordSet
		adoRS.Open adoCmd, , adOpenStatic, adLockReadOnly

		If adoRS.EOF = False Then
		
			adoRS.PageSize = iPageSize 
			intPageCount = adoRS.PageCount
			adoRS.AbsolutePage = intCurrentPage 

		Else
			intPageCount = 1
	
		End If



	  	If err.number <> 0 Then
		   pad="&nbsp;&nbsp;&nbsp;&nbsp;"
		   response.write "<b>1VBScript Errors Occured!<br>"
		   response.write "</b><br>"
		   response.write pad & "Error Number= #<b>" & err.number & "</b><br>"
		   response.write pad & "Error Desc.= <b>" & err.description & "</b><br>"
		   response.write pad & "Help Context= <b>" & err.HelpContext & "</b><br>"
		   response.write pad & "Help File Path=<b>" & err.helpfile & "</b><br>"
		   response.write pad & "Error Source= <b>" & err.source & "</b><br><hr>"

		End If

		strPages = 0 'adoCmd.Parameters("@pages").Value  

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

  	If err.number <> 0 Then
	   pad="&nbsp;&nbsp;&nbsp;&nbsp;"
	   response.write "<b>2VBScript Errors Occured!<br>"
	   response.write "</b><br>"
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
<meta http-equiv="Content-Language" content="en-us" >
<META HTTP-EQUIV="refresh" CONTENT="600;URL=default_session.asp" >
<title>Rate-Monitor by Rate-Highway, Inc. | Search Profiles</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252" >
<link rel="stylesheet" type="text/css" href="inc/rh_standard.css" >
<link rel="stylesheet" type="text/css" href="inc/rh_report.css" >
<link rel="stylesheet" type="text/css" href="inc/sitewide.css" >
<script type="text/javascript" language="javascript" src="inc/sitewide.js" ></script>
<script type="text/javascript" language="javascript" src="inc/header_menu_support.js" ></script>
<script type="text/javascript" language="javascript"  >
<!-- Begin
function checkAll(field)
{
	//alert(field.name);

	//alert(field[0].disabled);	

	for (i = 0; i < field.length; i++) 
		{
	//	if (field[i].disabled == "disabled")
			
	//	else
			field[i].checked = true ;
		}
}

function uncheckAll(field)
{

	for (i = 0; i < field.length; i++)
		field[i].checked = false ;
}
//  End -->
</script>
<script type="text/javascript" language="javascript" >
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
<script type='text/javascript' language='javascript' >
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//
// Page submition section
//
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

function confirmSubmit(SubmitType) {

	if (SubmitType == 'enable'){
		document.profiles.action = 'enable_profiles_car.asp'
		//document.search_criteria.submit;
		//document.queue.submit;
		//return true;   
		}

	if (SubmitType == 'run'){
		document.profiles.action = 'run_profiles_car.asp'
		//document.search_criteria.submit;
		//document.queue.submit;
		//return true;   
		}


}

</script>
</head>

<body leftmargin="0" topmargin="0" onload="MM_preloadImages('images/b_search_pro_on.gif','images/b_search_que_on.gif','images/b_search_cri_on.gif','images/b_rate_on.gif','images/b_alert_on.gif','images/b_user_on.gif','images/b_system_on.gif')">

<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td background="images/top_middle.jpg">
    <img src="images/top_left.jpg" width="423" height="91" alt=""></td>
    <td background="images/top_middle.jpg"></td>
    <td background="images/top_tile.gif">
    <img src="images/top_right.jpg" width="365" height="91" alt=""></td>
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
    <img src="images/med_bar.gif" width="12" height="8" alt=""></td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td background="images/user_tile.gif">
    <table width="100" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td><img src="images/user_left.gif" width="580" height="31" alt=""></td>
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
                <td><img src="images/separator.gif" width="183" height="6" alt=""></td>
              </tr>
            </table>
            </td>
            <td><img src="images/user_tile.gif" width="7" height="31" alt=""></td>
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
        <td><img src="images/h_search_profiles.gif" width="368" height="31" alt=""></td>
        <td><map name="logout_map">
        <area alt="Click to log out of Rate-Monitor" href="http://www.rate-monitor.com/" shape="rect" coords="337, 10, 394, 25">
        </map>
        <img src="images/bottom_right.gif" width="402" height="31" usemap="#logout_map" border="0" alt=""></td>
      </tr>
    </table>
    </td>
  </tr>
</table>
<p align="right">&nbsp;</p>
<!-- 
                <img border="0" alt="Click to view Help" src="images/help_button.jpg" width="32" height="32" onclick="centerPopUp('help_search_profile.htm', 'help', 650, 400, 1)"></font></p>
-->                
<form method="POST" action="search_profiles_car.asp" name="search_profiles" class="search">
  <table border="0" style="border-collapse: collapse" bordercolor="#FFFFFF" width="1110" cellspacing="0" height="4" id="table1">
    <tr>
      <td >&nbsp;</td>
      <td >&nbsp;</td>
      <td >&nbsp;</td>
      <td >&nbsp;</td>
      <td >&nbsp;</td>
      <td >&nbsp;</td>
      <td >&nbsp;</td>
      <td >&nbsp;</td>
      <td >&nbsp;</td>
      <td >&nbsp;</td>
      <td >&nbsp;</td>
      <td >&nbsp;</td>
      <td >
          <font color="#080000">
          <a href="http://wiki.rate-monitor.com/index.php?title=Creating_Search_Profiles&amp;user=Rate-Monitor&amp;pass=online_user" target="_blank">
		  <img border="0" alt="Click to view Help" src="images/help_button.jpg" width="32" height="32" style="float: right" ></a>
          </font>
      </td>
    </tr>
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
      <td width="182"><a name="top"></a></td>
      <td width="302" height="18" colspan="2">
      <font size="2" >To search for 
      a profile, enter the last name, or a portion of. You may also enter the car 
      type and/or the car companies.</font></td>
      <td width="608" height="18">&nbsp;</td>
    </tr>
    <tr>
      <td width="19" height="26">&nbsp;</td>
      <td width="182"><img border="0" src="images/search.GIF" alt="Search"></td>
      <td width="134" height="26"><font size="2">Profile</font><font size="2" >:
      </font></td>
      <td width="168" height="26">
      <input type="text" name="profile_desc" size="20" value="<%=strProfileDesc %>" onfocus="this.className='focus';cl(this,'<%=strClientUserid %>');" onblur="this.className='';fl(this,'<%=strClientUserid %>');"></td>
      <td width="608" height="26">
      <font size="2" >
      <input type="submit" value="  Display  " name="submit0" class="rh_button"></font></td>
    </tr>
    <tr>
      <td width="19" height="22">&nbsp;</td>
      <td width="182">&nbsp;</td>
      <td width="134" height="22">
      <font size="2" >Car Type:</font></td>
      <td width="168" height="22">
      <input type="text" name="profile_car_type" size="20" value="<%=strProfileCarType %>" onfocus="this.className='focus';cl(this,'<%=strCarType %>');" onblur="this.className='';fl(this,'<%=strCarType %>');"></td>
      <td width="608" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="19" height="22">&nbsp;</td>
      <td width="182">&nbsp;</td>
      <td width="134" height="22">
      <font size="2" >Company:</font></td>
      <td width="168" height="22">
      <input type="text" name="profile_car_co" size="20" value="<%=strProfileCarCo  %>" onfocus="this.className='focus';cl(this,'<%=strCompany %>');" onblur="this.className='';fl(this,'<%=strCompany %>');"></td>
      <td width="608" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="19" height="22">&nbsp;</td>
      <td width="182">&nbsp;</td>
      <td width="134" height="22">
      <font size="2" >City:</font></td>
      <td width="168" height="22">
      <input type="text" name="profile_city" size="20" value="<%=strProfileCity  %>" onfocus="this.className='focus';cl(this,'<%=strCompany %>');" onblur="this.className='';fl(this,'<%=strCompany %>');"></td>
      <td width="608" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="19" height="18">&nbsp;</td>
      <td width="182">&nbsp;</td>

          <td width="177" height="22">
          <font size="2" >Profile 
          Status:</font></td>
          <td width="80" height="22">
          <select size="1" name="profile_status" style="border:1px solid #000000; width:145; background-color:#FF9933">
		  <% Select Case intProfileStatus %>
		  <%	Case "0" %>
          <option  value="2" >All types</option>
          <option value="1">Enabled only</option>
          <option selected value="0">Disabled only</option>
		  <%	Case "1" %>
          <option  value="2" >All types</option>
          <option selected value="1">Enabled only</option>
          <option value="0">Disabled only</option>
		  <%	Case "2" %>
          <option selected value="2" >All types</option>
          <option value="1">Enabled only</option>
          <option value="0">Disabled only</option>
		  <%	Case Else %>
          <option  value="2" >All types</option>
          <option selected value="1">Enabled only</option>
          <option value="0">Disabled only</option>
          
          <% End Select %>
          </select></td>

      <td width="608" height="18">&nbsp;</td>
    </tr>
    <tr>
      <td width="19" height="18">&nbsp;</td>
      <td width="182">&nbsp;</td>
      <td width="134" height="18">
      &nbsp;</td>
      <td width="168" height="18">
      &nbsp;</td>
      <td width="608" height="18">&nbsp;</td>
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
<form name='profiles' method="post" action="search_profiles_car.asp">
<table width="1108" border="0" cellpadding="2" cellspacing="0" background="images/alt_color.gif" style="border-collapse: collapse" bordercolor="#111111">
  <tr valign="bottom">
    <td style="width: 52%" ><font size="2">| <a href="javascript:checkAll(document.profiles.profile_id)"  title="Select and check all the profiles on this page" >Select All</a> | <a  href="javascript:uncheckAll(document.profiles.profile_id)" title="Unselect and un-check all the profiles on this page" >Unselect All</a> |
    <a title="Select one or more profiles and click to toggle the enabled/disabled status" href="javascript:confirmSubmit('enable');document.profiles.submit();">Enable</a> 
    | 
    <a title="Select one or more profiles and click to toggle the enabled/disabled status" href="javascript:confirmSubmit('enable');document.profiles.submit();">Disable</a> |
    <a title="Select one or more profiles and click to run the profiles and perform the searches" href="javascript:confirmSubmit('run');document.profiles.submit();">Run</a> |
    <a href="search_profiles_car_ex.asp">Extended view</a> | Standard view
    </font></td>
    <td width="50%" >
    <font size="2" >

    <% If intCurrentPage <= 1 Then %>
	    |&lt;
	    &lt; 
	<% Else  %>
<!--
	    <a href="<%=sPageURL %>?days_back=<%=intDaysBack %>&profile_desc=<%=strProfileDesc %>&city=<%=strCity%>&profile_car_type=<%=strProfileCarType %>&profile_car_co=<%=strProfileCarCo %>&profile_status=<%=intProfileStatus %>&page=1">|&lt;</a>
	    <a href="<%=sPageURL %>?days_back=<%=intDaysBack %>&profile_desc=<%=strProfileDesc %>&city=<%=strCity%>&profile_car_type=<%=strProfileCarType %>&profile_car_co=<%=strProfileCarCo %>&profile_status=<%=intProfileStatus %>&page=<%=intCurrentPage - 1%>">&lt;</a> 
-->	    
	    <a href="search_profiles_car.asp?days_back=<%=intDaysBack %>&profile_desc=<%=strProfileDesc %>&city=<%=strCity%>&profile_car_type=<%=strProfileCarType %>&profile_car_co=<%=strProfileCarCo %>&profile_status=<%=intProfileStatus %>&page=1">|&lt;</a>
	    <a href="search_profiles_car.asp?days_back=<%=intDaysBack %>&profile_desc=<%=strProfileDesc %>&city=<%=strCity%>&profile_car_type=<%=strProfileCarType %>&profile_car_co=<%=strProfileCarCo %>&profile_status=<%=intProfileStatus %>&page=<%=intCurrentPage - 1%>">&lt;</a> 

    <% End If %>
    Page <%=intCurrentPage %> of <%=intPageCount %>
    <% If CInt(intCurrentPage) >= CInt(intPageCount) Then %>
	    &gt;
	    &gt;|
    <% Else %>
    <!--
	    <a href="<%=sPageURL %>?days_back=<%=intDaysBack %>&profile_desc=<%=strProfileDesc %>&city=<%=strCity%>&profile_car_type=<%=strProfileCarType %>&profile_car_co=<%=strProfileCarCo %>&profile_status=<%=intProfileStatus %>&page=<%=intCurrentPage + 1%>">&gt;</a>
	    <a href="<%=sPageURL %>?days_back=<%=intDaysBack %>&profile_desc=<%=strProfileDesc %>&city=<%=strCity%>&profile_car_type=<%=strProfileCarType %>&profile_car_co=<%=strProfileCarCo %>&profile_status=<%=intProfileStatus %>&page=<%=intPageCount %>">&gt;|</a>
-->
	    <a href="search_profiles_car.asp?days_back=<%=intDaysBack %>&profile_desc=<%=strProfileDesc %>&city=<%=strCity%>&profile_car_type=<%=strProfileCarType %>&profile_car_co=<%=strProfileCarCo %>&profile_status=<%=intProfileStatus %>&page=<%=intCurrentPage + 1%>">&gt;</a>
	    <a href="search_profiles_car.asp?days_back=<%=intDaysBack %>&profile_desc=<%=strProfileDesc %>&city=<%=strCity%>&profile_car_type=<%=strProfileCarType %>&profile_car_co=<%=strProfileCarCo %>&profile_status=<%=intProfileStatus %>&page=<%=intPageCount %>">&gt;|</a>

    <% End If %>

    </font></td>
  </tr>
</table>
<table border="0" style="border-collapse: collapse" bordercolor="#FFFFFF" width="1110" cellspacing="0" height="4">
  <tr>
    <td background="images/ruler.gif"></td>
  </tr>
</table>


<table border="1" style="border-collapse: collapse" bordercolor="#FFFFFF" width="1110" id="profiles">
  <tr>
    <td align="left" valign="bottom" bgcolor="#879AA2" height="45" width="26">&nbsp;</td>
    <td class="profile_header" width="63" style="background-color: #E07D1A" height="45">Selected</td>
    <td class="profile_header" width="57" height="45">User</td>
    <td class="profile_header" height="45" style="width: 303px">Profile Description </td>
    <!--
    <td class="profile_header" width="46" height="45">Action</td>
	-->
    <td class="profile_header" height="45" style="width: 65px">City</td>
    <!--
    <td class="profile_header" width="58" height="45">Rtrn<br>
    City</td>
    -->
    <td class="profile_header" height="45" style="width: 30px">LOR</td>
    <td class="profile_header" height="45" style="width: 62px">Begin/<br>1st Day</td>
    <td class="profile_header" height="45" style="width: 61px">End/<br>Days Out</td>
    <td class="profile_header" width="172" height="45">Car Type Codes</td>
    <td class="profile_header" height="45" style="width: 58px">Data Sources</td>
    <td class="profile_header" width="109" height="45">Car Companies</td>
    <td class="profile_header" width="20" height="45">&nbsp;</td>
  </tr>
  <% Dim strClass %>
  <% Dim intCount %>
  <% intCount = 0 %>
   
  <% If adoRS.State = adStateOpen Then %> 
  <% While (adoRS.EOF = False) And (intCount < iPageSize) %>
  <%	intCount = intCount + 1
   		If strClass <> "profile_dark" Then
        			strClass = "profile_dark" 'background-color:#B2BEC4
        		Else
        			strClass = "profile_light" 'background-color:#CFD7DB
        		End If
        		
			'OnClick="this.cell(,3).background = '#E07D1A'        	
			'onMouseOver="this.style.background = '#E07D1A'" onMouseOut="this.style = 'profile_light'"
        	%>
  <tr onclick="this.cells(3).innertext = 'abcd' " class="<%=strClass %>">
    <td width="26" class="<%=strClass %>" height="20"><a target="_blank" title="Click to view utilization report" href="profile_utilization_rpt.asp?profile_id=<%=adoRS.Fields("profile_id").Value %>" ><%=intCount %><%=lcase(adoRS.Fields("profile_status").Value) %></a></td>
    <td width="63" bgcolor="#FDC677" align="center" height="20">
    <input type="checkbox" name="profile_id" value="<%=adoRS.Fields("profile_id").Value %>" onclick="UpdateProfileForMaint(this.value)" alt="<%=adoRS.Fields("desc").Value %>" ></td>
    <td width="57" height="20" title="<%=adoRS.Fields("last_name").Value %>"><%=Left(adoRS.Fields("last_name").Value, 7) %></td>
    <td height="20" title="<%=adoRS.Fields("desc").Value %>" style="width: 303px"><a href="search_criteria_car.asp?profile=<%=adoRS.Fields("profile_id").Value %>&profile_status=<%=intProfileStatus %>">
    							 <% If Len(adoRS.Fields("desc").Value) > 33 Then %>
    									<%=Left(adoRS.Fields("desc").Value, 33) & "..." %>
								 <%	Else %>   									
										<%=adoRS.Fields("desc").Value %>
								 <%	End If %>
								 </a></td>
    <!--								 
    <td width="46" class="<%=strClass %>" height="20"><%=Left(adoRS.Fields("action_desc").Value, 7) %></td>
    -->
    <td  class="<%=strClass %>" height="20" title="<%=adoRS.Fields("city_cd").Value %>" style="width: 65px">
    <% If Len(adoRS.Fields("city_cd").Value) > 18 Then %>
    	<%=Left(adoRS.Fields("city_cd").Value, 16) & "..." %>
	<% Else %>
    	<%=adoRS.Fields("city_cd").Value %>
	<% End If %>
	</td> 
	<!--
    <td width="58" class="<%=strClass %>" height="20"><%=adoRS.Fields("rtrn_city_cd").Value %></td>
    -->
    <td class="<%=strClass %>" height="20" style="width: 30px">
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
    <td class="<%=strClass %>" height="20" style="width: 62px">
    <% If adoRS.Fields("exact_dates").Value = True Then %>
    <%=FormatDateTime(adoRS.Fields("begin_arv_dt").Value, 2 ) %>
    <% Else %>
    <%=adoRS.Fields("days_out").Value %>
    <% End If %>
    </td>
    <td class="<%=strClass %>" height="20" style="width: 61px">
    <% If adoRS.Fields("exact_dates").Value = True Then %>
    <%=FormatDateTime(adoRS.Fields("end_arv_dt").Value, 2 ) %>
    <% Else %>
    <% If adoRS.Fields("days_long").Value = 1 Then %>
	    <%="1 day" %>
	<% Else %>
	    <%=adoRS.Fields("days_long").Value & " days" %>
	<% End If %>
    <% End If %>
    </td>
    <td width="172" class="<%=strClass %>" height="20" title="<%=adoRS.Fields("shop_car_type_cds").Value %>">
    <% If Len(adoRS.Fields("shop_car_type_cds").Value) > 25 Then %>
    	<%=Left(adoRS.Fields("shop_car_type_cds").Value, 25) & "..." %>
	<% Else %>
    	<%=adoRS.Fields("shop_car_type_cds").Value %>
	<% End If %>
	</td> 
 
   <td class="<%=strClass %>" height="20" style="width: 58px">
      <%=adoRS.Fields("data_sources").Value %>
   </td>
    <td width="109" class="<%=strClass %>" height="20"><font title="<%=adoRS.Fields("vend_cds").Value %>">
    <% If (Len(adoRS.Fields("vend_cds").Value) > 17) Then %>
    <%=Left(adoRS.Fields("vend_cds").Value, 17) & "..." %>
    <% Else %>
    <%=adoRS.Fields("vend_cds").Value %>
    <% End If %>
    </font></td>
    <td class="<%=strClass %>" height="20"><a target="_blank" title="Click to view utilization report" href="car_rate_rule_util_maint.asp?profile_id=<%=adoRS.Fields("profile_id").Value %>" ><%=intCount %><%=lcase(adoRS.Fields("profile_status").Value) %></a></td>
  </tr>
  <% adoRS.MoveNext %> <% Wend %> <% End If %>
</table>
<table border="0" style="border-collapse: collapse" bordercolor="#FFFFFF" width="1110" cellspacing="0" height="4">
  <tr>
    <td background="images/ruler.gif"></td>
  </tr>
</table>
</form>
<table width="1108" border="0" cellpadding="2" cellspacing="0" background="images/alt_color.gif" style="border-collapse: collapse" bordercolor="#111111">
  <tr valign="bottom">
    <td style="width: 62%" ><font size="2">| <a href="javascript:checkAll(document.profiles.profile_id)"  title="Select and check all the profiles on this page" >Select All</a> | <a  href="javascript:uncheckAll(document.profiles.profile_id)" title="Unselect and un-check all the profiles on this page" >Unselect All</a> |
    <a title="Select one or more profiles and click to toggle the enabled/disabled status" href="javascript:confirmSubmit('enable');document.profiles.submit();">Enable</a> 
    | 
    <a title="Select one or more profiles and click to toggle the enabled/disabled status" href="javascript:confirmSubmit('enable');document.profiles.submit();">Disable</a> |
    <a title="Select one or more profiles and click to run the profiles and perform the searches" href="javascript:confirmSubmit('run');document.profiles.submit();">Run</a> |
    <a href="search_profiles_car_ex.asp">Extended view</a> | Standard view
    | <a href="#top">return to top</a></font></td>
    <td width="50%" >
    <font size="2" >

    <% If intCurrentPage <= 1 Then %>
	    |&lt;
	    &lt; 
	<% Else  %>
<!--
	    <a href="<%=sPageURL %>?days_back=<%=intDaysBack %>&profile_desc=<%=strProfileDesc %>&city=<%=strCity%>&profile_car_type=<%=strProfileCarType %>&profile_car_co=<%=strProfileCarCo %>&profile_status=<%=intProfileStatus %>&page=1">|&lt;</a>
	    <a href="<%=sPageURL %>?days_back=<%=intDaysBack %>&profile_desc=<%=strProfileDesc %>&city=<%=strCity%>&profile_car_type=<%=strProfileCarType %>&profile_car_co=<%=strProfileCarCo %>&profile_status=<%=intProfileStatus %>&page=<%=intCurrentPage - 1%>">&lt;</a> 
-->	    
	    <a href="search_profiles_car.asp?days_back=<%=intDaysBack %>&profile_desc=<%=strProfileDesc %>&city=<%=strCity%>&profile_car_type=<%=strProfileCarType %>&profile_car_co=<%=strProfileCarCo %>&profile_status=<%=intProfileStatus %>&page=1">|&lt;</a>
	    <a href="search_profiles_car.asp?days_back=<%=intDaysBack %>&profile_desc=<%=strProfileDesc %>&city=<%=strCity%>&profile_car_type=<%=strProfileCarType %>&profile_car_co=<%=strProfileCarCo %>&profile_status=<%=intProfileStatus %>&page=<%=intCurrentPage - 1%>">&lt;</a> 
    <% End If %>
    Page <%=intCurrentPage %> of <%=intPageCount %>
    <% If CInt(intCurrentPage) >= CInt(intPageCount) Then %>
	    &gt;
	    &gt;|
    <% Else %>
<!--    
	    <a href="<%=sPageURL %>?days_back=<%=intDaysBack %>&profile_desc=<%=strProfileDesc %>&city=<%=strCity%>&profile_car_type=<%=strProfileCarType %>&profile_car_co=<%=strProfileCarCo %>&profile_status=<%=intProfileStatus %>&page=<%=intCurrentPage + 1%>">&gt;</a>
	    <a href="<%=sPageURL %>?days_back=<%=intDaysBack %>&profile_desc=<%=strProfileDesc %>&city=<%=strCity%>&profile_car_type=<%=strProfileCarType %>&profile_car_co=<%=strProfileCarCo %>&profile_status=<%=intProfileStatus %>&page=<%=intPageCount %>">&gt;|</a>
-->	   
	    <a href="search_profiles_car.asp?days_back=<%=intDaysBack %>&profile_desc=<%=strProfileDesc %>&city=<%=strCity%>&profile_car_type=<%=strProfileCarType %>&profile_car_co=<%=strProfileCarCo %>&profile_status=<%=intProfileStatus %>&page=<%=intCurrentPage + 1%>">&gt;</a>
	    <a href="search_profiles_car.asp?days_back=<%=intDaysBack %>&profile_desc=<%=strProfileDesc %>&city=<%=strCity%>&profile_car_type=<%=strProfileCarType %>&profile_car_co=<%=strProfileCarCo %>&profile_status=<%=intProfileStatus %>&page=<%=intPageCount %>">&gt;|</a>
    <% End If %>

    </font></td>
  </tr>
</table>
<form method="POST" action="search_profiles_maint_car.asp" name="maintenance">
<table border="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#111111" width="1110" id="AutoNumber1" background="images/alt_color.gif">
  <tr>
    <td width="11">&nbsp;</td>
    <td width="247">&nbsp;</td>
    <td width="715"><font size="2"><b>NOTE</b>: For the below actions, you may 
    only select one of the above profiles at a time</font></td>
    <td width="326">&nbsp;</td>
  </tr>
  <tr>
    <td width="11">&nbsp;</td>
    <td width="247"><img border="0" src="images/maintenance.GIF"></td>
    <td width="715">
    <input type="radio" name="maint_action" value="enable" checked id="enable" ><label for="enable">Enable Profile</label></td>
    <td width="326">&nbsp;</td>
  </tr>
  <tr>
    <td width="11">&nbsp;</td>
    <td width="247">&nbsp;</td>
    <td width="715">
    <input type="radio" name="maint_action" value="disable"  id="enable" ><label for="enable">Disable Profile</label></td>
    <td width="326">&nbsp;</td>
  </tr>
  <tr>
    <td width="11">&nbsp;</td>
    <td width="247">&nbsp;</td>
    <td width="715">
    <input type="radio" name="maint_action" value="delete" id="delete" ><label for="delete">Delete Profile</label></td>
    <td width="326">&nbsp;</td>
  </tr>
  <tr>
    <td width="11">&nbsp;</td>
    <td width="247">&nbsp;</td>
    <td width="715">
    <input type="radio" name="maint_action" value="copy" id="copy" ><label for="copy">Copy 
    Profile (same user)</label></td>
    <td width="326">&nbsp;</td>
  </tr>
  <tr>
    <td width="11">&nbsp;</td>
    <td width="247">&nbsp;</td>
    <td width="715">
    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <font size="2">New name: <input type="text" name="new_name" size="33" style="width: 250"></font></td>
    <td width="326">&nbsp;</td>
  </tr>
  <tr>
    <td width="11">&nbsp;</td>
    <td width="247">&nbsp;</td>
    <td width="715">
    <input type="radio" name="maint_action" value="rename" id="rename" ><label for="rename">Rename 
    Profile</label></td>
    <td width="326">&nbsp;</td>
  </tr>
  <tr>
    <td width="11">&nbsp;</td>
    <td width="247">&nbsp;</td>
    <td width="715">
    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <font size="2">New 
    name: <input type="text" name="new_name2" size="33" style="width: 250"></font></td>
    <td width="326">&nbsp;</td>
  </tr>

  <tr>
    <td width="11">&nbsp;</td>
    <td width="247">&nbsp;</td>
    <td width="715">
    <input type="radio" name="maint_action" value="copy_new_user" id="copy_new_user" ><label for="copy_new_user">Copy 
    Profile (to another user)</label></td>
    <td width="326">&nbsp;</td>
  </tr>
  <tr>
    <td width="11">&nbsp;</td>
    <td width="247">&nbsp;</td>
    <td width="715">
    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <font size="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
    User</font><font size="2" >: 
    <select size="1" name="users" style="width: 250">
    <% While adoUsers.EOF = False			%>
	    
	    <option value='<%=adoUsers.Fields("user_id").Value %>'><%=adoUsers.Fields("client_userid").Value %></option>
	    <% adoUsers.MoveNext %>

	<% Wend									%>
    </select></font></td>
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
    <td width="715"><font size="2" >
    <input type="submit" value="Submit" name="maint_submit" class="rh_button"></font></td>
    <td width="326">&nbsp;</td>
  </tr>
  <tr>
    <td width="11">&nbsp;</td>
    <td width="247">&nbsp;</td>
    <td width="715">&nbsp;</td>
    <td width="326">&nbsp;</td>
  </tr>
  </table>
  <table border="0" style="border-collapse: collapse" bordercolor="#FFFFFF" width="1110" cellspacing="0" height="4">
  <tr>
    <td background="images/ruler.gif"></td>
  </tr>
</table>
<table border="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#111111" width="1110" id="AutoNumber1" background="images/alt_color.gif">
  <tr>
    <td width="11">&nbsp;</td>
    <td width="247"><font point-size="9" size="2">
	1
<a href="search_profiles_car_excel_export.asp">Export 
profile list to Excel</a> </font></td>
    <td width="715"><font size="2"><strong>Link 1:</strong> is for standard 
	profiles</font></td>
    <td width="326"><font point-size="9" size="2">
 <a href="search_profiles_car_calc_export.asp">Export 
profile list for a rate calculator</a></font></td>
  </tr>
  <tr>
    <td width="11">&nbsp;</td>
    <td width="247"><font point-size="9" size="2">
	2
<a href="search_profiles_car_excel_export2.asp">Export 
profile list to Excel</a> </font></td>
    <td width="715"><font size="2"><strong>Link 2:</strong> is for standard 
	profiles with advanced scheduler information</font></td>
    <td width="326">&nbsp;</td>
  </tr>
  <tr>
    <td width="11">&nbsp;</td>
    <td width="247"><font point-size="9" size="2">
	3
<a href="search_profiles_car_excel_export3.asp">Export 
profile list to Excel</a> </font></td>
    <td width="715"><font size="2"><strong>Link 3:</strong> is for standard 
	profiles in the upload format</font></td>
    <td width="326"><font point-size="9" size="2">
<%if Session("org_id") = 148 then %><a href="search_profiles_locations_export.asp">Export 
profile list with locations</a> </font><%end if %></td>
  </tr>
  <tr>
    <td width="11">&nbsp;</td>
    <td width="247">&nbsp;</td>
    <td width="715"><font size="2"><strong>Directions:</strong> Right click the 
	link and select &quot;Save Target As&quot; to save the Excel file to your machine. If 
	you do not have Excel you can download an
	<a href="http://www.microsoft.com/downloads/details.aspx?FamilyID=c8378bf4-996c-4569-b547-75edbd03aaf0&amp;displaylang=EN">
	Excel viewer</a> from the Microsoft Web site.</font></td>
    <td width="326">&nbsp;</td>
  </tr>
</table>
<input type="hidden" name="profile_id" value="0">
</form>

<!--#INCLUDE FILE="footer.asp"-->	
</body>
</html>
<%		Set adoRS = Nothing %>