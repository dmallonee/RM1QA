<%@ Language=VBScript %>
<!-- #INCLUDE FILE="inc/login_check.asp" -->
<!-- #INCLUDE FILE="inc/adovbs.asp" -->
<% Response.Expires = -1
   Response.cachecontrol="private" 
   Response.AddHeader "pragma", "no-cache"

   Server.ScriptTimeout = 180


	'On Error Resume Next

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
	Dim strAlertDesc
	Dim datBeginDate


	strUserId = Request.Cookies("rate-monitor.com")("user_id")
	intRuleId = Request.QueryString("rateruleid")	
	
	strConn = Session("pro_con")
	
	Rem Get the data sources
	Set adoCmd1 = CreateObject("ADODB.Command")

	adoCmd1.ActiveConnection =  strConn
	adoCmd1.CommandText = "data_source_select"
	adoCmd1.CommandType = 4
		
	Set adoRS1 = adoCmd1.Execute

	Rem Get the vendors
	Set adoCmd2 = CreateObject("ADODB.Command")

	adoCmd2.ActiveConnection =  strConn
	adoCmd2.CommandText = "car_rate_rule_tester2"
	adoCmd2.CommandType = 4

	adoCmd2.Parameters.Append adoCmd2.CreateParameter("@vendor_cd",   200, 1,  2, Null)
	adoCmd2.Parameters.Append adoCmd2.CreateParameter("@vendor_name", 200, 1, 50, Null)
	adoCmd2.Parameters.Append adoCmd2.CreateParameter("@user_id",       3, 1,  0, strUserId)
	
	
	adoCmd2.Parameters.Append adoCmd2.CreateParameter("@situation_cd", 3, 1, 0)
	adoCmd2.Parameters.Append adoCmd2.CreateParameter("@rt_amt", 6, 1, 0)
	adoCmd2.Parameters.Append adoCmd2.CreateParameter("@situation_amt", 6, 1, 0)
	adoCmd2.Parameters.Append adoCmd2.CreateParameter("@max_rt_amt", 6, 1, 0)
	adoCmd2.Parameters.Append adoCmd2.CreateParameter("@min_rt_amt", 6, 1, 0)
	adoCmd2.Parameters.Append adoCmd2.CreateParameter("@is_dollar", 11, 1, 0)

		
	Set adoRS2 = adoCmd2.Execute
	
	Rem Get the vendors
	'Set adoCmd3 = CreateObject("ADODB.Command")

	'adoCmd3.ActiveConnection =  strConn
	'adoCmd3.CommandText = "vendor_select"
	'adoCmd3.CommandType = 4
		
	'Set adoRS3 = adoCmd3.Execute
	Set adoRS3 = adoCmd2.Execute


	Rem Get the vendors
	Set adoCmd4 = CreateObject("ADODB.Command")

	adoCmd4.ActiveConnection =  strConn
	adoCmd4.CommandText = "car_rate_rule_select"
	adoCmd4.CommandType = 4
	adoCmd4.Parameters.Append adoCmd4.CreateParameter("@rate_rule_id",      3, 1, 0, Null)
	adoCmd4.Parameters.Append adoCmd4.CreateParameter("@user_id",           3, 1, 0, strUserId)
	adoCmd4.Parameters.Append adoCmd4.CreateParameter("@rate_rule_type_cd", 3, 1, 0, 1)
	adoCmd4.Parameters.Append adoCmd4.CreateParameter("@include_disabled", 11, 1, 0, Null)
				
	Set adoRS4 = adoCmd4.Execute
	Set adoRS18a = adoCmd4.Execute
	Set adoRS18b = adoCmd4.Execute
			
	Rem Get the cities
	Set adoCmd6 = CreateObject("ADODB.Command")

	adoCmd6.ActiveConnection =  strConn
	adoCmd6.CommandText = "user_city_select"
	adoCmd6.CommandType = 4

	adoCmd6.Parameters.Append adoCmd6.CreateParameter("@user_id", 3, 1, 0, strUserId)
		
	Set adoRS6 = adoCmd6.Execute

	Rem Get the car types
	Set adoCmd7 = CreateObject("ADODB.Command")

	adoCmd7.ActiveConnection =  strConn
	adoCmd7.CommandText = "car_type_select"
	adoCmd7.CommandType = 4
	
	adoCmd7.Parameters.Append adoCmd7.CreateParameter("@user_id", 3, 1, 0, strUserId)
		
	Set adoRS7 = adoCmd7.Execute
	Set adoRS8 = adoCmd7.Execute
	
	Set adoCmd9 = CreateObject("ADODB.Command")

	adoCmd9.ActiveConnection =  strConn
	adoCmd9.CommandText = "car_shop_profile_select"
	adoCmd9.CommandType = 4

	adoCmd9.Parameters.Append adoCmd9.CreateParameter("@desc",              200, 1, 255)
	adoCmd9.Parameters.Append adoCmd9.CreateParameter("@shop_car_type_cds", 200, 1, 4096)
	adoCmd9.Parameters.Append adoCmd9.CreateParameter("@vend_cds",          200, 1, 1024)
	adoCmd9.Parameters.Append adoCmd9.CreateParameter("@user_id",             3, 1, 0, strUserId)
	adoCmd9.Parameters.Append adoCmd9.CreateParameter("@profile_id",          3, 1, 0, Null)
		
	Set adoRS9 = adoCmd9.Execute
	



	If intRuleId > 0 Then

		Rem Get the specific rule
		Set adoCmd5 = CreateObject("ADODB.Command")

		adoCmd5.ActiveConnection = strConn
		adoCmd5.CommandText = "car_rate_rule_select"
		adoCmd5.CommandType = 4
		
		adoCmd5.Parameters.Append adoCmd5.CreateParameter("@rate_rule_id", 3, 1, 0, intRuleId)
	
		Set adoRS5 = adoCmd5.Execute

		intRuleId = adoRS5.Fields("rate_rule_id").Value
		strAlertDesc = adoRS5.Fields("alert_desc").Value
		
		datChangeDate = adoRS5.Fields("change_dttm").Value
		
		blnRollingDates = adoRS5.Fields("rolling_date").Value
		
		Dim datHolder
		Dim intDateDiff
		
		If IsDate(adoRS5.Fields("begin_dt").Value) Then
			If blnRollingDates Then
				intDateDiff = DateDiff("d", datChangeDate , adoRS5.Fields("begin_dt").Value)
				datHolder = DateAdd("d", intDateDiff, Now)
				datBeginDate = FormatDateTime(datHolder, 2)
			Else
				datBeginDate = FormatDateTime(adoRS5.Fields("begin_dt").Value, 2)
			End If
		Else
			datBeginDate = "continuous"
		End If
		If IsDate(adoRS5.Fields("end_dt").Value) Then
			datEndDate = FormatDateTime(adoRS5.Fields("end_dt").Value, 2)
		Else
			datEndDate = "continuous"
		End If
		If IsDate(adoRS5.Fields("first_pickup_dt").Value) Then
			datFirstPickupDate = FormatDateTime(adoRS5.Fields("first_pickup_dt").Value, 2)
		Else
			datFirstPickupDate = "continuous"
		End If
		If IsDate(adoRS5.Fields("last_pickup_dt").Value) Then
			datSecondPickupDate = FormatDateTime(adoRS5.Fields("last_pickup_dt").Value, 2)
		Else
			datSecondPickupDate = "continuous"
		End If
		strClientRateCode = adoRS5.Fields("client_sys_rate_cd").Value
		strCityCd = adoRS5.Fields("city_cd").Value
		strVendCd = adoRS5.Fields("vend_cd1").Value
		strSelfCd = adoRS5.Fields("vend_cd2").Value
		intLOR = adoRS5.Fields("lor").Value
		strCarTypeCd1 = adoRS5.Fields("car_type_cd1").Value
		strCarTypeCd2 = adoRS5.Fields("car_type_cd2").Value
		strDataSource = adoRS5.Fields("data_source").Value
		intSituationCd = adoRS5.Fields("situation_cd").Value
		strSituationAmt = adoRS5.Fields("situation_amt").Value
		blnIsDollar = adoRS5.Fields("is_dollar").Value
		strQuantityPeriodCd =  adoRS5.Fields("quantity_period_cd").Value
		intEventCount =  adoRS5.Fields("event_count").Value
		intEventTimeCount =  adoRS5.Fields("event_time_count").Value
		strEventTimeType = adoRS5.Fields("event_time_type").Value
		intResponseCd = adoRS5.Fields("response_cd").Value
		strResponseAmt = adoRS5.Fields("response_amt").Value
		blnResponseDollar = adoRS5.Fields("is_response_dollar").Value
		strRangeMax = adoRS5.Fields("range_max").Value
		strRangeMin = adoRS5.Fields("range_min").Value
		intSuccessId = adoRS5.Fields("on_success_id").Value
		intFailureId = adoRS5.Fields("on_failure_id").Value
		strProfileId = adoRS5.Fields("profile_id").Value
		intSearchProfile = adoRS5.Fields("search_profile").Value
		intUtilizationProfileId = adoRS5.Fields("utilization_profile_id").Value
		strUtilMax0 = adoRS5.Fields("util_max_0").Value & ""
		strUtilMin0 = adoRS5.Fields("util_min_0").Value & ""
		strUtilMax1 = adoRS5.Fields("util_max_1").Value & ""
		strUtilMin1 = adoRS5.Fields("util_min_1").Value & ""
		strUtilMax2 = adoRS5.Fields("util_max_2").Value & ""
		strUtilMin2 = adoRS5.Fields("util_min_2").Value & ""
		strUtilMax3 = adoRS5.Fields("util_max_3").Value & ""
		strUtilMin3 = adoRS5.Fields("util_min_3").Value & ""
		strUtilMax4 = adoRS5.Fields("util_max_4").Value & ""
		strUtilMin4 = adoRS5.Fields("util_min_4").Value & ""
		strUtilMax5 = adoRS5.Fields("util_max_5").Value & ""
		strUtilMin5 = adoRS5.Fields("util_min_5").Value & ""
		strUtilMax6 = adoRS5.Fields("util_max_6").Value & ""
		strUtilMin6 = adoRS5.Fields("util_min_6").Value & ""
		strUtilMax7 = adoRS5.Fields("util_max_7").Value & ""
		strUtilMin7 = adoRS5.Fields("util_min_7").Value & ""
		blnIgnoreClosed = adoRS5.Fields("ignore_closed").Value
		
		strButton = "Update"


	Else
	
		intRuleId = "(New - one will be assigned)"
		strAlertDesc = ""
		datBeginDate = "continuous"
		datEndDate = "continuous"
		datFirstPickupDate = "continuous"
		datSecondPickupDate = "continuous"
		strClientRateCode = ""
		strCityCd = ""
		strVendCd = ""
		strSelfCd = ""
		intLOR = "1"
		strCarTypeCd1 = "CCAR"
		strCarTypeCd2 = "CCAR"
		strDataSource = "ORB"
		intSituationCd = 0
		strSituationAmt = ""
		blnIsDollar = 1
		strQuantityPeriodCd  = 0
		intEventCount =  0
		intEventTimeCount =  0
		strEventTimeType = ""
		intResponseCd = 0
		strResponseAmt = 0
		blnResponseDollar = True
		strRangeMax = ""
		strRangeMin = ""
		intSuccessId = 0
		intFailureId = 0
		strProfileId = "XX"
		intSearchProfile = 1
		intUtilizationProfileId = 0
		strUtilMax0 = ""
		strUtilMin0 = ""
		strUtilMax1 = ""
		strUtilMin1 = ""
		strUtilMax2 = ""
		strUtilMin2 = ""
		strUtilMax3 = ""
		strUtilMin3 = ""
		strUtilMax4 = ""
		strUtilMin4 = ""
		strUtilMax5 = ""
		strUtilMin5 = ""
		strUtilMax6 = ""
		strUtilMin6 = ""
		strUtilMax7 = ""
		strUtilMin7 = ""
		blnIgnoreClosed = True
		
		
		strButton = "Create"

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



<!doctype HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>

<head>
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Language" content="en-us">
<title>Rate-Monitor by Rate-Highway, Inc. | Alerts! | Rate Management</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<script language="JavaScript" type="text/JavaScript">
<!--

function CreateAlert()
{ 
	var valid_form = true;
	var numSelected = 0;
	var i;
	
	//alert("testing");
	
	if (document.create_alert.alert_desc.value == '') 
		{
		alert("Please select a descriptve name for your rule.");  
		document.create_alert.alert_desc.focus();
		valid_form = false;
		}

	if (valid_form)
		{
		if((document.create_alert.vend_cd1.value == "ANY") && (document.create_alert.vend_cd1.value == document.create_alert.vend_cd2.value))
			{
			alert("You may not use ANY for both companies. Please input valid company. \n Please only use ANY for the " );  
			document.create_alert.vend_cd2.focus();
			valid_form = false ;
			}
		}

	if (valid_form)
		{
		numSelected = 0;

		// check if less than 1 options are selected
		for (i = 0;  i < document.create_alert.vend_cd1.length;  i++)
		{
		if (document.create_alert.vend_cd1.options[i].selected)
			numSelected++;
		}
		if (numSelected < 1)
			{
			alert("Please select at least one car company to create the competitive set");
			document.create_alert.vend_cd1.focus();
			valid_form = false;
			}
		}

	if (valid_form)
		{
		numSelected = 0;

		for (i = 0;  i < document.create_alert.vend_cd2.length;  i++)
		{
		if (document.create_alert.vend_cd2.options[i].selected)
			numSelected++;
		}
		if (numSelected < 1)
			{
			alert("Please select one car company to compare against the competitive set");
			document.create_alert.vend_cd2.focus();
			valid_form = false;
			}

		}

	if (valid_form)
		{
		numSelected = 0;

		for (i = 0;  i < document.create_alert.city_cd.length;  i++)
		{
		if (document.create_alert.city_cd.options[i].selected)
			numSelected++;
		}
		if (numSelected < 1)
			{
			alert("Please select at least one city for this rule to apply to. \n It can be one city or all of them or any number in-between.");
			document.create_alert.city_cd.focus();
			valid_form = false;
			}
		
		}

	if (valid_form) {
		document.create_alert.action = "car_rate_rule_insert.asp";
		document.create_alert.submit();
		return true;
		}
	else {
		return false;
		}
				
}

function CheckResponse()
{

	if (document.create_alert.response_type.selectedIndex != 0)
	{
		document.create_alert.response_amount.disabled = false;
	}
	else if (document.create_alert.response_type.selectedIndex == 0)
	{
		document.create_alert.response_amount.disabled = true;
		document.create_alert.response_amount.value = '';
		
	}

}

function maint_action(action_id)
{
	document.maint.action.value = action_id;
	document.maint.submit();

}

//-->
</script>

<style>
<!--
.profile_header { height="48" text-align:left; padding-left:3; padding-right:3; padding-top:0; background-color:#879AA2; font-family:Verdana, Arial, Helvetica, sans-serif; font-size:10pt; vertical-align:bottom; text-align:left }
.profile_light { height="68" text-align:left; padding-left:3; padding-right:3; padding-top:3; background-color:#CFD7DB; font-family:Verdana, Arial, Helvetica, sans-serif; font-size:10pt; vertical-align:bottom; text-align:left }
.profile_dark { height="48" text-align:left; padding-left:3; padding-right:3; padding-top:3; background-color:#B2BEC4; font-family:Verdana, Arial, Helvetica, sans-serif; font-size:10pt; vertical-align:bottom; text-align:left }
-->
</style>
<link rel="stylesheet" type="text/css" href="inc/rh_standard.css">
<link rel="stylesheet" type="text/css" href="inc/rh_report.css">
<link rel="stylesheet" type="text/css" href="inc/sitewide.css">
<link rel="stylesheet" type="text/css" href="inc/css_calendar_v2.css" id="calendarcss" media="all" >
<script language="Javascript" type="text/javascript" src="inc/sitewide.js"></script>
<script language="Javascript" type="text/javascript" src="inc/header_menu_support.js" ></script>
<script language="Javascript" type="text/javascript" src="inc/js_calendar_v2.js"></script>	
</head>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="MM_preloadImages('images/b_search_pro_on.gif','images/b_search_que_on.gif','images/b_search_cri_on.gif','images/b_rate_on.gif','images/b_alert_on.gif','images/b_user_on.gif','images/b_system_on.gif')">

<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td background="images/top_tile.gif">
    <img src="images/top.jpg" width="770" height="91"></td>
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

<p align="right">&nbsp;</p>
<div align="center">
<table cellpadding="0" cellspacing="0" border="0" width="1310" bgcolor="#FFFFFF">
<tr height="1">
<td colspan="1" width="1">&nbsp;</td>
<td rowspan="2" width="169"><img src="images/ratemanagementalerts2_a.JPG" width="169" height="25" hspace="0" vspace="0" border="0" alt="Rate Management" description=""></td>
<td colspan="1" >&nbsp;</td>
</tr>
<tr height="1">
<td bgcolor="#000000" colspan="1" height="1"><img src=pixel.gif width="1" height="1"></td>
<td bgcolor="#000000" colspan="1" height="1"><img src=pixel.gif width="1" height="1"></td>
<td bgcolor="#000000" colspan="1" height="1"><img src=pixel.gif width="1" height="1"></td>
</tr>
</table>
</div>
<table cellpadding="0" cellspacing="0" border="0" ALIGN="CENTER" width="100%" bgcolor="#FFFFFF">
<tr >
<td  width="1" bgcolor="#000000"><img src=pixel.gif width="1" height="1"></td>
<td colspan=3 bgcolor="#D9DEE1">
<table border="0" cellspacing="5" cellpadding="5">
<tr><td>
<font color="#080000">
<P>
<!-- JUSTTABS TOP OPEN-END -->
&nbsp;
<form method="POST" action="searched_alerts.asp" name="search_alerts" class="search">
<table border="0" style="border-collapse: collapse" bordercolor="#FFFFFF" width="100%" cellspacing="0" height="4">
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
 
  <table width="745" border="0" cellspacing="0" cellpadding="0" background="images/alt_color.gif">
    <tr>
      <td>
      <table width="1108" border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" background="images/alt_color.gif">
        <tr valign="bottom">
          <td width="10" height="51">&nbsp;</td>
          <td width="179" height="51">
          Testing a Situation</td>
          <td width="583" colspan="3" height="51">
          <p><font size="2" face="Verdana, Arial, Helvetica, sans-serif">To search 
          for an Alert, enter a login id, or a portion of the 
          address. You may also enter the alert type.</font></td>
          <td width="336" height="51">&nbsp;</td>
        </tr>
        <tr valign="bottom">
          <td width="10" height="26">&nbsp;</td>
          <td width="179" height="26">&nbsp;</td>
          <td width="177" height="26">
          <font face="Verdana, Arial, Helvetica, sans-serif" size="2">Owner:</font> </td>
          <td width="80" height="26">
          <font size="2" face="Verdana, Arial, Helvetica, sans-serif">
          <input type="text" name="name" size="20" style="width:150" style="width:150" onfocus="this.className='focus';cl(this,'<%=strClientUserid %>');" onblur="this.className='';fl(this,'<%=strClientUserid %>');"></font></td>
          <td width="662" colspan="2" height="26">
          <font size="2" face="Verdana, Arial, Helvetica, sans-serif">
          <input name="search" type="submit" id="Open2224" value="    Search    " class="rh_button"></font></td>
        </tr>
        <tr valign="bottom">
          <td width="10" height="22">&nbsp;</td>
          <td width="179" height="22">&nbsp;</td>
          <td width="177" height="22">
          <font size="2" face="Verdana, Arial, Helvetica, sans-serif">Alert 
          Type:</font></td>
          <td width="80" height="22">
          <select size="1" name="type" style="border:1px solid #000000; width:150; background-color:#FF9933">
          <option selected value="0">Any type</option>
          <option value="4">Rate Mgmt</option>
          <option value="1">Email notification</option>
          <option value="2">Pager notification</option>
          <option value="3">Login list</option>
          </select></td>
          <td width="662" colspan="2" height="22">&nbsp;</td>
        </tr>
        <tr valign="bottom">
          <td width="10" height="22">&nbsp;</td>
          <td width="179" height="22">&nbsp;</td>
          <td width="177" height="22">&nbsp;</td>
          <td width="80" height="22">&nbsp;</td>
          <td width="662" colspan="2" height="22">&nbsp;</td>
        </tr>
      </table>
      </td>
    </tr>
  </table>
  </form>
  <table width="1108" border="0" cellpadding="2" cellspacing="0" background="images/alt_color.gif" style="border-collapse: collapse" bordercolor="#111111">
    </table>
<form name="create_alert" method="POST" action="" OnSubmit="return CreateAlert()"  >
  <table border="0" style="border-collapse: collapse" bordercolor="#FFFFFF" width="1310" height="4">
    <tr>
      <td width="8" height="25">&nbsp;</td>
      <td width="217" height="25">
      &nbsp;</td>
      <td width="210" height="25" style="padding-top: 2px">
      &nbsp;</td>
      <td width="510" height="25" colspan="8">
      &nbsp;</td>
      <td width="262" height="25">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="25">&nbsp;</td>
      <td width="217" height="25">
      &nbsp;</td>
      <td width="210" height="25" style="vertical-align: top; padding-top: 2px">
      &nbsp;</td>
      <td width="510" height="25" colspan="8">
      &nbsp;</td>
      <td width="262" height="25">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="19">&nbsp;</td>
      <td width="217" height="19">&nbsp;</td>
      <td width="210" height="19">&nbsp;</td>
      <td width="510" height="19" colspan="8">&nbsp;</td>
      <td width="262" height="19">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="22">&nbsp;</td>
      <td width="217" height="22">&nbsp;</td>
      <td width="210" height="22"><font face="Verdana" size="2">13. Situation:</font></td>
      <td width="510" height="22" colspan="8">
      <select size="1" name="situation_cd" style="width:370; font-family:Verdana; font-size:10pt; height:24">
      
      <% If intSituationCd = 0 Then	%>
	      <option selected value="0">(None selected)</option>
	  <% Else %>
	      <option value="0">(None selected)</option>
	  <% End If %>
 
      <% If intSituationCd = 1 Then	%>
	      <option selected value="1">NONE - Set rate to the response amount</option>
	  <% Else %>
	      <option value="1">NONE - Set rate to the response amount</option>
	  <% End If %>
	  <!-- 	  
      <% If intSituationCd = 2 Then	%>
	      <option selected value="2">> (any competitor)</option>
	  <% Else %>
	      <option value="2">> (any competitor)</option>
	  <% End If %>
	  
      <% If intSituationCd = 3 Then	%>
	      <option selected value="3">> (custom)</option>
	  <% Else %>
	      <option value="3">> (custom)</option>
	  <% End If %>
	  -->
      <% If intSituationCd = 4 Then	%>
	      <option selected value="4">If rate is not more than (all competitors) by at least</option>
	  <% Else %>
	      <option value="4">If rate is not more than (all competitors) by at least</option>
	  <% End If %>
	  
      <% If intSituationCd = 5 Then	%>
	      <option selected value="5">If rate is not more than (any competitors) by at least</option>
	  <% Else %>
	      <option value="5">If rate is not more than (any competitors) by at least</option>
	  <% End If %>
	  
      <% If intSituationCd = 6 Then	%>
	      <option selected value="6">If rate is not more than (custom) by at least</option>
	  <% Else %>
	      <option value="6">If rate is not more than (custom) by at least</option>
	  <% End If %>
	  
      <% If intSituationCd = 7 Then	%>
	      <option selected value="7">If rate is equal to (all competitors)</option>
	  <% Else %>
	      <option value="7">If rate is equal to (all competitors)</option>
	  <% End If %>

      <% If intSituationCd = 8 Then	%>
	      <option selected value="8">If rate is equal to (any competitor)</option>
	  <% Else %>
	      <option value="8">If rate is equal to (any competitor)</option>
	  <% End If %>

      <% If intSituationCd = 9 Then	%>
	      <option selected value="9">If rate is equal to (custom)</option>
	  <% Else %>
	      <option value="9">If rate is equal to (custom)</option>
	  <% End If %>
	  <!--
      <% If intSituationCd = 11 Then	%>
	      <option selected value="11">< = (all competitors)</option>
	  <% Else %>
	      <option value="11">< = (all competitors)</option>
	  <% End If %>

      <% If intSituationCd = 12 Then	%>
	      <option selected value="12"><= (any competitor)</option>
	  <% Else %>
	      <option value="12"><= (any competitor)</option>
	  <% End If %>
	  
      <% If intSituationCd = 13 Then	%>
	      <option selected value="13">&lt;= (custom)</option>
	  <% Else %>
	      <option value="13">&lt;= (custom)</option>
	  <% End If %>
	  -->
      <% If intSituationCd = 14 Then	%>
	      <option selected value="14">If rate is not less than (all competitors) by at least</option>
	  <% Else %>
	      <option value="14">If rate is not less than (all competitors) by at least</option>
	  <% End If %>
	  
      <% If intSituationCd = 16 Then	%>
	      <option selected value="16">If rate is not less than (any competitor) by at least</option>
	  <% Else %>
	      <option value="16">If rate is not less than (any competitor) by at least</option>
	  <% End If %>

      <% If intSituationCd = 15 Then	%>
	      <option selected value="15">If rate is not less than (custom) by at least</option>
	  <% Else %>
	      <option value="15">If rate is not less than (custom) by at least</option>
	  <% End If %>
	  
      <% If intSituationCd = 17 Then	%>
	      <option selected value="17">If rate is not equal to (any competitor)</option>
	  <% Else %>
	      <option value="17">If rate is not equal to (any competitor)</option>
	  <% End If %>
	  
      <% If intSituationCd = 18 Then	%>
	      <option selected value="18">If rate is not equal to (all competitors)</option>
	  <% Else %>
	      <option value="18">If rate is not equal to (all competitors)</option>
	  <% End If %>
	  
      <% If intSituationCd = 19 Then	%>
	      <option selected value="19">If rate is not equal to (custom)</option>
	  <% Else %>
	      <option value="19">If rate is not equal to (custom)</option>
	  <% End If %>
	  <!--
      <% If intSituationCd = 20 Then	%>
	      <option selected value="20">If rate is not equal to  (all competitors) + diff</option>
	  <% Else %>
	      <option value="20">If rate is not equal to  (all competitors) + diff</option>
	  <% End If %>
 	  -->
      <% If intSituationCd = 30 Then	%>
	      <option selected value="30">If the diff. between (all competitors) is at least</option>
	  <% Else %>
	      <option value="30">If the diff. between (all competitors) is at least</option>
	  <% End If %>
	  
      <% If intSituationCd = 31 Then	%>
	      <option selected value="31">If the diff. between (all competitors) is less than</option>
	  <% Else %>
	      <option value="31">If the diff. between (all competitors) is less than</option>
	  <% End If %>

      <% If intSituationCd = 32 Then	%>
	      <option selected value="32">If the diff. between (any competitor) is at least</option>
	  <% Else %>
	      <option value="32">If the diff. between (any competitor) is at least</option>
	  <% End If %>
	  
      <% If intSituationCd = 33 Then	%>
	      <option selected value="33">If the diff. between (any competitor) is less than</option>
	  <% Else %>
	      <option value="33">If the diff. between (any competitor) is less than</option>
	  <% End If %>

      <% If intSituationCd = 34 Then	%>
	      <option selected value="34">If rate is not less than (all competitors) by exactly</option>
	  <% Else %>
	      <option value="34">If rate is not less than (all competitors) by exactly</option>
	  <% End If %>
	  
      <% If intSituationCd = 35 Then	%>
	      <option selected value="35">If rate is not less than (any competitor) by exactly</option>
	  <% Else %>
	      <option value="35">If rate is not less than (any competitor) by exactly</option>
	  <% End If %>

      <% If intSituationCd = 40 Then	%>
	      <option selected value="40">If (any competitor) rate is less than</option>
	  <% Else %>
	      <option value="40">If (any competitor) rate is less than</option>
	  <% End If %>

      <% If intSituationCd = 41 Then	%>
	      <option selected value="41">If (all competitor) rates are less than</option>
	  <% Else %>
	      <option value="41">If (all competitor) rate are less than</option>
	  <% End If %>

      <% If intSituationCd = 42 Then	%>
	      <option selected value="42">If (any competitor) rate is equal to</option>
	  <% Else %>
	      <option value="42">If (any competitor) rate is equal to</option>
	  <% End If %>

      <% If intSituationCd = 43 Then	%>
	      <option selected value="43">If (all competitor) rates are equal to</option>
	  <% Else %>
	      <option value="43">If (all competitor) rate are equal to</option>
	  <% End If %>

      <% If intSituationCd = 44 Then	%>
	      <option selected value="44">If (any competitor) rate is greater than</option>
	  <% Else %>
	      <option value="44">If (any competitor) rate is greater than</option>
	  <% End If %>

      <% If intSituationCd = 45 Then	%>
	      <option selected value="45">If (all competitor) rates are greater than</option>
	  <% Else %>
	      <option value="45">If (all competitor) rate are greater than</option>
	  <% End If %>

      </select></td>
      <td width="262" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="23">&nbsp;</td>
      <td width="217" height="23">&nbsp;</td>
      <td width="210" height="23"><font face="Verdana" size="2">&nbsp;&nbsp;&nbsp; 
      by amount:<br>
&nbsp;&nbsp;&nbsp; <br>
      </font></td>
      <td width="510" height="23" colspan="8">
      <input type="text" name="situation_amt" size="20" style="width:127; font-family:Verdana; font-size:10pt; text-align:right; height:22" value="<%=strSituationAmt %>">
      <br>
      <% If blnIsDollar Then %>
      <input type="radio" value="1" name="is_dollar" style="font-family:Verdana; font-size:10pt" checked id="is_dollar1">    
      <font face="Verdana" size="2"><label for="is_dollar1">Whole Dollar</label><br>
      <input type="radio" value="0" name="is_dollar" style="font-family:Verdana; font-size:10pt" id="is_dollar2"><font face="Verdana" size="2"><label for="is_dollar2">Percentage</label><br>
      <% Else %>
      <input type="radio" value="1" name="is_dollar" style="font-family:Verdana; font-size:10pt" id="is_dollar1">    
      <font face="Verdana" size="2"><label for="is_dollar1">Whole Dollar</label><br>
      <input type="radio" value="0" name="is_dollar" style="font-family:Verdana; font-size:10pt" checked  id="is_dollar2"><font face="Verdana" size="2"><label for="is_dollar2">Percentage</label><br>
      <% End If %>
      <% If blnIgnoreClosed Then %>
      <input type="checkbox" name="ignore_closed" value="True" id="ignore_closed" checked><font face="Verdana" size="2"><label for="ignore_closed">Ignore closed rates</label>
	  <% Else %>
      <input type="checkbox" name="ignore_closed" value="True" id="ignore_closed" ><font face="Verdana" size="2"><label for="ignore_closed">Ignore closed rates</label>
	  <% End If %>

      <input type="checkbox" name="totalprice" value="True" disabled >Use &quot;Total Price&quot; 
      vs. base rate</td>
      
      
      <td width="262" height="23">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="19">&nbsp;</td>
      <td width="217" height="19">&nbsp;</td>
      <td width="210" height="19">&nbsp;</td>
      <td width="510" height="19" colspan="8">&nbsp;</td>
      <td width="262" height="19">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="22">&nbsp;</td>
      <td width="217" height="22">&nbsp;</td>
      <td width="210" height="22"><font face="Verdana" size="2">14. Quantity &amp; 
      Period:</font></td>
      <td width="510" height="22" colspan="8">
      <select size="1" name="quantity_period_cd" style="width:200; font-family:Verdana; font-size:10pt">
      <% If intQuantityPeriodCd = 0 Then %>
	      <option selected value="0">Each time it occurs</option>
	  <% Else %>
	      <option value="0">Each time it occurs</option>
	  <% End If %>
      
      <% If intQuantityPeriodCd = 1 Then %>
	      <option selected  value="1">After X Events</option>
	  <% Else %>
	      <option value="1">After X Events</option>
	  <% End If %>
	  
      <% If intQuantityPeriodCd = 2 Then %>
	      <option selected value="2">After X Events in X Hours</option>
	  <% Else %>
	      <option value="2">After X Events in X Hours</option>
	  <% End If %>

      </select></td>
      <td width="262" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="22">&nbsp;</td>
      <td width="217" height="22">&nbsp;</td>
      <td width="210" height="22"><font face="Verdana" size="2">&nbsp;&nbsp;&nbsp; 
      No. of Events:</font></td>
      <td width="510" height="22" colspan="8">
      <input type="text" name="event_count" size="20" style="width:200; font-family:Verdana; font-size:10pt; text-align:right" value="<%=intEventCount %>"></td>
      <td width="262" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="22">&nbsp;</td>
      <td width="217" height="22">&nbsp;</td>
      <td width="210" height="22"><font face="Verdana" size="2">&nbsp;&nbsp;&nbsp; 
      That happen in:</font></td>
      <td width="510" height="22" colspan="8">
      <input type="text" name="event_time_count" size="20" style="width:137; font-family:Verdana; font-size:10pt; text-align:right; height:22" value="<%=intEventTimeCount %>">
      <select size="1" name="event_time_type">
      <% If strEventTimeType  = "Hours" Or strEventTimeType = "" Then %>
      <option selected>Hours</option>
      <option>Days</option>
	  <% Else %>
      <option>Hours</option>
      <option selected >Days</option>
	  <% End If %>

      </select></td>
      <td width="262" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="19">&nbsp;</td>
      <td width="217" height="19">&nbsp;</td>
      <td width="210" height="19">&nbsp;</td>
      <td width="510" height="19" colspan="8">&nbsp;</td>
      <td width="262" height="19">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="22">&nbsp;</td>
      <td width="217" height="22">&nbsp;</td>
      <td width="210" height="22"><font face="Verdana" size="2">15. Response:</font></td>
      <td width="510" height="22" colspan="8">
      <select size="1" name="response_cd" style="width:370; font-family:Verdana; font-size:10pt; height:24" >
 <!--     <select size="1" name="response_cd" style="width:200; font-family:Verdana; font-size:10pt" onclick="CheckResponse()" >
 -->
	  <% If intResponseCd = 0 Then	%>
      <option selected value="0">(None selected)</option>
	  <% Else						%>
      <option value="0">(None selected)</option>
	  <% End If 					%>	  
      
	  <% If intResponseCd = 1 Then	%>
      <option selected value="1">Set my rate to lowest competitor's rate minus $</option>
	  <% Else						%>
      <option value="1">Set my rate to lowest competitor's rate minus $</option>
	  <% End If 					%>	  

	  <% If intResponseCd = 2 Then	%>
      <option selected value="2">Set my rate to lowest competitor's rate plus $</option>
	  <% Else						%>
      <option value="2">Set my rate to lowest competitor's rate plus $</option>
	  <% End If 					%>	  

	  <% If intResponseCd = 4 Then	%>
      <option selected value="4">Set my rate to lowest competitor's rate minus %</option>
	  <% Else						%>
      <option value="4">Set my rate to lowest competitor's rate minus %</option>
	  <% End If 					%>	  

	  <% If intResponseCd = 5 Then	%>
      <option selected value="5">Set my rate to lowest competitor's rate plus %</option>
	  <% Else						%>
      <option value="5">Set my rate to lowest competitor's rate plus %</option>
	  <% End If 					%>	  

	  <% If intResponseCd = 3 Then	%>
      <option selected value="3">Set my rate to amount</option>
	  <% Else						%>
      <option value="3">Set my rate to amount</option>
	  <% End If 					%>	  

	  <% If intResponseCd = 7 Then	%>
      <option selected value="7">Set my rate to highest competitor's rate minus $</option>
	  <% Else						%>
      <option value="7">Set my rate to highest competitor's rate minus $</option>
	  <% End If 					%>	  

	  <% If intResponseCd = 6 Then	%>
      <option selected value="6">Set my rate to highest competitor's rate plus $</option>
	  <% Else						%>
      <option value="6">Set my rate to highest competitor's rate plus $</option>
	  <% End If 					%>	  
  
	  <% If intResponseCd = 8 Then	%>
      <option selected value="8">Set my rate to highest competitor's rate minus %</option>
	  <% Else						%>
      <option value="8">Set my rate to highest competitor's rate minus %</option>
	  <% End If 					%>	  

	  <% If intResponseCd = 9 Then	%>
      <option selected value="9">Set my rate to highest competitor's rate plus %</option>
	  <% Else						%>
      <option value="9">Set my rate to highest competitor's rate plus %</option>
	  <% End If 					%>	  



	  <% If intResponseCd = 10 Then	%>
      <option selected value="10">Set my rate to competitor avg rate minus $</option>
	  <% Else						%>
      <option value="10">Set my rate to competitor avg rate minus $</option>
	  <% End If 					%>	  

	  <% If intResponseCd = 11 Then	%>
      <option selected value="11">Set my rate to competitor avg rate plus $</option>
	  <% Else						%>
      <option value="11">Set my rate to competitor avg rate plus $</option>
	  <% End If 					%>	  
  
	  <% If intResponseCd = 12 Then	%>
      <option selected value="12">Set my rate to competitor avg rate minus %</option>
	  <% Else						%>
      <option value="12">Set my rate to competitor avg rate minus %</option>
	  <% End If 					%>	  

	  <% If intResponseCd = 13 Then	%>
      <option selected value="13">Set my rate to competitor avg rate plus %</option>
	  <% Else						%>
      <option value="13">Set my rate to competitor avg rate plus %</option>
	  <% End If 					%>	  




      </select>
      
      </td>
      <td width="262" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="19">&nbsp;</td>
      <td width="217" height="19">&nbsp;</td>
      <td width="210" height="19"><font face="Verdana" size="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Amount:<br>
      <br></font></td>
      <td width="510" height="19" colspan="8">
      <input type="text" name="response_amt" size="20" style="width:127; font-family:Verdana; font-size:10pt; text-align:right; height:22" value="<%=strResponseAmt %>" >
      <br>
      <% If blnResponseDollar Then %>
	      <input type="radio" value="1" name="is_response_dollar" id="is_response_dollar1" style="font-family:Verdana; font-size:10pt" checked><font face="Verdana" size="2"><label for="is_response_dollar1">Whole Dollar</label><br>
	      <input type="radio" value="0" name="is_response_dollar" id="is_response_dollar2"  style="font-family:Verdana; font-size:10pt"><label for="is_response_dollar2">Percentage</label> </td>
	  <% Else					   %>
	      <input type="radio" value="1" name="is_response_dollar" id="is_response_dollar1" style="font-family:Verdana; font-size:10pt" ><font face="Verdana" size="2"><label for="is_response_dollar1">Whole Dollar</label><br>
	      <input type="radio" value="0" name="is_response_dollar" id="is_response_dollar2" style="font-family:Verdana; font-size:10pt" checked><label for="is_response_dollar2">Percentage</label></td>
	  <% End If					   %>
	  
	  
      <td width="262" height="19">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="19">&nbsp;</td>
      <td width="217" height="19">&nbsp;</td>
      <td width="210" height="19">&nbsp;</td>
      <td width="510" height="19" colspan="8">&nbsp;</td>
      <td width="262" height="19">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="22">&nbsp;</td>
      <td width="217" height="22">&nbsp;</td>
      <td width="210" height="22"><font face="Verdana" size="2">16. Search Type:</font></td>
      <td width="510" height="22" colspan="8">
      <select size="1" name="search_profile" style="width:200; font-family:Verdana; font-size:10pt" >
      <!--
      <select size="1" name="search_profile" style="width:200; font-family:Verdana; font-size:10pt" onchange="showhide('profiles');">
		-->
      <% If intSearchProfile = 1 Then	%>
      <option selected value="1">Link to Profile(s)</option>
      <option value="2">General Rule - Not Linked</option>
      <% End If						%>

      <% If intSearchProfile = 2 Then	%>
      <option value="1">Link to Profile(s)</option>
      <option selected value="2">General Rule - Not Linked</option>
      <% End If						%>

<!--      
      <% If intSearchProfile = 0 Then	%>
      <option selected value="0">As searched (all searches)</option>
      <option value="1">Link to Profile</option>
      <% Else						%>
      <option value="0">As searched (all searches)</option>
      <option selected value="1">Link to Profile</option>
      <% End If						%>
-->      
      </select></td>
      <td width="262" height="22">&nbsp;</td>
    </tr>
	<div id="profiles" style="display: none;"> 
    <tr>
      <td width="8" height="22">&nbsp;</td>
      <td width="217" height="22">&nbsp;</td>
      <td width="210" height="22"><font face="Verdana" size="2">&nbsp;&nbsp;&nbsp;&nbsp; 
      Profile(s):<br>
      <br></font></td>
      <td width="510" height="22" colspan="8">
		<select name="profile_id" size="4" multiple style="width:373; font-family:Verdana; font-size:10pt; height:70">
                <% While adoRS9.EOF = False %> 
                	<% If adoRS9.Fields("profile_id").Value = 0 Then %>
	                	<% adoRS9.MoveNext %>
	                <% end If %>
                	<% If InStr(1, strProfileID, adoRS9.Fields("profile_id").Value) > 0 Then %>
	                	<option selected value="<%=adoRS9.Fields("profile_id").Value %>"><%=adoRS9.Fields("desc").Value %></option>
	                <% Else %>
		               <% If adoRS9.Fields("profile_status").Value = "E" Then %>
		                	<option value="<%=adoRS9.Fields("profile_id").Value %>"><%=adoRS9.Fields("desc").Value %></option>
		                <% End If %> 
					<% End If %>                
                <%	adoRS9.MoveNext
				   Wend
					   
				   Set adoRS9 = Nothing
				%>
				</select></td>
      <td width="262" height="22">&nbsp;</td>
    </tr>
    </div>
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
      <td width="210" height="22">
      <font face="Verdana" size="2">17. Rate Range:</font></td>
      <td width="510" height="22" colspan="8"><font face="Verdana" size="2">
      (leave blank or enter a zero to indicate no limit)</font></td>
      <td width="262" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="22">&nbsp;</td>
      <td width="217" height="22">&nbsp;</td>
      <td width="210" height="22"><font face="Verdana" size="2">&nbsp;&nbsp;&nbsp;&nbsp; 
      Maximum:</font></td>
      <td width="510" height="22" colspan="8">
      <input type="text" name="rate_maximum" size="20" style="width:127; font-family:Verdana; font-size:10pt; text-align:right; height:21" value="<%=strRangeMax %>" 
      onkeyup='this.onchange();' onchange='tV=(this.value.replace(/[^\d\.]/g,"")).replace(/[\.]{2,}/g,".");if(tV!=this.value){this.value=tV;}'></td>
      <td width="262" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="22">&nbsp;</td>
      <td width="217" height="22">
      &nbsp;</td>
      <td width="210" height="25"><font face="Verdana" size="2">&nbsp;&nbsp;&nbsp;&nbsp; 
      Minimum:</font></td>
      <td width="510" height="25" colspan="8">
      <input type="text" name="rate_minimum" size="20" style="width:127; font-family:Verdana; font-size:10pt; text-align:right; height:22" value="<%=strRangeMin %>"></td>
      <td width="262" height="22">&nbsp;</td>
    </tr>
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
      <td width="210" height="22"><font face="Verdana" size="2">18. Evaluate on 
      Success:</font></td>
      <td width="510" height="22" colspan="8">
      <select size="1" name="on_success_id" style="width:370; font-family:Verdana; font-size:10pt; height:24">
      <option selected value="0">(No follow-on rule)</option>
      
                <% While adoRS18a.EOF = False %> 
                	<% If intSuccessId = adoRS18a.Fields("rate_rule_id").Value Then %>
	                	<option selected value="<%=adoRS18a.Fields("rate_rule_id").Value %>"><%=adoRS18a.Fields("alert_desc").Value %></option>
	                <% Else %>
		               <% If adoRS18a.Fields("rule_status").Value = "E" Then %>
		                	<option value="<%=adoRS18a.Fields("rate_rule_id").Value %>"><%=adoRS18a.Fields("alert_desc").Value %></option>
		                <% End If %> 
					<% End If %>                
                <%	adoRS18a.MoveNext
				   Wend
					   
				   Set adoRS18a = Nothing
				%>
      
      </select></td>
      <td width="262" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="22">&nbsp;</td>
      <td width="217" height="22">&nbsp;</td>
      <td width="210" height="22"><font face="Verdana" size="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
      Evaluate on Failure:</font></td>
      <td width="510" height="22" colspan="8">
		<select name="on_failure_id" style="width:370; font-family:Verdana; font-size:10pt; height:24" size="1">
     <option selected value="0">(No follow-on rule)</option>
      
                <% While adoRS18b.EOF = False %> 
                	<% If intFailureId = adoRS18b.Fields("rate_rule_id").Value Then %>
	                	<option selected value="<%=adoRS18b.Fields("rate_rule_id").Value %>"><%=adoRS18b.Fields("alert_desc").Value %></option>
	                <% Else %>
		               <% If adoRS18b.Fields("rule_status").Value = "E" Then %>
		                	<option value="<%=adoRS18b.Fields("rate_rule_id").Value %>"><%=adoRS18b.Fields("alert_desc").Value %></option>
		                <% End If %> 
					<% End If %>                
                <%	adoRS18b.MoveNext
				   Wend
					   
				   Set adoRS18a = Nothing
				%>
				</select></td>
      <td width="262" height="22">&nbsp;</td>
    </tr>
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
      <td width="210" height="25"><font face="Verdana" size="2">
      19. Utilization Range:</td>
      <td width="510" height="25" colspan="8"><font face="Verdana" size="2">
      (leave blank or enter a zero to indicate no limit)</td>
      <td width="262" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="22">&nbsp;</td>
      <td width="217" height="22">&nbsp;</td>
      <td width="210" height="22"><font face="Verdana" size="2">&nbsp;&nbsp;&nbsp;&nbsp; Days out:</font></td>
      <td width="64" height="22" align="center"><font face="Verdana" size="2">Same</font></td>
      <td width="64" height="22" align="center"><font face="Verdana" size="2">Next</font></td>
      <td width="64" height="22" align="center"><font face="Verdana" size="2">2 - 4</font></td>
      <td width="64" height="22" align="center"><font face="Verdana" size="2">5 - 7</font></td>
      <td width="64" height="22" align="center"><font face="Verdana" size="2">8 - 14</font></td>
      <td width="64" height="22" align="center"><font face="Verdana" size="2">15 - 30</font></td>
      <td width="64" height="22" align="center"><font face="Verdana" size="2">31 - 50</font></td>
      <td width="63" height="22" align="center"><font face="Verdana" size="2">51 +</font></td>
	  <td width="262" height="22" >&nbsp;</td>
      <td height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="22">&nbsp;</td>
      <td width="217" height="22">&nbsp;</td>
      <td width="210" height="22"><font face="Verdana" size="2">&nbsp;&nbsp;&nbsp;&nbsp; Maximum:</td>
      <td width="64" height="22"><font face="Verdana" size="2"><input type="text"  name="util_max_0" size="20" style="width:59; font-family:Verdana; font-size:10pt; text-align:right; height:22" value="<%=strUtilMax0 %>"></font></td>
      <td width="64" height="22"><font face="Verdana" size="2"><input type="text"  name="util_max_1" size="20" style="width:59; font-family:Verdana; font-size:10pt; text-align:right; height:22" value="<%=strUtilMax1 %>"></font></td>
      <td width="64" height="22"><font face="Verdana" size="2"><input type="text"  name="util_max_2" size="20" style="width:59; font-family:Verdana; font-size:10pt; text-align:right; height:22" value="<%=strUtilMax2 %>"></font></td>
      <td width="64" height="22"><font face="Verdana" size="2"><input type="text"  name="util_max_3" size="20" style="width:59; font-family:Verdana; font-size:10pt; text-align:right; height:22" value="<%=strUtilMax3 %>"></font></td>
      <td width="64" height="22"><font face="Verdana" size="2"><input type="text"  name="util_max_4" size="20" style="width:59; font-family:Verdana; font-size:10pt; text-align:right; height:22" value="<%=strUtilMax4 %>"></font></td>
      <td width="64" height="22"><font face="Verdana" size="2"><input type="text"  name="util_max_5" size="20" style="width:59; font-family:Verdana; font-size:10pt; text-align:right; height:22" value="<%=strUtilMax5 %>"></font></td>
      <td width="64" height="22"><font face="Verdana" size="2"><input type="text"  name="util_max_6" size="20" style="width:59; font-family:Verdana; font-size:10pt; text-align:right; height:22" value="<%=strUtilMax6 %>"></font></td>
      <td width="63" height="22"><font face="Verdana" size="2"><input type="text"  name="util_max_7" size="20" style="width:59; font-family:Verdana; font-size:10pt; text-align:right; height:22" value="<%=strUtilMax7 %>">&nbsp; </font></td>
	  <td width="262" height="22" ><font face="Verdana" size="2"></font></td>
	  <td height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="22">&nbsp;</td>
      <td width="217" height="22">&nbsp;</td>
      <td width="210" height="25"><font face="Verdana" size="2">&nbsp;&nbsp;&nbsp;&nbsp; Minimum:</td>
      <td width="64" height="25"><font face="Verdana" size="2"><input type="text" name="util_min_0" size="20" style="width:59; font-family:Verdana; font-size:10pt; text-align:right; height:22" value="<%=strUtilMin0 %>"></font></td>
      <td width="64" height="25"><font face="Verdana" size="2"><input type="text" name="util_min_1" size="20" style="width:59; font-family:Verdana; font-size:10pt; text-align:right; height:22" value="<%=strUtilMin1 %>"></font></td>
      <td width="64" height="25"><font face="Verdana" size="2"><input type="text" name="util_min_2" size="20" style="width:59; font-family:Verdana; font-size:10pt; text-align:right; height:22" value="<%=strUtilMin2 %>"></font></td>
      <td width="64" height="25"><font face="Verdana" size="2"><input type="text" name="util_min_3" size="20" style="width:59; font-family:Verdana; font-size:10pt; text-align:right; height:22" value="<%=strUtilMin3 %>"></font></td>
      <td width="64" height="25"><font face="Verdana" size="2"><input type="text" name="util_min_4" size="20" style="width:59; font-family:Verdana; font-size:10pt; text-align:right; height:22" value="<%=strUtilMin4 %>"></font></td>
      <td width="64" height="25"><font face="Verdana" size="2"><input type="text" name="util_min_5" size="20" style="width:59; font-family:Verdana; font-size:10pt; text-align:right; height:22" value="<%=strUtilMin5 %>"></font></td>
      <td width="64" height="25"><font face="Verdana" size="2"><input type="text" name="util_min_6" size="20" style="width:59; font-family:Verdana; font-size:10pt; text-align:right; height:22" value="<%=strUtilMin6 %>"></font></td>
      <td width="63" height="25"><font face="Verdana" size="2"><input type="text" name="util_min_7" size="20" style="width:59; font-family:Verdana; font-size:10pt; text-align:right; height:22" value="<%=strUtilMin7 %>"></font></td>
	  <td width="262" height="22"><font face="Verdana" size="2"></font></td>
      <td height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="22">&nbsp;</td>
      <td width="217" height="22">&nbsp;</td>
      <td width="210" height="22">
      &nbsp;</td>
</font>
      <td width="510" height="22" colspan="8">
      <input type="checkbox" name="util_in_percent" id="util_in_percent" value="True" checked disabled ><label for="util_in_percent"><font size="2">values 
      are listed as percentages (please do not include percent signs)</font></label></td>
      <td width="262" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="22">&nbsp;</td>
      <td width="217" height="22">&nbsp;</td>
      <td width="210" height="22" align="center"><font face="Verdana" size="2">OR</font></td>
      <td width="510" height="22" colspan="8">
      &nbsp;</td>
      <td width="262" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="22">&nbsp;</td>
      <td width="217" height="22">&nbsp;</td>
      <td width="210" height="22">
      &nbsp;&nbsp;&nbsp; <font size="2">Utilization Schedule</font></td>
      <td width="510" height="22" colspan="8">
      <font face="Verdana" size="2">
		<select name="utilization_profile_id" style="width:370; font-family:Verdana; font-size:10pt; height:24" size="1">
     <option selected value="0">(None - Use above buckets)</option>
				</select></font></td>
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
      <td width="510" height="22" colspan="8">
      <input type="checkbox" name="automatic" value="ON" disabled ><font size="2">Run this rule in automatic mode 
      (no user review or approval req.)</font></td>
      <td width="262" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="22">&nbsp;</td>
      <td width="217" height="22">&nbsp;</td>
      <td width="210" height="23">
      <font size="2" face="Verdana, Arial, Helvetica, sans-serif">
          <input name="<%=strButton %>" type="submit" id="submit" value="    <%=strButton %>   " class="rh_button"></font></td>
      <td width="510" height="23" colspan="8">
      &nbsp;</td>
      <td width="262" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="8"  height="61">&nbsp;</td>
      <td width="217" height="61">&nbsp;</td>
      <td width="720" height="61" colspan="9">
      <!--
      <font size="1">closed = <%=blnIgnoreClosed  %></font>
      -->
      </td>
      <td width="262" height="61">&nbsp;</td>
    </tr>
    <tr>
      <td width="8"  height="19">&nbsp;</td>
      <td width="217" height="19">&nbsp;</td>
      <td width="720" height="19" colspan="9" >&nbsp;</td>
      <td width="262" height="19">&nbsp;</td>
    </tr>
  </table>
  <table border="0" style="border-collapse: collapse" bordercolor="#FFFFFF" width="1210" height="4">
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
  <input type="hidden" name="refresh_from" value="create">
  <input type="hidden" name="rule_status" value="E">
</form>
<!-- Content goes before this comment -->
<!-- JUSTTABS BOTTOM OPEN -->
</font></td></tr></table>
</td>
<td  width="1" bgcolor="#000000"><img src=pixel.gif width="1" height="1"></td>
</tr>
<tr bgcolor="#000000" height="1">
<td colspan=5><img src=pixel.gif width="1" height="1"></td>
</tr>
</table>
<!-- JUSTTABS BOTTOM CLOSE -->
<!--#INCLUDE FILE="footer.asp"-->
<div id="calbox" class="calboxoff"></div>	
</body>

</html>

<%	
	Rem Clean up
	Set adoCmd = Nothing 
	Set adoCmd1 = Nothing 
	Set adoCmd2 = Nothing 
	Set adoCmd3 = Nothing 
	Set adoRS5 = Nothing
	
%>

<script language="javascript">
	document.create_alert.alert_desc.focus();
</script>