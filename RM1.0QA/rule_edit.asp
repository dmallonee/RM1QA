<%@ Language=VBScript %>
<!-- #INCLUDE FILE="inc/login_check.asp" -->
<!-- #INCLUDE FILE="inc/adovbs.asp" -->
<% Response.Expires = -1
   Response.cachecontrol="private" 
   Response.AddHeader "pragma", "no-cache"

   Server.ScriptTimeout = 30

	On Error Resume Next

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
	Dim intComparisonRate 
	Dim strRateAmtTolerance
	Dim strRuleStatus
	
	'Declare variables
	Dim iCurrentPage
	Dim intPageSize
	Dim i
	Dim oConnection
	Dim oRecordSet
	Dim oTableField
	Dim sPageURL
	Dim strEditMode 

	'Retrieve the name of the current ASP document
	sPageURL = Request.ServerVariables("SCRIPT_NAME")

	'Retrieve the current page number from the QueryString
	intCurrentPage= Request.QueryString("page")
	If intCurrentPage= "" Or intCurrentPage= 0 Then intCurrentPage= 1

	'Retrieve the current page number from the QueryString
	strRuleStatus = Request.QueryString("rule_status")
	If strRuleStatus = "" Then strRuleStatus = "E"
	
	'Set the number of records to be displayed on each page
	intPageSize = 100

	'Retrieve the current edit style from the QueryString
	strEditMode = Request.QueryString("edit_mode")

	strUserId = Request.Cookies("rate-monitor.com")("user_id")
	intRuleId = Request.QueryString("rate_rule_id")	
	
	strConn = Session("pro_con")
	
	Rem Get the data sources
	Set adoCmd1 = CreateObject("ADODB.Command")

	adoCmd1.ActiveConnection =  strConn
	adoCmd1.CommandText = "data_source_select"
	adoCmd1.CommandType = adCmdStoredProc
		
	Set adoRS1 = adoCmd1.Execute

	Rem Get the vendors
	Set adoCmd2 = CreateObject("ADODB.Command")

	adoCmd2.ActiveConnection =  strConn
	adoCmd2.CommandText = "vendor_select_ex"
	adoCmd2.CommandType = adCmdStoredProc

	adoCmd2.Parameters.Append adoCmd2.CreateParameter("@vendor_cd",   200, 1,  2, Null)
	adoCmd2.Parameters.Append adoCmd2.CreateParameter("@vendor_name", 200, 1, 50, Null)
	adoCmd2.Parameters.Append adoCmd2.CreateParameter("@user_id",       3, 1,  0, strUserId)
		
	Set adoRS2 = adoCmd2.Execute
	
	Rem Get the vendors
	'Set adoCmd3 = CreateObject("ADODB.Command")

	'adoCmd3.ActiveConnection =  strConn
	'adoCmd3.CommandText = "vendor_select"
	'adoCmd3.CommandType = adCmdStoredProc
		
	'Set adoRS3 = adoCmd3.Execute
	Set adoRS3 = adoCmd2.Execute

   'response.write "Error test - 2 <br>"
	
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

	Rem Get the rule
	Set adoCmd4 = CreateObject("ADODB.Command")

	adoCmd4.ActiveConnection =  strConn
	adoCmd4.CommandText = "car_rate_rule_select"
	adoCmd4.CommandType = adCmdStoredProc
	adoCmd4.Parameters.Append adoCmd4.CreateParameter("@rate_rule_id",      3, 1, 0, Null)
	adoCmd4.Parameters.Append adoCmd4.CreateParameter("@user_id",           3, 1, 0, strUserId)
	adoCmd4.Parameters.Append adoCmd4.CreateParameter("@rate_rule_type_cd", 3, 1, 0, 1)
	adoCmd4.Parameters.Append adoCmd4.CreateParameter("@rule_status",     200, 1, 1, strRuleStatus)

	'Create an ADO RecordSet object
	Set adoRS4 = Server.CreateObject("ADODB.Recordset")

	Set adoRS18a = adoCmd4.Execute
	Set adoRS18b = adoCmd4.Execute
	
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

			
	Rem Get the cities
	Set adoCmd6 = CreateObject("ADODB.Command")

	adoCmd6.ActiveConnection =  strConn
	adoCmd6.CommandText = "user_city_select"
	adoCmd6.CommandType = adCmdStoredProc

	adoCmd6.Parameters.Append adoCmd6.CreateParameter("@user_id", 3, 1, 0, strUserId)
		
	Set adoRS6 = adoCmd6.Execute

	Rem Get the car types
	Set adoCmd7 = CreateObject("ADODB.Command")

	adoCmd7.ActiveConnection =  strConn
	adoCmd7.CommandText = "car_type_select"
	adoCmd7.CommandType = adCmdStoredProc
	
	adoCmd7.Parameters.Append adoCmd7.CreateParameter("@user_id", 3, 1, 0, strUserId)
		
	Set adoRS7 = adoCmd7.Execute
	Set adoRS8 = adoCmd7.Execute
	
	Set adoCmd9 = CreateObject("ADODB.Command")

	adoCmd9.ActiveConnection =  strConn
	adoCmd9.CommandText = "car_shop_profile_select"
	adoCmd9.CommandType = adCmdStoredProc

	adoCmd9.Parameters.Append adoCmd9.CreateParameter("@desc",              200, 1, 255)
	adoCmd9.Parameters.Append adoCmd9.CreateParameter("@shop_car_type_cds", 200, 1, 4096)
	adoCmd9.Parameters.Append adoCmd9.CreateParameter("@vend_cds",          200, 1, 1024)
	adoCmd9.Parameters.Append adoCmd9.CreateParameter("@user_id",             3, 1, 0, strUserId)
	adoCmd9.Parameters.Append adoCmd9.CreateParameter("@profile_id",          3, 1, 0, Null)
		
	Set adoRS9 = adoCmd9.Execute

	Set adoCmd9 = CreateObject("ADODB.Command")

	adoCmd9.ActiveConnection =  strConn
	adoCmd9.CommandText = "car_rate_rule_schedule_select"
	adoCmd9.CommandType = adCmdStoredProc

	adoCmd9.Parameters.Append adoCmd9.CreateParameter("@org_id",              3, 1, 0, Null)
	adoCmd9.Parameters.Append adoCmd9.CreateParameter("@schedule_type_id",    3, 1, 0, 1)
	adoCmd9.Parameters.Append adoCmd9.CreateParameter("@user_id",             3, 1, 0, strUserId)
		
	Set adoRS10 = adoCmd9.Execute

	Set adoCmd9 = CreateObject("ADODB.Command")

	adoCmd9.ActiveConnection =  strConn
	adoCmd9.CommandText = "car_rate_rule_schedule_select"
	adoCmd9.CommandType = adCmdStoredProc

	adoCmd9.Parameters.Append adoCmd9.CreateParameter("@org_id",              3, 1, 0, Null)
	adoCmd9.Parameters.Append adoCmd9.CreateParameter("@schedule_type_id",    3, 1, 0, 4)
	adoCmd9.Parameters.Append adoCmd9.CreateParameter("@user_id",             3, 1, 0, strUserId)
		
	Set adoRS11 = adoCmd9.Execute

	Set adoCmd9 = Nothing

	If intRuleId > 0 Then

		Rem Get the specific rule
		Set adoCmd5 = CreateObject("ADODB.Command")

		adoCmd5.ActiveConnection = strConn
		adoCmd5.CommandText = "car_rate_rule_select"
		adoCmd5.CommandType = adCmdStoredProc
		
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


		Rem Really really lame, use the code of 99999 to indicate a schedule
		If strRangeMin = 99999 Then
			intMaxMinProfileId = strRangeMax
			strRangeMax = ""
			strRangeMin = ""
		Else
			intMaxMinProfileId = 0
		End If

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
		intComparisonRate = adoRS5.Fields("comparison_rate").value
		strRateAmtTolerance  = adoRS5.Fields("rt_amt_tolerance").value
		strDowList = adoRS5.Fields("dow_list").value
		blnAutomatic = adoRS5.Fields("automatic").value
		curExtraDayRt = adoRS5.Fields("extra_day_rt").value
		curExtraDayMiles = adoRS5.Fields("extra_day_miles").value
		curExtraDayRtPerMile = adoRS5.Fields("extra_day_rt_per_mile").value
		curExtraHrRt = adoRS5.Fields("extra_hr_rt").value
		curExtraHrMiles = adoRS5.Fields("extra_hr_miles").value
		curExtraHrRtPerMile = adoRS5.Fields("extra_hr_rt_per_mile").value
		intRulePostActionId1 = adoRS5.Fields("rule_post_action_id1").Value
		If IsNull(adoRS5.Fields("rule_post_action_amt1").Value) Then
			strRulePostActionAmt1 = ""
		Else
			'strRulePostActionAmt1 = FormatCurrency(adoRS5.Fields("rule_post_action_amt1").Value)
			Rem Only format currency if the item is a dollar amt  i.e. 1 or 2
			If (intRulePostActionId1 < 3) Then
				strRulePostActionAmt1 = FormatCurrency(adoRS5.Fields("rule_post_action_amt1").Value)
			Else
				strRulePostActionAmt1 = adoRS5.Fields("rule_post_action_amt1").Value
			End If
		End If
		intRulePostActionId2 = adoRS5.Fields("rule_post_action_id2").Value
		If IsNull(adoRS5.Fields("rule_post_action_amt2").Value) Then
			strRulePostActionAmt2 = ""
		Else
			Rem Only format currency if the item is a dollar amt  i.e. 1 or 2
			If (intRulePostActionId2 < 3) Then
				strRulePostActionAmt2 = FormatCurrency(adoRS5.Fields("rule_post_action_amt2").Value)
			Else
				strRulePostActionAmt2 = adoRS5.Fields("rule_post_action_amt2").Value
			End If
		End If
		intRulePostActionId3 = adoRS5.Fields("rule_post_action_id3").Value
		If IsNull(adoRS5.Fields("rule_post_action_amt3").Value) Then
			strRulePostActionAmt3 = ""
		Else
			'strRulePostActionAmt3 = FormatCurrency(adoRS5.Fields("rule_post_action_amt3").Value)
			Rem Only format currency if the item is a dollar amt  i.e. 1 or 2
			If (intRulePostActionId3 < 3) Then
				strRulePostActionAmt3 = FormatCurrency(adoRS5.Fields("rule_post_action_amt3").Value)
			Else
				strRulePostActionAmt3 = adoRS5.Fields("rule_post_action_amt3").Value
			End If

		End If
		intRulePostActionId4 = adoRS5.Fields("rule_post_action_id4").Value
		If IsNull(adoRS5.Fields("rule_post_action_amt4").Value) Then
			strRulePostActionAmt4 = ""
		Else
			'strRulePostActionAmt4 = FormatCurrency(adoRS5.Fields("rule_post_action_amt4").Value)
			Rem Only format currency if the item is a dollar amt  i.e. 1 or 2
			If (intRulePostActionId4 < 3) Then
				strRulePostActionAmt4 = FormatCurrency(adoRS5.Fields("rule_post_action_amt4").Value)
			Else
				strRulePostActionAmt4 = adoRS5.Fields("rule_post_action_amt4").Value
			End If
			
		End If


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
		intSituationCd = 1
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
		intComparisonRate = 1
		strRateAmtTolerance = "0"
		strDowList = "1,2,3,4,5,6,7"
		curExtraDayRt = ""
		curExtraDayMiles = ""
		curExtraDayRtPerMile = ""
		curExtraHrRt = ""
		curExtraHrMiles = ""
		curExtraHrRtPerMile = ""
		intMaxMinProfileId = 0 
		intRulePostActionId1 = 0
		strRulePostActionAmt1 = ""
		intRulePostActionId2 = 0
		strRulePostActionAmt2 = ""
		intRulePostActionId3 = 0
		strRulePostActionAmt3 = ""
		intRulePostActionId4 = 0
		strRulePostActionAmt4 = ""
		blnAutomatic = False
		
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
<meta name="GENERATOR" content="Microsoft FrontPage 12.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Language" content="en-us">
<title>Rate-Monitor by Rate-Highway, Inc. | Rules | Rule Management</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<script language="Javascript" type="text/javascript" src="inc/sitewide.js"></script>
<script language="Javascript" type="text/javascript" src="inc/header_menu_support.js" ></script>
<script language="Javascript" type="text/javascript" src="inc/js_calendar_v2.js"></script>
<script language="Javascript" type="text/javascript" src="inc/validate2.js"></script>
<script language="JavaScript" type="text/JavaScript">
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


function centerPopUp( url, name, width, height, scrollbars ) { 
 
	if( scrollbars == null ) scrollbars = "0" 
 
	str  = ""; 
	str += "resizable=1,"; 
	str += "scrollbars=" + scrollbars + ","; 
	str += "width=" + width + ","; 
	str += "height=" + height + ","; 
 
	if ( window.screen ) { 
		var ah = screen.availHeight - 30; 
		var aw = screen.availWidth - 10; 
 
		var xc = ( aw - width ) / 2; 
		var yc = ( ah - height ) / 2; 
 
		str += ",left=" + xc + ",screenX=" + xc; 
		str += ",top=" + yc + ",screenY=" + yc; 
	} 
	window.open( url, name, str ); 
} 

// Hide and display the addtional Rate information layer
function toggleLayer(whichLayer)
{
if (document.getElementById)
{
// this is the way the standards work
var style2 = document.getElementById(whichLayer).style;
style2.display = style2.display? "":"block";
}
else if (document.all)
{
// this is the way old msie versions work
var style2 = document.all[whichLayer].style;
style2.display = style2.display? "":"block";
}
else if (document.layers)
{
// this is the way nn4 works
var style2 = document.layers[whichLayer].style;
style2.display = style2.display? "":"block";
}
}
</script>
<script type='text/javascript' language='javascript' > 
<!-- Begin
function formatCurrency(num) {
	if (num.toString() == "") {
		return "";
		}
	else {
		num = num.toString().replace(/\$|\,/g,'');
		if(isNaN(num)) {
			return "";
			}
		else {	
			sign = (num == (num = Math.abs(num)));
			num = Math.floor(num*100+0.50000000001);
			cents = num%100;
			num = Math.floor(num/100).toString();
			if(cents<10)
				cents = "0" + cents;
			for (var i = 0; i < Math.floor((num.length-(1+i))/3); i++)
				num = num.substring(0,num.length-(4*i+3))+','+
			num.substring(num.length-(4*i+3));
			return (((sign)?'':'-') + '$' + num + '.' + cents);
			}
		}
}
function formatNumber(num) {
	if (num.toString() == "") {
		return "";
		}
	else {
		num = num.toString().replace(/\$|\,/g,'');
		if(isNaN(num)) {
			return "";
			}
		else {	
			sign = (num == (num = Math.abs(num)));
			num = Math.floor(num*100+0.50000000001);
			cents = num%100;
			num = Math.floor(num/100).toString();
			if (cents<10)
				cents = "0" + cents;
			for (var i = 0; i < Math.floor((num.length-(1+i))/3); i++)
				num = num.substring(0,num.length-(4*i+3))+','+
			num.substring(num.length-(4*i+3));
			if (cents==0)
			return (((sign)?'':'-') + num);
			else
				return (((sign)?'':'-') + num + '.' + cents);
			}
		}
}
// If the user selects a situation the "ignore closed" does not apply 
function showCodeOption(obj) {
      if (obj.options[obj.selectedIndex].value == '47') {
        document.getElementById('ignore_closed').checked = false ;
      	}
      else {
        document.getElementById('ignore_closed').checked = true ;
        
        }
    }


//  End -->
</script>

<style>
<!--
.profile_header { height="48" text-align:left; padding-left:3; padding-right:3; padding-top:0; background-color:#879AA2; font-family:Verdana, Arial, Helvetica, sans-serif; font-size:10pt; vertical-align:bottom; text-align:left }
.profile_light { height="68" text-align:left; padding-left:3; padding-right:3; padding-top:3; background-color:#CFD7DB; font-family:Verdana, Arial, Helvetica, sans-serif; font-size:10pt; vertical-align:bottom; text-align:left }
.profile_dark { height="48" text-align:left; padding-left:3; padding-right:3; padding-top:3; background-color:#C0C0C0; font-family:Verdana, Arial, Helvetica, sans-serif; font-size:10pt; vertical-align:bottom; text-align:left }
div#ExtraDay { margin: 0px 20px 0px 20px; display: none; }
.style1 {
	font-size: x-small;
}
.style2 {
	font-size: x-small;
	text-align: left;
}
-->
</style>
<link rel="stylesheet" type="text/css" href="inc/rh_standard.css">
<link rel="stylesheet" type="text/css" href="inc/rh_report.css">
<link rel="stylesheet" type="text/css" href="inc/sitewide.css">
<link rel="stylesheet" type="text/css" href="inc/css_calendar_v2.css" id="calendarcss" media="all" >
	
</head>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="MM_preloadImages('images/b_search_pro_on.gif','images/b_search_que_on.gif','images/b_search_cri_on.gif','images/b_rate_on.gif','images/b_alert_on.gif','images/b_user_on.gif','images/b_system_on.gif')">

<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td background="images/top_tile.gif">
    <a target="_blank" href="http://www.rate-highway.com">
    <img src="images/top.jpg" width="770" height="91" border="0" ></a></td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td background="images/b_tile.gif">
    
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
                  </div>
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
        <td><img src="images/h_alerts.gif" width="368" height="31" alt=""></td>
        <td>
        <img src="images/bottom_right_blank.gif" width="402" height="31" border="0" alt=""></td>
      </tr>
    </table>
    </td>
  </tr>
</table>

<p align="right"><font color="#080000">
                <img border="0" alt="Click to view Help" src="images/help_button.jpg" width="32" height="32" onclick="centerPopUp('help_rate_management.htm', 'help', 650, 400, 1)"></font></p>
<div align="center">
<table cellpadding="0" cellspacing="0" border="0" width="1310" bgcolor="#FFFFFF">
<tr height="1">
<td colspan="1" width="1">&nbsp;</td>
<td rowspan="2" width="169"><img src="images/ratemanagementalerts2_a.JPG" width="169" height="25" hspace="0" vspace="0" border="0" alt="Rate Management" description=""></td>
<td colspan="1" >&nbsp;</td>
</tr>
<tr height="1">
<td bgcolor="#000000" colspan="1" height="1"><img src="images/pixel.gif" width="1" height="1"></td>
<td bgcolor="#000000" colspan="1" height="1"><img src="images/pixel.gif" width="1" height="1"></td>
<td bgcolor="#000000" colspan="1" height="1"><img src="images/pixel.gif" width="1" height="1"></td>
</tr>
</table>
</div>
<table cellpadding="0" cellspacing="0" border="0" ALIGN="CENTER" width="100%" bgcolor="#FFFFFF">
<tr >
<td  width="1" bgcolor="#000000"><img src="images/pixel.gif" width="1" height="1"></td>
<td colspan=3 bgcolor="#D9DEE1">
<table border="0" cellspacing="5" cellpadding="5">
<tr><td>
<font color="#080000">
<P>
&nbsp;
  <form name="create_alert" method="POST" action="" OnSubmit="return CreateAlert()">
  <table border="0" style="border-collapse: collapse" bordercolor="#FFFFFF" width="1310" height="4">
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
  <table border="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#111111" width="1210" id="new_alert" background="images/alt_color.gif" height="561">
    <tr>
      <td width="8" height="19">&nbsp;</td>
      <td width="217" height="19">&nbsp;</td>
      <td width="210" height="19">&nbsp;</td>
      <td width="510" height="19" colspan="8">&nbsp;</td>
      <td width="262" height="19">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="25">&nbsp;</td>
      <td width="217" height="25">
      <img border="0" src="images/maintenance.GIF" width="162" height="25"></td>
      <td width="210" height="25"><font face="Verdana" size="2">1. Rate Change 
		Alert No.:</font></td>
      <td width="510" height="25" colspan="8">
      <input type="text" name="rate_rule_id" size="20" style="width:200; font-family:Verdana; font-size:10pt; text-align:right; background-image:url('images/alt_color.gif')" value="<%=intRuleId %>" READONLY></td>
      <td width="262" height="25">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="25">&nbsp;</td>
      <td width="217" height="25">
      &nbsp;</td>
      <td width="210" height="25">&nbsp;</td>
      <td width="510" height="25" colspan="8">
      &nbsp;<font face="Verdana" size="2" color="#080000"><input type="checkbox" name="copy" id="copy" value="true"><font face="Verdana" size="2">Save 
		as a copy (leaves the original unchanged)</font></td>
      <td width="262" height="25">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="25">&nbsp;</td>
      <td width="217" height="25">
      &nbsp;</td>
      <td width="210" height="25"><font face="Verdana" size="2">2. Description:</font></td>
      <td width="510" height="25" colspan="8">
      <input type="text" name="alert_desc" size="20" style="width:439; font-family:Verdana; font-size:10pt; height:21" value="<%=strAlertDesc %>"></td>
      <td width="262" height="25">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="25">&nbsp;</td>
      <td width="217" height="25">
      &nbsp;</td>
      <td width="210" height="25">&nbsp;</td>
      <td width="510" height="25" colspan="8">
      &nbsp;</td>
      <td width="262" height="25">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="25">&nbsp;</td>
      <td width="217" height="25">
      &nbsp;</td>
      <td width="210" height="25"><font face="Verdana" size="2">3. Rule Begin 
		Date:</font></td>
      <td width="510" height="25" colspan="8">
		<div class="cbrow" id="rule_begin_dt">
		<input type="text" name="begin_dt" id="begin_dt" class="cb_txtdate" value="<%=datBeginDate %>" onfocus="openCal(this,'begin_dt','end_dt','calbox','rule_begin_dt','us','vertical');if(this.value=='mm/dd/yyyy')this.value=''" title="mm/dd/yyyy" style="width:200"  size='20' maxlength="10" >
		</div>
      </td>
      <td width="262" height="25">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="25">&nbsp;</td>
      <td width="217" height="25">
      &nbsp;</td>
      <td width="210" height="25">&nbsp;&nbsp;&nbsp;</td>
      <td width="510" height="25" colspan="8">
      <font face="Verdana" size="2">(enter 'continuous' or blank for no begin 
		date) </font></td>
      <td width="262" height="25">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="25">&nbsp;</td>
      <td width="217" height="25">
      &nbsp;</td>
      <td width="210" height="25"><font face="Verdana" size="2">4. Rule End 
		Date:</font></td>
      <td width="510" height="25" colspan="8">
		<div class="cbrow" id="rule_end_dt">
		<input type="text" name="end_dt" id="end_dt" class="cb_txtdate" value="<%=datEndDate %>" onfocus="openCal(this,'begin_dt','end_dt','calbox','rule_end_dt','us','vertical');if(this.value=='mm/dd/yyyy')this.value=''" title="mm/dd/yyyy" style="width:200"  size='20' maxlength="10" >
		</div>
	  </td>
      <td width="262" height="25">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="25">&nbsp;</td>
      <td width="217" height="25">
      &nbsp;</td>
      <td width="210" height="25">&nbsp;</td>
      <td width="510" height="25" colspan="8">
      <font face="Verdana" size="2">(enter 'continuous' or blank for no end 
		date)</font></td>
      <td width="262" height="25">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="25">&nbsp;</td>
      <td width="217" height="25">
      &nbsp;</td>
      <td width="210" height="25"><font face="Verdana" size="2">5. First 
		Pick-up:</font></td>
      <td width="510" height="25" colspan="8">
		<div class="cbrow" id="rule_first_pickup_dt">
		<input type="text" name="first_pickup_dt" id="first_pickup_dt" class="cb_txtdate" value="<%=datFirstPickupDate %>" onfocus="openCal(this,'first_pickup_dt','last_pickup_dt','calbox','rule_first_pickup_dt','us','vertical');if(this.value=='mm/dd/yyyy')this.value=''" title="mm/dd/yyyy" style="width:200"  size='20' maxlength="10" >
		</div>
	  </td>
      <td width="262" height="25">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="25">&nbsp;</td>
      <td width="217" height="25">
      &nbsp;</td>
      <td width="210" height="25">&nbsp; &nbsp;
      <% If blnRollingDates Then %>
      <input type="checkbox" name="rolling_date" value="True" id="rolling_date" checked >
      <% Else %>
      <input type="checkbox" name="rolling_date" value="True" id="rolling_date">
      <% End If %>
      <label for="rolling_dates"><font size="2">Rolling begin &amp; end</font></label>      
      </td>
      <td width="510" height="25" colspan="8">
      <font face="Verdana" size="2">(enter 'continuous' or blank for no pick-up 
		date)</font></td>
      <td width="262" height="25">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="25">&nbsp;</td>
      <td width="217" height="25">
      &nbsp;</td>
      <td width="210" height="25"><font face="Verdana" size="2">6. Last Pick-up:</font></td>
      <td width="510" height="25" colspan="8">
		<div class="cbrow" id="rule_last_pickup_dt">
		<input type="text" name="last_pickup_dt" id="last_pickup_dt" class="cb_txtdate" value="<%=datSecondPickupDate %>" onfocus="openCal(this,'first_pickup_dt','last_pickup_dt','calbox','rule_last_pickup_dt','us','vertical');if(this.value=='mm/dd/yyyy')this.value=''" title="mm/dd/yyyy" style="width:200"  size='20' maxlength="10" >
		</div>
      </td>
      <td width="262" height="25">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="25">&nbsp;</td>
      <td width="217" height="25">
      &nbsp;</td>
      <td width="210" height="25">&nbsp;</td>
      <td width="510" height="25" colspan="8">
      <font face="Verdana" size="2">(enter 'continuous' or blank for no pick-up 
		date)</font></td>
      <td width="262" height="25">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="25">&nbsp;</td>
      <td width="217" height="25">
      &nbsp;</td>
      <td width="210" height="25"><font face="Verdana" size="2">7. System Rate 
		Code:</font></td>
      <td width="510" height="25" colspan="8">
      <input type="text" name="client_sys_rate_cd" size="20" style="width:200; font-family:Verdana; font-size:10pt" value="<%=strClientRateCode %>"></td>
      <td width="262" height="25">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="25">&nbsp;</td>
      <td width="217" height="25">
      &nbsp;</td>
      <td width="210" height="25">&nbsp;</td>
      <td width="510" height="25" colspan="8"><font face="Verdana" size="2">
      (the rate code used within your system, i.e. Daily, Weekly, etc.)</td>
      <td width="262" height="25">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="25">&nbsp;</td>
      <td width="217" height="25">
      &nbsp;</td>
      <td width="210" height="25"><font face="Verdana" size="2">8. Car 
		Companies:</font></td>
      <td width="510" height="25" colspan="8">&nbsp;</td>
      <td width="262" height="25">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="25">&nbsp;</td>
      <td width="217" height="25">
      &nbsp;</td>
      <td width="210" height="25" style="padding-top: 2px" >
      <font face="Verdana" size="2">&nbsp;&nbsp;&nbsp;Competitive Set:</font></td>
      <td width="256" height="25" colspan="4">
      <select size="4" name="vend_cd1" multiple style="width:200; font-family:Verdana; font-size:10pt"  >
      <% If strVendCd = "XX" Then %>
      <option selected value="XX"><%="All comp. set" %></option>
	  <% Else                     %>
      <option value="XX"><%="All comp. set" %></option>
      <% End  If                  %>
					<% Dim intLoopCount         %>
	 				<% While (adoRS2.EOF = False) And (intLoopCount < 100) %>
	 				<% If (InStr(1, strVendCd, adoRS2.Fields("vendor_cd").Value)) And (strVendCd <> "") Then %>
			              <option selected value="<%=adoRS2.Fields("vendor_cd").Value %>"><%=adoRS2.Fields("vendor_name").Value %></option>
	 				<% Else                     %>
			              <option value="<%=adoRS2.Fields("vendor_cd").Value %>"><%=adoRS2.Fields("vendor_name").Value %></option>
	 				<% End If                   %>
					<% 
						 adoRS2.MoveNext
						 intLoopCount = intLoopCount + 1
					   Wend
					   Set adoRS2 = Nothing
					%>


      </select>


	              </select>      
      </td>
      <td width="125" height="25" colspan="2" bgcolor="#FFFFFF" bordercolor="#555566" bordercolorlight="#999999" bordercolordark="#777777" style="border: 1px solid #40618F; padding: 2px">
      <p align="left">
      <img border="0" src="images/tip_ballon.gif" width="23" height="22"><font size="2"><b> 
      Note</b>: Your competitive set can be one or more comp. set. Just 
		CTRL+Click to select multiple companies </font>      
      </td>
      <td width="127" height="25" colspan="2" bordercolor="#555566">
      &nbsp;</td>
      <td width="262" height="25">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="25">&nbsp;</td>
      <td width="217" height="25">
      &nbsp;</td>
      <td width="210" height="25" valign="top" style="padding-top: 2px">
      &nbsp;</td>
      <td width="510" height="25" colspan="8">
      &nbsp;</td>
      <td width="262" height="25">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="25">&nbsp;</td>
      <td width="217" height="25">
      &nbsp;</td>
      <td width="210" height="25" style="padding-top: 2px">
      <font face="Verdana" size="2">&nbsp;&nbsp; Comparison Company<br>
		&nbsp;&nbsp; (usually self):</font></td>
      <td width="256" height="25" colspan="4">
      <select size="4" name="vend_cd2" multiple style="width:200; font-family:Verdana; font-size:10pt; background-image:url('images/alt_color.gif'); background-repeat:repeat" >
      <% If strSelfCd = "XX" Then %>
      <option selected value="XX"><%="N/A" %></option>
	  <% Else                     %>
      <option value="XX"><%="N/A" %></option>
      <% End  If                  %>

					<% intLoopCount = 0 %>
	 				<% While (adoRS3.EOF = False) And (intLoopCount < 100)  %>
	 				<% If (InStr(1, strSelfCd, adoRS3.Fields("vendor_cd").Value)) And (strVendCd <> "") Then %>
			              <option selected value="<%=adoRS3.Fields("vendor_cd").Value %>"><%=adoRS3.Fields("vendor_name").Value %></option>
	 				<% Else                     %>
			              <option value="<%=adoRS3.Fields("vendor_cd").Value %>"><%=adoRS3.Fields("vendor_name").Value %></option>
	 				<% End If                   %>
					<% 
						 adoRS3.MoveNext
						 intLoopCount = intLoopCount + 1
					   Wend
					   Set adoRS3 = Nothing
					%>


      </select></td>
<td width="125" height="25" colspan="2" bgcolor="#FFFFFF" bordercolor="#555566" bordercolorlight="#999999" bordercolordark="#777777" style="border: 1px solid #40618F; padding: 2px">
      <p align="left">
      <img border="0" src="images/tip_ballon.gif" width="23" height="22"><font size="2"><b> 
      Note</b>: If you select a group as the comparison company, the lowest rate 
		of the group will be used </font>      
      </td>
      <td width="127" height="25" colspan="2" bordercolor="#555566">
      &nbsp;</td>      
      <td width="262" height="25">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="25">&nbsp;</td>
      <td width="217" height="25">
      &nbsp;</td>
      <td width="210" height="25" valign="top" style="padding-top: 2px">
      &nbsp;</td>
      <td width="510" height="25" colspan="8">
      <font face="Verdana" size="2">(use &quot;any&quot; to denote any company, you may 
		not use &quot;any&quot; for <br>
      both items)</font></td>
      <td width="262" height="25">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="25">&nbsp;</td>
      <td width="217" height="25">
      &nbsp;</td>
      <td width="210" height="25" valign="top" style="padding-top: 2px">
      <font face="Verdana" size="2">9. LOR(s):</font></td>
      <td width="510" height="25" colspan="8">
      <select size="4" name="lor" style="width:200; font-family:Verdana; font-size:10pt" locked >
      <% Select Case intLOR	%>
      
      <%	Case 1			%>
      <option selected value="1">1</option>
      <option value="2">2</option>
      <option value="3">3</option>
      <option value="4">4</option>
      <option value="5">5</option>
      <option value="6">6</option>
      <option value="7">7</option>
      <option value="8">8</option>
      
      <%	Case 2			%>
      <option value="1">1</option>
      <option selected value="2">2</option>
      <option value="3">3</option>
      <option value="4">4</option>
      <option value="5">5</option>
      <option value="6">6</option>
      <option value="7">7</option>
      <option value="8">8</option>
      
      <%	Case 3			%>
      <option value="1">1</option>
      <option value="2">2</option>
      <option selected value="3">3</option>
      <option value="4">4</option>
      <option value="5">5</option>
      <option value="6">6</option>
      <option value="7">7</option>
      <option value="8">8</option>

      <%	Case 4			%>
      <option value="1">1</option>
      <option value="2">2</option>
      <option value="3">3</option>
      <option selected value="4">4</option>
      <option value="5">5</option>
      <option value="6">6</option>
      <option value="7">7</option>
      <option value="8">8</option>

      <%	Case 5			%>
      <option value="1">1</option>
      <option value="2">2</option>
      <option value="3">3</option>
      <option value="4">4</option>
      <option selected value="5">5</option>
      <option value="6">6</option>
      <option value="7">7</option>
      <option value="8">8</option>
      
      <%	Case 6			%>
      <option value="1">1</option>
      <option value="2">2</option>
      <option value="3">3</option>
      <option value="4">4</option>
      <option value="5">5</option>
      <option selected value="6">6</option>
      <option value="7">7</option>
      <option value="8">8</option>

      <%	Case 7			%>
      <option value="1">1</option>
      <option value="2">2</option>
      <option value="3">3</option>
      <option value="4">4</option>
      <option value="5">5</option>
      <option value="6">6</option>
      <option selected value="7">7</option>
      <option value="8">8</option>

      <%	Case 8			%>
      <option value="1">1</option>
      <option value="2">2</option>
      <option value="3">3</option>
      <option value="4">4</option>
      <option value="5">5</option>
      <option value="6">6</option>
      <option value="7">7</option>
      <option selected value="8">8</option>

      
      <%	Case Else			%>
      <option selected value="1">1</option>
      <option value="5">5</option>
      
      
      <% End Select			%>
      </select></td>
      <td width="262" height="25">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="25">&nbsp;</td>
      <td width="217" height="25">
      &nbsp;</td>
      <td width="210" height="25" valign="top" style="padding-top: 2px">&nbsp;</td>
      <td width="510" height="25" colspan="8">
      &nbsp;</td>
      <td width="262" height="25">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="71">&nbsp;</td>
      <td width="217" height="71">
      &nbsp;</td>
      <td width="210" height="71" valign="top" style="padding-top: 2px">
      <font face="Verdana" size="2">10. Location(s):</font></td>
      <td width="510" height="71" valign="top" colspan="8">
      <select size="4" name="city_cd" style="width:200; font-family:Verdana; font-size:10pt" multiple >
     	 <% intLoopCount = 0                                     %>
         <% While (adoRS6.EOF = False) And (intLoopCount < 100)  %>
         <% 	If (InStr(strCityCd, adoRS6.Fields("city_cd").Value) = 0) And (strCityCd <> "") Then %>
         			<option value="<%=adoRS6.Fields("city_cd").Value %>"><%=adoRS6.Fields("city_cd").Value %></option>
         <% 	Else 											 %>		                    
         			<option selected value="<%=adoRS6.Fields("city_cd").Value %>"><%=adoRS6.Fields("city_cd").Value %></option>
						
		 <% 	End If 											 %>
		 <%    adoRS6.MoveNext 								     %>
		 <%    intLoopCount = intLoopCount + 1                   %>
		 <% Wend												 %> 
		 <% Set adoRS6 = Nothing 								 %>
      </select></td>
      <td width="262" height="71">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="24">&nbsp;</td>
      <td width="217" height="24">
      &nbsp;</td>
      <td width="210" height="24" valign="top" style="padding-top: 2px">&nbsp;</td>
      <td width="510" height="24" colspan="8">
      <font face="Verdana" size="2">(select airport/city codes &quot;any&quot; for any 
		location, edit to edit custom)</font></td>
      <td width="262" height="24">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="25">&nbsp;</td>
      <td width="217" height="25">
      &nbsp;</td>
      <td width="210" height="25" style="vertical-align: top; padding-top: 2px">
      <font face="Verdana" size="2">11. Car Types:</font></td>
      <td width="510" height="25" colspan="8">
      &nbsp;</td>
      <td width="262" height="25">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="25">&nbsp;</td>
      <td width="217" height="25">
      &nbsp;</td>
      <td width="210" height="25" style="padding-top: 2px">
      <font face="Verdana" size="2">&nbsp;&nbsp;&nbsp;&nbsp; Car Type(s):</font></td>
      <td width="510" height="25" colspan="8">
      <p style="margin-top: 2px">
      <font size="2" face="Vendana, Arial, Helvetica, sans-serif">
      <select name="car_type_cd1" size="4" multiple style="width:200; font-family:Verdana; font-size:10pt">
	 <% If strCarTypeCd1 = "XXXX" Then %>
      <option selected value="XXXX"><%="N/A" %></option>
	  <% Else                     %>
      <option value="XXXX"><%="N/A" %></option>
      <% End  If                  %>

		 <% intLoopCount = 0                                    %>
         <% While (adoRS7.EOF = False) And (intLoopCount < 100) %>
         <% 	If (InStr(strCarTypeCd1 , adoRS7.Fields("car_type_cd").Value) = 0) Or (strCarTypeCd1 = "") Then %>
         			<option value="<%=adoRS7.Fields("car_type_cd").Value %>"><%=adoRS7.Fields("car_type_cd").Value %></option>
         <% 	Else 											 %>		                    
         			<option selected value="<%=adoRS7.Fields("car_type_cd").Value %>"><%=adoRS7.Fields("car_type_cd").Value %></option>
						
		 <% 	End If 											 %>
		 <%    adoRS7.MoveNext 								     %>
		 <%    intLoopCount = intLoopCount + 1                   %>
		 <% Wend												 %> 
		 <% Set adoRS7 = Nothing 								 %>
      
      
<!--
      <option value="CCAR">CCAR</option>
      <option value="CFAR">CFAR</option>
      <option value="CPAR">CPAR</option>
      <option value="ECAR">ECAR</option>
      <option value="EDAR">EDAR</option>
      <option value="FCAR">FCAR</option>
      <option value="FDAR">FDAR</option>
      <option value="FPAR">FPAR</option>
      <option value="FVAN">FVAN</option>
      <option value="FVAR">FVAR</option>
      <option value="FVMN">FVMN</option>
      <option value="FVMR">FVMR</option>
      <option value="FWAR">FWAR</option>
      <option value="FWMN">FWMN</option>
      <option value="FWMR">FWMR</option>
      <option value="ICAN">ICAN</option>
      <option value="ICAR">ICAR</option>
      <option value="ICMN">ICMN</option>
      <option value="ICMR">ICMR</option>
      <option value="IDAN">IDAN</option>
      <option value="IDAR">IDAR</option>
      <option value="IDMN">IDMN</option>
      <option value="IDMR">IDMR</option>
      <option value="IFAR">IFAR</option>
      <option value="IJAR">IJAR</option>
      <option value="IPAR">IPAR</option>
      <option value="IVMN">IVMN</option>
      <option value="IVMR">IVMR</option>
      <option value="IWAN">IWAN</option>
      <option value="IWAR">IWAR</option>
      <option value="IWMN">IWMN</option>
      <option value="IWMR">IWMR</option>
      <option value="IXMN">IXMN</option>
      <option value="IXMR">IXMR</option>
      <option value="LCAR">LCAR</option>
      <option value="LDAR">LDAR</option>
      <option value="LDMR">LDMR</option>
      <option value="LFAR">LFAR</option>
      <option value="LTAR">LTAR</option>
      <option value="LWAR">LWAR</option>
      <option value="LXAR">LXAR</option>
      <option value="MBMN">MBMN</option>
      <option value="MCAR">MCAR</option>
      <option value="MCMN">MCMN</option>
      <option value="MCMR">MCMR</option>
      <option value="MVAR">MVAR</option>
      <option value="PCAR">PCAR</option>
      <option value="PCMR">PCMR</option>
      <option value="PDAR">PDAR</option>
      <option value="PDMR">PDMR</option>
      <option value="PFAR">PFAR</option>
      <option value="PSAR">PSAR</option>
      <option value="PVAN">PVAN</option>
      <option value="PVMN">PVMN</option>
      <option value="PVMR">PVMR</option>
      <option value="PWAR">PWAR</option>
      <option value="PXAR">PXAR</option>
      <option value="SCAN">SCAN</option>
      <option value="SCAR">SCAR</option>
      <option value="SCMN">SCMN</option>
      <option value="SCMR">SCMR</option>
      <option value="SDAN">SDAN</option>
      <option value="SDAR">SDAR</option>
      <option value="SDMN">SDMN</option>
      <option value="SDMR">SDMR</option>
      <option value="SFAR">SFAR</option>
      <option value="SJAR">SJAR</option>
      <option value="SPAR">SPAR</option>
      <option value="SSAR">SSAR</option>
      <option value="STAR">STAR</option>
      <option value="SVAR">SVAR</option>
      <option value="SVMN">SVMN</option>
      <option value="SVMR">SVMR</option>
      <option value="SWAN">SWAN</option>
      <option value="SWAR">SWAR</option>
      <option value="SWMN">SWMN</option>
      <option value="SWMR">SWMR</option>
      <option value="XCAR">XCAR</option>
      <option value="XCMN">XCMN</option>
      <option value="XDAR">XDAR</option>
      <option value="XDMN">XDMN</option>
      <option value="XFAR">XFAR</option>
      <option value="XRAR">XRAR</option>
      <option value="XSAR">XSAR</option>
      <option value="XXAR">XXAR</option>
-->
      </select> </font></td>
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
      <td width="210" height="19"><font face="Verdana" size="2">&nbsp;&nbsp;&nbsp;&nbsp; Comparison 
		Car&nbsp;&nbsp;&nbsp;<br>
		&nbsp;&nbsp;&nbsp;&nbsp; Type(s):</font></td>
      <td width="510" height="19" colspan="8">
      <font size="2" face="Vendana, Arial, Helvetica, sans-serif">
      <select name="car_type_cd2" size="4" multiple style="width:200; font-family:Verdana; font-size:10pt">
	 <% If strCarTypeCd2 = "XXXX" Then %>
      <option selected value="XXXX"><%="N/A" %></option>
	  <% Else                     %>
      <option value="XXXX"><%="N/A" %></option>
      <% End  If                  %>

		 <% intLoopCount = 0                                     %>
         <% While (adoRS8.EOF = False) And (intLoopCount < 100)  %>
         <% 	If (InStr(strCarTypeCd2 , adoRS8.Fields("car_type_cd").Value) = 0) Or (strCarTypeCd2 = "") Then %>
         			<option value="<%=adoRS8.Fields("car_type_cd").Value %>"><%=adoRS8.Fields("car_type_cd").Value %></option>
         <% 	Else 											 %>		                    
         			<option selected value="<%=adoRS8.Fields("car_type_cd").Value %>"><%=adoRS8.Fields("car_type_cd").Value %></option>
						
		 <% 	End If 											 %>
		 <%    adoRS8.MoveNext 								     %>
		 <%    intLoopCount = intLoopCount + 1                   %>
		 <% Wend												 %> 
		 <% Set adoRS8 = Nothing 								 %>
      </select> </font></td>
      <td width="262" height="19">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="19">&nbsp;</td>
      <td width="217" height="19">&nbsp;</td>
      <td width="210" height="19">&nbsp;</td>
      <td width="510" height="19" colspan="8">
      <font face="Verdana" size="2">(use &quot;n/a&quot; to denote any car type, you can 
		use &quot;n/a&quot; for<br>
		&nbsp;both items to have the system match car types)</font></td>
      <td width="262" height="19">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="19">&nbsp;</td>
      <td width="217" height="19">&nbsp;</td>
      <td width="210" height="19"><font face="Verdana" size="2">12. Rate 
		Source::</td>
      <td width="510" height="19" colspan="8">
      <select size="1" name="data_source" style="width:200; font-family:Verdana; font-size:10pt">
	  <option selected value="ALL">Determined by profile</option>      
<!--

      <% Select Case strDataSource	%>
      
      <%	Case "CAR"	%>
	      <option selected value="CAR">Brand sites (avis.com, etc.)</option>
	      <option value="ORB">Orbitz</option>
	      <option value="EXP">Expedia</option>
	      <option value="TRV">Travelocity</option>
	      <option value="ATV">All Travel sites</option>
	      <option value="ALL">All sites</option>

      <%	Case "ORB"	%>
	      <option value="CAR">Brand sites (avis.com, etc.)</option>
	      <option selected value="ORB">Orbitz</option>
	      <option value="EXP">Expedia</option>
	      <option value="TRV">Travelocity</option>
	      <option value="ATV">All Travel sites</option>
	      <option value="ALL">All sites</option>

      <%	Case "EXP"	%>
	      <option value="CAR">Brand sites (avis.com, etc.)</option>
	      <option value="ORB">Orbitz</option>
	      <option selected value="EXP">Expedia</option>
	      <option value="TRV">Travelocity</option>
	      <option value="ATV">All Travel sites</option>
	      <option value="ALL">All sites</option>

      <%	Case "TRV"	%>
	      <option value="CAR">Brand sites (avis.com, etc.)</option>
	      <option value="ORB">Orbitz</option>
	      <option value="EXP">Expedia</option>
	      <option selected value="TRV">Travelocity</option>
	      <option value="ATV">All Travel sites</option>
	      <option value="ALL">All sites</option>

      <%	Case "ATV"	%>
	      <option value="CAR">Brand sites (avis.com, etc.)</option>
	      <option value="ORB">Orbitz</option>
	      <option value="EXP">Expedia</option>
	      <option value="TRV">Travelocity</option>
	      <option selected value="ATV">All Travel sites</option>
	      <option value="ALL">All sites</option>
      
      <%	Case "ALL"	%>
	      <option value="CAR">Brand sites (avis.com, etc.)</option>
	      <option value="ORB">Orbitz</option>
	      <option value="EXP">Expedia</option>
	      <option value="TRV">Travelocity</option>
	      <option value="ATV">All Travel sites</option>
	      <option selected value="ALL">All sites</option>


      <%	Case Else	%>
	      <option value="CAR">Brand sites (avis.com, etc.)</option>
	      <option selected value="ORB">Orbitz</option>
	      <option value="EXP">Expedia</option>
	      <option value="TRV">Travelocity</option>
	      <option value="ATV">All Travel sites</option>
	      <option value="ALL">All sites</option>


      <% End Select 				%>
-->
      </select>&nbsp;</td>
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
      <td width="210" height="22"><font face="Verdana" size="2">13. Situation:</font></td>
      <td width="510" height="22" colspan="8">
      <select size="1" name="situation_cd" style="width:370; font-family:Verdana; font-size:10pt; height:24"  onchange="showCodeOption(this)">
    <!--   
      <% If intSituationCd = 0 Then	%>
	      <option selected value="0">(None selected)</option>
	  <% Else %>
	      <option value="0">(None selected)</option>
	  <% End If %>
    -->
      <% If intSituationCd = 1 Then	%>
	      <option selected value="1">NONE - Set rate to the response amount</option>
	  <% Else %>
	      <option value="1">NONE - Set rate to the response amount</option>
	  <% End If %>
	  <!-- 	  
      <% If intSituationCd = 2 Then	%>
	      <option selected value="2">> (any comp. set)</option>
	  <% Else %>
	      <option value="2">> (any comp. set)</option>
	  <% End If %>
	  
      <% If intSituationCd = 3 Then	%>
	      <option selected value="3">> (custom)</option>
	  <% Else %>
	      <option value="3">> (custom)</option>
	  <% End If %>
	  -->
      <% If intSituationCd = 4 Then	%>
	      <option selected value="4">If rate is not more than (all comp. set) by 
			at least</option>
	  <% Else %>
	      <option value="4">If rate is not more than (all comp. set) by at least</option>
	  <% End If %>
	  
      <% If intSituationCd = 5 Then	%>
	      <option selected value="5">If rate is not more than (any comp. set) by 
			at least</option>
	  <% Else %>
	      <option value="5">If rate is not more than (any comp. set) by at least</option>
	  <% End If %>
	  
      <% If intSituationCd = 6 Then	%>
	      <option selected value="6">If rate is not more than (custom) by at 
			least</option>
	  <% Else %>
	      <option value="6">If rate is not more than (custom) by at least</option>
	  <% End If %>
	  
      <% If intSituationCd = 7 Then	%>
	      <option selected value="7">If rate is equal to (all comp. set)</option>
	  <% Else %>
	      <option value="7">If rate is equal to (all comp. set)</option>
	  <% End If %>

      <% If intSituationCd = 8 Then	%>
	      <option selected value="8">If rate is equal to (any comp. set)</option>
	  <% Else %>
	      <option value="8">If rate is equal to (any comp. set)</option>
	  <% End If %>

      <% If intSituationCd = 9 Then	%>
	      <option selected value="9">If rate is equal to (custom)</option>
	  <% Else %>
	      <option value="9">If rate is equal to (custom)</option>
	  <% End If %>
	  <!--
      <% If intSituationCd = 11 Then	%>
	      <option selected value="11">< = (all comp. set)</option>
	  <% Else %>
	      <option value="11">< = (all comp. set)</option>
	  <% End If %>

      <% If intSituationCd = 12 Then	%>
	      <option selected value="12"><= (any comp. set)</option>
	  <% Else %>
	      <option value="12"><= (any comp. set)</option>
	  <% End If %>
	  
      <% If intSituationCd = 13 Then	%>
	      <option selected value="13">&lt;= (custom)</option>
	  <% Else %>
	      <option value="13">&lt;= (custom)</option>
	  <% End If %>
	  -->
      <% If intSituationCd = 14 Then	%>
	      <option selected value="14">If rate is not less than (all comp. set) 
			by at least</option>
	  <% Else %>
	      <option value="14">If rate is not less than (all comp. set) by at 
			least</option>
	  <% End If %>
	  
      <% If intSituationCd = 16 Then	%>
	      <option selected value="16">If rate is not less than (any comp. set) 
			by at least</option>
	  <% Else %>
	      <option value="16">If rate is not less than (any comp. set) by at 
			least</option>
	  <% End If %>

      <% If intSituationCd = 15 Then	%>
	      <option selected value="15">If rate is not less than (custom) by at 
			least</option>
	  <% Else %>
	      <option value="15">If rate is not less than (custom) by at least</option>
	  <% End If %>
	  
      <% If intSituationCd = 17 Then	%>
	      <option selected value="17">If rate is not equal to (any comp. set)</option>
	  <% Else %>
	      <option value="17">If rate is not equal to (any comp. set)</option>
	  <% End If %>
	  
      <% If intSituationCd = 18 Then	%>
	      <option selected value="18">If rate is not equal to (all comp. set)</option>
	  <% Else %>
	      <option value="18">If rate is not equal to (all comp. set)</option>
	  <% End If %>
	  
      <% If intSituationCd = 19 Then	%>
	      <option selected value="19">If rate is not equal to (custom)</option>
	  <% Else %>
	      <option value="19">If rate is not equal to (custom)</option>
	  <% End If %>
	  <!--
      <% If intSituationCd = 20 Then	%>
	      <option selected value="20">If rate is not equal to  (all comp. set) + diff</option>
	  <% Else %>
	      <option value="20">If rate is not equal to  (all comp. set) + diff</option>
	  <% End If %>
 	  -->
      <% If intSituationCd = 30 Then	%>
	      <option selected value="30">If the diff. between (all comp. set) is at 
			least</option>
	  <% Else %>
	      <option value="30">If the diff. between (all comp. set) is at least</option>
	  <% End If %>
	  
      <% If intSituationCd = 31 Then	%>
	      <option selected value="31">If the diff. between (all comp. set) is 
			less than</option>
	  <% Else %>
	      <option value="31">If the diff. between (all comp. set) is less than</option>
	  <% End If %>

      <% If intSituationCd = 32 Then	%>
	      <option selected value="32">If the diff. between (any comp. set) is at 
			least</option>
	  <% Else %>
	      <option value="32">If the diff. between (any comp. set) is at least</option>
	  <% End If %>
	  
      <% If intSituationCd = 33 Then	%>
	      <option selected value="33">If the diff. between (any comp. set) is 
			less than</option>
	  <% Else %>
	      <option value="33">If the diff. between (any comp. set) is less than</option>
	  <% End If %>

      <% If intSituationCd = 34 Then	%>
	      <option selected value="34">If rate is not less than (all comp. set) 
			by exactly</option>
	  <% Else %>
	      <option value="34">If rate is not less than (all comp. set) by exactly</option>
	  <% End If %>
	  
      <% If intSituationCd = 35 Then	%>
	      <option selected value="35">If rate is not less than (any comp. set) 
			by exactly</option>
	  <% Else %>
	      <option value="35">If rate is not less than (any comp. set) by exactly</option>
	  <% End If %>

      <% If intSituationCd = 40 Then	%>
	      <option selected value="40">If (any comp. set) rate is less than</option>
	  <% Else %>
	      <option value="40">If (any comp. set) rate is less than</option>
	  <% End If %>

      <% If intSituationCd = 41 Then	%>
	      <option selected value="41">If (all comp. set) rates are less than</option>
	  <% Else %>
	      <option value="41">If (all comp. set) rate are less than</option>
	  <% End If %>

      <% If intSituationCd = 42 Then	%>
	      <option selected value="42">If (any comp. set) rate is equal to</option>
	  <% Else %>
	      <option value="42">If (any comp. set) rate is equal to</option>
	  <% End If %>

      <% If intSituationCd = 43 Then	%>
	      <option selected value="43">If (all comp. set) rates are equal to</option>
	  <% Else %>
	      <option value="43">If (all comp. set) rate are equal to</option>
	  <% End If %>

      <% If intSituationCd = 44 Then	%>
	      <option selected value="44">If (any comp. set) rate is greater than</option>
	  <% Else %>
	      <option value="44">If (any comp. set) rate is greater than</option>
	  <% End If %>

      <% If intSituationCd = 45 Then	%>
	      <option selected value="45">If (all comp. set) rates are greater than</option>
	  <% Else %>
	      <option value="45">If (all comp. set) rate are greater than</option>
	  <% End If %>

      <% If intSituationCd = 47 Then	%>
	      <option selected value="47">Is comparison rate closed?</option>
	  <% Else %>
	      <option value="47">Is comparison rate closed?</option>
	  <% End If %>



      </select>&nbsp;</td>
      <td width="262" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="23">&nbsp;</td>
      <td width="217" height="23">&nbsp;</td>
      <td width="210" height="23"><font face="Verdana" size="2">&nbsp;&nbsp;&nbsp; by amount:<br>
		&nbsp;&nbsp;&nbsp; <br>
      </font></td>
      <td width="510" height="23" colspan="8">
      <input type="text" name="situation_amt" size="20" style="width:127; font-family:Verdana; font-size:10pt; text-align:right; height:22" value="<%=strSituationAmt %>">
      <font face="Verdana" size="2">
<font color="#080000">
      <a href="javascript:centerPopUp( 'rule_situation_tester.asp', 'test', 620, 500 )">
		test situations</a>
</font></font>
<font color="#FF0000" face="Courier New" size="2">
      (beta)</font><br>
      <!--
      <% If blnIsDollar Then %>
      <input type="radio" value="1" name="is_dollar" style="font-family:Verdana; font-size:10pt" checked id="is_dollar1">    
      <font face="Verdana" size="2"><label for="is_dollar1">Dollar amount</label><br>
      <input type="radio" value="0" name="is_dollar" style="font-family:Verdana; font-size:10pt" id="is_dollar2"><font face="Verdana" size="2"><label for="is_dollar2">Percentage</label><br>
      <% Else %>
      <input type="radio" value="1" name="is_dollar" style="font-family:Verdana; font-size:10pt" id="is_dollar1">    
      <label for="is_dollar1">Dollar amount</label><br>
      <input type="radio" value="0" name="is_dollar" style="font-family:Verdana; font-size:10pt" checked  id="is_dollar2"><font face="Verdana" size="2"><label for="is_dollar2">Percentage</label><br>
      <% End If %>
      -->
      <% If blnIgnoreClosed Then %>
      <input type="checkbox" name="ignore_closed" value="True" id="ignore_closed" checked><font face="Verdana" size="2"><label for="ignore_closed">Ignore 
		closed rates</label>
	  <% Else %>
      <input type="checkbox" name="ignore_closed" value="True" id="ignore_closed" ><font face="Verdana" size="2"><label for="ignore_closed">Ignore 
		closed rates</label>
	  <% End If %><br>
      Rate to use in this rule:

      <select size="1" name="comparison_rate">
      
      <% Select Case intComparisonRate %>

		<% Case 1 %>
	      <option selected value="1">Base rate amount</option>
	      <option value="2">Total rate amount</option>
	      <option value="3">Total Price</option>

		<% Case 2 %>
	      <option value="1">Base rate amount</option>
	      <option selected value="2">Total rate amount</option>
	      <option value="3">Total Price</option>

		<% Case 3 %>
	      <option value="1">Base rate amount</option>
	      <option value="2">Total rate amount</option>
	      <option selected value="3">Total Price</option>

	  <% End Select %>


      </select>&nbsp; Tolerance: </font></font></font></font></font>
<font color="#080000">
      <input type="text" name="rt_amt_tolerance" size="20" style="width:79; font-family:Verdana; font-size:10pt; text-align:right; height:22" value="<%=strRateAmtTolerance %>"></font></td>
      
      
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
      <td width="210" height="22"><font face="Verdana" size="2">&nbsp;&nbsp;&nbsp; No. of 
		Events:</font></td>
      <td width="510" height="22" colspan="8">
      <input type="text" name="event_count" size="20" style="width:200; font-family:Verdana; font-size:10pt; text-align:right" value="<%=intEventCount %>"></td>
      <td width="262" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="22">&nbsp;</td>
      <td width="217" height="22">&nbsp;</td>
      <td width="210" height="22"><font face="Verdana" size="2">&nbsp;&nbsp;&nbsp; That happen 
		in:</font></td>
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
      <option selected value="1">Set my rate to lowest comp. set's rate minus $</option>
	  <% Else						%>
      <option value="1">Set my rate to lowest comp. set's rate minus $</option>
	  <% End If 					%>	  

	  <% If intResponseCd = 2 Then	%>
      <option selected value="2">Set my rate to lowest comp. set's rate plus $</option>
	  <% Else						%>
      <option value="2">Set my rate to lowest comp. set's rate plus $</option>
	  <% End If 					%>	  

	  <% If intResponseCd = 4 Then	%>
      <option selected value="4">Set my rate to lowest comp. set's rate minus %</option>
	  <% Else						%>
      <option value="4">Set my rate to lowest comp. set's rate minus %</option>
	  <% End If 					%>	  

	  <% If intResponseCd = 5 Then	%>
      <option selected value="5">Set my rate to lowest comp. set's rate plus %</option>
	  <% Else						%>
      <option value="5">Set my rate to lowest comp. set's rate plus %</option>
	  <% End If 					%>	  

	  <% If intResponseCd = 3 Then	%>
      <option selected value="3">Set my rate to amount</option>
	  <% Else						%>
      <option value="3">Set my rate to amount</option>
	  <% End If 					%>	  

	  <% If intResponseCd = 7 Then	%>
      <option selected value="7">Set my rate to highest comp. set's rate minus $</option>
	  <% Else						%>
      <option value="7">Set my rate to highest comp. set's rate minus $</option>
	  <% End If 					%>	  

	  <% If intResponseCd = 6 Then	%>
      <option selected value="6">Set my rate to highest comp. set's rate plus $</option>
	  <% Else						%>
      <option value="6">Set my rate to highest comp. set's rate plus $</option>
	  <% End If 					%>	  
  
	  <% If intResponseCd = 8 Then	%>
      <option selected value="8">Set my rate to highest comp. set's rate minus %</option>
	  <% Else						%>
      <option value="8">Set my rate to highest comp. set's rate minus %</option>
	  <% End If 					%>	  

	  <% If intResponseCd = 9 Then	%>
      <option selected value="9">Set my rate to highest comp. set's rate plus %</option>
	  <% Else						%>
      <option value="9">Set my rate to highest comp. set's rate plus %</option>
	  <% End If 					%>	  

	  <% If intResponseCd = 10 Then	%>
      <option selected value="10">Set my rate to comp. set avg rate minus $</option>
	  <% Else						%>
      <option value="10">Set my rate to comp. set avg rate minus $</option>
	  <% End If 					%>	  

	  <% If intResponseCd = 11 Then	%>
      <option selected value="11">Set my rate to comp. set avg rate plus $</option>
	  <% Else						%>
      <option value="11">Set my rate to comp. set avg rate plus $</option>
	  <% End If 					%>	  
  
	  <% If intResponseCd = 12 Then	%>
      <option selected value="12">Set my rate to comp. set avg rate minus %</option>
	  <% Else						%>
      <option value="12">Set my rate to comp. set avg rate minus %</option>
	  <% End If 					%>	  

	  <% If intResponseCd = 13 Then	%>
      <option selected value="13">Set my rate to comp. set avg rate plus %</option>
	  <% Else						%>
      <option value="13">Set my rate to comp. set avg rate plus %</option>
	  <% End If 					%>	  

	  <% If intResponseCd = 14 Then	%>
      <option selected value="14">Set my rate to 2nd lowest comp. set plus $</option>
	  <% Else						%>
      <option value="14">Set my rate to 2nd lowest comp. set plus $</option>
	  <% End If 					%>	 

	  <% If intResponseCd = 15 Then	%>
      <option selected value="15">Set my rate to avg. of 2 lowest comp. set plus 
		$</option>
	  <% Else						%>
      <option value="15">Set my rate to avg. of 2 lowest comp. set plus $</option>
	  <% End If 					%>	 

	  <% If strUserId = 33 Then %>

	  <% If intResponseCd = 16 Then	%>
      <option selected value="16">Gap/Cluster type 1 [5/2/1] plus $</option>
	  <% Else						%>
      <option value="16">Gap/Cluster type 1 [5/2/1] plus $</option>
	  <% End If 					%>	 

	  <% End If %>

      </select> <a href="alerts_rate_management_car_help_15.asp" onclick="window.open('alerts_rate_management_car_help_15.asp','window_name','toolbar=no,status=no,scrollbars=yes,resizable=no,width=450,height=255'); return false;"><img src="images/question.gif" border="0" alt="About response options" class="cbtip"></a></td>
      <td width="262" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="19">&nbsp;</td>
      <td width="217" height="19">&nbsp;</td>
      <td width="210" height="19"><font size="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Amount:<br>
      <br>&nbsp;</font></td>
      <td width="510" height="19" colspan="8">
      <input type="text" name="response_amt" size="20" style="width:127; font-family:Verdana; font-size:10pt; text-align:right; height:22" value="<%=strResponseAmt %>" >
      <br>
      <% If blnResponseDollar Then %>
	      <input type="radio" value="1" name="is_response_dollar" id="is_response_dollar1" style="font-family:Verdana; font-size:10pt" checked><font face="Verdana" size="2">Dollar 
		amount<br>
	      <input type="radio" value="0" name="is_response_dollar" id="is_response_dollar2"  style="font-family:Verdana; font-size:10pt">Percentage</td>
	  <% Else					   %>
	      <input type="radio" value="1" name="is_response_dollar" id="is_response_dollar1" style="font-family:Verdana; font-size:10pt" ><font face="Verdana" size="2">Whole 
		Dollar<br>
	      <input type="radio" value="0" name="is_response_dollar" id="is_response_dollar2" style="font-family:Verdana; font-size:10pt" checked>Percentage</td>
	  <% End If					   %>
      <td width="262" height="19">&nbsp;&nbsp;
      </td>
    </tr>
    <tr>
      <td width="8" height="19">&nbsp;</td>
      <td width="217" height="19">&nbsp;</td>
      <td width="210" height="19">&nbsp;</td>
      <td width="510" height="19" colspan="8">
		<font face="Verdana" size="2" color="#080000">&nbsp;<a href="javascript:toggleLayer('ExtraDay');" title="Add extra day and hour rates to this rule">Show 
		Extra Rate detail</a></font></td>
      <td width="262" height="19">&nbsp;</td>
    </tr>
</table>
	<div id="ExtraDay">
  <table border="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#111111" width="1210" id="new_alert_extra_day" background="images/alt_color.gif">
  	<tr>
	  <td width="8" height="19">&nbsp;</td>
      <td width="217" height="19">&nbsp;</td>
      <td width="188" height="19">&nbsp;</td>
      <td width="204" height="19">&nbsp;</td>
      <td width="274" height="19">&nbsp;</td>
      <td width="319" height="19">&nbsp;</td>
  	
  	</tr>
    <tr>
      <td width="8" height="19">&nbsp;</td>
      <td width="217" height="19">&nbsp;</td>
      <td width="188" height="19"></td>
      <td width="204" height="19" bgcolor="#C0C0C0">
      <font size="2">15a. Extra Day rate:</font></td>
      <td width="274" height="19" bgcolor="#C0C0C0">
      <input type="text" name="extra_day_rt" size="20" value="<%=curExtraDayRt %>" style="text-align: right" onBlur="this.value=formatCurrency(this.value);"></td>
      <td width="319" height="19">&nbsp;&nbsp;
      </td>
    </tr>
    <tr>
      <td width="8" height="19">&nbsp;</td>
      <td width="217" height="19">&nbsp;</td>
      <td width="188" height="19">&nbsp;</td>
      <td width="204" height="19" bgcolor="#C0C0C0">
      <font size="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Extra Day free miles:</font></td>
      <td width="274" height="19" bgcolor="#C0C0C0">
      <font face="Verdana" size="2" color="#080000">
		<input type="text" name="extra_day_miles" size="20" value="<%=curExtraDayMiles %>" style="text-align: right" onBlur="this.value=formatNumber(this.value);"></font><font face="Verdana" size="2" color="#080000">
		(blank = unlimited)</font></td>
      <td width="319" height="19">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="19">&nbsp;</td>
      <td width="217" height="19">&nbsp;</td>
      <td width="188" height="19">&nbsp;</td>
      <td width="204" height="19" bgcolor="#C0C0C0">
      <font size="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Extra Day $/extra mile:</font></td>
      <td width="274" height="19" bgcolor="#C0C0C0">
      <font face="Verdana" size="2" color="#080000">
		<input type="text" name="extra_day_rt_per_mile" size="20" value="<%=curExtraDayRtPerMile %>" style="text-align: right" onBlur="this.value=formatCurrency(this.value);"></font></td>
      <td width="319" height="19">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="19">&nbsp;</td>
      <td width="217" height="19">&nbsp;</td>
      <td width="188" height="19">&nbsp;</td>
      <td width="204" height="19">&nbsp;</td>
      <td width="274" height="19">&nbsp;</td>
      <td width="319" height="19">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="19">&nbsp;</td>
      <td width="217" height="19">&nbsp;</td>
      <td width="188" height="19">&nbsp;</td>
      <td width="204" height="19" bgcolor="#C0C0C0">
      <font size="2">15b. Extra Hour rate:</font></td>
      <td width="274" height="19" bgcolor="#C0C0C0">
      <font face="Verdana" size="2" color="#080000">
		<input type="text" name="extra_hr_rt" size="20" value="<%=curExtraHrRt %>" style="text-align: right" onBlur="this.value=formatCurrency(this.value);"></font></td>
      <td width="319" height="19">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="19">&nbsp;</td>
      <td width="217" height="19">&nbsp;</td>
      <td width="188" height="19">&nbsp;</td>
      <td width="204" height="19" bgcolor="#C0C0C0">
      <font size="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Extra Hour free miles:</font></td>
      <td width="274" height="19" bgcolor="#C0C0C0">
      <font face="Verdana" size="2" color="#080000">
		<input type="text" name="extra_hr_miles" size="20" value="<%=curExtraHrMiles %>" style="text-align: right" onBlur="this.value=formatNumber(this.value);"> 
		(blank = unlimited)</font></td>
      <td width="319" height="19">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="19">&nbsp;</td>
      <td width="217" height="19">&nbsp;</td>
      <td width="188" height="19">&nbsp;</td>
      <td width="204" height="19" bgcolor="#C0C0C0">
      <font size="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Extra Hour $/extra mile:</font></td>
      <td width="274" height="19" bgcolor="#C0C0C0">
      <font face="Verdana" size="2" color="#080000">
		<input type="text" name="extra_hr_rt_per_mile" size="20" value="<%=curExtraHrRtPerMile %>" style="text-align: right" onBlur="this.value=formatCurrency(this.value);"></font></td>
      <td width="319" height="19">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="19">&nbsp;</td>
      <td width="217" height="19">&nbsp;</td>
      <td width="188" height="19">&nbsp;</td>
      <td width="204" height="19">
<!--      <font size="2" ><input type="reset" name="reset" value="Hide Extra Rate detail" onclick="javascript:toggleLayer('ExtraDay');" style="float: right" /></font>
-->
		</td>
      <td width="274" height="19">
      &nbsp;</td>
      <td width="319" height="19">&nbsp;</td>
    </tr>
</table> 
</div>
  <table border="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#111111" width="1210" id="new_alert_part2" background="images/alt_color.gif" height="561">
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
      <td width="210" height="22"><font face="Verdana" size="2">&nbsp;&nbsp;&nbsp;&nbsp; Profile(s):<br>
      <br></font></td>
      <td width="510" height="22" colspan="8">
		<select name="profile_id" size="4" multiple style="width:373; font-family:Verdana; font-size:10pt; height:70">
				<% strProfileID = "," & strProfileID & "," %>
				<% strProfileID = Replace(strProfileID, " ", "") %>
				<% intLoopCount = 0 %>
                <% While (adoRS9.EOF = False) And (intLoopCount < 100)  %> 
                	<% If adoRS9.Fields("profile_id").Value = 0 Then %>
	                	<% adoRS9.MoveNext %>
	                <% End If %>
                	<% If InStr(1, strProfileID, "," & adoRS9.Fields("profile_id").Value & ",") > 0 Then %>
	                	<option selected value="<%=adoRS9.Fields("profile_id").Value %>"><%=adoRS9.Fields("desc").Value %></option>
	                <% Else %>
		               <% If adoRS9.Fields("profile_status").Value = "E" Then %>
		                	<option value="<%=adoRS9.Fields("profile_id").Value %>"><%=adoRS9.Fields("desc").Value %></option>
		                <% End If %> 
					<% End If %>                
                <%	adoRS9.MoveNext
                    intLoopCount = intLoopCount + 1
				   Wend
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
      <td width="210" height="22"><font face="Verdana" size="2">&nbsp;&nbsp;&nbsp;&nbsp; Maximum:</font></td>
      <td width="510" height="22" colspan="8">
      <input type="text" name="rate_maximum" size="20" style="width:127; font-family:Verdana; font-size:10pt; text-align:right; height:21" value="<%=strRangeMax %>" 
      onkeyup='this.onchange();' onchange='tV=(this.value.replace(/[^\d\.]/g,"")).replace(/[\.]{2,}/g,".");if(tV!=this.value){this.value=tV;}'></td>
      <td width="262" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="22">&nbsp;</td>
      <td width="217" height="22">
      &nbsp;</td>
      <td width="210" height="25"><font face="Verdana" size="2">&nbsp;&nbsp;&nbsp;&nbsp; Minimum:</font></td>
      <td width="510" height="25" colspan="8">
      <input type="text" name="rate_minimum" size="20" style="width:127; font-family:Verdana; font-size:10pt; text-align:right; height:22" value="<%=strRangeMin %>"></td>
      <td width="262" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="22">&nbsp;</td>
      <td width="217" height="22">
      &nbsp;</td>
      <td width="210" height="22" align="center"><font face="Verdana" size="2">
		OR</font></td>
      <td width="510" height="22" colspan="8">
      &nbsp;</td>
      <td width="262" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="22">&nbsp;</td>
      <td width="217" height="22">
      &nbsp;</td>
      <td width="210" height="22">
      &nbsp;&nbsp;&nbsp; <font size="2">Max. / Min. Schedule</font></td>
      <td width="510" height="22" colspan="8">
      <font face="Verdana" size="2">
		<select name="maxmin_profile_id" style="width:370; font-family:Verdana; font-size:10pt; height:24" size="1">
			<% If intMaxMinProfileId = 0 Then %>
     		<option selected value="0">(None - Use above buckets)</option>
     		<% Else %>
     		<option value="0">(None - Use above buckets)</option>
			<% End If %>
			<% intLoopCount = 0                                      %>
     		<% While (adoRS11.EOF = False) And (intLoopCount < 100)  %>
				<% If intMaxMinProfileId = adoRS11.Fields("car_rate_rule_schedule_id").Value Then %>
		   		<option selected value="<%=adoRS11.Fields("car_rate_rule_schedule_id").Value %>"><%=adoRS11.Fields("car_rate_rule_schedule_desc").Value  %></option>
	     		<% Else %>
		   		<option value="<%=adoRS11.Fields("car_rate_rule_schedule_id").Value %>"><%=adoRS11.Fields("car_rate_rule_schedule_desc").Value  %></option>
				<% End If %>
	   			<% adoRS11.MoveNext
	   			   intLoopCount = intLoopCount + 1
	   		    %>
     		<% Wend             %>
     		
		</select>&nbsp;
		<a title="Click to edit or create a rule schedule" href="rate_rule_maxmin_schedule_a.asp" onclick="window.open('rate_rule_maxmin_schedule_a.asp','manage_schedule','toolbar=no,status=no,scrollbars=yes,resizable=no,width=820,height=600'); return false;">
		Manage schedules</a></font></td>
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
      <td width="210" height="22"><font face="Verdana" size="2">18. When 
		situation = true:</font></td>
      <td width="510" height="22" colspan="8">
      <select size="1" name="on_success_id" style="width:370; font-family:Verdana; font-size:10pt; height:24">
      <option selected value="0">(No follow-on rule)</option>
      			<% intLoopCount = 0 %>
                <% While (adoRS18a.EOF = False) And (intLoopCount < 100)  %> 
                	<% If intSuccessId = adoRS18a.Fields("rate_rule_id").Value Then %>
	                	<option selected value="<%=adoRS18a.Fields("rate_rule_id").Value %>"><%=adoRS18a.Fields("alert_desc").Value %></option>
	                <% Else %>
		               <% If adoRS18a.Fields("rule_status").Value = "E" Then %>
		                	<option value="<%=adoRS18a.Fields("rate_rule_id").Value %>"><%=adoRS18a.Fields("alert_desc").Value %></option>
		                <% End If %> 
					<% End If %>                
                <%	adoRS18a.MoveNext
                    intLoopCount = intLoopCount + 1
				   Wend
					   
				   Set adoRS18a = Nothing
				%>
      
      </select></td>
      <td width="262" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="22">&nbsp;</td>
      <td width="217" height="22">&nbsp;</td>
      <td width="210" height="22"><font face="Verdana" size="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;When 
		situation = false:</font></td>
      <td width="510" height="22" colspan="8">
		<select name="on_failure_id" style="width:370; font-family:Verdana; font-size:10pt; height:24" size="1">
     <option selected value="0">(No follow-on rule)</option>
                <% intLoopCount = 0 %>
                <% While (adoRS18b.EOF = False) And (intLoopCount < 100)  %> 
                	<% If intFailureId = adoRS18b.Fields("rate_rule_id").Value Then %>
	                	<option selected value="<%=adoRS18b.Fields("rate_rule_id").Value %>"><%=adoRS18b.Fields("alert_desc").Value %></option>
	                <% Else %>
		               <% If adoRS18b.Fields("rule_status").Value = "E" Then %>
		                	<option value="<%=adoRS18b.Fields("rate_rule_id").Value %>"><%=adoRS18b.Fields("alert_desc").Value %></option>
		                <% End If %> 
					<% End If %>                
                <%	adoRS18b.MoveNext
                    intLoopCount = intLoopCount + 1
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
      <td width="64" height="22" align="center"><font face="Verdana" size="2">
		Same</font></td>
      <td width="64" height="22" align="center"><font face="Verdana" size="2">
		Next</font></td>
      <td width="64" height="22" align="center"><font face="Verdana" size="2">2 
		- 4</font></td>
      <td height="22" align="center"><font face="Verdana" size="2">5 - 7</font></td>
      <td width="64" height="22" align="center"><font face="Verdana" size="2">8 
		- 14</font></td>
      <td width="64" height="22" align="center"><font face="Verdana" size="2">15 
		- 30</font></td>
      <td width="64" height="22" align="center"><font face="Verdana" size="2">31 
		- 50</font></td>
      <td width="63" height="22" align="center"><font face="Verdana" size="2">51 
		+</font></td>
	  <td width="262" height="22" >&nbsp;</td>
      <td height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" style="height: 7px"></td>
      <td width="217" style="height: 7px"></td>
      <td width="210" style="height: 7px"><font face="Verdana" size="2">&nbsp;&nbsp;&nbsp;&nbsp; 
		Maximum:</td>
      <td width="64" style="height: 7px"><font face="Verdana" size="2"><input type="text"  name="util_max_0" size="20" style="width:59; font-family:Verdana; font-size:10pt; text-align:right; height:22" value="<%=strUtilMax0 %>"></font></td>
      <td width="64" style="height: 7px"><font face="Verdana" size="2"><input type="text"  name="util_max_1" size="20" style="width:59; font-family:Verdana; font-size:10pt; text-align:right; height:22" value="<%=strUtilMax1 %>"></font></td>
      <td width="64" style="height: 7px"><font face="Verdana" size="2"><input type="text"  name="util_max_2" size="20" style="width:59; font-family:Verdana; font-size:10pt; text-align:right; height:22" value="<%=strUtilMax2 %>"></font></td>
      <td style="height: 7px"><font face="Verdana" size="2"><input type="text"  name="util_max_3" size="20" style="width:59; font-family:Verdana; font-size:10pt; text-align:right; height:22" value="<%=strUtilMax3 %>"></font></td>
      <td width="64" style="height: 7px"><font face="Verdana" size="2"><input type="text"  name="util_max_4" size="20" style="width:59; font-family:Verdana; font-size:10pt; text-align:right; height:22" value="<%=strUtilMax4 %>"></font></td>
      <td width="64" style="height: 7px"><font face="Verdana" size="2"><input type="text"  name="util_max_5" size="20" style="width:59; font-family:Verdana; font-size:10pt; text-align:right; height:22" value="<%=strUtilMax5 %>"></font></td>
      <td width="64" style="height: 7px"><font face="Verdana" size="2"><input type="text"  name="util_max_6" size="20" style="width:59; font-family:Verdana; font-size:10pt; text-align:right; height:22" value="<%=strUtilMax6 %>"></font></td>
      <td width="63" style="height: 7px"><font face="Verdana" size="2"><input type="text"  name="util_max_7" size="20" style="width:59; font-family:Verdana; font-size:10pt; text-align:right; height:22" value="<%=strUtilMax7 %>"></font></td>
	  <td width="262" style="height: 7px" ><font face="Verdana" size="2"></font></td>
	  <td style="height: 7px"></td>
    </tr>
    <tr>
      <td width="8" style="height: 24px"></td>
      <td width="217" style="height: 24px"></td>
      <td width="210" style="height: 24px"><font face="Verdana" size="2">&nbsp;&nbsp;&nbsp;&nbsp; 
		Minimum:</td>
      <td width="64" style="height: 24px"><font face="Verdana" size="2"><input type="text" name="util_min_0" size="20" style="width:59; font-family:Verdana; font-size:10pt; text-align:right; height:22" value="<%=strUtilMin0 %>"></font></td>
      <td width="64" style="height: 24px"><font face="Verdana" size="2"><input type="text" name="util_min_1" size="20" style="width:59; font-family:Verdana; font-size:10pt; text-align:right; height:22" value="<%=strUtilMin1 %>"></font></td>
      <td width="64" style="height: 24px"><font face="Verdana" size="2"><input type="text" name="util_min_2" size="20" style="width:59; font-family:Verdana; font-size:10pt; text-align:right; height:22" value="<%=strUtilMin2 %>"></font></td>
      <td style="height: 24px"><font face="Verdana" size="2"><input type="text" name="util_min_3" size="20" style="width:59; font-family:Verdana; font-size:10pt; text-align:right; height:22" value="<%=strUtilMin3 %>"></font></td>
      <td width="64" style="height: 24px"><font face="Verdana" size="2"><input type="text" name="util_min_4" size="20" style="width:59; font-family:Verdana; font-size:10pt; text-align:right; height:22" value="<%=strUtilMin4 %>"></font></td>
      <td width="64" style="height: 24px"><font face="Verdana" size="2"><input type="text" name="util_min_5" size="20" style="width:59; font-family:Verdana; font-size:10pt; text-align:right; height:22" value="<%=strUtilMin5 %>"></font></td>
      <td width="64" style="height: 24px"><font face="Verdana" size="2"><input type="text" name="util_min_6" size="20" style="width:59; font-family:Verdana; font-size:10pt; text-align:right; height:22" value="<%=strUtilMin6 %>"></font></td>
      <td width="63" style="height: 24px"><font face="Verdana" size="2"><input type="text" name="util_min_7" size="20" style="width:59; font-family:Verdana; font-size:10pt; text-align:right; height:22" value="<%=strUtilMin7 %>"></font></td>
	  <td width="262" style="height: 24px"><font face="Verdana" size="2"></font></td>
      <td style="height: 24px"></td>
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
      <td width="210" height="22" align="center"><font face="Verdana" size="2">
		OR</font></td>
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
		
			<% If intUtilizationProfileId = 0 Then %>
     		<option selected value="0">(None - Use above buckets)</option>
     		<% Else %>
     		<option value="0">(None - Use above buckets)</option>
			<% End If %>
            <% intLoopCount = 0 %>
     		<% While (adoRS10.EOF = False) And (intLoopCount < 100)  %>
     		
				<% If intUtilizationProfileId = adoRS10.Fields("car_rate_rule_schedule_id").Value Then %>
		   		<option selected value="<%=adoRS10.Fields("car_rate_rule_schedule_id").Value %>"><%=adoRS10.Fields("car_rate_rule_schedule_desc").Value  %></option>
	     		<% Else %>
		   		<option value="<%=adoRS10.Fields("car_rate_rule_schedule_id").Value %>"><%=adoRS10.Fields("car_rate_rule_schedule_desc").Value  %></option>
				<% End If %>
     		
     		
	   			<% adoRS10.MoveNext 
	   			   intLoopCount = intLoopCount + 1
	   			%>
     		<% Wend             %>
		</select>&nbsp;
		<a title="Click to edit or create a rule schedule" href="rate_rule_schedule_a.asp" onclick="window.open('rate_rule_schedule_a.asp','manage_schedule','toolbar=no,status=no,scrollbars=yes,resizable=no,width=820,height=600'); return false;">
		Manage schedules</a></font></td>
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
      <font face="Verdana" size="2" color="#080000">
      20. Pickup Days of Week:</td>
      <td width="510" height="22" colspan="8">
      <% If InStr(1, strDowList, "2") Then %>
      <input type="checkbox" name="dow_list" value="2" checked>
      <% Else %>
      <input type="checkbox" name="dow_list" value="2" >
      <% End If %>
      <font size="2">Mon&nbsp; </font>

      <% If InStr(1, strDowList, "3") Then %>
      <input type="checkbox" name="dow_list" value="3" checked>
      <% Else %>
      <input type="checkbox" name="dow_list" value="3" >
      <% End If %>
	  <font size="2">Tue&nbsp; </font>
	  
      <% If InStr(1, strDowList, "4") Then %>
      <input type="checkbox" name="dow_list" value="4" checked>
      <% Else %>
      <input type="checkbox" name="dow_list" value="4" >
      <% End If %>
	  <font size="2">Wed&nbsp; </font>

      <% If InStr(1, strDowList, "5") Then %>
      <input type="checkbox" name="dow_list" value="5" checked>
      <% Else %>
      <input type="checkbox" name="dow_list" value="5" >
      <% End If %>
	  <font size="2">Thu&nbsp; </font>

      <% If InStr(1, strDowList, "6") Then %>
      <input type="checkbox" name="dow_list" value="6" checked>
      <% Else %>
      <input type="checkbox" name="dow_list" value="6" >
      <% End If %>
	  <font size="2">Fri&nbsp; </font>

      <% If InStr(1, strDowList, "7") Then %>
      <input type="checkbox" name="dow_list" value="7" checked>
      <% Else %>
      <input type="checkbox" name="dow_list" value="7" >
      <% End If %>
	  <font size="2">Sat&nbsp; </font>

      <% If InStr(1, strDowList, "1") Then %>
      <input type="checkbox" name="dow_list" value="1" checked>
      <% Else %>
      <input type="checkbox" name="dow_list" value="1" >
      <% End If %>
	  <font size="2">Sun&nbsp; </font>
		  
	  
	</td>
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
      <td width="210" height="22" class="style1">
      21. Optional Post Actions:</td>
      <td height="22" colspan="4" class="style2">
      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Action Description</td>
      <td width="510" height="22" colspan="4" style="width: 255px" class="style1">
      &nbsp;&nbsp;&nbsp;&nbsp; Action Amt.</td>
      <td width="262" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" style="height: 22px"></td>
      <td width="217" style="height: 22px"></td>
      <td width="210" style="height: 22px">
      &nbsp;</td>
      <td colspan="2" style="height: 22px">
	<font color="#080000">
      <span class="style1">First Action:</span></font></td>
      <td colspan="2" style="height: 22px">
	<font color="#080000">
      <select name="rule_post_action_id1">
        <% Select Case intRulePostActionId1 %>
		<% Case 1 %>
		<option value="0">(none)</option>
		<option selected="" value="1">Subtract $ amt.</option>
		<option value="2">Add $ amt.</option>
		<option value="3">Subtract %</option>
		<option value="4">Add %</option>
		<% Case 2 %>
		<option value="0">(none)</option>
		<option value="1">Subtract $ amt.</option>
		<option selected="" value="2">Add $ amt.</option>
		<option value="3">Subtract %</option>
		<option value="4">Add %</option>
		<% Case 3 %>
		<option value="0">(none)</option>
		<option value="1">Subtract $ amt.</option>
		<option value="2">Add $ amt.</option>
		<option selected="" value="3">Subtract %</option>
		<option value="4">Add %</option>
		<% Case 4 %>
		<option value="0">(none)</option>
		<option value="1">Subtract $ amt.</option>
		<option value="2">Add $ amt.</option>
		<option value="3">Subtract %</option>
		<option selected="" value="4">Add %</option>
		<% Case Else %>
		<option selected="" value="0">(none)</option>
		<option value="1">Subtract $ amt.</option>
		<option value="2">Add $ amt.</option>
		<option value="3">Subtract %</option>
		<option value="4">Add %</option>
		<% End Select %>
	  </select></font></td>
      <td width="510" colspan="4" style="height: 22px; width: 255px">
		<font color="#080000">
		<input	name="rule_post_action_amt1" 
				type="text" 
				value="<%=strRulePostActionAmt1 %>" 
				style="width:127; font-family:Verdana; font-size:10pt; text-align:right; height:22"
				onblur="extractNumber(this,4,true);"
				onkeyup="extractNumber(this,4,true);"
				onkeypress="return blockNonNumbers(this, event, true, true);">
		</font></td>
      <td width="262" style="height: 22px"></td>
    </tr>
    <tr>
      <td width="8" height="22">&nbsp;</td>
      <td width="217" height="22">&nbsp;</td>
      <td width="210" height="22">
      &nbsp;</td>
      <td height="22" colspan="2">
      <font color="#080000"><span class="style1">Second Action:</span></font></td>
      <td height="22" colspan="2">
      <font color="#080000">
		<select name="rule_post_action_id2">
        <% Select Case intRulePostActionId2 %>
		<% Case 1 %>
		<option value="0">(none)</option>
		<option selected="" value="1">Subtract $ amt.</option>
		<option value="2">Add $ amt.</option>
		<option value="3">Subtract %</option>
		<option value="4">Add %</option>
		<% Case 2 %>
		<option value="0">(none)</option>
		<option value="1">Subtract $ amt.</option>
		<option selected="" value="2">Add $ amt.</option>
		<option value="3">Subtract %</option>
		<option value="4">Add %</option>
		<% Case 3 %>
		<option value="0">(none)</option>
		<option value="1">Subtract $ amt.</option>
		<option value="2">Add $ amt.</option>
		<option selected="" value="3">Subtract %</option>
		<option value="4">Add %</option>
		<% Case 4 %>
		<option value="0">(none)</option>
		<option value="1">Subtract $ amt.</option>
		<option value="2">Add $ amt.</option>
		<option value="3">Subtract %</option>
		<option selected="" value="4">Add %</option>
		<% Case Else %>
		<option selected="" value="0">(none)</option>
		<option value="1">Subtract $ amt.</option>
		<option value="2">Add $ amt.</option>
		<option value="3">Subtract %</option>
		<option value="4">Add %</option>
		<% End Select %>
		</select></font></td>
      <td width="510" height="22" colspan="4" style="width: 255px">
      <font color="#080000"> 
		<input	name="rule_post_action_amt2" 
				type="text" 
				value="<%=strRulePostActionAmt2 %>" 
				style="width:127; font-family:Verdana; font-size:10pt; text-align:right; height:22"
				onblur="extractNumber(this,4,true);"
				onkeyup="extractNumber(this,4,true);"
				onkeypress="return blockNonNumbers(this, event, true, true);">
	  </font></td>
      <td width="262" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="8"></td>
      <td width="217"></td>
      <td width="210">
      </td>
      <td colspan="2">
      <font color="#080000"><span class="style1">Third Action:</span></font></td>
      <td colspan="2">
      <font color="#080000">
		<select name="rule_post_action_id3">
        <% Select Case intRulePostActionId3 %>
		<% Case 1 %>
		<option value="0">(none)</option>
		<option selected="" value="1">Subtract $ amt.</option>
		<option value="2">Add $ amt.</option>
		<option value="3">Subtract %</option>
		<option value="4">Add %</option>
		<% Case 2 %>
		<option value="0">(none)</option>
		<option value="1">Subtract $ amt.</option>
		<option selected="" value="2">Add $ amt.</option>
		<option value="3">Subtract %</option>
		<option value="4">Add %</option>
		<% Case 3 %>
		<option value="0">(none)</option>
		<option value="1">Subtract $ amt.</option>
		<option value="2">Add $ amt.</option>
		<option selected="" value="3">Subtract %</option>
		<option value="4">Add %</option>
		<% Case 4 %>
		<option value="0">(none)</option>
		<option value="1">Subtract $ amt.</option>
		<option value="2">Add $ amt.</option>
		<option value="3">Subtract %</option>
		<option selected="" value="4">Add %</option>
		<% Case Else %>
		<option selected="" value="0">(none)</option>
		<option value="1">Subtract $ amt.</option>
		<option value="2">Add $ amt.</option>
		<option value="3">Subtract %</option>
		<option value="4">Add %</option>
		<% End Select %>
		</select></font></td>
      <td width="510" colspan="4" style="width: 255px">
      <font color="#080000"> 
			<input	name="rule_post_action_amt3" 
				type="text" 
				value="<%=strRulePostActionAmt3 %>" 
				style="width:127; font-family:Verdana; font-size:10pt; text-align:right; height:22"
				onblur="extractNumber(this,4,true);"
				onkeyup="extractNumber(this,4,true);"
				onkeypress="return blockNonNumbers(this, event, true, true);"></font></td>
      <td width="262"></td>
    </tr>
    <tr>
      <td width="8" height="22">&nbsp;</td>
      <td width="217" height="22">&nbsp;</td>
      <td width="210" height="22">
      &nbsp;</td>
      <td height="22" colspan="2">
      <font color="#080000"><span class="style1">Forth Action:</span></font></td>
      <td height="22" colspan="2">
      <font color="#080000">
		<select name="rule_post_action_id4">
        <% Select Case intRulePostActionId4 %>
		<% Case 1 %>
		<option value="0">(none)</option>
		<option selected="" value="1">Subtract $ amt.</option>
		<option value="2">Add $ amt.</option>
		<option value="3">Subtract %</option>
		<option value="4">Add %</option>
		<% Case 2 %>
		<option value="0">(none)</option>
		<option value="1">Subtract $ amt.</option>
		<option selected="" value="2">Add $ amt.</option>
		<option value="3">Subtract %</option>
		<option value="4">Add %</option>
		<% Case 3 %>
		<option value="0">(none)</option>
		<option value="1">Subtract $ amt.</option>
		<option value="2">Add $ amt.</option>
		<option selected="" value="3">Subtract %</option>
		<option value="4">Add %</option>
		<% Case 4 %>
		<option value="0">(none)</option>
		<option value="1">Subtract $ amt.</option>
		<option value="2">Add $ amt.</option>
		<option value="3">Subtract %</option>
		<option selected="" value="4">Add %</option>
		<% Case Else %>
		<option selected="" value="0">(none)</option>
		<option value="1">Subtract $ amt.</option>
		<option value="2">Add $ amt.</option>
		<option value="3">Subtract %</option>
		<option value="4">Add %</option>
		<% End Select %>
		</select></font></td>
      <td width="510" height="22" colspan="4" style="width: 255px">
      <font color="#080000"> 
 	  <input	name="rule_post_action_amt4" 
				type="text" 
				value="<%=strRulePostActionAmt4 %>" 
				style="width:127; font-family:Verdana; font-size:10pt; text-align:right; height:22"
				onblur="extractNumber(this,4,true);"
				onkeyup="extractNumber(this,4,true);"
				onkeypress="return blockNonNumbers(this, event, true, true);">
				</font></td>
      <td width="262" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="22">&nbsp;</td>
      <td width="217" height="22">&nbsp;</td>
      <td width="210" height="22">
      &nbsp;</td>
      <td width="510" height="22" colspan="8" class="style1">
      <em>(Note: please do not use a % sign. For example 10.5% would be 10.5)</em></td>
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
      <font face="Verdana" size="2" color="#080000">
      22. Auto-pilot mode:</font></td>
      <td width="510" height="22" colspan="8">
      <% If blnAutomatic Then %>
      <input type="checkbox" name="automatic" value="True"  checked="True" >
      <% Else %>
      <input type="checkbox" name="automatic" value="True" >
      <% End If %>
      <font size="2">Run this rule in automatic mode (no user review or approval 
		req.)</font></td>
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
      <td width="210" height="23">
      <font size="2" >
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
      <font size="2">Rate rule tester 
		<a target="_blank" href="rate_rule_tester_20130718.asp">click here</a></font>
      <font color="#FF0000" face="Courier New" size="2">
      (beta)</font></td>
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
<td  width="1" bgcolor="#000000"><img src="images/pixel.gif" width="1" height="1"></td>
</tr>
<tr bgcolor="#000000" height="1">
<td colspan=5><img src="images/pixel.gif" width="1" height="1"></td>
</tr>
</table>
<!-- JUSTTABS BOTTOM CLOSE -->
<!--#INCLUDE FILE="footer.asp"-->
<div id="calbox" class="calboxoff"></div>	
<%
	Rem Debug
		   
'	Response.Write("Your search has returned ")
'	Response.Write(intRecordCount)
'	Response.Write(" records and ")
'	Response.Write(intPageCount)
'	Response.Write(" pages for userid " & strUserId & "<br>")
'	Response.Write strProfileID 

%>
</body>

</html>

<%	
	Rem Clean up
	Set adoCmd = Nothing 
	Set adoCmd1 = Nothing 
	Set adoCmd2 = Nothing 
	Set adoCmd3 = Nothing 
	Set adoCmd9 = Nothing 

		'adoRS4.Close
	Set adoRS4 = Nothing
	Set adoCmd4 = Nothing

	Set adoRS5 = Nothing
 
    Set adoRS9 = Nothing
	
%>

<script language="javascript">
	document.create_alert.alert_desc.focus();
</script>