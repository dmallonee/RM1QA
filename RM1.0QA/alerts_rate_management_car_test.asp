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
	Dim blnExtraday
	
	'Declare variables
	Dim iCurrentPage
	Dim intPageSize
	Dim i
	Dim oConnection
	Dim oRecordSet
	Dim oTableField
	Dim sPageURL
	Dim strEditMode 
	Dim strCustomResponse
	Dim strCustomSituation
	

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
	intRuleId = Request.QueryString("rateruleid")	
	
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

	Rem If this page has been called for edit mode, then don't both with gettign the rules since we won't show
	Rem them - they are instead show on the new ASPX version of this page.
	If strEditMode <> "1" Then

		Rem Get the rule
		Set adoCmd4 = CreateObject("ADODB.Command")
	
		adoCmd4.ActiveConnection =  strConn
		adoCmd4.CommandText = "car_rate_rule_select"
		adoCmd4.CommandType = adCmdStoredProc
		adoCmd4.Parameters.Append adoCmd4.CreateParameter("@rate_rule_id",      3, 1, 0, Null)
		adoCmd4.Parameters.Append adoCmd4.CreateParameter("@user_id",           3, 1, 0, strUserId)
		adoCmd4.Parameters.Append adoCmd4.CreateParameter("@rate_rule_type_cd", 3, 1, 0, 1)
		adoCmd4.Parameters.Append adoCmd4.CreateParameter("@rule_status",     200, 1, 1, strRuleStatus)
		adoCmd4.Parameters.Append adoCmd4.CreateParameter("@sandbox",          11, 1, 0, False)
		
		'Create an ADO RecordSet object
		Set adoRS4 = Server.CreateObject("ADODB.Recordset")
	
		'If adoRS4.State = adStateOpen Then
				'Set the RecordSet PageSize property
				adoRS4.PageSize = intPageSize
				adoRS4.CursorLocation = adUseClient 
	
				'Set the RecordSet CacheSize property to the
				'number of records that are returned on each page of results
				adoRS4.CacheSize = intPageSize
		
				'Open the RecordSet
				adoRS4.Open adoCmd4, , adOpenStatic, adLockReadOnly
	
				If adoRS4.EOF = False Then
					adoRS4.PageSize = intPageSize
					adoRS4.MoveFirst
					intPageCount = adoRS4.PageCount
					intRecordCount = adoRS4.RecordCount
					adoRS4.AbsolutePage = intCurrentPage
				Else
					intPageCount = 1
					intRecordCount = 0
	
				End If	
						
				'Set adoRS4 = adoCmd4.Execute
		
		'Else
		'	intPageCount = 100
		'	intRecordCount = 0
		
		'End If
	
	
		'   response.write "Error test - 3 <br>"
		
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

	Else
		intPageCount = 1
		intRecordCount = 0
	
	End If

	Set adoRS18a = adoCmd4.Execute
	Set adoRS18b = adoCmd4.Execute
			
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

	Set adoCmd9 = CreateObject("ADODB.Command")

	adoCmd9.ActiveConnection =  strConn
	adoCmd9.CommandText = "car_rate_rule_schedule_select"
	adoCmd9.CommandType = adCmdStoredProc

	adoCmd9.Parameters.Append adoCmd9.CreateParameter("@org_id",              3, 1, 0, Null)
	adoCmd9.Parameters.Append adoCmd9.CreateParameter("@schedule_type_id",    3, 1, 0, 6)
	adoCmd9.Parameters.Append adoCmd9.CreateParameter("@user_id",             3, 1, 0, strUserId)
		
	Set adoRS12 = adoCmd9.Execute

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
		blnExtraday = adoRS5.Fields("xday_processing").value
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
		
		intAdditionalDays = adoRS5.Fields("additional_days").Value
		strCustomResponse = adoRS5.Fields("custom_response").Value
		strCustomSituation = adoRS5.Fields("custom_situation").Value
		intAdditionalDaysPrior = adoRS5.Fields("additional_days_prior").Value
		bolSetToSelf = CBool(adoRS5.Fields("set_to_self").Value)

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
		strCarTypeCd1 = "XXXX"
		strCarTypeCd2 = "XXXX"
		strDataSource = "EXP"
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
		intAdditionalDays = 0
		intAdditionalDaysPrior = 0
		blnExtraday = False
		strCustomResponse = ""
		strCustomSituation = ""
		bolSetToSelf = False
		
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
<META HTTP-EQUIV="refresh" CONTENT="900;URL=default_session.asp"> 
<title>Rate-Monitor by Rate-Highway, Inc. | Alerts! | Rate Management</title>
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
		document.create_alert.action = "car_rate_rule_insert.asp?debug=false";
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

function showLayerResponse(SelectString)
	{
	var myLayer = document.getElementById("CustomResponse").style.display;
	if (SelectString == 500)
		{
		document.getElementById("CustomResponse").style.display="block";
		} 
	else 
		{ 
		document.getElementById("CustomResponse").style.display="none";
		}
		
		
	}
	
	
function showLayerSituation(SelectString)
	{
	var myLayer = document.getElementById("CustomSituation").style.display;
	if (SelectString == 500)
		{
		document.getElementById("CustomSituation").style.display="block";
		} 
	else 
		{ 
		document.getElementById("CustomSituation").style.display="none";
		}
		
		
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
      showLayerSituation(obj.options[obj.selectedIndex].value);
    }


//  End -->
</script>
<style type="text/css" >
<!--
.profile_header { height="48" text-align:left; padding-left:3; padding-right:3; padding-top:0; background-color:#879AA2; font-family:Verdana, Arial, Helvetica, sans-serif; font-size:10pt; vertical-align:bottom; text-align:left }
.profile_light { height="68" text-align:left; padding-left:3; padding-right:3; padding-top:3; background-color:#CFD7DB; font-family:Verdana, Arial, Helvetica, sans-serif; font-size:10pt; vertical-align:bottom; text-align:left }
.profile_dark { height="48" text-align:left; padding-left:3; padding-right:3; padding-top:3; background-color:#C0C0C0; font-family:Verdana, Arial, Helvetica, sans-serif; font-size:10pt; vertical-align:bottom; text-align:left }
div#ExtraDay { margin: 0px 20px 0px 20px; display: none; }
div#CustomResponse { margin: 0px 0px 0px 0px; display: none; }
div#CustomSituation { margin: 0px 0px 0px 0px; display: none; }
<link rel="stylesheet" type="text/css" href="inc/rh_standard.css">
<link rel="stylesheet" type="text/css" href="inc/rh_report.css">
<link rel="stylesheet" type="text/css" href="inc/sitewide.css">
<link rel="stylesheet" type="text/css" href="inc/css_calendar_v2.css" id="calendarcss" media="all" >
.link {
	font-size: small;
}
.style2 {
	font-size: small;
	text-align: left;
}
td {
	font-family: Verdana;
	font-size: small;
	color: #000000;
	font-weight: normal;
}
body {
	font-family: Verdana;
	font-size: small;
	color: #000000;
	font-weight: normal;
}
-->
</style>
</head>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="MM_preloadImages('images/b_search_pro_on.gif','images/b_search_que_on.gif','images/b_search_cri_on.gif','images/b_rate_on.gif','images/b_alert_on.gif','images/b_user_on.gif','images/b_system_on.gif')">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td background="images/top_tile.gif">
    <a target="_blank" href="http://www.rate-highway.com">
    <img alt=""src="images/top.jpg" width="770" height="91" border="0" ></a></td>
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
    <img alt=""src="images/med_bar.gif" width="12" height="8"></td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td background="images/user_tile.gif">
    <table width="100" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td><img alt=""src="images/user_left.gif" width="580" height="31"></td>
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
                <td><img alt=""src="images/separator.gif" width="183" height="6"></td>
              </tr>
            </table>
            </td>
            <td><img alt=""src="images/user_tile.gif" width="7" height="31"></td>
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
        <td><img alt=""src="images/h_blank.gif" width="368" height="31"></td>
        <td><map name="logout_map">
        <area alt="Click to log out of Rate-Monitor"  href="http://www.rate-monitor.com/" shape="rect" coords="337, 10, 394, 25">
        </map>
        <img alt=""src="images/bottom_right.gif" width="402" height="31" usemap="#logout_map" border="0"></td>
      </tr>
    </table>
    </td>
  </tr>
</table>

<p align="right">&nbsp;</p>
    <div align="center">
<table cellpadding="0" cellspacing="0" border="0" width="100%" bgcolor="#FFFFFF">
<tr height="1">
<td >&nbsp;</td>
</tr>
<tr height="1">
<td bgcolor="#000000" colspan="1" height="1"><img alt=""src="images/pixel.gif" width="1" height="1"></td>
</tr>
</table>
</div>
<table cellpadding="0" cellspacing="0" border="0" ALIGN="CENTER" width="100%" bgcolor="#FFFFFF">
<tr >
<td  width="1" bgcolor="#000000"><img alt=""src="images/pixel.gif" width="1" height="1"></td>
<td colspan=3 bgcolor="#D9DEE1">
<table border="0" cellspacing="5" cellpadding="5">
<tr><td>
<font color="#080000">
<br>
<!-- JUSTTABS TOP OPEN-END -->
&nbsp;
<% If strEditMode <> "1" Then %>
<form method="GET" name="search_alerts" class="search">

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
          <img alt=""border="0" src="images/search.GIF"></td>
          <td width="583" colspan="3" height="51">
          To search for an Rule, enter a login id, or a portion of the 
          address. You may also enter the alert type.</td>
          <td width="336" height="51">&nbsp;</td>
        </tr>
        <tr valign="bottom">
          <td width="10" height="26">&nbsp;</td>
          <td width="179" height="26">&nbsp;</td>
          <td width="177" height="26">Owner:</td>
          <td width="80" height="26">
          
          <input type="text" name="name" size="20" style="width:150" style="width:150" onfocus="this.className='focus';cl(this,'<%=strClientUserid %>');" onblur="this.className='';fl(this,'<%=strClientUserid %>');"></td>          <td width="662" colspan="2" height="26">
          
          <input name="search" type="submit" id="Open2224" value="    Search    " class="rh_button"></td>        </tr>
        <tr valign="bottom">
          <td width="10" height="22">&nbsp;</td>
          <td width="179" height="22">&nbsp;</td>
          <td width="177" height="22">
          Rule 
          Status:</td>          <td width="80" height="22">
          <select size="1" name="rule_status" style="border:1px solid #000000; width:150; background-color:#FF9933">
		  <% Select Case strRuleStatus %>
		  <%	Case "D" %>
          <option  value="A" >All types</option>
          <option value="E">Enabled only</option>
          <option selected value="D">Disabled only</option>
		  <%	Case "E" %>
          <option  value="A" >All types</option>
          <option selected value="E">Enabled only</option>
          <option value="D">Disabled only</option>
		  <%	Case "A" %>
          <option selected value="A" >All types</option>
          <option value="E">Enabled only</option>
          <option value="D">Disabled only</option>
		  <%	Case Else %>
          <option  value="A" >All types</option>
          <option selected value="E">Enabled only</option>
          <option value="D">Disabled only</option>
          
          <% End Select %>
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
    <tr valign="bottom">
      <td >
    <% If intCurrentPage<= 1 Then %>
	    |&lt;
	    &lt; 
	<% Else  %>
	    <a href="alerts_rate_management_car.asp?user_id=<%=strClientUserid %>&city=<%=strCity%>&car_type=<%=strCarType%>&rule_status=<%=strRuleStatus %>&page=1">|&lt;</a>
	    <a href="alerts_rate_management_car.asp?user_id=<%=strClientUserid %>&city=<%=strCity%>&car_type=<%=strCarType%>&rule_status=<%=strRuleStatus %>&page=<%=intCurrentPage - 1%>">&lt;</a> 
    <% End If %>
    Page <%=intCurrentPage%> of <%=intPageCount %>
    <% If CInt(intCurrentPage) >= CInt(intPageCount) Then %>
	    &gt;
	    &gt;|
    <% Else %>
	    <a href="alerts_rate_management_car.asp?user_id=<%=strClientUserid %>&city=<%=strCity%>&car_type=<%=strCarType%>&rule_status=<%=strRuleStatus %>&page=<%=intCurrentPage + 1%>">&gt;</a>
	    <a href="alerts_rate_management_car.asp?user_id=<%=strClientUserid %>&city=<%=strCity%>&car_type=<%=strCarType%>&rule_status=<%=strRuleStatus %>&page=<%=intPageCount %>">&gt;|</a>
    <% End If %>
      
      
      
      
      
      </td>
    </tr>
  </table>
  <table border="0" style="border-collapse: collapse" bordercolor="#FFFFFF" width="1310" height="4">
    <tr>
      <td background="images/ruler.gif"></td>
    </tr>
  </table>
     <form name="maint" method="POST" action="car_rate_rule_maint.asp">
  <input type="hidden" name="action" value="1"><input type="hidden" name="refresh_from" value="search">
  <table border="1" style="border-collapse: collapse" bordercolor="#FFFFFF" width="1310" id="profiles">
    <tr>
      <td align="left" valign="bottom" bgcolor="#879AA2" height="45" width="30">&nbsp;</td>
      <td class="profile_header" width="58" style="background-color: #E07D1A" height="45">
      Selected</td>      <td class="profile_header" width="291" height="45">
		Description</td>      <td class="profile_header" width="46" height="45">Rate Code</td>      <td class="profile_header" width="288" height="45">Situation 
      (if you situation is listed as a number, please contact 
      support@rate-highway.com)</td>      <td class="profile_header" width="102" height="45">Location(s)</td>      <td class="profile_header" width="100" height="45">
      Recipient or Response</td>      <td class="profile_header" width="342" height="45">Search Type / Profile</td>    </tr>
    
  <%
        
        Dim strClass
        Dim strOrange
        Dim intCount

		If adoRS4.State = adStateOpen Then

		While (adoRS4.EOF = False) And (intCount <= intPageSize)
		
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
    <td class="<%=strClass %>" height="20">
	<%=adoRS4.Fields("rate_rule_id").Value  & LCASE(adoRS4.Fields("rule_status").Value)%></td>
    <td bgcolor="#FDC677" align="center" height="20">
    <input type="checkbox" value="<%=adoRS4.Fields("rate_rule_id").Value %>" name="rate_rule_id"></td>
    <td class="<%=strClass %>" height="20" width="360" nowrap >
    <a target="_self" title="<%=adoRS4.Fields("alert_desc").Value %>" href="alerts_rate_management_car.asp?rateruleid=<%=adoRS4.Fields("rate_rule_id").Value %>#rule_maint">
   	<% If Len(adoRS4.Fields("alert_desc").Value) > 35 Then %>
	  <%=Left(adoRS4.Fields("alert_desc").Value, 35) & "..." %>
	<% Else %>
	  <%=adoRS4.Fields("alert_desc").Value %>
	<% End If %>
    </a></td>
    <td class="<%=strClass %>" height="20" width="46">
    <%=adoRS4.Fields("client_sys_rate_cd").Value %></td>    <td class="<%=strClass %>" height="20" width="288">
    <% Select Case adoRS4.Fields("situation_cd").Value %>

   <% Case 1 %>
    <font face="Verdana" size="1">NONE - Set rate to the response amount</font>

   <% Case 4 %>
    <font face="Verdana" size="1">If rate is not more than (all comp. set) by at least</font>

   <% Case 5 %>
    <font face="Verdana" size="1">If rate is not more than (any comp. set) by at least</font>

    <% Case 14 %>
    <font face="Verdana" size="1">If rate is not less than (all comp. set) by at least</font>

   <% Case 18 %>
    <font face="Verdana" size="1">If rate is not equal to (all comp. set)</font>

   <% Case 16 %>
    <font face="Verdana" size="1">If rate is not less than (any comp. set) by at least</font>

   <% Case 34 %>
    <font face="Verdana" size="1">If rate is not less than (all comp. set) by exactly</font>
	  
   <% Case 35 %>
    <font face="Verdana" size="1">If rate is not less than (any comp. set) by exactly</font>

   <% Case 40 %>
    <font face="Verdana" size="1">If (any comp. set's) rate is less than</font>

   <% Case 43 %>
    <font face="Verdana" size="1">If (all comp. set) rates are equal to</font>

   <% Case 32 %>
    <font face="Verdana" size="1">If the diff. between (any comp. set) is at least</font>

   <% Case 42 %>
    <font face="Verdana" size="1">If (any comp. set) rate is equal to</font>

   <% Case 44 %>
    <font face="Verdana" size="1">If (any comp. set) rate is greater than</font>
    
   <% Case 45 %>
    <font face="Verdana" size="1">If (all comp. set) rate is greater than</font>
    
   <% Case 46 %>
    <font face="Verdana" size="1">If gap between two lowest competitors</font>

   <% Case 47 %>
    <font face="Verdana" size="1">Is comparison rate closed?</font>

   <% Case 48 %>
    <font face="Verdana" size="1">Is comp set rate closed?</font>

   <% Case 50 %>
    <font face="Verdana" size="1">Are two or more competitors open?</font>
    
   <% Case 51 %>
    <font face="Verdana" size="1">Is competitive car greater than comparison?</font>
    
    <% Case Else %>
    <%=adoRS4.Fields("situation_cd").Value %>
    <% End Select %>
    
    
    </td>
    <td class="<%=strClass %>" height="20" width="102">
	<% If adoRS4.Fields("city_cd").Value = "" Then %>
	 Any
	<% Else %>
		<% If Len(adoRS4.Fields("city_cd").Value) > 8 Then %>
		  <%=Left(adoRS4.Fields("city_cd").Value, 8) & "..." %>
		<% Else %>
		  <%=adoRS4.Fields("city_cd").Value %>
		<% End If %>
	<% End If %>
	</td>
    <td class="<%=strClass %>" height="20">Display</td>
    <td class="<%=strClass %>" height="20">
    <% If IsNull(adoRS4.Fields("shop_profile_desc").Value) Then %>
		All searched rates - no profile assoc.
	<% Else %>
	    <%=adoRS4.Fields("shop_profile_desc").Value %>
	<% End If %>
	</td>
   
  </tr>
  <%
        
        	adoRS4.MoveNext
        	
        Wend
        
        End If
        

  %>
    
    
    </table>
 
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
  <p>&nbsp;| <a href="javascript:maint_action(1);">Delete</a> 
  | <!-- <a href="javascript:maint_action(2)">Copy</a> | -->
  <a href="javascript:maint_action(3)">Enable</a> |
  <a href="javascript:maint_action(4)">Disable</a> |
  <a href="javascript:maint_action(5)">Move to Sandbox</a> | 
  <a target="_blank" href="alerts_rate_management_export.asp">Download cross-reference</a>
  | <a target="_blank" href="alerts_rate_management_export_worksheet.asp">Download rule worksheet for upload</a>
  | <a target="_blank" href="rule_upload.asp">Upload rules</a> 
  | Utilization Levels Rpt. 
  |	<a target="_self" href="alerts_rate_management_car_grid.asp">New Grid</a> </form>
<form name="display_disabled">
<br>
&nbsp;</form>
<% End If %>
<form name="create_alert" method="POST" action="" OnSubmit="return CreateAlert()"  >
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
      <td width="8" height="19"><a name="rule_maint"></a>&nbsp;</td>
      <td width="217" height="19">&nbsp;</td>
      <td width="210" height="19">&nbsp;</td>
      <td height="19" colspan="3">&nbsp;</td>
      <td width="262" height="19">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="25">&nbsp;</td>
      <td width="217" height="25"><img border="0" src="images/maintenance.GIF" width="162" height="25" alt=""></td>
      <td width="210" height="25">1. Rate Change Rule No.:</td>
      <td height="25" colspan="3">
        <input type="text" name="rate_rule_id" size="20" style="width:200; font-family:Verdana; font-size:10pt; text-align:right; background-image:url('images/alt_color.gif')" value="<%=intRuleId %>" READONLY>&nbsp; 
        <input type="checkbox" name="copy" id="copy" value="true"><label for="copy">Save as a copy (leaves the original unchanged)</label></td>
      <td width="262" height="25">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="25">&nbsp;</td>
      <td width="217" height="25">
      &nbsp;</td>
      <td width="210" height="25">&nbsp;</td>
      <td height="25" colspan="3">
      &nbsp;</td>
      <td width="262" height="25">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="25">&nbsp;</td>
      <td width="217" height="25">
      &nbsp;</td>
      <td width="210" height="25">2. Description:</td>      
		<td height="25" colspan="3">
      <input type="text" name="alert_desc" size="20" style="width:439; font-family:Verdana; font-size:10pt; height:21" value="<%=strAlertDesc %>"></td>
      <td width="262" height="25">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="25">&nbsp;</td>
      <td width="217" height="25">
      &nbsp;</td>
      <td width="210" height="25">&nbsp;</td>
      <td height="25" colspan="3">
      &nbsp;</td>
      <td width="262" height="25">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="25">&nbsp;</td>
      <td width="217" height="25">
      &nbsp;</td>
      <td width="210" height="25">3. Rule Begin Date:</td>      
		<td height="25" colspan="3">
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
      <td height="25" colspan="3">
      (enter 'continuous' or blank for no begin date) </td>      <td width="262" height="25">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="25">&nbsp;</td>
      <td width="217" height="25">
      &nbsp;</td>
      <td width="210" height="25">4. Rule End Date:</td>      
		<td height="25" colspan="3">
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
      <td height="25" colspan="3">
      (enter 'continuous' or blank for no end date)</td>      <td width="262" height="25">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="25">&nbsp;</td>
      <td width="217" height="25">
      &nbsp;</td>
      <td width="210" height="25">5. First 
      Pick-up:</td>      <td height="25" colspan="3">
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
      <label for="rolling_dates">Rolling begin &amp; end</label>      
      </td>
      <td height="25" colspan="3">
      (enter 'continuous' or blank for no pick-up date)</td>      <td width="262" height="25">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="25">&nbsp;</td>
      <td width="217" height="25">
      &nbsp;</td>
      <td width="210" height="25">6. Last Pick-up:</td>      
		<td height="25" colspan="3">
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
      <td height="25" colspan="3">
      (enter 'continuous' or blank for no pick-up date)</td>      <td width="262" height="25">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="25">&nbsp;</td>
      <td width="217" height="25">
      &nbsp;</td>
      <td width="210" height="25">7. System Rate 
      Code:</td>      <td height="25" colspan="3">
      <input type="text" name="client_sys_rate_cd" size="20" style="width:200; font-family:Verdana; font-size:10pt" value="<%=strClientRateCode %>"></td>
      <td width="262" height="25">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="25">&nbsp;</td>
      <td width="217" height="25">
      &nbsp;</td>
      <td width="210" height="25">&nbsp;</td>
      <td height="25" colspan="3">
      (the rate code used within your system, i.e. Daily, Weekly, etc.)</td>
      <td width="262" height="25">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="25">&nbsp;</td>
      <td width="217" height="25">
      &nbsp;</td>
      <td width="210" height="25">8. Car 
      Companies:</td>      <td height="25" colspan="3">&nbsp;</td>
      <td width="262" height="25">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="25">&nbsp;</td>
      <td width="217" height="25">
      &nbsp;</td>
      <td width="210" height="25" style="padding-top: 2px" >
      &nbsp;&nbsp;&nbsp;Competitive Set:</td>      <td width="256" height="25">
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


     
      </td>
      <td width="125" height="25" bgcolor="#FFFFFF" bordercolor="#555566" bordercolorlight="#999999" bordercolordark="#777777" style="border: 1px solid #40618F; padding: 2px">
      <p align="left">
      <img alt=""border="0" src="images/tip_ballon.gif" width="23" height="22"><b> 
      Note</b>: Your competitive set can be one or more comp. set. Just 
      CTRL+Click to select multiple companies      
      </td>
      <td height="25" bordercolor="#555566" style="width: 191px">
      &nbsp;</td>
      <td width="262" height="25">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="25">&nbsp;</td>
      <td width="217" height="25">
      &nbsp;</td>
      <td width="210" height="25" valign="top" style="padding-top: 2px">
      &nbsp;</td>
      <td height="25" colspan="3">
      &nbsp;</td>
      <td width="262" height="25">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="25">&nbsp;</td>
      <td width="217" height="25">
      &nbsp;</td>
      <td width="210" height="25" style="padding-top: 2px">
      &nbsp;&nbsp; Comparison Company<br>
&nbsp;&nbsp; (usually self):</td>      <td width="256" height="25">
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
<td width="125" height="25" bgcolor="#FFFFFF" bordercolor="#555566" bordercolorlight="#999999" bordercolordark="#777777" style="border: 1px solid #40618F; padding: 2px">
      <p align="left">
      <img alt=""border="0" src="images/tip_ballon.gif" width="23" height="22"><b> 
      Note</b>: If you select a group as the comparison company, the lowest rate 
      of the group will be used </td>
      <td height="25" bordercolor="#555566" style="width: 191px">
      &nbsp;</td>      
      <td width="262" height="25">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="25">&nbsp;</td>
      <td width="217" height="25">
      &nbsp;</td>
      <td width="210" height="25" valign="top" style="padding-top: 2px">
      &nbsp;</td>
      <td height="25" colspan="3">
      (use &quot;any&quot; to 
      denote any company, you may not use &quot;any&quot; for <br>
      both items)</td>      <td width="262" height="25">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="25">&nbsp;</td>
      <td width="217" height="25">
      &nbsp;</td>
      <!--
      <td width="210" height="25" valign="top" style="padding-top: 2px">
      9. LOR(s):</td>      <td width="510" height="25" colspan="3">
      <select size="4" name="lor" style="width:200; font-family:Verdana; font-size:10pt" >
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
      -->
      <td width="262" height="25">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="25">&nbsp;</td>
      <td width="217" height="25">
      &nbsp;</td>
      <td width="210" height="25" valign="top" style="padding-top: 2px">&nbsp;</td>
      <td height="25" colspan="3">
      &nbsp;</td>
      <td width="262" height="25">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="71">&nbsp;</td>
      <td width="217" height="71">
      &nbsp;</td>
      <td width="210" height="71" valign="top" style="padding-top: 2px">
      10. Location(s):</td>      
		<td height="71" valign="top" colspan="3">
      <select size="4" name="city_cd" style="width:200; font-family:Verdana; font-size:10pt" multiple >
     	 <% intLoopCount = 0                                     %>
         <% While (adoRS6.EOF = False) And (intLoopCount < 800)  %>
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
      <td height="24" colspan="3">
      (select 
      airport/city codes &quot;any&quot; for any location, edit to edit custom)</td>      <td width="262" height="24">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="25">&nbsp;</td>
      <td width="217" height="25">
      &nbsp;</td>
      <td width="210" height="25" style="vertical-align: top; padding-top: 2px">
      11. Car Types:<br>&nbsp;&nbsp;&nbsp;&nbsp;
	  <a href="bulk_update_car_types.asp" target="_self">bulk update</a></td>      
		<td height="25" colspan="3">
      &nbsp;</td>
      <td width="262" height="25">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="25">&nbsp;</td>
      <td width="217" height="25">
      &nbsp;</td>
      <td width="210" height="25" style="padding-top: 2px">
      &nbsp;&nbsp;&nbsp;&nbsp; Competitive Car Type(s):</td>      
		<td height="25" colspan="3">
      <p style="margin-top: 2px">
      
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

	 <% If strCarTypeCd1 = "----" Then %>
      <option selected value="----"><%="Ignore" %></option>
	  <% Else                     %>
      <option value="----"><%="Ignore" %></option>
     <% End  If                  %> 

      
      </select> </td>      <td width="262" height="25">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="25">&nbsp;</td>
      <td width="217" height="25">
      &nbsp;</td>
      <td width="210" height="25" style="vertical-align: top; padding-top: 2px">
      &nbsp;</td>
      <td height="25" colspan="3">
      &nbsp;</td>
      <td width="262" height="25">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="19">&nbsp;</td>
      <td width="217" height="19">&nbsp;</td>
      <td width="210" height="19">&nbsp;&nbsp;&nbsp;&nbsp; 
		Suggestion Car<br>&nbsp;&nbsp;&nbsp;&nbsp; (Comparison)&nbsp;&nbsp;&nbsp;<br>
&nbsp;&nbsp;&nbsp;&nbsp; Type(s):</td>      
		<td height="19" colspan="3">
      
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
		 
		 
		 
     </select> </td>      <td width="262" height="19">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="19">&nbsp;</td>
      <td width="217" height="19">&nbsp;</td>
      <td width="210" height="19">&nbsp;</td>
      <td height="19" colspan="3">
      (use &quot;n/a&quot; to 
      denote any car type, you can use &quot;n/a&quot; for<br>
&nbsp;both items to have the system match car types)</td>      <td width="262" height="19">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="19">&nbsp;</td>
      <td width="217" height="19">&nbsp;</td>
      <td width="210" height="19">12. Rate Source::</td>
      <td height="19" colspan="3">
      <select size="1" name="data_source" style="width:200; font-family:Verdana; font-size:10pt">
	  <option selected value="ALL">Determined by profile</option>      
      </select>&nbsp;</td>
      <td width="262" height="19">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="19">&nbsp;</td>
      <td width="217" height="19">&nbsp;</td>
      <td width="210" height="19">&nbsp;</td>
      <td height="19" colspan="3">&nbsp;</td>
      <td width="262" height="19">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="22">&nbsp;</td>
      <td width="217" height="22">&nbsp;</td>
      <td width="210" height="22">13. Situation:</td>      
		<td height="22" colspan="3">
      <select size="1" name="situation_cd" style="width:370; font-family:Verdana; font-size:10pt; height:24" onchange="showCodeOption(this)">
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
	      <option selected value="4">If rate is not more than (all comp. set) by at least</option>
	  <% Else %>
	      <option value="4">If rate is not more than (all comp. set) by at least</option>
	  <% End If %>
	  
      <% If intSituationCd = 5 Then	%>
	      <option selected value="5">If rate is not more than (any comp. set) by at least</option>
	  <% Else %>
	      <option value="5">If rate is not more than (any comp. set) by at least</option>
	  <% End If %>
	  <!-- 
      <% If intSituationCd = 6 Then	%>
	      <option selected value="6">If rate is not more than (custom) by at least</option>
	  <% Else %>
	      <option value="6">If rate is not more than (custom) by at least</option>
	  <% End If %>
	   -->
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
	      <option selected value="14">If rate is not less than (all comp. set) by at least</option>
	  <% Else %>
	      <option value="14">If rate is not less than (all comp. set) by at least</option>
	  <% End If %>
	  
      <% If intSituationCd = 16 Then	%>
	      <option selected value="16">If rate is not less than (any comp. set) by at least</option>
	  <% Else %>
	      <option value="16">If rate is not less than (any comp. set) by at least</option>
	  <% End If %>

      <% If intSituationCd = 15 Then	%>
	      <option selected value="15">If rate is not less than (custom) by at least</option>
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
	      <option selected value="30">If the diff. between (all comp. set) is at least</option>
	  <% Else %>
	      <option value="30">If the diff. between (all comp. set) is at least</option>
	  <% End If %>
	  
      <% If intSituationCd = 31 Then	%>
	      <option selected value="31">If the diff. between (all comp. set) is less than</option>
	  <% Else %>
	      <option value="31">If the diff. between (all comp. set) is less than</option>
	  <% End If %>

      <% If intSituationCd = 32 Then	%>
	      <option selected value="32">If the diff. between (any comp. set) is at least</option>
	  <% Else %>
	      <option value="32">If the diff. between (any comp. set) is at least</option>
	  <% End If %>
	  
      <% If intSituationCd = 33 Then	%>
	      <option selected value="33">If the diff. between (any comp. set) is less than</option>
	  <% Else %>
	      <option value="33">If the diff. between (any comp. set) is less than</option>
	  <% End If %>

      <% If intSituationCd = 34 Then	%>
	      <option selected value="34">If rate is not less than (all comp. set) by exactly</option>
	  <% Else %>
	      <option value="34">If rate is not less than (all comp. set) by exactly</option>
	  <% End If %>
	  
      <% If intSituationCd = 35 Then	%>
	      <option selected value="35">If rate is not less than (any comp. set) by exactly</option>
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
	  
      <% If intSituationCd = 46 Then	%>
	      <option selected value="46">If gap between two lowest competitors is greater than</option>
	  <% Else %>
	      <option value="46">If gap between two lowest competitors is greater than</option>
	  <% End If %>

      <% If intSituationCd = 47 Then	%>
	      <option selected value="47">Is comparison rate closed?</option>
	  <% Else %>
	      <option value="47">Is comparison rate closed?</option>
	  <% End If %>

      <% If intSituationCd = 48 Then	%>
	      <option selected value="48">Is comp set rate closed?</option>
	  <% Else %>
	      <option value="48">Is comp set rate closed?</option>
	  <% End If %>

      <!-- What the heck does this one do?-->
      <% If intSituationCd = 49 Then	%>
	      <option selected value="49">Create suggestion for extra car type</option>
	  <% Else %>
	      <option value="49">Create suggestion for extra car type</option>
	  <% End If %>
	  
      <% If intSituationCd = 50 Then	%>
	      <option selected value="50">Are two or more competitors open?</option>
	  <% Else %>
	      <option value="50">Are two or more competitors open?</option>
	  <% End If %>

      <% If intSituationCd = 51 Then	%>
	      <option selected value="51">Is competitive car greater than comparison?</option>
	  <% Else %>
	      <option value="51">Is competitive car greater than comparison?</option>
	  <% End If %>

      <% If intSituationCd >= 500 Then	%>
	      <option selected value="500">Custom rule sitution</option>
	  <% Else %>
	      <option value="500">Custom rule situation</option>
	  <% End If %>


      </select>&nbsp;</td>
      <td width="262" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="23">&nbsp;</td>
      <td width="217" height="23">&nbsp;</td>
      <td width="210" height="23">&nbsp;&nbsp;&nbsp; 
      by amount:<br>
&nbsp;&nbsp;&nbsp; <br>
      </td>      <td height="23" colspan="3">
      <input type="text" name="situation_amt" size="20" style="width:127; font-family:Verdana; font-size:10pt; text-align:right; height:22" value="<%=strSituationAmt %>">
      <a href="javascript:centerPopUp( 'rule_situation_tester.asp', 'test', 620, 500 )">test situations</a>
<font color="#FF0000" face="Courier New" size="2">
      (beta)</font><br>
      <% If blnIsDollar Then %>
      <input type="radio" value="1" name="is_dollar" style="font-family:Verdana; font-size:10pt" checked id="is_dollar1">    
      <label for="is_dollar1">Dollar amount</label><br>
      <input type="radio" value="0" name="is_dollar" style="font-family:Verdana; font-size:10pt" id="is_dollar2"><label for="is_dollar2">Percentage</label><br>
      <% Else %>
      <input type="radio" value="1" name="is_dollar" style="font-family:Verdana; font-size:10pt" id="is_dollar1">    
      <label for="is_dollar1">Dollar amount</label><br>
      <input type="radio" value="0" name="is_dollar" style="font-family:Verdana; font-size:10pt" checked  id="is_dollar2"><label for="is_dollar2">Percentage</label><br>
      <% End If %>
	  <% If blnIgnoreClosed Then %>
      <input type="checkbox" name="ignore_closed" value="True" id="ignore_closed" checked>
	  <% Else %>
      <input type="checkbox" name="ignore_closed" value="True" id="ignore_closed" >
      <% End If %><label for="ignore_closed">Do not suggest when comparison 
		company is closed</label>
	  <br>
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


      </select>&nbsp; Tolerance: 
      <input type="text" name="rt_amt_tolerance" size="20" style="width:79; font-family:Verdana; font-size:10pt; text-align:right; height:22" value="<%=strRateAmtTolerance %>"></td>            
      <td width="262" height="23">&nbsp;</td>
    </tr>
    </table>
    <div id="CustomSituation" name="CustomSituation" >
    <table border="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#111111" width="1210" id="custom_situation_table" background="images/alt_color.gif">
    <tr>
      <td width="8" height="19">&nbsp;</td>
      <td width="217" height="19">&nbsp;</td>
      <td width="210" height="19"><font color="#080000">&nbsp;&nbsp;&nbsp;Custom:</font></td>
		<td width="510" height="19" colspan="3"><font color="#080000">
		<input type="text" name="custom_Situation" size="20" style="width:370; font-family:Verdana; font-size:10pt; text-align:left; height:22" value="<%=strCustomSituation %>" ></font><td width="262" height="19">
		</td>
    </tr>
    </table>
    </div>
   <table border="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#111111" width="1210" id="filler_table1" background="images/alt_color.gif">
    <tr>
      <td width="8" height="19">&nbsp;</td>
      <td width="217" height="19">&nbsp;</td>
      <td width="210" height="19">&nbsp;</td>
      <td width="510" height="19" colspan="3">&nbsp;</td>
      <td width="262" height="19">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="22">&nbsp;</td>
      <td width="217" height="22">&nbsp;</td>
      <td width="210" height="22">14. Quantity &amp; 
      Period:</td>      <td width="510" height="22" colspan="3">
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
      <td width="210" height="22">&nbsp;&nbsp;&nbsp; 
      No. of Events:</td>      <td width="510" height="22" colspan="3">
      <input type="text" name="event_count" size="20" style="width:200; font-family:Verdana; font-size:10pt; text-align:right" value="<%=intEventCount %>"></td>
      <td width="262" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="22">&nbsp;</td>
      <td width="217" height="22">&nbsp;</td>
      <td width="210" height="22">&nbsp;&nbsp;&nbsp; 
      That happen in:</td>      <td width="510" height="22" colspan="3">
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
      <td width="510" height="19" colspan="3">&nbsp;</td>
      <td width="262" height="19">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="22">&nbsp;</td>
      <td width="217" height="22">&nbsp;</td>
      <td width="210" height="22">15. Response:</td>      
		<td width="510" height="22" colspan="3">
      <select size="1" name="response_cd" style="width:370; font-family:Verdana; font-size:10pt; height:24" onChange="showLayerResponse(this.options[this.selectedIndex].value);">  <!--     <select size="1" name="response_cd" style="width:200; font-family:Verdana; font-size:10pt" onclick="CheckResponse()" >
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
	  
	  <% If intResponseCd = 24 Then	%>
      <option selected value="24">Set my rate to 2nd lowest comp. set plus %</option>
	  <% Else						%>
      <option value="24">Set my rate to 2nd lowest comp. set plus %</option>
	  <% End If 					%>		  
	  
	  <% If intResponseCd = 23 Then	%>
      <option selected value="23">Set my rate to 3rd lowest comp. set plus $</option>
	  <% Else						%>
      <option value="23">Set my rate to 3rd lowest comp. set plus $</option>
	  <% End If 					%>	  

	  <% If intResponseCd = 25 Then	%>
      <option selected value="25">Set my rate to 3rd lowest comp. set plus %</option>
	  <% Else						%>
      <option value="25">Set my rate to 3rd lowest comp. set plus %</option>
	  <% End If 					%>	
	  
	  <% If intResponseCd = 15 Then	%>
      <option selected value="15">Set my rate to avg. of 2 lowest comp. set plus $</option>
	  <% Else						%>
      <option value="15">Set my rate to avg. of 2 lowest comp. set plus $</option>
	  <% End If 					%>	 

	  <% If intResponseCd = 26 Then	%>
      <option selected value="26">Set my rate to avg. of 2 lowest comp. set plus %</option>
	  <% Else						%>
      <option value="26">Set my rate to avg. of 2 lowest comp. set plus %</option>
	  <% End If 					%>	

	  <% If strUserId = 33 Then %>

	  <% If intResponseCd = 16 Then	%>
      <option selected value="16">Gap/Cluster type 1 [5/2/1] plus $</option>
	  <% Else						%>
      <option value="16">Gap/Cluster type 1 [5/2/1] plus $</option>
	  <% End If 					%>	 

	  <% End If %>

	  <% If intResponseCd = 17  Then	%>
      <option selected value="17">Close date</option>
	  <% Else						%>
      <option value="17">Close date</option>
	  <% End If 					%>	 

	  <% If intResponseCd = 18 Then	%>
      <option selected value="18">Open date</option>
	  <% Else						%>
      <option value="18">Open date</option>
	  <% End If 					%>	 

	  <% If intResponseCd = 19 Then	%>
      <option selected value="19">Set to max</option>
	  <% Else						%>
      <option value="19">Set to max</option>
	  <% End If 					%>	 

	  <% If intResponseCd = 20 Then	%>
      <option selected value="20">Generate drop charge</option>
	  <% Else						%>
      <option value="20">Generate drop charge</option>
	  <% End If 					%>	 
	  
	  <% If intResponseCd = 21 Then	%>
      <option selected value="21">Set my rate to avg. of 2nd &amp; 3rd lowest comp. set plus $</option>
	  <% Else						%>
      <option value="21">Set my rate to avg. of 2nd &amp; 3rd lowest comp set plus $</option>
	  <% End If 					%>	  

	  <% If intResponseCd = 22 Then	%>
      <option selected value="22">No Change</option>
	  <% Else						%>
      <option value="22">No Change</option>
	  <% End If 					%>	  


	  <% If intResponseCd >= 500 Then %>
      <option selected value="500">Custom Rule Response</option>
	  <% Else						%>
      <option value="500">Custom Rule Response</option>
	  <% End If 					%>	 


      </select> <a href="alerts_rate_management_car_help_15.asp" onclick="window.open('alerts_rate_management_car_help_15.asp','window_name','toolbar=no,status=no,scrollbars=yes,resizable=no,width=450,height=255'); return false;"><img alt=""src="images/question.gif" border="0" alt="About response options" class="cbtip"></a></td>
      <td width="262" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="19"></td>
      <td width="217" height="19"></td>
      <td width="210" height="19">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Amount:&nbsp; <br>
      					  </td><td width="510" height="19" colspan="3"><input type="text" name="response_amt" size="20" style="width:127; font-family:Verdana; font-size:10pt; text-align:right; height:22" value="<%=strResponseAmt %>" ><td width="262" height="19">
							&nbsp;</td>
    </tr>
    </table>
    
    <div id="CustomResponse" name="CustomResponse" >
    <table border="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#111111" width="1210" id="custom_response_table" background="images/alt_color.gif">
    <tr>
      <td width="8" height="19">&nbsp;</td>
      <td width="217" height="19">&nbsp;</td>
      <td width="210" height="19">
<font color="#080000">
&nbsp;&nbsp;&nbsp;&nbsp; Custom:</font></td>
		<td width="510" height="19" colspan="3"><font color="#080000">
							<input type="text" name="custom_response" size="20" style="width:370; font-family:Verdana; font-size:10pt; text-align:left; height:22" value="<%=strCustomResponse %>" ></font><td width="262" height="19">
		&nbsp;</td>
    </tr>
    </table>
    </div>
   <table border="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#111111" width="1210" id="filler_table2" background="images/alt_color.gif">
    <tr>
      <td width="8" height="19">&nbsp;</td>
      <td width="217" height="19">&nbsp;</td>
      <td width="210" height="19">&nbsp;</td><td width="510" height="19" colspan="3">
		&nbsp;<td width="262" height="19">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="19">&nbsp;</td>
      <td width="217" height="19">&nbsp;</td>
      <td width="210" height="19">&nbsp;</td>
      <td width="510" height="19" colspan="3">
		&nbsp;<a href="javascript:toggleLayer('ExtraDay');" title="Add extra day and hour rates to this rule">Show Extra Rate 
		detail</a></td>      <td width="262" height="19">&nbsp;</td>
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
      15a. Extra Day rate:</td>      <td width="274" height="19" bgcolor="#C0C0C0">
      <input type="text" name="extra_day_rt" size="20" value="<%=curExtraDayRt %>" style="text-align: right" onBlur="this.value=formatCurrency(this.value);"></td>
      <td width="319" height="19">&nbsp;&nbsp;
      </td>
    </tr>
    <tr>
      <td width="8" height="19">&nbsp;</td>
      <td width="217" height="19">&nbsp;</td>
      <td width="188" height="19">&nbsp;</td>
      <td width="204" height="19" bgcolor="#C0C0C0">
      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Extra Day free miles:</td>      <td width="274" height="19" bgcolor="#C0C0C0">
      	<input type="text" name="extra_day_miles" size="20" value="<%=curExtraDayMiles %>" style="text-align: right" onBlur="this.value=formatNumber(this.value);"><font face="Verdana" size="2" color="#080000">
		(blank = unlimited)</font></td>      <td width="319" height="19">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="19">&nbsp;</td>
      <td width="217" height="19">&nbsp;</td>
      <td width="188" height="19">&nbsp;</td>
      <td width="204" height="19" bgcolor="#C0C0C0">
      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Extra Day $/extra mile:</td>      <td width="274" height="19" bgcolor="#C0C0C0">
      	<input type="text" name="extra_day_rt_per_mile" size="20" value="<%=curExtraDayRtPerMile %>" style="text-align: right" onBlur="this.value=formatCurrency(this.value);"></td>      <td width="319" height="19">&nbsp;</td>
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
      15b. Extra Hour rate:</td>      <td width="274" height="19" bgcolor="#C0C0C0">
		<input type="text" name="extra_hr_rt" size="20" value="<%=curExtraHrRt %>" style="text-align: right" onBlur="this.value=formatCurrency(this.value);"></td>      <td width="319" height="19">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="19">&nbsp;</td>
      <td width="217" height="19">&nbsp;</td>
      <td width="188" height="19">&nbsp;</td>
      <td width="204" height="19" bgcolor="#C0C0C0">
      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Extra Hour free miles:</td>      <td width="274" height="19" bgcolor="#C0C0C0">
		<input type="text" name="extra_hr_miles" size="20" value="<%=curExtraHrMiles %>" style="text-align: right" onBlur="this.value=formatNumber(this.value);"> 
		(blank = unlimited)</td>      <td width="319" height="19">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="19">&nbsp;</td>
      <td width="217" height="19">&nbsp;</td>
      <td width="188" height="19">&nbsp;</td>
      <td width="204" height="19" bgcolor="#C0C0C0">
      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Extra Hour $/extra mile:</td>      <td width="274" height="19" bgcolor="#C0C0C0">
		<input type="text" name="extra_hr_rt_per_mile" size="20" value="<%=curExtraHrRtPerMile %>" style="text-align: right" onBlur="this.value=formatCurrency(this.value);"></td>      <td width="319" height="19">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="19">&nbsp;</td>
      <td width="217" height="19">&nbsp;</td>
      <td width="188" height="19">&nbsp;</td>
      <td width="204" height="19">
<!--      <input type="reset" name="reset" value="Hide Extra Rate detail" onclick="javascript:toggleLayer('ExtraDay');" style="float: right" /></font>
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
      <td width="210" height="22">16. Search Type:</td>      <td width="510" height="22" colspan="8">
      <select size="1" name="search_profile" style="width:200; font-family:Verdana; font-size:10pt" >
      <!--
      <select size="1" name="search_profile" style="width:200; font-family:Verdana; font-size:10pt" onchange="showhide('profiles');">
		-->
      <% If intSearchProfile = 1 Then	%>
      <option selected value="1">Link to Profile(s)</option>
      <option value="2">General Rule - Not Linked</option>
      <% Else						%>
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
      <td width="210" height="22">&nbsp;&nbsp;&nbsp;&nbsp; 
      Profile(s):<br>
      <br></td>      <td width="510" height="22" colspan="8">
		<select name="profile_id" size="4" multiple style="width:373; font-family:Verdana; font-size:10pt; height:70">
				<% strProfileID = "," & strProfileID & "," %>
				<% strProfileID = Replace(strProfileID, " ", "") %>
				<% intLoopCount = 0 %>
                <% While (adoRS9.EOF = False) And (intLoopCount < 2000)  %> 
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
      17. Rate Range:</td>      <td width="510" height="22" colspan="8">
      (leave blank or enter a zero to indicate no limit)</td>      <td width="262" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="22">&nbsp;</td>
      <td width="217" height="22">&nbsp;</td>
      <td width="210" height="22">&nbsp;&nbsp;&nbsp;&nbsp; 
      Maximum:</td>      <td width="510" height="22" colspan="8">
      <input type="text" name="rate_maximum" size="20" style="width:127; font-family:Verdana; font-size:10pt; text-align:right; height:21" value="<%=strRangeMax %>" 
      onkeyup='this.onchange();' onchange='tV=(this.value.replace(/[^\d\.]/g,"")).replace(/[\.]{2,}/g,".");if(tV!=this.value){this.value=tV;}'></td>
      <td width="262" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="22">&nbsp;</td>
      <td width="217" height="22">
      &nbsp;</td>
      <td width="210" height="25">&nbsp;&nbsp;&nbsp;&nbsp; 
      Minimum:</td>      <td width="510" height="25" colspan="8">
      <input type="text" name="rate_minimum" size="20" style="width:127; font-family:Verdana; font-size:10pt; text-align:right; height:22" value="<%=strRangeMin %>"></td>
      <td width="262" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="22">&nbsp;</td>
      <td width="217" height="22">
      &nbsp;</td>
      <td width="210" height="22" align="center">OR</td>      <td width="510" height="22" colspan="8">
      &nbsp;</td>
      <td width="262" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="22">&nbsp;</td>
      <td width="217" height="22">
      &nbsp;</td>
      <td width="210" height="22">
      &nbsp;&nbsp;&nbsp; Max. / Min. Schedule</td>      <td width="510" height="22" colspan="8">
      
		<select name="maxmin_profile_id" style="width:370; font-family:Verdana; font-size:10pt; height:24" size="1">
			<% If intMaxMinProfileId = 0 Then %>
     		<option selected value="0">(None - Use above buckets)</option>
     		<% Else %>
     		<option value="0">(None - Use above buckets)</option>
			<% End If %>
			<% intLoopCount = 0                                      %>
     		<% While (adoRS11.EOF = False) And (intLoopCount < 200)  %>
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
		<a title="Click to edit or create a rule schedule" href="rate_rule_maxmin_schedule_a.aspx" onclick="window.open('rate_rule_maxmin_schedule_a.aspx?user_id=<%=strUserId %>','manage_schedule','toolbar=no,status=no,scrollbars=yes,resizable=no,width=820,height=600'); return false;">Manage schedules</a></td>      <td width="262" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="22">&nbsp;</td>
      <td width="217" height="22">
      &nbsp;</td>
      <td width="210" height="22">
      &nbsp;&nbsp;&nbsp; Threshold Schedule</td>      <td width="510" height="22" colspan="8">
      
		<select name="threshold_profile_id" style="width:370; font-family:Verdana; font-size:10pt; height:24" size="1">
			<% If intThresholdProfileId = 0 Then %>
     		<option selected value="0">(None - Use above buckets)</option>
     		<% Else %>
     		<option value="0">(None - Use above buckets)</option>
			<% End If %>
			<% intLoopCount = 0                                      %>
     		<% While (adoRS12.EOF = False) And (intLoopCount < 200)  %>
				<% If intThresholdProfileId = adoRS12.Fields("car_rate_rule_schedule_id").Value Then %>
		   		<option selected value="<%=adoRS12.Fields("car_rate_rule_schedule_id").Value %>"><%=adoRS12.Fields("car_rate_rule_schedule_desc").Value  %></option>
	     		<% Else %>
		   		<option value="<%=adoRS12.Fields("car_rate_rule_schedule_id").Value %>"><%=adoRS12.Fields("car_rate_rule_schedule_desc").Value  %></option>
				<% End If %>
	   			<% adoRS12.MoveNext
	   			   intLoopCount = intLoopCount + 1
	   		    %>
     		<% Wend             %>
     		
		</select>&nbsp;
		<a title="Click to edit or create a threshold schedule" href="threshold_maxmin_schedule_a.aspx" onclick="window.open('threshold_maxmin_schedule_a.aspx?user_id=<%=strUserId %>','manage_schedule','toolbar=no,status=no,scrollbars=yes,resizable=no,width=820,height=600'); return false;">Manage Thresholds</a></td>      <td width="262" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="22">&nbsp;</td>
      <td width="217" height="22">&nbsp;</td>
      <td width="210" height="22">&nbsp;</td>
      <td width="510" height="22" colspan="8">
      <input name="set_to_self" type="checkbox" value="True" id="set_to_self"
      <% If bolSetToSelf Then %>
      checked
      <% End If %>
      ><label for="set_to_self">Do not suggest max when all 
	  competitors are closed</label></td>
      <td width="262" height="22">&nbsp;</td>
    </tr>
    <!-- Begin follow-on rules -->
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
      <td width="210" height="22">18. When 
      situation = true:</td>      <td width="510" height="22" colspan="8">
      <select size="1" name="on_success_id" style="width:370; font-family:Verdana; font-size:10pt; height:24">
      <option selected value="0">(No follow-on rule)</option>
      			<% intLoopCount = 0 %>
                <% While (adoRS18a.EOF = False) And (intLoopCount < 400)  %>
                	<% If adoRS18a.Fields("search_profile").Value = 2 Then %>
	                	<% If intSuccessId = adoRS18a.Fields("rate_rule_id").Value Then %>
		                	<option selected value="<%=adoRS18a.Fields("rate_rule_id").Value %>"><%=adoRS18a.Fields("alert_desc").Value %></option>
		                <% Else %>
			               <% If adoRS18a.Fields("rule_status").Value = "E" Then %>
			                	<option value="<%=adoRS18a.Fields("rate_rule_id").Value %>"><%=adoRS18a.Fields("alert_desc").Value %></option>
			                <% End If %> 
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
      <td width="210" height="22">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;When 
      situation = false:</td>      <td width="510" height="22" colspan="8">
		<select name="on_failure_id" style="width:370; font-family:Verdana; font-size:10pt; height:24" size="1">
     <option selected value="0">(No follow-on rule)</option>
                <% intLoopCount = 0 %>
                <% While (adoRS18b.EOF = False) And (intLoopCount < 400)  %> 
                   	<% If adoRS18b.Fields("search_profile").Value = 2 Then %>
	                	<% If intFailureId = adoRS18b.Fields("rate_rule_id").Value Then %>
		                	<option selected value="<%=adoRS18b.Fields("rate_rule_id").Value %>"><%=adoRS18b.Fields("alert_desc").Value %></option>
		                <% Else %>
			               <% If adoRS18b.Fields("rule_status").Value = "E" Then %>
			                	<option value="<%=adoRS18b.Fields("rate_rule_id").Value %>"><%=adoRS18b.Fields("alert_desc").Value %></option>
			                <% End If %> 
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
    <!-- End follow-on rules -->
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
      <td width="210" height="25">
      19. Utilization Range:</td>
      <td width="510" height="25" colspan="8">
      (leave blank or enter a zero to indicate no limit)</td>
      <td width="262" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="22">&nbsp;</td>
      <td width="217" height="22">&nbsp;</td>
      <td width="210" height="22">&nbsp;&nbsp;&nbsp;&nbsp; Days out:</td>      <td width="64" height="22" align="center">Same</td>      <td width="64" height="22" align="center">Next</td>      <td width="64" height="22" align="center">2 - 4</td>      <td height="22" align="center">5 - 7</td>      <td width="64" height="22" align="center">8 - 14</td>      <td width="64" height="22" align="center">15 - 30</td>      <td width="64" height="22" align="center">31 - 50</td>      <td width="63" height="22" align="center">51 +</td>	  <td width="262" height="22" >&nbsp;</td>
      <td height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" style="height: 7px"></td>
      <td width="217" style="height: 7px"></td>
      <td width="210" style="height: 7px">&nbsp;&nbsp;&nbsp;&nbsp; Maximum:</td>
      <td width="64" style="height: 7px"><input type="text"  name="util_max_0" size="20" style="width:59; font-family:Verdana; font-size:10pt; text-align:right; height:22" value="<%=strUtilMax0 %>"></td>      <td width="64" style="height: 7px"><input type="text"  name="util_max_1" size="20" style="width:59; font-family:Verdana; font-size:10pt; text-align:right; height:22" value="<%=strUtilMax1 %>"></td>      <td width="64" style="height: 7px"><input type="text"  name="util_max_2" size="20" style="width:59; font-family:Verdana; font-size:10pt; text-align:right; height:22" value="<%=strUtilMax2 %>"></td>      <td style="height: 7px"><input type="text"  name="util_max_3" size="20" style="width:59; font-family:Verdana; font-size:10pt; text-align:right; height:22" value="<%=strUtilMax3 %>"></td>      <td width="64" style="height: 7px"><input type="text"  name="util_max_4" size="20" style="width:59; font-family:Verdana; font-size:10pt; text-align:right; height:22" value="<%=strUtilMax4 %>"></td>      <td width="64" style="height: 7px"><input type="text"  name="util_max_5" size="20" style="width:59; font-family:Verdana; font-size:10pt; text-align:right; height:22" value="<%=strUtilMax5 %>"></td>      <td width="64" style="height: 7px"><input type="text"  name="util_max_6" size="20" style="width:59; font-family:Verdana; font-size:10pt; text-align:right; height:22" value="<%=strUtilMax6 %>"></td>      <td width="63" style="height: 7px"><input type="text"  name="util_max_7" size="20" style="width:59; font-family:Verdana; font-size:10pt; text-align:right; height:22" value="<%=strUtilMax7 %>"></td>	  <td width="262" style="height: 7px" ></td>	  <td style="height: 7px"></td>
    </tr>
    <tr>
      <td width="8" style="height: 24px"></td>
      <td width="217" style="height: 24px"></td>
      <td width="210" style="height: 24px">&nbsp;&nbsp;&nbsp;&nbsp; Minimum:</td>
      <td width="64" style="height: 24px"><input type="text" name="util_min_0" size="20" style="width:59; font-family:Verdana; font-size:10pt; text-align:right; height:22" value="<%=strUtilMin0 %>"></td>      <td width="64" style="height: 24px"><input type="text" name="util_min_1" size="20" style="width:59; font-family:Verdana; font-size:10pt; text-align:right; height:22" value="<%=strUtilMin1 %>"></td>      <td width="64" style="height: 24px"><input type="text" name="util_min_2" size="20" style="width:59; font-family:Verdana; font-size:10pt; text-align:right; height:22" value="<%=strUtilMin2 %>"></td>      <td style="height: 24px"><input type="text" name="util_min_3" size="20" style="width:59; font-family:Verdana; font-size:10pt; text-align:right; height:22" value="<%=strUtilMin3 %>"></td>      <td width="64" style="height: 24px"><input type="text" name="util_min_4" size="20" style="width:59; font-family:Verdana; font-size:10pt; text-align:right; height:22" value="<%=strUtilMin4 %>"></td>      <td width="64" style="height: 24px"><input type="text" name="util_min_5" size="20" style="width:59; font-family:Verdana; font-size:10pt; text-align:right; height:22" value="<%=strUtilMin5 %>"></td>      <td width="64" style="height: 24px"><input type="text" name="util_min_6" size="20" style="width:59; font-family:Verdana; font-size:10pt; text-align:right; height:22" value="<%=strUtilMin6 %>"></td>      <td width="63" style="height: 24px"><input type="text" name="util_min_7" size="20" style="width:59; font-family:Verdana; font-size:10pt; text-align:right; height:22" value="<%=strUtilMin7 %>"></td>	  <td width="262" style="height: 24px"></td>      <td style="height: 24px"></td>
    </tr>
    <tr>
      <td width="8" height="22">&nbsp;</td>
      <td width="217" height="22">&nbsp;</td>
      <td width="210" height="22">&nbsp;&nbsp;&nbsp;&nbsp; <a target="_blank" href="bulk_update_utilization.asp">bulk update</a>&nbsp;</td>
      <td width="510" height="22" colspan="8">
      <input type="checkbox" name="util_in_percent" id="util_in_percent" value="True" checked disabled ><label for="util_in_percent">values 
      are listed as percentages (please do not include percent signs)</label></td>
      <td width="262" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="22">&nbsp;</td>
      <td width="217" height="22">&nbsp;</td>
      <td width="210" height="22" align="center">OR</td>      <td width="510" height="22" colspan="8">
      &nbsp;</td>
      <td width="262" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="22">&nbsp;</td>
      <td width="217" height="22">&nbsp;</td>
      <td width="210" height="22">
      &nbsp;&nbsp;&nbsp; Utilization Schedule</td>      <td width="510" height="22" colspan="8">
      
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
		<a title="Click to edit or create a rule schedule" href="rate_rule_schedule_a.asp" onclick="window.open('rate_rule_schedule_a.asp','manage_schedule','toolbar=no,status=no,scrollbars=yes,resizable=no,width=820,height=600'); return false;">Manage schedules</a></td>      <td width="262" height="22">&nbsp;</td>
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
      20. Pickup Days of Week:</font></td>
      <td width="510" height="22" colspan="8">
      <% If InStr(1, strDowList, "2") Then %>
      <input type="checkbox" name="dow_list" value="2" checked>
      <% Else %>
      <input type="checkbox" name="dow_list" value="2" >
      <% End If %>
      Mon&nbsp; 

      <% If InStr(1, strDowList, "3") Then %>
      <input type="checkbox" name="dow_list" value="3" checked>
      <% Else %>
      <input type="checkbox" name="dow_list" value="3" >
      <% End If %>
	  Tue&nbsp; 
	  
      <% If InStr(1, strDowList, "4") Then %>
      <input type="checkbox" name="dow_list" value="4" checked>
      <% Else %>
      <input type="checkbox" name="dow_list" value="4" >
      <% End If %>
	  Wed&nbsp;

      <% If InStr(1, strDowList, "5") Then %>
      <input type="checkbox" name="dow_list" value="5" checked>
      <% Else %>
      <input type="checkbox" name="dow_list" value="5" >
      <% End If %>
	  Thu&nbsp; 

      <% If InStr(1, strDowList, "6") Then %>
      <input type="checkbox" name="dow_list" value="6" checked>
      <% Else %>
      <input type="checkbox" name="dow_list" value="6" >
      <% End If %>
	  Fri&nbsp; 

      <% If InStr(1, strDowList, "7") Then %>
      <input type="checkbox" name="dow_list" value="7" checked>
      <% Else %>
      <input type="checkbox" name="dow_list" value="7" >
      <% End If %>
	  Sat&nbsp;

      <% If InStr(1, strDowList, "1") Then %>
      <input type="checkbox" name="dow_list" value="1" checked>
      <% Else %>
      <input type="checkbox" name="dow_list" value="1" >
      <% End If %>
	  Sun&nbsp; 
		  
	  
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
      <td width="210" height="22" class="link">
      21. Optional Post Actions:</td>
      <td height="22" colspan="4" class="style2">
      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Action Description</td>
      <td width="510" height="22" colspan="4" style="width: 255px" class="link">
      &nbsp;&nbsp;&nbsp;&nbsp;
      Action Amt.</td>
      <td width="262" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" style="height: 22px"></td>
      <td width="217" style="height: 22px"></td>
      <td width="210" style="height: 22px">&nbsp;&nbsp;&nbsp;&nbsp;<a target="_blank" href="bulk_update_post_actions.asp">bulk update</a></td>
      <td colspan="2" style="height: 22px">First Action:</td>
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
	  </select></font></td>      <td width="510" colspan="4" style="height: 22px; width: 255px">
		<font color="#080000">
		<input	name="rule_post_action_amt1" 
				type="text" 
				value="<%=strRulePostActionAmt1 %>" 
				style="width:127; font-family:Verdana; font-size:10pt; text-align:right; height:22"
				onblur="extractNumber(this,4,true);"
				onkeyup="extractNumber(this,4,true);"
				onkeypress="return blockNonNumbers(this, event, true, true);">
		</font></td>      <td width="262" style="height: 22px"></td>
    </tr>
    <tr>
      <td width="8" height="22">&nbsp;</td>
      <td width="217" height="22">&nbsp;</td>
      <td width="210" height="22">
      &nbsp;</td>
      <td height="22" colspan="2">
      <font color="#080000"><span class="style1">Second Action:</span></font></td>      <td height="22" colspan="2">
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
		</select></font></td>      <td width="510" height="22" colspan="4" style="width: 255px">
      <font color="#080000"> 
		<input	name="rule_post_action_amt2" 
				type="text" 
				value="<%=strRulePostActionAmt2 %>" 
				style="width:127; font-family:Verdana; font-size:10pt; text-align:right; height:22"
				onblur="extractNumber(this,4,true);"
				onkeyup="extractNumber(this,4,true);"
				onkeypress="return blockNonNumbers(this, event, true, true);">
	  </font></td>      <td width="262" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="8"></td>
      <td width="217"></td>
      <td width="210">
      </td>
      <td colspan="2">
      <font color="#080000"><span class="style1">Third Action:</span></font></td>      <td colspan="2">
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
		</select></font></td>      <td width="510" colspan="4" style="width: 255px">
      <font color="#080000"> 
			<input	name="rule_post_action_amt3" 
				type="text" 
				value="<%=strRulePostActionAmt3 %>" 
				style="width:127; font-family:Verdana; font-size:10pt; text-align:right; height:22"
				onblur="extractNumber(this,4,true);"
				onkeyup="extractNumber(this,4,true);"
				onkeypress="return blockNonNumbers(this, event, true, true);"></font></td>      <td width="262"></td>
    </tr>
    <tr>
      <td width="8" height="22">&nbsp;</td>
      <td width="217" height="22">&nbsp;</td>
      <td width="210" height="22">
      &nbsp;</td>
      <td height="22" colspan="2">
      <font color="#080000"><span class="style1">Forth Action:</span></font></td>      <td height="22" colspan="2">
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
		</select></font></td>      <td width="510" height="22" colspan="4" style="width: 255px">
      <font color="#080000"> 
 	  <input	name="rule_post_action_amt4" 
				type="text" 
				value="<%=strRulePostActionAmt4 %>" 
				style="width:127; font-family:Verdana; font-size:10pt; text-align:right; height:22"
				onblur="extractNumber(this,4,true);"
				onkeyup="extractNumber(this,4,true);"
				onkeypress="return blockNonNumbers(this, event, true, true);">
				</font></td>      <td width="262" height="22">&nbsp;</td>
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
      22. Duplicate Suggestion:</td>      <td width="510" height="22" colspan="8">
		<font color="#080000">
		<select name="additional_days_prior">
		<% intCount = 0        %>
		<% While intCount < 32 %>
		<%   If intCount = intAdditionalDaysPrior Then %>
				<option selected="selected" ><%=intCount %></option>
		<%   Else   %>
				<option ><%=intCount %></option>
		<%   End If  
		     intCount = intCount + 1
		   Wend
		   
		%>
		
		
		</select> days prior, and </font>
		<select name="additional_days">
		<% intCount = 0        %>
		<% While intCount < 32 %>
		<%   If intCount = intAdditionalDays Then %>
				<option selected="selected" ><%=intCount %></option>
		<%   Else   %>
				<option ><%=intCount %></option>
		<%   End If  
		     intCount = intCount + 1
		   Wend
		   
		%>
		
		
		</select> days after</td>      <td width="262" height="22">
		&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="22">&nbsp;</td>
      <td width="217" height="22">&nbsp;</td>
      <td width="210" height="22">
      &nbsp;</td>      <td width="510" height="22" colspan="8">
        &nbsp;</td>      <td width="262" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="22">&nbsp;</td>
      <td width="217" height="22">&nbsp;</td>
      <td width="210" height="22">
      23. Extra day Processing</td>      <td width="510" height="22" colspan="8">
      <% If blnExtraday Then %>
      <input type="checkbox" name="extraday"  id="extraday" value="True"  checked="true" disabled="disabled" title="This is automatic now - no need to check this">
      <% Else %>
      <input type="checkbox" name="extraday" id="extraday" value="True" disabled="disabled" title="This is automatic now - no need to check this">
      <% End If %>
      <label for="extraday">Calculate the rate while factoring in the extra day(s) (weekly rates only)</label></td>      <td width="262" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="22">&nbsp;</td>
      <td width="217" height="22">&nbsp;</td>
      <td width="210" height="22">
      &nbsp;</td>      <td width="510" height="22" colspan="8">
        &nbsp;</td>      <td width="262" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="22">&nbsp;</td>
      <td width="217" height="22">&nbsp;</td>
      <td width="210" height="22">
      <font face="Verdana" size="2" color="#080000">
      24. Auto-pilot mode:</font></td>      <td width="510" height="22" colspan="8">
      <% If blnAutomatic Then %>
      <input type="checkbox" name="automatic" id="automatic" value="True"  checked="true" >
      <% Else %>
      <input type="checkbox" name="automatic" id="automatic" value="True" >
      <% End If %>
      <label for="automatic">Run this rule in automatic mode (no user review or approval req.)</label></td>      <td width="262" height="22">
		&nbsp;</td>
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
      
          <input name="<%=strButton %>" type="submit" id="submit" value="    <%=strButton %>   " class="rh_button"></td>      <td width="510" height="23" colspan="8">
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
      Rate rule tester 
		<a target="_blank" href="rate_rule_tester.asp">click here</a>
      </td>      <td width="262" height="61">&nbsp;</td>
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
	<input name="sandbox" type="hidden" value="0">
</form>
<!-- Content goes before this comment -->
<!-- JUSTTABS BOTTOM OPEN -->
</font></td>
</tr>
</table>
</td>
<td  width="1" bgcolor="#000000"><img alt=""src="images/pixel.gif" width="1" height="1"></td>
</tr>
<tr bgcolor="#000000" height="1">
<td colspan=5><img alt=""src="images/pixel.gif" width="1" height="1"></td>
</tr>
</table>
<!-- JUSTTABS BOTTOM CLOSE -->
<p>&nbsp;&nbsp;[<a target="_self" href="alerts_rate_management_car_sandbox.asp">Sandbox Rules</a>]&nbsp;[<a target="_self" href="rezcentral_tethering_20130715.asp">RezCentral tethering settings</a>]</p>
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