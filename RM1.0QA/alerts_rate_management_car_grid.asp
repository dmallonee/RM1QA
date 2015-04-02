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
	Dim strOnSuccessId
	Dim strOnFailureId

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
		adoCmd4.CommandText = "car_rate_rule_select_grid"
		adoCmd4.CommandType = adCmdStoredProc
		adoCmd4.Parameters.Append adoCmd4.CreateParameter("@rate_rule_id",      3, 1, 0, Null)
		adoCmd4.Parameters.Append adoCmd4.CreateParameter("@user_id",           3, 1, 0, strUserId)
		adoCmd4.Parameters.Append adoCmd4.CreateParameter("@rate_rule_type_cd", 3, 1, 0, 1)
		adoCmd4.Parameters.Append adoCmd4.CreateParameter("@rule_status",     200, 1, 1, strRuleStatus)
	
	
		'Create an ADO RecordSet object
		Set adoRS4 = Server.CreateObject("ADODB.Recordset")
	
		adoRS4.CursorLocation = adUseClient 
	
		'Open the RecordSet
		adoRS4.Open adoCmd4, , adOpenStatic, adLockReadOnly
		
		
		Set adoCmd10 = CreateObject("ADODB.Command")
	
		adoCmd10.ActiveConnection =  strConn
		adoCmd10.CommandText = "car_rate_rule_select_names"
		adoCmd10.CommandType = adCmdStoredProc
		adoCmd10.Parameters.Append adoCmd4.CreateParameter("@rate_rule_id",      3, 1, 0, Null)
		adoCmd10.Parameters.Append adoCmd4.CreateParameter("@user_id",           3, 1, 0, strUserId)
		adoCmd10.Parameters.Append adoCmd4.CreateParameter("@rate_rule_type_cd", 3, 1, 0, 1)
		adoCmd10.Parameters.Append adoCmd4.CreateParameter("@rule_status",     200, 1, 1, strRuleStatus)
	
		Set adoRS18a = adoCmd10.Execute
		Set adoRS18b = adoCmd10.Execute
		
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

		Rem Get the success follow-on rules if any
		Set adoCmdOnSuccess = CreateObject("ADODB.Command")

		adoCmdOnSuccess.ActiveConnection = strConn
		adoCmdOnSuccess.CommandText = "car_rate_rule_on_success_select"
		adoCmdOnSuccess.CommandType = adCmdStoredProc
		
		adoCmdOnSuccess.Parameters.Append adoCmdOnSuccess.CreateParameter("@rate_rule_id", 3, 1, 0, intRuleId)
	
		Set adoRSOnSuccess = adoCmdOnSuccess.Execute
		
		While (adoRSOnSuccess.EOF = False) 
			strOnSuccessId = strOnSuccessId & "," & adoRSOnSuccess.Fields("on_success_id").Value 
	    	adoRSOnSuccess.MoveNext
		Wend

		Rem Get the failure follow-on rules if any
		Set adoCmdOnFailure = CreateObject("ADODB.Command")

		adoCmdOnFailure.ActiveConnection = strConn
		adoCmdOnFailure.CommandText = "car_rate_rule_on_failure_select"
		adoCmdOnFailure.CommandType = adCmdStoredProc
		
		adoCmdOnFailure.Parameters.Append adoCmdOnFailure.CreateParameter("@rate_rule_id", 3, 1, 0, intRuleId)
	
		Set adoRSOnFailure = adoCmdOnFailure.Execute

		While (adoRSOnFailure.EOF = False) 
			strOnFailureId = strOnFailureId & "," & adoRSOnFailure.Fields("on_failure_id").Value 
	    	adoRSOnFailure.MoveNext
		Wend

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
		intAdditionalDays = 0
		blnExtraday = False
		
		strButton = "Create"

	End If
	


' Build connection string to aspgrid.mdb
'strConnect = Session("pro_con")	

' Create an instance of AspGrid
Set Grid = Server.CreateObject("Persits.Grid")
'Grid.LoadParameters Server.MapPath("usergrid.xml")

' Connect to the database
'Grid.Connect strConnect, "", ""

' Specify SQL Recordset
Grid.Recordset = adoRS4

' Specify location of button images
Grid.ImagePath = "../images/aspgrid/"

Rem Grid Level Properties
Grid.MaxRows = 25 
Grid.Table.Width = "2310"
Grid.MethodGet = False
Grid.CanEdit = False
Grid.CanDelete = False
Grid.CanAppend = False

Rem Table level properties
Grid.Table.CellSpacing = 0
Grid.Table.CellPadding = 1 
Grid.Table.Class = "grid_table"
 
Rem Column Widths
Grid.Cols(1).Header.Width = "50"
Grid.Cols(2).Header.Width = "58"
Grid.Cols(3).Header.Width = "251"
Grid.Cols(4).Header.Width = "100"
Grid.Cols(5).Header.Width = "120"

Rem Hidden columns
Grid.Cols("Long Description").Hidden = True
Grid.Cols("Comp Set").Hidden = True
Grid.Cols("Car type").Hidden = True
Grid.Cols("Tol. Amt").Hidden = True
Grid.Cols("Status").Hidden = True

Rem Headers
'Grid.Cols(1).Header.Font.Class = "table_header_first"
Grid.ColRange(1, 17).Header.Font.Class = "table_header"
Grid.ColRange(1,  1).Header.BGColor = "#879AA2"
Grid.ColRange(2,  2).Header.BGColor = "#E07D1A"
Grid.ColRange(3, 17).Header.BGColor = "#879AA2"

Rem Rows
Grid.ColRange(1,  1).Cell.AltBGColor = "#CFD7DB"
Grid.ColRange(1,  1).Cell.BGColor = "#B2BEC4"
Grid.ColRange(1,  1).Cell.Class = "cell_data"
Grid.ColRange(1,  1).CanSort = True

Grid.ColRange(2, 2).Cell.AltBGColor = "#FDC677"
Grid.ColRange(2, 2).Cell.BGColor = "#FDC677"
Grid.ColRange(2, 2).Cell.Class = "cell_data"
Grid.ColRange(2, 2).CanSort = False

Grid.ColRange(3, 17).Cell.AltBGColor = "#CFD7DB"
Grid.ColRange(3, 17).Cell.BGColor = "#B2BEC4"
Grid.ColRange(3, 17).Cell.Class = "cell_data"
Grid.ColRange(3, 17).CanSort = True 

Rem Misc
Grid.Cols(0).Footer.BGColor = "#879AA2"
Grid.Cols(0).Header.BGColor = "#879AA2"

Rem Cells
Grid("ID").Cell.Align = "CENTER"

Grid("Selected").Cell.Align = "CENTER"
Grid("selected").AttachExpression "<input type=""checkbox"" value=""{{ID}}"" name=""rate_rule_id"" onclick=""update_selected_list();"" >"

Grid("Description").AttachExpression "<a target=""_self"" title=""{{Long Description}}"" href=""alerts_rate_management_car_grid.asp?rateruleid={{ID}}"">{{Description}}</A>"

Grid("Sit. Amt").FormatNumeric 2, True, False, True, "$"
Grid("Sit. Amt").Cell.Align = "RIGHT"

Grid("Res. Amt").FormatNumeric 2, True, False, True, "$"
Grid("Res. Amt").Cell.Align = "RIGHT"

Grid("Tol. Amt").FormatNumeric 2, True, False, True, "$"
Grid("Tol. Amt").Cell.Align = "RIGHT"

Grid("Rate Ceiling").FormatNumeric 2, True, False, True, "$"
Grid("Rate Ceiling").Cell.Align = "RIGHT"

Grid("Rate Floor").FormatNumeric 2, True, False, True, "$"
Grid("Rate Floor").Cell.Align = "RIGHT"

Grid("Status").Cell.Align = "CENTER"
Grid("Auto").Cell.Align = "CENTER"


'Grid("expiration").FormatDate "%m/%d/%y"
'Grid("last_login").FormatDate "%m/%d/%y" '"%b %d, %Y"
'Grid("modified").FormatDate "%m/%d/%y"

'Grid("lob_id").Array = Array("Air", "Car", "Hotel")
'Grid("lob_id").VArray = Array(3, 2, 1)


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
<html xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office">
<head>
<meta name="GENERATOR" content="Microsoft FrontPage 12.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Language" content="en-us">
<title>Rate-Monitor by Rate-Highway, Inc. | Alerts! | Rate Management</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<script language="Javascript" type="text/javascript" src="inc/sitewide.js"></script>
<script language="Javascript" type="text/javascript" src="inc/header_menu_support.js" ></script>
<script language="Javascript" type="text/javascript" src="inc/js_calendar_v2.js"></script>
<script language="Javascript" type="text/javascript" src="inc/validate2.js"></script>
<script language="JavaScript" type="text/javascript" src="inc/multiple_select_support.js"></script>
<script language="JavaScript" type="text/javascript" src="inc/multiple_select_support2.js"></script>
<script language="JavaScript" type="text/javascript">

//window.onload=function(){enableTooltips("content")};

function CreateAlert()
{ 
	var valid_form = true;
	var numSelected = 0;
	var i;
	
	//alert("testing");
	SelectOptions1();
	SelectOptions2();
	SelectOptions3();
	SelectOptions4();
		
	
	if (document.create_alert.alert_desc.value == '') 
		{
		alert("Please select a descriptive name for your rule.");  
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
		//// change debug to true for debug messages
		//alert("1about to transfer to " + window.document.create_alert.action.value);
		////window.document.create_alert.action = "car_rate_rule_insert1.asp?debug=false";
		//window.document.create_alert.txtaction.value = "car_rate_rule_insert1.asp?debug=true";
		//window.document.create_alert.action.value = "car_rate_rule_insert1.asp?debug=true";
		//alert("2about to transfer to " + window.document.create_alert.action.value);
		////window.document.create_alert.txtaction.value = window.document.create_alert.action.value ;
		window.document.create_alert.submit();
		return true;
		}
	else {
		return false;
		}
}


function SelectOptions1(){
	var listBox = document.create_alert.profile_id;
	var len = listBox.length;
	for(var x=0;x<len;x++){
		listBox.options[x].selected= true;
	}
}


function SelectOptions2(){
	var listBox = document.create_alert.vend_cd1;
	var len = listBox.length;
	for(var x=0;x<len;x++){
		listBox.options[x].selected= true;
	}
}


function SelectOptions3(){
	var listBox = document.create_alert.on_success_id_selected;
	var len = listBox.length;
	for(var x=0;x<len;x++){
		listBox.options[x].selected= true;
	}
}


function SelectOptions4(){
	var listBox = document.create_alert.on_failure_id_selected;
	var len = listBox.length;
	for(var x=0;x<len;x++){
		listBox.options[x].selected= true;
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


function update_selected_list() { 
// each time the user clicks on a checkbox, add it to the list of checkboxes checked, so that they
// can be disabled or enabled or whatever.

	var list = new String("")
	var inputs = window.document.getElementsByTagName('input');
	var myTextField = window.document.getElementById('rate_rule_id_list');


	for(var i=0; i < inputs.length; i++){ //iterate through all input elements
		if (inputs[i].type.toLowerCase() == 'checkbox' && inputs[i].name.toLowerCase() == 'rate_rule_id') { //if the element is a checkbox
			if (inputs[i].checked == true)
				list = list + ',' + inputs[i].value;
		}			
	}

	//window.document.maint.rate_rule_id_list.value = list;
	myTextField.value = list;
    return true 
} 

//  End -->
</script>
<link rel="stylesheet" type="text/css" href="inc/rh_standard.css">
<link rel="stylesheet" type="text/css" href="inc/rh_report.css">
<link rel="stylesheet" type="text/css" href="inc/sitewide.css">
<link rel="stylesheet" type="text/css" href="inc/css_calendar_v2.css" id="calendarcss" media="all" >
<link rel="stylesheet" type="text/css" href="aspgrid.css">
<style type="text/css" >
<!--
.profile_header { height="48" text-align:left; padding-left:3; padding-right:3; padding-top:0; background-color:#879AA2; font-family:Verdana, Arial, Helvetica, sans-serif; font-size:12pt; vertical-align:bottom; text-align:left }
.profile_light { height="68" text-align:left; padding-left:3; padding-right:3; padding-top:3; background-color:#CFD7DB; font-family:Verdana, Arial, Helvetica, sans-serif; font-size:12pt; vertical-align:bottom; text-align:left }
.profile_dark { height="48" text-align:left; padding-left:3; padding-right:3; padding-top:3; background-color:#C0C0C0; font-family:Verdana, Arial, Helvetica, sans-serif; font-size:12pt; vertical-align:bottom; text-align:left }
div#ExtraDay { margin: 0px 20px 0px 20px; display: none; }
.style1 {
	font-size: small;
}
.style2 {
	font-size: small;
	text-align: left;
}
.grid_table {
	border-collapse: collapse;
	border-color: #FFFFFF;
}
.style3 {
	border-style: solid;
	border-width: 0;
}
.style4 {
	text-align: center;
	border-style: solid;
	border-width: 0;
}
.style5 {
	border-width: 0;
}
.style6 {
	border-width: 0;
	text-decoration: line-through;
}
.style7 {
	border-collapse: collapse;
	background-color: #FFFFFF;
}
.style9 {
	border-collapse: collapse;
}
.style10 {
	margin-top: 2px;
	margin-bottom: 2px;
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
    <img src="images/top.jpg" width="770" height="91" border="0" alt="Rate Automation"></a></td>
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
        <td><map name="logout_map">
        <area alt="Click to log out of Rate-Monitor" href="http://www.rate-monitor.com/" shape="rect" coords="337, 10, 394, 25">
        </map>
        <img src="images/bottom_right.gif" width="402" height="31" usemap="#logout_map" border="0"></td>
      </tr>
    </table>
    </td>
  </tr>
</table>
<p align="right">

    &nbsp;</p>
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
<tr>
<td>

<!-- JUSTTABS TOP OPEN-END -->
&nbsp;
<% If strEditMode <> "1" Then %>
<form method="GET" name="search_alerts" class="search">
<table border="0" style="border-collapse: collapse" bordercolor="#FFFFFF" width="100%" cellspacing="0" height="4">
    <tr>
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
          <img border="0" src="images/search.GIF"></td>
          <td width="583" colspan="3" height="51">
          To search 
          for an Alert, enter a login id, or a portion of the 
          address. You may also enter the alert type.</td>
          <td width="336" height="51">&nbsp;</td>
        </tr>
        <tr valign="bottom">
          <td width="10" height="26">&nbsp;</td>
          <td width="179" height="26">&nbsp;</td>
          <td width="177" height="26">
          Owner: </td>
          <td width="80" height="26">
          
          <input type="text" name="name" size="20" style="width:150" style="width:150" onfocus="this.className='focus';cl(this,'<%=strClientUserid %>');" onblur="this.className='';fl(this,'<%=strClientUserid %>');"></td>
          <td width="662" colspan="2" height="26">
          
          <input name="search" type="submit" id="Open2224" value="    Search    " class="rh_button"></td>
        </tr>
        <tr valign="bottom">
          <td width="10" height="22">&nbsp;</td>
          <td width="179" height="22">&nbsp;</td>
          <td width="177" height="22">
          Alert 
          Status:</td>
          <td width="80" height="22">
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
          <td width="179" height="22"><a name="grid_top">&nbsp;</a></td>
          <td width="177" height="22">&nbsp;</td>
          <td width="80" height="22">&nbsp;</td>
          <td width="662" colspan="2" height="22">&nbsp;</td>
        </tr>
      </table>
      </td>
    </tr>
  </table>
  </form>
	<table border="0" style="border-collapse: collapse" bordercolor="#FFFFFF" width="2310" height="4">
    	<tr>
    		<td background="images/ruler.gif"></td>
    	</tr>
  	</table>
	<%
	' Display grid
	Grid.Display
	Grid.Disconnect
	%>
	<!-- 
	<table border="0" style="border-collapse: collapse" bordercolor="#FFFFFF" width="2310" height="4">
	    <tr>
			<td background="images/ruler.gif"></td>
	    </tr>
	</table>
  	-->
  <form name="maint" method="post" action="car_rate_rule_maint.asp">
  	<input type="hidden" name="action" value="0">
  	<input type="hidden" name="refresh_from" value="search">
  &nbsp;| <a href="javascript:maint_action(1);">Delete</a> 
  | <!-- <a href="javascript:maint_action(2)">Copy</a> | -->
  <a href="javascript:maint_action(3)">Enable</a> |
  <a href="javascript:maint_action(4)">Disable</a> | 
  <a target="_blank" href="alerts_rate_management_export.asp">Download cross-reference</a>
  | <a target="_blank" href="alerts_rate_management_export_worksheet.asp">Download rule worksheet for upload</a>
  | <a target="_blank" href="rule_upload.asp">Upload rules</a>
  |	<a target="_self" href="alerts_rate_management_car_grid.asp">Prior Version</a>
  
  
		<input name="rate_rule_id" type="hidden"  id="rate_rule_id_list">
	</form>
    <br>
    <br>
<form name="create_alert" method="post" OnSubmit="return CreateAlert();" action="car_rate_rule_insert1.asp?debug=false" >
  <table border="0" style="border-collapse: collapse" bordercolor="#FFFFFF" width="2310" height="4">
    <tr>
      <td background="images/ruler.gif"></td>
    </tr>
  </table>
  <table border="0" cellpadding="0" bordercolor="#111111" width="1210" id="new_alert" background="images/alt_color.gif" class="style9">
    <tr>
      <td width="8" height="19">&nbsp;</td>
      <td width="217" height="19">&nbsp;</td>
      <td width="210" height="19">
		&nbsp;</td>
      <td width="510" height="19">&nbsp;</td>
      <td width="262" height="19">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="25">&nbsp;</td>
      <td width="217" height="25">
      <img border="0" src="images/maintenance.GIF" width="162" height="25"></td>
      <td width="210" height="25">1. Rate Change Alert No.:</td>
      <td width="510" height="25">
      <input type="text" name="rate_rule_id" size="20" style="width:200; font-family:Verdana; font-size:10pt; text-align:right; background-image:url('images/alt_color.gif')" value="<%=intRuleId %>" READONLY><a name="maint_top">&nbsp;</a>
      <input type="checkbox" name="copy" id="copy" value="true"><label for="copy">Save as a copy 
		<br>
		</label>

		<label for="copy">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
		(leaves the original unchanged)</label></td>
      <td width="262" height="25"></td>
    </tr>
    <tr>
      <td width="8" height="25">&nbsp;</td>
      <td width="217" height="25">
      &nbsp;</td>
      <td width="210" height="25">&nbsp;</td>
      <td width="510" height="25">
      &nbsp;</td>
      <td width="262" height="25">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="25">&nbsp;</td>
      <td width="217" height="25">
      &nbsp;</td>
      <td width="210" height="25">2. Description:</td>
      <td width="510" height="25">
      <input type="text" name="alert_desc" size="20" style="width:439; font-family:Verdana; font-size:10pt; height:21" value="<%=strAlertDesc %>"></td>
      <td width="262" height="25">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="25">&nbsp;</td>
      <td width="217" height="25">
      &nbsp;</td>
      <td width="210" height="25">&nbsp;</td>
      <td width="510" height="25">
      &nbsp;</td>
      <td width="262" height="25">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="25">&nbsp;</td>
      <td width="217" height="25">
      &nbsp;</td>
      <td width="210" height="25">3. Rule Begin Date:</td>
      <td width="510" height="25">
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
      <td width="510" height="25">
      (enter 'continuous' or blank for no begin date) </td>
      <td width="262" height="25">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="25">&nbsp;</td>
      <td width="217" height="25">
      &nbsp;</td>
      <td width="210" height="25">4. Rule End Date:</td>
      <td width="510" height="25">
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
      <td width="510" height="25">
      (enter 'continuous' or blank for no end date)</td>
      <td width="262" height="25">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="25">&nbsp;</td>
      <td width="217" height="25">
      &nbsp;</td>
      <td width="210" height="25">5. First 
      Pick-up:</td>
      <td width="510" height="25">
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
      <td width="510" height="25">
      (enter 'continuous' or blank for no pick-up date)</td>
      <td width="262" height="25">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="25">&nbsp;</td>
      <td width="217" height="25">
      &nbsp;</td>
      <td width="210" height="25">6. Last Pick-up:</td>
      <td width="510" height="25">
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
      <td width="510" height="25">
      (enter 'continuous' or blank for no pick-up date)</td>
      <td width="262" height="25">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="25">&nbsp;</td>
      <td width="217" height="25">
      &nbsp;</td>
      <td width="210" height="25">7. System Rate 
      Code:</td>
      <td width="510" height="25">
      <input type="text" name="client_sys_rate_cd" size="20" style="width:200; font-family:Verdana; font-size:10pt" value="<%=strClientRateCode %>"></td>
      <td width="262" height="25">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="25">&nbsp;</td>
      <td width="217" height="25">
      &nbsp;</td>
      <td width="210" height="25">&nbsp;</td>
      <td width="510" height="25">
      (the rate code used within your system, i.e. Daily, Weekly, etc.)</td>
      <td width="262" height="25">&nbsp;</td>
    </tr>
   </table>
   <table border="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#111111" width="1210" id="new_alert0" background="images/alt_color.gif" height="561">
      <tr>
      <td width="8" height="25">&nbsp;</td>
      <td style="width: 217px" class="style5">
      &nbsp;</td>
      <td style="width: 210px" class="style5">8. Car 
      Companies:</td>
      <td height="25" colspan="3">&nbsp;</td>
      <td width="262" height="25">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="25">&nbsp;</td>
      <td style="width: 217px" class="style5">
      &nbsp;</td>
      <td style="padding-top: 2px; width: 210px;" class="style5" >
      &nbsp;</td>
      <td height="25" style="width: 211px">
	  
      Un-selected:</td>
      <td height="25" style="width: 39px" >
	  &nbsp;</td>
      <td height="25" bordercolor="#555566" style="width: 49px">
	  
	  Selected:</td>
      <td >
      &nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="25">&nbsp;</td>
      <td style="width: 217px" class="style5">
      &nbsp;</td>
      <td style="padding-top: 2px; width: 210px;" class="style5" >
      &nbsp;&nbsp;&nbsp;Competitive Set:</td>
      <td height="25" style="width: 211px">
      <select size="4" name="un_selected_vend_cd1" multiple style="width:200; font-family:Verdana; font-size:10pt; height: 90px;"  >
      <!-- 
      <% If strVendCd = "XX" Then %>
      	<option selected value="XX"><%="All comp. set" %></option>
	  <% Else                     %>
      	<option value="XX"><%="All comp. set" %></option>
      <% End  If                  %>
      -->
		<% Dim intLoopCount         %>
		<% Dim strCSVendCd          %>
		<% Dim strCSVendName        %>

		<% While (adoRS2.EOF = False) And (intLoopCount < 100) %>
		<% If ((InStr(1, strVendCd, adoRS2.Fields("vendor_cd").Value)) And (strVendCd <> "")) Or (strVendCd = "XX") Then %>
			  <% strCSVendCd = strCSVendCd & adoRS2.Fields("vendor_cd").Value & ","		    %>       
			  <% strCSVendName = strCSVendName & adoRS2.Fields("vendor_name").Value & ","	%>	          
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
      <td style="width: 39px"  >
		<img border="0" src="images/move_right.GIF" width="24" height="22"  onclick="moveDualList( document.create_alert.un_selected_vend_cd1, document.create_alert.vend_cd1, false );return false" ><br>
		<img border="0" src="images/move_right_all.GIF" width="24" height="22"  onclick="moveDualList( document.create_alert.un_selected_vend_cd1, document.create_alert.vend_cd1, true );return false"  ><br>
		<img border="0" src="images/move_left.GIF" width="24" height="22"  onclick="moveDualList( document.create_alert.vend_cd1, document.create_alert.un_selected_vend_cd1, false );return false"  ><br>
		<img border="0" src="images/move_left_all.GIF" width="24" height="22"  onclick="moveDualList( document.create_alert.vend_cd1, document.create_alert.un_selected_vend_cd1, true );return false"  >
      </td>
      <td height="25" bordercolor="#555566" style="width: 49px">
	  <select size="4" name="vend_cd1" multiple style="width:200; font-family:Verdana; font-size:10pt; height: 90px;" >
      <% Dim strCSVendCds        %>
      <% Dim strCSVendNames      %>
      <% strCSVendCds = Split(strCSVendCd, ",")       %>
      <% strCSVendNames = Split(strCSVendName, ",")   %>
		<% For intLoopCount = LBound(strCSVendCds) To UBound(strCSVendCds)  %> 
      		<% If strCSVendCds(intLoopCount) <> "" Then %>
        		<option value="<%=strCSVendCds(intLoopCount) %>"><%=strCSVendNames(intLoopCount) %></option>
            <% End If %>
		<% Next %>  
      </select>
      </td>
      <td width="262" >
      
      &nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="25">&nbsp;</td>
      <td style="width: 217px" class="style5">
      &nbsp;</td>
      <td valign="top" style="padding-top: 2px; width: 210px;" class="style5">
      &nbsp;</td>
      <td height="25" colspan="3">
      </td>
      <td width="262" height="25">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="25">&nbsp;</td>
      <td style="width: 217px" class="style5">
      &nbsp;</td>
      <td style="padding-top: 2px; width: 210px;" class="style5">
      &nbsp;&nbsp; Comparison Company<br>
&nbsp;&nbsp; (usually self):</td>
      <td height="25" style="width: 211px">
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
	  <td height="25" >
      &nbsp;</td>
      <td height="25" bgcolor="#FFFFFF" bordercolor="#555566" bordercolorlight="#999999" bordercolordark="#777777" style="border: 1px solid #40618F; padding: 2px;" >
      <p align="left">
      <img alt="" border="0" src="images/tip_ballon.gif" width="23" height="22"><b> 
      Note</b>: If you select a group as the comparison company, the lowest rate 
		of the group will be used      
	  </p>
      </td>      
      <td width="262" height="25">
      </td>
    </tr>
    <tr>
      <td width="8" height="25">&nbsp;</td>
      <td style="width: 217px" class="style5">
      &nbsp;</td>
      <td valign="top" style="padding-top: 2px; width: 210px;" class="style5">
      &nbsp;</td>
      <td height="25" colspan="3">
      &nbsp;</td>
      <td width="262" height="25">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="25">&nbsp;</td>
      <td style="width: 217px" class="style5">
      &nbsp;</td>
      <td valign="top" style="padding-top: 2px; width: 210px;" class="style6">
      9. LOR(s): <br>
		</td>
      <td height="25" colspan="3">
      <!-- 
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
      </select>
      -->
      <span class="style1"><em>Removed - no longer necessary</em></span></td>
      <td width="262" height="25">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="25">&nbsp;</td>
      <td style="width: 217px" class="style5">
      &nbsp;</td>
      <td valign="top" style="padding-top: 2px; width: 210px;" class="style5">&nbsp;</td>
      <td height="25" colspan="3">
      &nbsp;</td>
      <td width="262" height="25">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="71">&nbsp;</td>
      <td style="width: 217px" class="style5">
      &nbsp;</td>
      <td valign="top" style="padding-top: 2px; width: 210px;" class="style5">
      10. Location(s):</td>
      <td height="71" valign="top" colspan="3">
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
      <td style="width: 217px" class="style5">
      &nbsp;</td>
      <td valign="top" style="padding-top: 2px; width: 210px;" class="style5">&nbsp;</td>
      <td height="24" colspan="3">
      (select 
      airport/city codes &quot;any&quot; for any location, edit to edit custom)</td>
      <td width="262" height="24">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="25">&nbsp;</td>
      <td style="width: 217px" class="style5">
      &nbsp;</td>
      <td style="vertical-align: top; padding-top: 2px; width: 210px;" class="style5">
      11. Car Types:</td>
      <td height="25" colspan="3">
      &nbsp;</td>
      <td width="262" height="25">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="25">&nbsp;</td>
      <td style="width: 217px" class="style5">
      &nbsp;</td>
      <td style="padding-top: 2px; width: 210px;" class="style5">
      &nbsp;&nbsp;&nbsp;&nbsp; Compare Car Type(s):</td>
      <td height="25" colspan="3">
      
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
      </select> </td>
      <td width="262" height="25">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="25">&nbsp;</td>
      <td style="width: 217px" class="style5">
      &nbsp;</td>
      <td style="vertical-align: top; padding-top: 2px; width: 210px;" class="style5">
      &nbsp;</td>
      <td height="25" colspan="3">
      &nbsp;</td>
      <td width="262" height="25">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="19">&nbsp;</td>
      <td style="width: 217px" class="style5">&nbsp;</td>
      <td style="width: 210px" class="style5">&nbsp;&nbsp;&nbsp;&nbsp; 
		Suggestion Car&nbsp;&nbsp;&nbsp;<br>
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
      </select> </td>
      <td width="262" height="19">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="19">&nbsp;</td>
      <td style="width: 217px" class="style5">&nbsp;</td>
      <td style="width: 210px" class="style5">&nbsp;</td>
      <td height="19" colspan="3">
      (use &quot;n/a&quot; to 
      denote any car type, you can use &quot;n/a&quot; for<br>
	  &nbsp;both items to have the system match car types)</td>
      <td width="262" height="19">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="19">&nbsp;</td>
      <td style="width: 217px" class="style5">&nbsp;</td>
      <td style="width: 210px" class="style6">12. Rate Source::</td>
      <td height="19" colspan="3" class="style1">
<!--
      <select size="1" name="data_source" style="width:200; font-family:Verdana; font-size:10pt">
	  <option selected value="ALL">Determined by profile</option>      


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

      </select>
-->      
      <em>Removed - no longer necessary</em></td>
      <td width="262" height="19">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="19">&nbsp;</td>
      <td style="width: 217px" class="style5">&nbsp;</td>
      <td style="width: 210px" class="style5">&nbsp;</td>
      <td height="19" colspan="3">&nbsp;</td>
      <td width="262" height="19">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="22">&nbsp;</td>
      <td style="width: 217px" class="style5">&nbsp;</td>
      <td style="width: 210px" class="style5">13. Situation:</td>
      <td height="22" colspan="3">
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
	      <option selected value="4">If rate is not more than (all comp. set) by at least</option>
	  <% Else %>
	      <option value="4">If rate is not more than (all comp. set) by at least</option>
	  <% End If %>
	  
      <% If intSituationCd = 5 Then	%>
	      <option selected value="5">If rate is not more than (any comp. set) by at least</option>
	  <% Else %>
	      <option value="5">If rate is not more than (any comp. set) by at least</option>
	  <% End If %>
	  
      <% If intSituationCd = 6 Then	%>
	      <option selected value="6">If rate is not more than (custom) by at least</option>
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



      </select>&nbsp;</td>
      <td width="262" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="23">&nbsp;</td>
      <td style="width: 217px" class="style5">&nbsp;</td>
      <td style="width: 210px" class="style5">&nbsp;&nbsp;&nbsp; 
      by amount:<br>
&nbsp;&nbsp;&nbsp; <br>
      </td>
      <td height="23" colspan="3">
      <input type="text" name="situation_amt" size="20" style="width:127; font-family:Verdana; font-size:10pt; text-align:right; height:22" value="<%=strSituationAmt %>">
      
      <a href="javascript:centerPopUp( 'rule_situation_tester.asp', 'test', 620, 500 )">test situations</a>

<font color="#FF0000" face="Courier New" size="2">
      (beta)</font><br>
      <!-- 
      <% If blnIsDollar Then %>
      <input type="radio" value="1" name="is_dollar" style="font-family:Verdana; font-size:10pt" checked id="is_dollar1">    
      <label for="is_dollar1">Dollar amount</label><br>
      <input type="radio" value="0" name="is_dollar" style="font-family:Verdana; font-size:10pt" id="is_dollar2"><label for="is_dollar2">Percentage</label><br>
      <% Else %>
      <input type="radio" value="1" name="is_dollar" style="font-family:Verdana; font-size:10pt" id="is_dollar1">    
      <label for="is_dollar1">Dollar amount</label><br>
      <input type="radio" value="0" name="is_dollar" style="font-family:Verdana; font-size:10pt" checked  id="is_dollar2"><label for="is_dollar2">Percentage</label><br>
      <% End If %>
	  -->
      <% If blnIgnoreClosed Then %>
      <input type="checkbox" name="ignore_closed" value="True" id="ignore_closed" checked><label for="ignore_closed">Ignore closed rates</label> 
	  <% Else %>
      <input type="checkbox" name="ignore_closed" value="True" id="ignore_closed" ><label for="ignore_closed">Ignore closed rates</label> 
	  <% End If %>
		(do not provide suggestions when 
		comparison is closed) 	
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
    <tr>
      <td width="8" height="19">&nbsp;</td>
      <td style="width: 217px" class="style5">&nbsp;</td>
      <td style="width: 210px" class="style5">&nbsp;</td>
      <td height="19" colspan="3">&nbsp;</td>
      <td width="262" height="19">&nbsp;</td>
    </tr>
    <!-- 
    <tr>
      <td width="8" height="22">&nbsp;</td>
      <td style="width: 217px" class="style5">&nbsp;</td>
      <td style="width: 210px" class="style5">14. Quantity &amp; 
      Period:</td>
      <td height="22" colspan="3">
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
      <td style="width: 217px" class="style5">&nbsp;</td>
      <td style="width: 210px" class="style5">&nbsp;&nbsp;&nbsp; 
      No. of Events:</td>
      <td height="22" colspan="3">
      <input type="text" name="event_count" size="20" style="width:200; font-family:Verdana; font-size:10pt; text-align:right" value="<%=intEventCount %>"></td>
      <td width="262" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="22">&nbsp;</td>
      <td style="width: 217px" class="style5">&nbsp;</td>
      <td style="width: 210px" class="style5">&nbsp;&nbsp;&nbsp; 
      That happen in:</td>
      <td height="22" colspan="3">
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
    -->
    <tr>
      <td width="8" height="19">&nbsp;</td>
      <td style="width: 217px" class="style5">&nbsp;</td>
      <td style="width: 210px" class="style5">&nbsp;</td>
      <td height="19" colspan="3">&nbsp;</td>
      <td width="262" height="19">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="22">&nbsp;</td>
      <td style="width: 217px" class="style5">&nbsp;</td>
      <td style="width: 210px" class="style5">15. Response:</td>
      <td height="22" colspan="3">
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
      <option selected value="15">Set my rate to avg. of 2 lowest comp. set plus $</option>
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


      </select> <a href="alerts_rate_management_car_help_15.asp" onclick="window.open('alerts_rate_management_car_help_15.asp','window_name','toolbar=no,status=no,scrollbars=yes,resizable=no,width=450,height=255'); return false;"><img src="images/question.gif" border="0" alt="About response options" class="cbtip"></a></td>
      <td width="262" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="19">&nbsp;</td>
      <td style="width: 217px" class="style5">&nbsp;</td>
      <td style="width: 210px" class="style5">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Amount:<br>
      <br>&nbsp;</td>
      <td height="19" colspan="3">
      <input type="text" name="response_amt" size="20" style="width:127; font-family:Verdana; font-size:10pt; text-align:right; height:22" value="<%=strResponseAmt %>" >
      <br>
      <!-- 
      <% If blnResponseDollar Then %>
	      <input type="radio" value="1" name="is_response_dollar" id="is_response_dollar1" style="font-family:Verdana; font-size:10pt" checked>Dollar amount<br>
	      <input type="radio" value="0" name="is_response_dollar" id="is_response_dollar2"  style="font-family:Verdana; font-size:10pt">Percentage</td>
	  <% Else					   %>
	      <input type="radio" value="1" name="is_response_dollar" id="is_response_dollar1" style="font-family:Verdana; font-size:10pt" >Whole Dollar<br>
	      <input type="radio" value="0" name="is_response_dollar" id="is_response_dollar2" style="font-family:Verdana; font-size:10pt" checked>Percentage</td>
	  <% End If					   %>
	  -->
      <td width="262" height="19">&nbsp;&nbsp;
      </td>
    </tr>
    <tr>
      <td width="8" height="19">&nbsp;</td>
      <td style="width: 217px" class="style5">&nbsp;</td>
      <td style="width: 210px" class="style5">&nbsp;</td>
      <td height="19" colspan="3">
		&nbsp;<a href="javascript:toggleLayer('ExtraDay');" title="Add extra day and hour rates to this rule">Hide/Show Extra Rate 
		detail</a></td>
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
      15a. Extra Day rate:</td>
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
      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Extra Day free miles:</td>
      <td width="274" height="19" bgcolor="#C0C0C0">
      
		<input type="text" name="extra_day_miles" size="20" value="<%=curExtraDayMiles %>" style="text-align: right" onBlur="this.value=formatNumber(this.value);">
		(blank = unlimited)</td>
      <td width="319" height="19">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="19">&nbsp;</td>
      <td width="217" height="19">&nbsp;</td>
      <td width="188" height="19">&nbsp;</td>
      <td width="204" height="19" bgcolor="#C0C0C0">
      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Extra Day $/extra mile:</td>
      <td width="274" height="19" bgcolor="#C0C0C0">
      
		<input type="text" name="extra_day_rt_per_mile" size="20" value="<%=curExtraDayRtPerMile %>" style="text-align: right" onBlur="this.value=formatCurrency(this.value);"></td>
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
      15b. Extra Hour rate:</td>
      <td width="274" height="19" bgcolor="#C0C0C0">
      
		<input type="text" name="extra_hr_rt" size="20" value="<%=curExtraHrRt %>" style="text-align: right" onBlur="this.value=formatCurrency(this.value);"></td>
      <td width="319" height="19">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="19">&nbsp;</td>
      <td width="217" height="19">&nbsp;</td>
      <td width="188" height="19">&nbsp;</td>
      <td width="204" height="19" bgcolor="#C0C0C0">
      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Extra Hour free miles:</td>
      <td width="274" height="19" bgcolor="#C0C0C0">
      
		<input type="text" name="extra_hr_miles" size="20" value="<%=curExtraHrMiles %>" style="text-align: right" onBlur="this.value=formatNumber(this.value);"> 
		(blank = unlimited)</td>
      <td width="319" height="19">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="19">&nbsp;</td>
      <td width="217" height="19">&nbsp;</td>
      <td width="188" height="19">&nbsp;</td>
      <td width="204" height="19" bgcolor="#C0C0C0">
      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Extra Hour $/extra mile:</td>
      <td width="274" height="19" bgcolor="#C0C0C0">
      
		<input type="text" name="extra_hr_rt_per_mile" size="20" value="<%=curExtraHrRtPerMile %>" style="text-align: right" onBlur="this.value=formatCurrency(this.value);"></td>
      <td width="319" height="19">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="19">&nbsp;</td>
      <td width="217" height="19">&nbsp;</td>
      <td width="188" height="19">&nbsp;</td>
      <td width="204" height="19">
<!--      <input type="reset" name="reset" value="Hide Extra Rate detail" onclick="javascript:toggleLayer('ExtraDay');" style="float: right" />
-->
		</td>
      <td width="274" height="19">
      &nbsp;</td>
      <td width="319" height="19">&nbsp;</td>
    </tr>
</table> 
</div>
  <table border="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#111111" width="1210" id="new_alert_part2" background="images/alt_color.gif" >
    <tr>
      <td width="8" height="19">&nbsp;</td>
      <td width="217" height="19">&nbsp;</td>
      <td width="210" height="19">&nbsp;</td>
      <td width="510" height="19">&nbsp;</td>
      <td width="262" height="19">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="22">&nbsp;</td>
      <td width="217" height="22">&nbsp;</td>
      <td width="210" height="22">16. Search Type:</td>
      <td width="510" height="22">
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
      <td width="262" height="22">
		&nbsp;</td>
    </tr>
	<div id="profiles" style="display: none;"> 
    <tr>
      <td width="8" height="22">&nbsp;</td>
      <td width="217" height="22">&nbsp;</td>
      <td width="210" height="22">&nbsp;&nbsp;&nbsp;&nbsp; Unselected 
		Profile(s):<br>
      <br></td>
      <td width="510" height="22" class="style3">
		<select name="un_selected_profile_id" size="4" multiple style="width:487px; font-family:Verdana; font-size:10pt; height:70">
				<% strProfileID = "," & strProfileID & "," %>
				<% strProfileID = Replace(strProfileID, " ", "") %>
				<% Dim strSelectProfileID  %>
				<% Dim strSelectProfileDesc  %>
				<% intLoopCount = 0 %>
                <% While (adoRS9.EOF = False) And (intLoopCount < 800)  %> 
                	<% If adoRS9.Fields("profile_id").Value = 0 Then %>
	                	<% adoRS9.MoveNext %>
	                <% End If %>
                	<% If InStr(1, strProfileID, "," & adoRS9.Fields("profile_id").Value & ",") > 0 Then %>
	                	<% strSelectProfileID = strSelectProfileID & adoRS9.Fields("profile_id").Value & "," %>
	                	<% strSelectProfileDesc = strSelectProfileDesc & adoRS9.Fields("desc").Value & "," %>
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
      <td width="510" height="22" class="style4">
                      <img border="0" src="images/down_button.GIF" width="24" height="22"  onclick="moveDualList( document.create_alert.un_selected_profile_id, document.create_alert.profile_id, false );return false" alt="Select profile" class="style10" >
                      <img border="0" src="images/up_button.GIF"   width="24" height="22"  onclick="moveDualList( document.create_alert.profile_id, document.create_alert.un_selected_profile_id, false );return false" alt="Un-select profile" class="style10" >&nbsp;&nbsp;&nbsp;
                    
		</td>
      <td width="262" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="22">&nbsp;</td>
      <td width="217" height="22">&nbsp;</td>
      <td width="210" height="22">&nbsp;&nbsp;&nbsp;&nbsp; 
		Selected Profile(s):</td>
      <td width="510" height="22" class="style3">
		<select name="profile_id" size="4" multiple style="width:487px; font-family:Verdana; font-size:10pt; height:70">
				<% Dim strProfileIDs   %>
				<% Dim strProfileDescs %>
				<% strProfileIDs = Split(strSelectProfileID, ",")%>
				<% strProfileDescs = Split(strSelectProfileDesc, ",")%>
                <% For intLoopCount = LBound(strProfileIDs) To UBound(strProfileIDs)  %> 
                	<% If IsNumeric(strProfileIDs(intLoopCount)) Then %>
                	<option value="<%=strProfileIDs(intLoopCount) %>"><%=strProfileDescs(intLoopCount) %></option>
                	<% End If %>
				<% Next %>                
		</select>
	  </td>
      <td width="262" height="22">&nbsp;</td>
    </tr>
        
</table>
  <table border="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#111111" width="1210" id="new_alert_part3" background="images/alt_color.gif" height="561">

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
      17. Rate Range:</td>
      <td width="510" height="22" colspan="8">
      (leave blank or enter a zero to indicate no limit)</td>
      <td width="262" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="22">&nbsp;</td>
      <td width="217" height="22">&nbsp;</td>
      <td width="210" height="22">&nbsp;&nbsp;&nbsp;&nbsp; 
      Maximum:</td>
      <td width="510" height="22" colspan="8">
      <input type="text" name="rate_maximum" size="20" style="width:127; font-family:Verdana; font-size:10pt; text-align:right; height:21" value="<%=strRangeMax %>" 
      onkeyup='this.onchange();' onchange='tV=(this.value.replace(/[^\d\.]/g,"")).replace(/[\.]{2,}/g,".");if(tV!=this.value){this.value=tV;}'></td>
      <td width="262" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="22">&nbsp;</td>
      <td width="217" height="22">
      &nbsp;</td>
      <td width="210" height="25">&nbsp;&nbsp;&nbsp;&nbsp; 
      Minimum:</td>
      <td width="510" height="25" colspan="8">
      <input type="text" name="rate_minimum" size="20" style="width:127; font-family:Verdana; font-size:10pt; text-align:right; height:22" value="<%=strRangeMin %>"></td>
      <td width="262" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="22">&nbsp;</td>
      <td width="217" height="22">
      &nbsp;</td>
      <td width="210" height="22" align="center">OR</td>
      <td width="510" height="22" colspan="8">
      &nbsp;</td>
      <td width="262" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="22">&nbsp;</td>
      <td width="217" height="22">
      &nbsp;</td>
      <td width="210" height="22">
      &nbsp;&nbsp;&nbsp; Max. / Min. Schedule</td>
      <td width="510" height="22" colspan="8">
      
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
		<a title="Click to edit or create a rule schedule" href="rate_rule_maxmin_schedule_a.asp" onclick="window.open('rate_rule_maxmin_schedule_a.asp','manage_schedule','toolbar=no,status=no,scrollbars=yes,resizable=no,width=820,height=600'); return false;">Manage schedules</a></td>
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
    <!-- Begin follow-on rules -->
    <tr>
      <td width="8" height="22">&nbsp;</td>
      <td width="217" height="22">&nbsp;</td>
      <td width="210" height="22">18. When situation = true:</td>
      <td width="510" height="22" colspan="8">
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
      <td width="210" height="22">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;When situation = false:</td>
      <td width="510" height="22" colspan="8">
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
      <td width="510" height="22" colspan="8">&nbsp;</td>
      <td width="262" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="22">&nbsp;</td>
      <td width="217" height="22">&nbsp;</td>
      <td width="210" height="25">
      19. Utilization Range: </td>
      <td width="510" height="25" colspan="8">
      (leave blank or enter a zero to indicate no limit)</td>
      <td width="262" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="22">&nbsp;</td>
      <td width="217" height="22">&nbsp;</td>
      <td width="210" height="22">&nbsp;&nbsp;&nbsp;&nbsp; Days out:</td>
      <td width="64" height="22" align="center" class="style5">Same</td>
      <td width="64" height="22" align="center" class="style5">Next</td>
      <td width="64" height="22" align="center" class="style5">2 - 4</td>
      <td height="22" align="center" style="width: 64px" class="style5">5 - 7</td>
      <td width="64" height="22" align="center" class="style5">8 - 14</td>
      <td width="64" height="22" align="center" class="style5">15 - 30</td>
      <td width="64" height="22" align="center" class="style5">31 - 50</td>
      <td height="22" align="center" style="width: 64px" class="style5">51 +</td>
	  <td width="262" height="22" >&nbsp;</td>
      <td height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" style="height: 7px"></td>
      <td width="217" style="height: 7px"></td>
      <td width="210" style="height: 7px">&nbsp;&nbsp;&nbsp;&nbsp; Maximum:</td>
      <td width="64" style="height: 7px" class="style5"><input type="text"  name="util_max_0" size="20" style="width:59; font-family:Verdana; font-size:10pt; text-align:right; height:22" value="<%=strUtilMax0 %>"></td>
      <td width="64" style="height: 7px" class="style5"><input type="text"  name="util_max_1" size="20" style="width:59; font-family:Verdana; font-size:10pt; text-align:right; height:22" value="<%=strUtilMax1 %>"></td>
      <td width="64" style="height: 7px" class="style5"><input type="text"  name="util_max_2" size="20" style="width:59; font-family:Verdana; font-size:10pt; text-align:right; height:22" value="<%=strUtilMax2 %>"></td>
      <td style="height: 7px; width: 64px;" class="style5"><input type="text"  name="util_max_3" size="20" style="width:59; font-family:Verdana; font-size:10pt; text-align:right; height:22" value="<%=strUtilMax3 %>"></td>
      <td width="64" style="height: 7px" class="style5"><input type="text"  name="util_max_4" size="20" style="width:59; font-family:Verdana; font-size:10pt; text-align:right; height:22" value="<%=strUtilMax4 %>"></td>
      <td width="64" style="height: 7px" class="style5"><input type="text"  name="util_max_5" size="20" style="width:59; font-family:Verdana; font-size:10pt; text-align:right; height:22" value="<%=strUtilMax5 %>"></td>
      <td width="64" style="height: 7px" class="style5"><input type="text"  name="util_max_6" size="20" style="width:59; font-family:Verdana; font-size:10pt; text-align:right; height:22" value="<%=strUtilMax6 %>"></td>
      <td style="height: 7px; width: 64px;" class="style5"><input type="text"  name="util_max_7" size="20" style="width:59; font-family:Verdana; font-size:10pt; text-align:right; height:22" value="<%=strUtilMax7 %>"></td>
	  <td width="262" style="height: 7px" >&nbsp;</td>
	  <td style="height: 7px"></td>
    </tr>
    <tr>
      <td width="8" style="height: 24px"></td>
      <td width="217" style="height: 24px"></td>
      <td width="210" style="height: 24px">&nbsp;&nbsp;&nbsp;&nbsp; Minimum: </td>
      <td width="64" style="height: 24px" class="style5"><input type="text" name="util_min_0" size="20" style="width:59; font-family:Verdana; font-size:10pt; text-align:right; height:22" value="<%=strUtilMin0 %>"></td>
      <td width="64" style="height: 24px" class="style5"><input type="text" name="util_min_1" size="20" style="width:59; font-family:Verdana; font-size:10pt; text-align:right; height:22" value="<%=strUtilMin1 %>"></td>
      <td width="64" style="height: 24px" class="style5"><input type="text" name="util_min_2" size="20" style="width:59; font-family:Verdana; font-size:10pt; text-align:right; height:22" value="<%=strUtilMin2 %>"></td>
      <td style="height: 24px; width: 64px;" class="style5"><input type="text" name="util_min_3" size="20" style="width:59; font-family:Verdana; font-size:10pt; text-align:right; height:22" value="<%=strUtilMin3 %>"></td>
      <td width="64" style="height: 24px" class="style5"><input type="text" name="util_min_4" size="20" style="width:59; font-family:Verdana; font-size:10pt; text-align:right; height:22" value="<%=strUtilMin4 %>"></td>
      <td width="64" style="height: 24px" class="style5"><input type="text" name="util_min_5" size="20" style="width:59; font-family:Verdana; font-size:10pt; text-align:right; height:22" value="<%=strUtilMin5 %>"></td>
      <td width="64" style="height: 24px" class="style5"><input type="text" name="util_min_6" size="20" style="width:59; font-family:Verdana; font-size:10pt; text-align:right; height:22" value="<%=strUtilMin6 %>"></td>
      <td style="height: 24px; width: 64px;" class="style5"><input type="text" name="util_min_7" size="20" style="width:59; font-family:Verdana; font-size:10pt; text-align:right; height:22" value="<%=strUtilMin7 %>"></td>
	  <td width="262" style="height: 24px">&nbsp;</td>
      <td style="height: 24px"></td>
    </tr>
    <tr>
      <td width="8" height="22">&nbsp;</td>
      <td width="217" height="22">&nbsp;</td>
      <td width="210" height="22">&nbsp;&nbsp;&nbsp;<a target="_blank" href="bulk_update_utilization.asp">bulk update</a></td>
      <td width="510" height="22" colspan="8">
      <input type="checkbox" name="util_in_percent" id="util_in_percent" value="True" checked disabled ><label for="util_in_percent">values are listed as percentages (please do not include percent signs)</label></td>
      <td width="262" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="22">&nbsp;</td>
      <td width="217" height="22">&nbsp;</td>
      <td width="210" height="22" align="center">OR</td>
      <td width="510" height="22" colspan="8">
      &nbsp;</td>
      <td width="262" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" height="22">&nbsp;</td>
      <td width="217" height="22">&nbsp;</td>
      <td width="210" height="22">
      &nbsp;&nbsp;&nbsp; Utilization Schedule</td>
      <td width="510" height="22" colspan="8">
      
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
		<a title="Click to edit or create a rule schedule" href="rate_rule_schedule_a.asp" onclick="window.open('rate_rule_schedule_a.asp','manage_schedule','toolbar=no,status=no,scrollbars=yes,resizable=no,width=820,height=600'); return false;">Manage schedules</a></td>
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
      
      20. Pickup Days of Week:</td>
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
      <td width="210" height="22" class="style1">
      21. Optional Post Actions:</td>
      <td height="22" colspan="4" class="style2">
      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Action Description</td>
      <td width="510" height="22" colspan="4" style="width: 255px" class="style1">
      &nbsp;&nbsp;&nbsp;&nbsp;
      Action Amt.</td>
      <td width="262" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="8" style="height: 22px"></td>
      <td width="217" style="height: 22px"></td>
      <td width="210" style="height: 22px">&nbsp;&nbsp;&nbsp;&nbsp;<a target="_blank" href="bulk_update_post_actions.asp">bulk update</a></td>
      <td colspan="2" style="height: 22px">
	
      <span class="style1">First Action:</span></td>
      <td colspan="2" style="height: 22px">
	
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
	  </select></td>
      <td width="510" colspan="4" style="height: 22px; width: 255px">
		
		<input	name="rule_post_action_amt1" 
				type="text" 
				value="<%=strRulePostActionAmt1 %>" 
				style="width:127; font-family:Verdana; font-size:10pt; text-align:right; height:22"
				onblur="extractNumber(this,4,true);"
				onkeyup="extractNumber(this,4,true);"
				onkeypress="return blockNonNumbers(this, event, true, true);">
		</td>
      <td width="262" style="height: 22px"></td>
    </tr>
    <tr>
      <td width="8" height="22">&nbsp;</td>
      <td width="217" height="22">&nbsp;</td>
      <td width="210" height="22">
      &nbsp;</td>
      <td height="22" colspan="2">
      <span class="style1">Second Action:</span></td>
      <td height="22" colspan="2">
      
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
		</select></td>
      <td width="510" height="22" colspan="4" style="width: 255px">
       
		<input	name="rule_post_action_amt2" 
				type="text" 
				value="<%=strRulePostActionAmt2 %>" 
				style="width:127; font-family:Verdana; font-size:10pt; text-align:right; height:22"
				onblur="extractNumber(this,4,true);"
				onkeyup="extractNumber(this,4,true);"
				onkeypress="return blockNonNumbers(this, event, true, true);">
	  </td>
      <td width="262" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="8"></td>
      <td width="217"></td>
      <td width="210">
      </td>
      <td colspan="2">
      <span class="style1">Third Action:</span></td>
      <td colspan="2">
      
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
		</select></td>
      <td width="510" colspan="4" style="width: 255px">
       
			<input	name="rule_post_action_amt3" 
				type="text" 
				value="<%=strRulePostActionAmt3 %>" 
				style="width:127; font-family:Verdana; font-size:10pt; text-align:right; height:22"
				onblur="extractNumber(this,4,true);"
				onkeyup="extractNumber(this,4,true);"
				onkeypress="return blockNonNumbers(this, event, true, true);"></td>
      <td width="262"></td>
    </tr>
    <tr>
      <td width="8" height="22">&nbsp;</td>
      <td width="217" height="22">&nbsp;</td>
      <td width="210" height="22">
      &nbsp;</td>
      <td height="22" colspan="2">
      <span class="style1">Forth Action:</span></td>
      <td height="22" colspan="2">
      
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
		</select></td>
      <td width="510" height="22" colspan="4" style="width: 255px">
       
 	  <input	name="rule_post_action_amt4" 
				type="text" 
				value="<%=strRulePostActionAmt4 %>" 
				style="width:127; font-family:Verdana; font-size:10pt; text-align:right; height:22"
				onblur="extractNumber(this,4,true);"
				onkeyup="extractNumber(this,4,true);"
				onkeypress="return blockNonNumbers(this, event, true, true);">
				</td>
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
      
      22. Duplicate Suggestion:</td>      <td width="510" height="22" colspan="8">
		<select name="additional_days">
		<% intCount = 0        %>
		<% While intCount < 31 %>
		<%   If intCount = intAdditionalDays Then %>
				<option selected="selected" ><%=intCount %></option>
		<%   Else   %>
				<option ><%=intCount %></option>
		<%   End If  
		     intCount = intCount + 1
		   Wend
		   
		%>
		
		
		</select> additional days (0 to 30)</td>      <td width="262" height="22">
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
      <input type="checkbox" name="extraday" value="True"  checked="true" >
      <% Else %>
      <input type="checkbox" name="extraday" value="True" >
      <% End If %>
      Calculate the rate while factoring in the extra day(s) (weekly rates only)</td>      <td width="262" height="22">&nbsp;</td>
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
      <input type="checkbox" name="automatic" value="True"  checked="true" >
      <% Else %>
      <input type="checkbox" name="automatic" value="True" >
      <% End If %>
      Run this rule in automatic mode (no user review or approval req.)</td>      <td width="262" height="22">
		&nbsp;</td>
    </tr>    <tr>
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
      
          <input name="<%=strButton %>" type="submit" id="submit" value="    <%=strButton %>   " class="rh_button">
          </td>
      <td width="510" height="23" colspan="8">
      &nbsp;  (<a href="#maint_top">back to top</a>)&nbsp;(<a href="#grid_top">back to grid</a>)</td>
      <td width="262" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="8"  height="61">&nbsp;</td>
      <td width="217" height="61">&nbsp;</td>
      <td width="720" height="61" colspan="9">
      <!--
      <font size="1">closed = <%=blnIgnoreClosed  %>
      -->
      Rate rule tester 
		<a target="_blank" href="rate_rule_tester_20130718.asp">click here</a>
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
     </tr>
  </table>
  <input type="hidden" name="refresh_from" value="create">
  <input type="hidden" name="rule_status" value="E">
  <input type="hidden" name="action" value="0">
</form>
<!-- Content goes before this comment -->
<!-- JUSTTABS BOTTOM OPEN -->

</td>
</tr>
</table>
<!-- JUSTTABS BOTTOM CLOSE -->
<p>&nbsp;&nbsp;[<a target="_self" href="rezcentral_tethering_20130715.asp">RezCentral tethering settings</a>]</p>
  <table border="0" bordercolor="#FFFFFF" width="1210" height="4" class="style7">
    <tr>
      <td>
<!--#INCLUDE FILE="footer.asp"-->
      </td>
    </tr>
  </table>
<div id="calbox" class="calboxoff"></div>
</body>
<script language="javascript" type="text/javascript">
	document.create_alert.alert_desc.focus();
</script>
</html>
<%	
	Rem Clean up
	Set adoCmd = Nothing 
	Set adoCmd1 = Nothing 
	Set adoCmd2 = Nothing 
	Set adoCmd3 = Nothing 
	Set adoCmd4 = Nothing
	Set adoCmd9 = Nothing 

	adoRS4.Close
	Set adoRS4 = Nothing
	Set adoRS5 = Nothing
    Set adoRS9 = Nothing
    
    Set adoCmdOnSuccess = Nothing
    Set adoRSOnSuccess = Nothing
    Set adoCmdOnFailure = Nothing
    Set adoRSOnFailure = Nothing
	
	End If
%>