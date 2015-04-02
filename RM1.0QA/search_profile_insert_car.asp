<%@ Language=VBScript %>
<!-- #INCLUDE FILE="inc/adovbs.asp" --> 

<% Response.Expires = -1
   Response.cachecontrol="private" 
   Response.AddHeader "pragma", "no-cache"

   Server.ScriptTimeout = 0
   
   On Error Resume Next
 
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
	Dim Whatever
	Dim strVendors
	Dim strSelectDataSource 

	'strClientUserid = Request.Form("userid")
	'strCity = Request.Form("city")
	'strCarType = Request.Form("car_type")
	'strCompany = Request.Form("company")
	
	
	strUserId = Request.Cookies("rate-monitor.com")("user_id")
	intRptLimit = CInt(Request.Cookies("rate-monitor.com")("rpt_limit"))
	
	
	strConn = Session("pro_con")
	
	Set adoCmd = CreateObject("ADODB.Command")

	adoCmd.ActiveConnection =  strConn
	adoCmd.CommandText = "car_shop_profile_insert"
	adoCmd.CommandType = 4


	If (Request.Form("profile") > 0) And (Request.Form("profile_save_as") = "") Then
		adoCmd.Parameters.Append adoCmd.CreateParameter("@profile_id", 3, 1, 0, Request.Form("profile"))

	Else
		adoCmd.Parameters.Append adoCmd.CreateParameter("@profile_id", 3, 1, 0, Null)

	End If

	adoCmd.Parameters.Append adoCmd.CreateParameter("@desc", 200, 1, 255, Request.Form("profile_save_as"))
	adoCmd.Parameters.Append adoCmd.CreateParameter("@org_id", 3, 1, 0, Session("org_id"))
	adoCmd.Parameters.Append adoCmd.CreateParameter("@city_cd", 200, 1, 1024, Request.Form("cities_list"))  'Request.Form("pickup_city"))
	Rem Based upon a change from Expedia, they chose multiple cities over one way, so null it out
	Rem Based upon a change from Expedia, they chose multiple cities over one way, so null it out
	If (Trim(Request.Form("return_city") & "") = "") Or (UCase(Request.Form("return_city") & "") = "(SAME)") Then
		adoCmd.Parameters.Append adoCmd.CreateParameter("@rtrn_city_cd", 200, 1, 5, Null)
	Else
		adoCmd.Parameters.Append adoCmd.CreateParameter("@rtrn_city_cd", 200, 1, 5, UCase(Request.Form("return_city")))
	End If
	adoCmd.Parameters.Append adoCmd.CreateParameter("@shop_car_type_cds", 200, 1, 1024, Request.Form("car_type_list"))

	strVendors = Request.Form("company_list")
	
	If Right(strVendors, 1) = "," Then
		strVendors = Mid(strVendors, 1, Len(strVendors) - 1)
	
	End If
	
	adoCmd.Parameters.Append adoCmd.CreateParameter("@vend_cds", 200, 1, 1024, strVendors)
	
	strSelectDataSource = Request.Form("data_source")
	Select Case strSelectDataSource
	
	'NOTE: IMPORTANT - if you add or modify these special cases
    '      you must update "search_criteria_car.asp" also
    '                       "search_request_insert_car.asp"
	
		Case "M01"
			Rem We have to translate the multi-site ones for selection on the site
			strSelectDataSource = "EXP,VAC" 
   		
		Case "M02"
			Rem We have to translate the multi-site ones for selection on the site
			strSelectDataSource = "EXP,VMW" 

		Case "M03"
			Rem We have to translate the multi-site ones for selection on the site
			strSelectDataSource = "EXS,VAC" 
   		
		Case "M04"
			Rem We have to translate the multi-site ones for selection on the site
			strSelectDataSource = "EXD,VAC" 
 			
		Case "M05"
			Rem We have to translate the multi-site ones for selection on the site
			strSelectDataSource = "EXP,VAD" 

		Case "M06"
			Rem We have to translate the multi-site ones for selection on the site
			strSelectDataSource = "WSR,VZE" 

		Case "M07"
			Rem We have to translate the multi-site ones for selection on the site
			strSelectDataSource = "WSG,VZE" 

		Case "M08"
			Rem We have to translate the multi-site ones for selection on the site
			strSelectDataSource = "WSA,VZE" 
   		
		Case "M09"
			Rem We have to translate the multi-site ones for selection on the site
			strSelectDataSource = "EXP,VZI" 

		Case "M10"
			Rem We have to translate the multi-site ones for selection on the site
			strSelectDataSource = "WSR,VAD" 

		Case "M11"
			Rem We have to translate the multi-site ones for selection on the site
			strSelectDataSource = "CRX,VET" 

		Case "M12"
			Rem We have to translate the multi-site ones for selection on the site
			strSelectDataSource = "WSR,VZI,VZD,VET"

		Case "M13"
			Rem We have to translate the multi-site ones for selection on the site
			strSelectDataSource = "WSR,VEZ"

		Case "M14"
			Rem We have to translate the multi-site ones for selection on the site
			strSelectDataSource = "FFR,VET,VZE,VZT,VZR,VAL,VZL"
   		
   	End Select

	
	adoCmd.Parameters.Append adoCmd.CreateParameter("@data_sources", 200, 1, 1024, strSelectDataSource)
	adoCmd.Parameters.Append adoCmd.CreateParameter("@shop_currency_cd", 200, 1, 3, Request.Form("display_currency"))
	adoCmd.Parameters.Append adoCmd.CreateParameter("@shop_location_categ", 200, 1, 9, "T,O")
	adoCmd.Parameters.Append adoCmd.CreateParameter("@shop_rt_categ", 200, 1, 9, "S,P")
	adoCmd.Parameters.Append adoCmd.CreateParameter("@shop_rt_type", 200, 1, 9, "D,E")
	adoCmd.Parameters.Append adoCmd.CreateParameter("@airline_arv_ind", 200, 1, 1, "N")
	adoCmd.Parameters.Append adoCmd.CreateParameter("@shop_pos_cd", 200, 1, 5, Request.Form("pos"))
	
	Rem ==========================================================================================
	Rem     If changes are made to extra_criteria you need to update search_request_insert_car.asp
	Rem     also since it is a twin of this process
	Rem ==========================================================================================
	
	Dim DiscountCompanies 
	Dim DiscountCodes
	Dim strExtraCriteria 
	Dim DiscountCodesArray

	DiscountCompanies = Split(Request.Form("company_list"), ",", -1, 1)
	DiscountCodes = Trim(Request.Form("discount_code"))

	Rem if the data source is not TRD and their are no discount codes then don't bother with them
	If Replace(DiscountCodes, ",", "") = "" Then
		DiscountCodes = ""
	End If
	
	Rem default it to blank
	strExtraCriteria = ""

	Rem Format for discount codes
	Rem DISC-ZE|123DBC|DISC-AL|34DSE|

	If Len(DiscountCodes) > 0 Then

		DiscountCodesArray = Split(DiscountCodes, ",", -1, 1)
		intArrayHolder = UBound(DiscountCodesArray)
	
		If IsNumeric(intArrayHolder) Then
	
			While intArrayHolder >= 0 
				strExtraCriteria = strExtraCriteria & "DISC-" & DiscountCompanies(intArrayHolder) & "|" & Trim(DiscountCodesArray(intArrayHolder)) & "|" 
				intArrayHolder = intArrayHolder - 1
			Wend
	
		End If
		
	End If


	adoCmd.Parameters.Append adoCmd.CreateParameter("@extra_criteria", 200, 1, 255, strExtraCriteria)
	adoCmd.Parameters.Append adoCmd.CreateParameter("@lor", 3, 1, 0, Request.Form("lor"))
	adoCmd.Parameters.Append adoCmd.CreateParameter("@arv_tm", 200, 1, 5, Request.Form("arrival_time"))
	adoCmd.Parameters.Append adoCmd.CreateParameter("@rtrn_tm", 200, 1, 5, Request.Form("return_time"))
	adoCmd.Parameters.Append adoCmd.CreateParameter("@begin_arv_dt", 135, 1, 0, Request.Form("begin_date"))
	
	Rem We need to limit end dates for users that exceed their limit
	datEndDate = Request.Form("end_date")
	
	If intRptLimit > 0 Then
		If (DateDiff("d", Request.Form("begin_date"), datEndDate) >= intRptLimit) Then
			datEndDate = DateAdd("d", (intRptLimit - 1), Request.Form("begin_date"))
		End If
'		adoCmd.Parameters.Append adoCmd.CreateParameter("@end_arv_dt", 135, 1, 0, DateAdd("d", (intRptLimit - 1), Request.Form("begin_date")))
'	Else
'		adoCmd.Parameters.Append adoCmd.CreateParameter("@end_arv_dt", 135, 1, 0, Request.Form("end_date"))
	End If

	adoCmd.Parameters.Append adoCmd.CreateParameter("@end_arv_dt", 135, 1, 0, datEndDate)
	
	adoCmd.Parameters.Append adoCmd.CreateParameter("@dow_list", 200, 1, 13, Replace(Request.Form("dow_list")," ",""))

	If Request.Form("scheduled_time") = "" Then
		adoCmd.Parameters.Append adoCmd.CreateParameter("@scheduled_dttm", 135, 1, 0, Null)
	Else
		adoCmd.Parameters.Append adoCmd.CreateParameter("@scheduled_dttm", 135, 1, 0, "1/1/1971 " & FormatDateTime(Request.Form("scheduled_time"), 4))
	End If

	adoCmd.Parameters.Append adoCmd.CreateParameter("@start_dttm", 135, 1, 0, Null)
	adoCmd.Parameters.Append adoCmd.CreateParameter("@complete_dttm", 135, 1, 0, Null)
	adoCmd.Parameters.Append adoCmd.CreateParameter("@change_dttm", 135, 1, 0, Now)
	adoCmd.Parameters.Append adoCmd.CreateParameter("@request_status", 200, 1, 1, "N")
	adoCmd.Parameters.Append adoCmd.CreateParameter("@client_userid", 200, 1, 99, Request.Cookies("rate-monitor.com")("client_userid"))
	adoCmd.Parameters.Append adoCmd.CreateParameter("@client_system_cd", 200, 1, 5, "SYM")
	adoCmd.Parameters.Append adoCmd.CreateParameter("@server_system_cd", 200, 1, 5, "SYM")
	adoCmd.Parameters.Append adoCmd.CreateParameter("@work_units", 3, 1, 0, 1)
	adoCmd.Parameters.Append adoCmd.CreateParameter("@request_msg", 200, 1, 99, "")
	adoCmd.Parameters.Append adoCmd.CreateParameter("@vend_cd", 200, 1, 255, "")
	adoCmd.Parameters.Append adoCmd.CreateParameter("@email_address", 200, 1, 255, Request.Form("recipient_address"))
	adoCmd.Parameters.Append adoCmd.CreateParameter("@user_id", 3, 1, 0, strUserId)
	adoCmd.Parameters.Append adoCmd.CreateParameter("@action_id", 3, 1, 0, Request.Form("search_action"))
	adoCmd.Parameters.Append adoCmd.CreateParameter("@highlight_vendor", 200, 1, 2, Request.Form("highlighted_company"))
	adoCmd.Parameters.Append adoCmd.CreateParameter("@schedule_dow_list", 200, 1, 13, Replace(Request.Form("schedule_dow_list")," ",""))
	adoCmd.Parameters.Append adoCmd.CreateParameter("@exact_dates", 11, 1, 0, Request.Form("exact_dates"))
	adoCmd.Parameters.Append adoCmd.CreateParameter("@output", 200, 1, 10, Request.Form("output"))
	Select Case Request.Form("output")
	 	 Case "html"
			adoCmd.Parameters.Append adoCmd.CreateParameter("@output_style", 3, 1, 0, Request.Form("html_style"))
			
		case "ftp"
			adoCmd.Parameters.Append adoCmd.CreateParameter("@output_style", 3, 1, 0, Request.Form("ftp_style"))

		Case Else
			adoCmd.Parameters.Append adoCmd.CreateParameter("@output_style", 3, 1, 0, 1)
		
	End Select		
	adoCmd.Parameters.Append adoCmd.CreateParameter("@display_rate_type", 3, 1,   0, Request.Form("display_rate_type"))	
	adoCmd.Parameters.Append adoCmd.CreateParameter("@alert_address",   200, 1, 255, Request.Form("alert_address"))
	If Request.Form("oneway_reverse") = "True" Then
		adoCmd.Parameters.Append adoCmd.CreateParameter("@oneway_reverse", 11, 1, 0, 1)
	Else
		adoCmd.Parameters.Append adoCmd.CreateParameter("@oneway_reverse", 11, 1, 0, 0)
	End If

	If Request.Form("search_schedules") > 0 Then
		adoCmd.Parameters.Append adoCmd.CreateParameter("@schedule_grp_id",        3, 1,   0, Request.Form("search_schedules")) 
	Else
		adoCmd.Parameters.Append adoCmd.CreateParameter("@schedule_grp_id",        3, 1,   0, Null) ' no schedule attached
	End If
    if(trim(Request.Form("boundary_start_date")) = "") then
	    adoCmd.Parameters.Append adoCmd.CreateParameter("@shop_boundary_start",        135, 1,   0, Null)
    else
	    adoCmd.Parameters.Append adoCmd.CreateParameter("@shop_boundary_start",        135, 1,   0, Request.Form("boundary_start_date"))
    end if
    if(trim(Request.Form("boundary_end_date")) = "") then
	    adoCmd.Parameters.Append adoCmd.CreateParameter("@shop_boundary_end",        135, 1,   0, Null)
    else
	    adoCmd.Parameters.Append adoCmd.CreateParameter("@shop_boundary_end",         135, 1,   0, Request.Form("boundary_end_date"))
    end if
	adoCmd.Parameters.Append adoCmd.CreateParameter("@division_id",         3, 1,   0, Session("division_id"))
	
	If err.number <> 0 Then
	   pad="&nbsp;&nbsp;&nbsp;&nbsp;"
	   response.write "<b>VBScript Errors Occured!<br>"
	   response.write "</b><br>"
	   response.write pad & "Error Number= #<b>" & err.number & "</b><br>"
	   response.write pad & "Error Desc.= <b>" & err.description & "</b><br>"
	   response.write pad & "Help Context= <b>" & err.HelpContext & "</b><br>"
	   response.write pad & "Help File Path=<b>" & err.helpfile & "</b><br>"
	   response.write pad & "Error Source= <b>" & err.source & "</b><br><hr>"

	Else
		If Request.Form("debug") = "true" Then
	
		Else
			Rem Set adoRS = adoCmd.Execute
			Call adoCmd.Execute()
	
			If err.number = 0 Then
				Server.Transfer "search_profiles_car.asp"	
			Else
			
			   pad="&nbsp;&nbsp;&nbsp;&nbsp;"
			   response.write "<b>VBScript Errors Occured!<br>"
			   response.write "</b><br>"
			   response.write pad & "Error Number= #<b>" & err.number & "</b><br>"
			   response.write pad & "Error Desc.= <b>" & err.description & "</b><br>"
			   response.write pad & "Help Context= <b>" & err.HelpContext & "</b><br>"
			   response.write pad & "Help File Path=<b>" & err.helpfile & "</b><br>"
			   response.write pad & "Error Source= <b>" & err.source & "</b><br><hr>"
			
				Session("error_msg") = "An error was encountered while request your search. Please contact Rate-Highway support"
				'Server.Transfer "search_criteria_car.asp"
			End If

		End If
	

	End If

	Set adoCmd = Nothing

	
	
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>search_request_insert</title>
</head>
<body>
We are sorry, an error has occurred, please <a href="search_criteria_car.asp">click 
here</a> to return to Search Criteria<p>Please print and fax this page to 
Rate-Highway Support at (888) 551-0029</p>
<p><font size="2">Debug information</font></p>
<p><font size="2">Request.Form("dow_list</font><font size="2">") = <%= Len(Request.Form("dow_list")) %><br>
strDowList = <%= Len(strDowList) %><br>


			   <%	For Each Whatever In Request.Form
						Response.Write Whatever & " = <b>" & Request.Form(Whatever) & "</b> <br>"

       
					Next
				%>

</font></p>
<!--#INCLUDE FILE="footer.asp"-->
</body>
</html>
