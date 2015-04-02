<%@ Language=VBScript %>
<!-- #INCLUDE FILE="inc/adovbs.asp" --> 
<!-- #INCLUDE FILE="inc/login_check_ex.asp" -->

<% Response.Expires = -1
   Response.cachecontrol = "private" 
   Response.AddHeader "pragma", "no-cache"

   Server.ScriptTimeout = 0
   
   On Error Resume Next
 
	Dim strConn	
	Dim adoCmd	
	Dim adoRS
	Dim adoCmd2	
	Dim adoRS2
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
	Dim varDiscounts()
	Dim Whatever
	Dim strDowList
	Dim strDataSourceList
	Dim datEndDate
	Dim strDiscountList
	Dim intWorkUnits
	Dim strArray
	Dim intArrayHolder
	Dim strSelectDataSource 
	
	intWorkUnits = intArrayHolder  

	strUserId = Request.Cookies("rate-monitor.com")("user_id")
	intRptLimit = CInt(Request.Cookies("rate-monitor.com")("rpt_limit"))
	
	Rem The checkboxes put spaces between the numbers and the commas - strip them out Danno
	strDowList = Request.Form("dow_list")
	strDowList = Replace(strDowList, " ", "")
	
	strConn = Session("pro_con")
	

	Set adoCmd = CreateObject("ADODB.Command")

	adoCmd.ActiveConnection =  strConn
	adoCmd.CommandText = "car_shop_request_insert"
	adoCmd.CommandType = 4

	adoCmd.Parameters.Append adoCmd.CreateParameter("@shop_request_id", 3, 1, 0, Null)
	adoCmd.Parameters.Append adoCmd.CreateParameter("@org_id", 3, 1, 0, Session("org_id"))
	adoCmd.Parameters.Append adoCmd.CreateParameter("@city_cd", 200, 1, 1024, Request.Form("cities_list"))  'Request.Form("pickup_city"))
	Rem Based upon a change from Expedia, they chose multiple cities over one way, so null it out

	If (Trim(Request.Form("return_city") & "") = "") Or (UCase(Request.Form("return_city") & "") = "(SAME)") Then
		adoCmd.Parameters.Append adoCmd.CreateParameter("@rtrn_city_cd", 200, 1, 6, Null)
	Else
		adoCmd.Parameters.Append adoCmd.CreateParameter("@rtrn_city_cd", 200, 1, 6, UCase(Request.Form("return_city")))
	End If
	adoCmd.Parameters.Append adoCmd.CreateParameter("@shop_car_type_cds", 200,  1, 1024, Request.Form("car_type_list"))
	
	Rem NOTE: IMPORTANT - if you add or modify these special cases
	Rem     you must update "search_profile_insert_car.asp" AND
	Rem                     "search_criteria_car.asp"
	
	strSelectDataSource = Request.Form("data_source")
	Select Case strSelectDataSource
	
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
	DiscountCodes = Request.Form("discount_code")
	
	Rem if the data source is not TRD and their are no discount codes then down't bother with them
	If Trim(Replace(DiscountCodes, ",", "")) = "" Then
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
				If Trim(DiscountCodesArray(intArrayHolder)) <> "" Then
					strExtraCriteria = strExtraCriteria & "DISC-" & DiscountCompanies(intArrayHolder) & "|" & Trim(DiscountCodesArray(intArrayHolder)) & "|" 
                End If					
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
	
	Dim intDateCount
	Dim datBegin
	Dim intIncriment
	
	If intRptLimit > 0 Then
	
		Rem Check the days and make sure the restricted users are restricted
		
		intDateCount = 0
		datBegin = Request.Form("begin_date")
		
		For intIncriment = 0 To DateDiff("d", datBegin, datEndDate)
			datEndDate = DateAdd("d", intIncriment, datBegin)
			If InStr(1, strDowList, Weekday(datEndDate)) > 0 Then
				intDateCount = intDateCount + 1
				If intDateCount = intRptLimit Then
					Exit For
				End If	
			
			'Else
			'	Response.Write(intDateCount & " of " & intRptLimit & " : " & datEndDate & " - not found <br>")
			End If
		
		Next
	
	'	If (DateDiff("d", Request.Form("begin_date"), datEndDate) >= intRptLimit) Then
	'		datEndDate = DateAdd("d", (intRptLimit - 1), Request.Form("begin_date"))
	'	End If
'		adoCmd.Parameters.Append adoCmd.CreateParameter("@end_arv_dt", 135, 1, 0, DateAdd("d", (intRptLimit - 1), Request.Form("begin_date")))
'	Else
'		adoCmd.Parameters.Append adoCmd.CreateParameter("@end_arv_dt", 135, 1, 0, Request.Form("end_date"))
	End If

	adoCmd.Parameters.Append adoCmd.CreateParameter("@end_arv_dt", 135, 1, 0, datEndDate)
	
	adoCmd.Parameters.Append adoCmd.CreateParameter("@dow_list", 200, 1, 13, strDowList)
	adoCmd.Parameters.Append adoCmd.CreateParameter("@scheduled_dttm", 135, 1, 0, Now)
	adoCmd.Parameters.Append adoCmd.CreateParameter("@start_dttm", 135, 1, 0, Null)
	adoCmd.Parameters.Append adoCmd.CreateParameter("@complete_dttm", 135, 1, 0, Null)
	adoCmd.Parameters.Append adoCmd.CreateParameter("@change_dttm", 135, 1, 0, Null)
	adoCmd.Parameters.Append adoCmd.CreateParameter("@request_status", 200, 1, 1, "N")
	adoCmd.Parameters.Append adoCmd.CreateParameter("@client_userid", 200, 1, 99, Request.Cookies("rate-monitor.com")("client_userid"))
	adoCmd.Parameters.Append adoCmd.CreateParameter("@client_system_cd", 200, 1, 5, "DFLT")
	adoCmd.Parameters.Append adoCmd.CreateParameter("@server_system_cd", 200, 1, 5, "DFLT")
	
	strArray = Split(Request.Form("selected_car_types"), ",", -1, 1)
	intArrayHolder = UBound(strArray)
	If intArrayHolder < 1 Then
		intArrayHolder = 1
	End If
	
	intWorkUnits = intArrayHolder  
	
	If (Request.Form("data_source") = "ALL") Then

		intArrayHolder = 0
	
		Rem Get the data sources since the user selected all
		Set adoCmd2 = CreateObject("ADODB.Command")

		adoCmd2.ActiveConnection =  strConn
		adoCmd2.CommandText = "car_all_data_source_select"
		adoCmd2.CommandType = 4
	
		Set adoRS2 = adoCmd2.Execute
	
		If adoRS2.EOF = False Then
			strDataSourceList = adoRS2.Fields("data_source_list").Value
		End If

		Set adoRS2 = Nothing
		Set	adoCmd2 = Nothing


	Else
		strDataSourceList = Request.Form("data_source")	
	
	End If

	
	strArray = Split(strDataSourceList, ",", -1, 1)
	intArrayHolder = UBound(strArray) 

	If intArrayHolder < 1 Then
		intArrayHolder = 1
	End If

	intWorkUnits = intWorkUnits * intArrayHolder  
	
	If (Left(Request.Form("cities_list"), 1) = "_") Then
	
		intArrayHolder = 0
	
		Rem Get the cities
		Set adoCmd2 = CreateObject("ADODB.Command")

		adoCmd2.ActiveConnection =  strConn
		adoCmd2.CommandText = "city_select"
		adoCmd2.CommandType = 4
	
		adoCmd2.Parameters.Append adoCmd2.CreateParameter("@city_grp_cd", 200, 1, 5, Request.Form("cities_list"))
		
		Set adoRS2 = adoCmd2.Execute
	
		While adoRS2.EOF = False
			intArrayHolder = intArrayHolder + 1
			adoRS2.MoveNext
		Wend
		Set adoRS2 = Nothing
		Set	adoCmd2 = Nothing
	
	Else

		strArray = Split(Request.Form("cities_list"), ",", -1, 1)
		intArrayHolder = UBound(strArray) 
		If intArrayHolder < 1 Then
			intArrayHolder = 1
		End If
	
	End If

	intWorkUnits = intWorkUnits * intArrayHolder  

	Dim CompanyList 
	
	CompanyList = Request.Form("company_list")
	
	Rem Remove the highlighted company from the middle or whereever and pre-pend it to the list
	Rem if it is in the middle of the list it will have a comma
	CompanyList = Replace(CompanyList, Request.Form("highlighted_company") & ",", "")
	Rem if it is at the end it wont
	CompanyList = Replace(CompanyList, Request.Form("highlighted_company"), "")
	CompanyList = Request.Form("highlighted_company") & "," & CompanyList
	
	If Right(CompanyList, 1) = "," Then
		CompanyList = Left(CompanyList, Len(CompanyList) - 1)
		
	End If

	strArray = Split(CompanyList, ",", -1, 1)
	intArrayHolder = UBound(strArray) 
	If intArrayHolder < 1 Then
		intArrayHolder = 1
	End If

	intWorkUnits = intWorkUnits * intArrayHolder  
	
	If intRptLimit > 0 Then
		intArrayHolder = DateDiff("d", Request.Form("begin_date"), DateAdd("d", intRptLimit, Request.Form("begin_date")))
	Else
		intArrayHolder = DateDiff("d", Request.Form("begin_date"), Request.Form("end_date"))
	End If
	
	intWorkUnits = intWorkUnits * intArrayHolder 
	
	adoCmd.Parameters.Append adoCmd.CreateParameter("@work_units", 3, 1, 0, intWorkUnits)
	adoCmd.Parameters.Append adoCmd.CreateParameter("@work_units_complete", 3, 1, 0, 0)
	adoCmd.Parameters.Append adoCmd.CreateParameter("@request_msg", 200, 1, 99, "")
	adoCmd.Parameters.Append adoCmd.CreateParameter("@vend_cd", 200, 1, 255, CompanyList)
	adoCmd.Parameters.Append adoCmd.CreateParameter("@highlight_vendor", 200, 1, 2, Request.Form("highlighted_company"))
	
	
	 Rem If Len(Request.Form(“profile_save_as”)) = 0 Then
	
	If Request.Form("profile_text") = "Default" Then
		adoCmd.Parameters.Append adoCmd.CreateParameter("@profile_desc", 200, 1, 255, "ad-hoc on demand report")
	
	Else
		adoCmd.Parameters.Append adoCmd.CreateParameter("@profile_desc", 200, 1, 255, "ad-hoc:" & Request.Form("profile_text"))
	
	End If
	
	adoCmd.Parameters.Append adoCmd.CreateParameter("@email_address", 200, 1, 255, Request.Form("recipient_address"))
	adoCmd.Parameters.Append adoCmd.CreateParameter("@action_id", 3, 1, 0, Request.Form("search_action"))
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
	adoCmd.Parameters.Append adoCmd.CreateParameter("@profile_id",        3, 1,   0, Null) ' This is a on-demand ad-hoc report
	adoCmd.Parameters.Append adoCmd.CreateParameter("@alert_address",   200, 1, 255, Request.Form("alert_address"))
	
	If Request.Form("oneway_reverse") = "True" Then
		adoCmd.Parameters.Append adoCmd.CreateParameter("@oneway_reverse",   11, 1, 0, 1)
	Else
		adoCmd.Parameters.Append adoCmd.CreateParameter("@oneway_reverse",   11, 1, 0, 0)
	End If 

	
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
		If (Request.Form("debug") = "true") Then
			Reponse.Write "debug mode<br>"
		Else
			adoCmd.Execute

			If err.number = 0 Then
				Server.Transfer "search_queue_car.asp"	
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
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>search_request_insert</title>
</head>
<body>
We are sorry, an error has occurred, please <a href="search_queue_car.asp">click 
here</a> to return to Search Criteria<p>Please print and fax this page to 
Rate-Highway Support at (888) 551-0029</p>
<p><font size="2">Debug information</font></p>
<br>
<font size="2">

			   <%	For Each Whatever In Request.Form
						Response.Write Whatever & " = <b>" & Request.Form(Whatever) & "</b> <br>"

       
					Next
				%>

Rpt Limit =<%=intRptLimit %><br>
datEndDate =<%=datEndDate %><br>
DateDiff("d", Request.Form("begin_date"), datEndDate) = <%=DateDiff("d", Request.Form("begin_date"), datEndDate) %>

</font>
<!--#INCLUDE FILE="footer.asp"-->
</body>
</html>

