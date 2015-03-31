<%@ Language=VBScript %>
<!-- #INCLUDE FILE="inc/adovbs.asp" --> 

<% Response.Expires = -1
   Response.cachecontrol="private" 
   Response.AddHeader "pragma", "no-cache"

   Server.ScriptTimeout = 0
 
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
	Dim Whatever
	Dim strDowList
	Dim strDataSourceList

	strUserId = Request.Cookies("rate-monitor.com")("user_id")
	
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
	adoCmd.Parameters.Append adoCmd.CreateParameter("@rtrn_city_cd", 200, 1, 5, Null) ' Request.Form("return_city"))
	adoCmd.Parameters.Append adoCmd.CreateParameter("@shop_car_type_cds", 200,  1, 1024, Request.Form("car_type_list"))
	adoCmd.Parameters.Append adoCmd.CreateParameter("@data_sources", 200, 1, 1024, Request.Form("data_source"))
	adoCmd.Parameters.Append adoCmd.CreateParameter("@shop_currency_cd", 200, 1, 3, Request.Form("display_currency"))
	adoCmd.Parameters.Append adoCmd.CreateParameter("@shop_location_categ", 200, 1, 9, "T,O")
	adoCmd.Parameters.Append adoCmd.CreateParameter("@shop_rt_categ", 200, 1, 9, "S,P")
	adoCmd.Parameters.Append adoCmd.CreateParameter("@shop_rt_type", 200, 1, 9, "D,E")
	adoCmd.Parameters.Append adoCmd.CreateParameter("@airline_arv_ind", 200, 1, 1, "N")
	adoCmd.Parameters.Append adoCmd.CreateParameter("@shop_pos_cd", 200, 1, 5, Request.Form("pos"))
	adoCmd.Parameters.Append adoCmd.CreateParameter("@extra_criteria", 200, 1, 99, "")
	adoCmd.Parameters.Append adoCmd.CreateParameter("@lor", 3, 1, 0, Request.Form("lor"))
	adoCmd.Parameters.Append adoCmd.CreateParameter("@arv_tm", 200, 1, 5, Request.Form("arrival_time"))
	adoCmd.Parameters.Append adoCmd.CreateParameter("@rtrn_tm", 200, 1, 5, Request.Form("return_time"))
	adoCmd.Parameters.Append adoCmd.CreateParameter("@begin_arv_dt", 135, 1, 0, Request.Form("begin_date"))
	adoCmd.Parameters.Append adoCmd.CreateParameter("@end_arv_dt", 135, 1, 0, Request.Form("end_date"))
	adoCmd.Parameters.Append adoCmd.CreateParameter("@dow_list", 200, 1, 13, strDowList)
	adoCmd.Parameters.Append adoCmd.CreateParameter("@scheduled_dttm", 135, 1, 0, Now)
	adoCmd.Parameters.Append adoCmd.CreateParameter("@start_dttm", 135, 1, 0, Null)
	adoCmd.Parameters.Append adoCmd.CreateParameter("@complete_dttm", 135, 1, 0, Null)
	adoCmd.Parameters.Append adoCmd.CreateParameter("@change_dttm", 135, 1, 0, Null)
	adoCmd.Parameters.Append adoCmd.CreateParameter("@request_status", 200, 1, 1, "N")
	adoCmd.Parameters.Append adoCmd.CreateParameter("@client_userid", 200, 1, 99, Request.Cookies("rate-monitor.com")("client_userid"))
	adoCmd.Parameters.Append adoCmd.CreateParameter("@client_system_cd", 200, 1, 5, "DFLT")
	adoCmd.Parameters.Append adoCmd.CreateParameter("@server_system_cd", 200, 1, 5, "DFLT")
	
	Dim intWorkUnits
	Dim strArray
	Dim intArrayHolder
	
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

	strArray = Split(Request.Form("company_list"), ",", -1, 1)
	intArrayHolder = UBound(strArray) 
	If intArrayHolder < 1 Then
		intArrayHolder = 1
	End If

	intWorkUnits = intWorkUnits * intArrayHolder  
	intArrayHolder = DateDiff("d", Request.Form("begin_date"), Request.Form("end_date"))
	
	intWorkUnits = intWorkUnits * intArrayHolder 
	
	adoCmd.Parameters.Append adoCmd.CreateParameter("@work_units", 3, 1, 0, intWorkUnits)
	adoCmd.Parameters.Append adoCmd.CreateParameter("@work_units_complete", 3, 1, 0, 0)
	adoCmd.Parameters.Append adoCmd.CreateParameter("@request_msg", 200, 1, 99, "")
	adoCmd.Parameters.Append adoCmd.CreateParameter("@vend_cd", 200, 1, 255, Request.Form("company_list"))
	adoCmd.Parameters.Append adoCmd.CreateParameter("@highlight_vendor", 200, 1, 2, Request.Form("highlighted_company"))
	adoCmd.Parameters.Append adoCmd.CreateParameter("@profile_desc", 200, 1, 255, "")
	adoCmd.Parameters.Append adoCmd.CreateParameter("@email_address", 200, 1, 255, Request.Form("recipient_address"))
	adoCmd.Parameters.Append adoCmd.CreateParameter("@action_id", 3, 1, 0, Request.Form("search_action"))
	adoCmd.Parameters.Append adoCmd.CreateParameter("@output", 200, 1, 10, Request.Form("output"))
	Select Case Request.Form("output")
	 	 Case "html"
			adoCmd.Parameters.Append adoCmd.CreateParameter("@output_style", 3, 1, 0, 1)
			
		case "ftp"
			adoCmd.Parameters.Append adoCmd.CreateParameter("@output_style", 3, 1, 0, Request.Form("ftp_style"))

		Case Else
			adoCmd.Parameters.Append adoCmd.CreateParameter("@output_style", 3, 1, 0, 1)
		
	End Select		


	
	If Request.Form("debug") = "true" Then
	
	Else
		adoCmd.Execute
	
		If err.number = 0 Then
			Server.Transfer "search_queue_car.asp"	
		Else
			Server.Transfer "search_criteria_car.asp"
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

Request.Form("dow_list") = <%= Len(Request.Form("dow_list")) %><br>
strDowList = <%= Len(strDowList) %><br>

			   <%	For Each Whatever In Request.Form
						Response.Write Whatever & " = <b>" & Request.Form(Whatever) & "</b> <br>"

       
					Next
				%>
<!--#INCLUDE FILE="footer.asp"-->
</body>
</html>
