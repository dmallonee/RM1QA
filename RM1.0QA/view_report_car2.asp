<%@ Language=VBScript %>
<!-- #INCLUDE FILE="inc/login_check.asp" -->
<!-- #INCLUDE FILE="inc/adovbs.asp" -->
<% Response.Expires = -1
   Response.cachecontrol="private" 
   Response.AddHeader "pragma", "no-cache"

	On Error Resume Next

   Server.ScriptTimeout = 360
 
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
	Dim varVendors()
	Dim varCities()
	Dim strSelectedVendor
	Dim strBgColor
	Dim blnDarkRow 
	Dim curRate
	Dim strCarList
	Dim strVendList
	Dim strCityList
	Dim bolRedoOff
	Dim strIPAddress
    Dim strDowString
    Dim strLOR

	strIPAddress = Request.Servervariables("REMOTE_ADDR") 

	If Request.QueryString("redo_off") = "true" Then
		bolRedoOff = True
	Else
		bolRedoOff = False
	End If

	strConn = Session("pro_con")
	
  	'Set adoCmd = CreateObject("ADODB.Command")
	
	'adoCmd.ActiveConnection = strConn 
	
	'adoCmd.CommandText = "car_shop_request_select_generic"
	'adoCmd.CommandType = 4
	
	'adoCmd.Parameters.Append adoCmd.CreateParameter("@shop_request_id", 3, 1, 0, Request("reportrequestid"))

	'Set adoRS1 = adoCmd.Execute

  	Set adoRS = CreateObject("ADODB.Recordset")
  	Set adoCmd = CreateObject("ADODB.Command")
	adoRS.CursorLocation = adUseClient
	
	adoCmd.ActiveConnection = strConn 

	adoCmd.CommandText = "car_shopped_rate_select_rpt2"
	adoCmd.CommandType = 4
	
	adoCmd.Parameters.Append adoCmd.CreateParameter("@shop_request_id", 3, 1, 0, Request("reportrequestid"))
	adoCmd.Parameters.Append adoCmd.CreateParameter("@report", 3, 1, 0, 1)
	
	If Len(Request.QueryString("vend_override")) = 2 Then
		adoCmd.Parameters.Append adoCmd.CreateParameter("@vend_override", 200, 1, 2, Request.QueryString("vend_override"))
	Else
		adoCmd.Parameters.Append adoCmd.CreateParameter("@vend_override", 200, 1, 2, Null)
	End If

	adoCmd.Parameters.Append adoCmd.CreateParameter("@ipaddress", 200, 1, 20, strIPAddress)
		
	adoCmd.Parameters.Append adoCmd.CreateParameter("@shop_car_type_cd", 200, 1, 4, "ALLL")	

    'adoRS.LockType = adLockOptimistic
    'adoRS.CursorLocation = adUseServer
    'adoRS.CursorType = adOpenForwardOnly
    'adoRS.Open "SET NOCOUNT ON"

	'adoRS.Open "car_shopped_rate_select_rpt 11, 1", strConn, adOpenStatic, adLockReadOnly  
	adoRS.Open adoCmd, , adOpenStatic, adLockReadOnly, adCmdStoredProc 

	'Set adoRS = adoCmd.Execute
	
	
	Dim intDateIndex
	Dim intCarTypeIndex
	Dim intDataSourceIndex
	Dim intCityIndex	

	While NOT adoRS.EOF 
		intDateIndex = adoRS.Fields("date_count").Value
		intDataSourceIndex = adoRS.Fields("data_source_count").Value
		intCarTypeIndex = adoRS.Fields("car_type_count").Value
		intVendorIndex = adoRS.Fields("vendor_count").Value
		intCityIndex = adoRS.Fields("city_count").Value
		strDowString = adoRS.Fields("dow_list").Value
		strLOR = adoRS.Fields("lor").Value

		adoRS.MoveNext

	Wend

	ReDim varCarTypes(intCarTypeIndex)
	ReDim varDataSources(intDataSourceIndex)
	ReDim varDates(intDateIndex)
	ReDim varVendors(intVendorIndex)
	ReDim varCities(intCityIndex)

	intDateIndex = 0
	intCarTypeIndex = 0
	intDataSourceIndex = 0
	intVendorIndex = 0
	intCityIndex = 0


	
	Set adoRS = adoRS.NextRecordset

	
	While adoRS.EOF = False
	
	   If intDateIndex < UBound(varDates) Then
		   If varDates(intDateIndex) <> adoRS.Fields("arv_dt").Value Then
			   intDateIndex= intDateIndex+ 1
	    	   If intDateIndex <= UBound(varDates) Then
					varDates(intDateIndex) = adoRS.Fields("arv_dt").Value
				End If
			End if
		End If
		
		If intCarTypeIndex < UBound(varCarTypes) Then
			If varCarTypes(intCarTypeIndex ) <> adoRS.Fields("shop_car_type_cd").Value Then
				intCarTypeIndex = intCarTypeIndex + 1
				If intCarTypeIndex <= UBound(varCarTypes) Then
					varCarTypes(intCarTypeIndex ) = adoRS.Fields("shop_car_type_cd").Value 
					strCarList = strCarList & adoRS.Fields("shop_car_type_cd").Value & ", " 
					'varHotelIds(intHotelIndex ) = adoRS.Fields("HotelID").Value 
				End If
			End If
		End If

		If intDataSourceIndex < UBound(varDataSources) Then
			If varDataSources(intDataSourceIndex) <> adoRS.Fields("data_source_name").Value Then
				intDataSourceIndex = intDataSourceIndex + 1
				If intDataSourceIndex <= UBound(varDataSources) Then
					varDataSources(intDataSourceIndex) = adoRS.Fields("data_source_name").Value
					strSourceList = strSourceList & adoRS.Fields("data_source_name").Value & ", " 
					'varSiteIds(intSiteIndex) = adoRS.Fields("SiteID").Value
				End If
			End If
		End If

		If intVendorIndex < UBound(varVendors) Then
			If varVendors(intVendorIndex) <> adoRS.Fields("vendor_name_rpt").Value Then
				intVendorIndex = intVendorIndex + 1
				If intVendorIndex <= UBound(varVendors) Then
					varVendors(intVendorIndex) = adoRS.Fields("vendor_name_rpt").Value
					strVendList = strVendList & adoRS.Fields("vendor_name_rpt").Value & ", " 
					'varSiteIds(intSiteIndex) = adoRS.Fields("SiteID").Value
				End If
			End If
		End If

		If intCityIndex < UBound(varCities) Then
			If varCities(intCityIndex ) <> adoRS.Fields("city_cd").Value Then
				intCityIndex = intCityIndex + 1
				If intCityIndex <= UBound(varVendors) Then
					varCities(intCityIndex) = adoRS.Fields("city_cd").Value
					strCityList = strCityList & adoRS.Fields("city_cd").Value & ", " 
					'varSiteIds(intSiteIndex) = adoRS.Fields("SiteID").Value
				End If
			End If
		End If

			
		adoRS.MoveNext
			
	Wend
	
	'adoRS.Close
	'Set adoRS = Nothing
	'Set adoCmd = Nothing
	adoRS.MoveFirst

	strCityCd = adoRS.Fields("city_cd").Value


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
<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Rate-Highway, Inc. | Rate-Monitor Report</title>
<link rel="stylesheet" type="text/css" href="inc/rh_standard.css">
<script language='Javascript'> 
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

</script> 

<style>
<!--
P {
	COLOR: navy; FONT-FAMILY: Verdana, Arial, sans-serif
}
.copyright {
	FONT-SIZE: 0.7em; TEXT-ALIGN: right
}

.report_detail_dark { 
	height="48" text-align:left; padding-left:3; padding-right:3; padding-top:3; background-color:#879AA2; font-family:Verdana, Arial, Helvetica, sans-serif; font-size:10pt; vertical-align:bottom; text-align:left; font-weight:bold  
}

.report_detail_light { 
	height="48" text-align:left; padding-left:3; padding-right:3; padding-top:3; background-color:#B2BEC4; font-family:Verdana, Arial, Helvetica, sans-serif; font-size:10pt; vertical-align:bottom; text-align:left
}

.report_rate_dark { 
	height="48" text-align:left; padding-left:3; padding-right:3; padding-top:3; background-color:#B2BEC4; font-family:Verdana, Arial, Helvetica, sans-serif; font-size:10pt; vertical-align:bottom; text-align:left 
	
}
-->
</style>
</head>

<body topmargin="0">

<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td background="http://www.rate-monitor.com/images/top_middle.jpg">
    <img src="images/top_left.jpg" width="423" height="91"></td>
    <td background="http://www.rate-monitor.com/images/top_middle.jpg"></td>
    <td background="http://www.rate-monitor.com/images/top_tile.gif"><img src="images/top_right.jpg" width="365" height="91"></td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td background="http://www.rate-monitor.com/images/med_bar_tile.gif"><img src="images/med_bar.gif"></td>
  </tr>
</table>
<p align="right"><font class="copyright">Copyright (c) 2001-2004, 
<a target="_blank" href="http://www.rate-highway.com">Rate-Highway, Inc.</a> (www.rate-highway.com) All Rights Reserved.<br>
Rate-Monitor is a product of Rate-Highway, Inc. - the leader in competitive 
market intelligence technology for the auto rental industry.</font></p>
<p>
<p><font size="-1"><br>
&nbsp;</font></p>


	
	<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#92393A" width="100%" id="AutoNumber2">
      <tr>
        <td width="100%" bgcolor="#92393A">&nbsp;<font color="#FFFFFF" size="4" face="Verdana">Rate-Monitor� 
        Report</font></td>
      </tr>
</table>


	
	<%



	While NOT adoRS.EOF 

		If (strCityCd <> adoRS.Fields("city_cd").Value) Or (strCarType <> adoRS.Fields("shop_car_type_cd").Value) Then
			strCityCd = adoRS.Fields("city_cd").Value
			strCarType = adoRS.Fields("shop_car_type_cd").Value
			%>
			<table cellSpacing='0' cellPadding='8' width='100%' border='0'>
			  <tr>
			    <td vAlign="top" bgColor="#CFD7DB" style="font-size: 0.8em; color: navy; font-family: Verdana, Arial, sans-serif">
			    <font size="+1"><b><%=strCityCd %>, <%=strCarType %></b></font>
			    <table cellSpacing="0" cellPadding="2" border="1">
		      <tr>
			<!--
		        <th align="right" bgColor="#B2BEC4" style="font-size: 0.8em; color: navy; font-family: Verdana, Arial, sans-serif">
		        &nbsp;</th>
			-->
		      	<%
		      	For intVendorIndex = LBound(varVendors) To UBound(varVendors) 
		      		blnMoved = False
		      	%>
		        <th align="middle" bgColor="#B2BEC4" style="font-size: 0.8em; color: navy; font-family: Verdana, Arial, sans-serif"><%=varVendors(intVendorIndex) %>&nbsp;</th>
		      	<%
		      	Next
		      	%>

				<% If bolRedoOff = False Then %>
		        <th align="middle" bgColor="#B2BEC4" style="font-size: 0.8em; color: navy; font-family: Verdana, Arial, sans-serif">
                Redo</th>
                <% End If %>
				
<!--		        <th noWrap align="middle" bgColor="#B2BEC4" style="font-size: 0.8em; color: navy; font-family: Verdana, Arial, sans-serif">
		        Details</th>
-->		        
		      </tr>
		      <%
		      
		      %>
		      <tr>
		      <% 
		      	blnDarkRow = True
		      	
				For intIndex = 1 To intDateIndex 'LBound(varDates) To UBound(varDates)	- 1	      
					blnMoved = False
				
					'Dim strSelectedVendor
					
					blnDarkRow = Not blnDarkRow
					
					If adoRS.EOF = True Then
						exit for
					end if
							
					curRate = adoRS.Fields("rt_amt").Value
				

		      %>
		      <th noWrap align="right" bgColor="#B2BEC4" style="font-size: 0.8em; color: navy; font-family: Verdana, Arial, sans-serif"><%=FormatDateTime(adoRS.Fields("arv_dt").Value ,1) %>&nbsp;</th>
			  
		      <% 
				For intVendorIndex = LBound(varVendors) To UBound(varVendors) '- 1	
				
					'If adoRS.Fields("data_source_name").Value = "Hertz" Then
					'	strBgColor = "#FFFFFF"
					'Else
						If blnDarkRow Then
							strBgColor = "#B2BEC4"
						Else
							strBgColor = "#CFD7DB"
						End If					
					'End If	   
					
					If adoRS.EOF = True Then
						Exit For
					End If   
					
					
					If intVendorIndex = UBound(varVendors) Then
					
					      %>
					      <% If bolRedoOff = False Then %>
		    			    <td noWrap align="center" bgColor="#B2BEC4" style="font-size: 0.8em; color: black; font-family: Verdana, Arial, sans-serif">
                            <a href="javascript:centerPopUp( 'row_rerun_car.asp?shop_request_id=<%=adoRS.Fields("shop_request_id").Value%>&city_cd=<%=adoRS.Fields("city_cd").Value%>&shop_car_type_cd=<%=adoRS.Fields("shop_car_type_cd").Value%>&data_source=<%=adoRS.Fields("data_source").Value%>&arv_dt=<%=Server.URLEncode(adoRS.Fields("arv_dt").Value)%>', 'rerequest', 400, 250, 0 );" >
                  
                  <!--          <a href='row_rerun_car.asp?shop_request_id=<%=adoRS.Fields("shop_request_id").Value%>&city_cd=<%=adoRS.Fields("city_cd").Value%>&shop_car_type_cd=<%=adoRS.Fields("shop_car_type_cd").Value%>&data_source=<%=adoRS.Fields("data_source").Value%>&arv_dt=<%=Server.URLEncode(adoRS.Fields("arv_dt").Value)%>' target="_blank">
                  -->
                            <img src="images/re_run.jpg" align="middle" width="18" height="18" border="0"></a></td>
                          <% End If %>
						  <%
					
		      
		      		ElseIf IsNumeric(adoRS.Fields("rt_amt").Value) Then
		      		
			      		If adoRS.Fields("rent_mi_alwnc_cd").Value = "Y" Then
			      			strMilage = "<sup>M </sup>"
			      		Else
			      			strMilage = ""
			      		End If
		      			
		      		

		      			If adoRS.Fields("rt_amt").Value = -1 Then
		      				Rem As in not searched - no rate should be displayed
					      %>
		    			    <td noWrap align="right" bgColor="<%=strBgColor  %>" style="font-size: 0.8em; color: black; font-family: Verdana, Arial, sans-serif">&nbsp;&nbsp;&nbsp;&nbsp;</td>
						  <%


		      			ElseIf curRate > adoRS.Fields("rt_amt").Value Then
					      %>
		    			    <td noWrap align="right" bgColor="<%=strBgColor  %>" style="font-size: 0.8em; color: red; font-family: Verdana, Arial, sans-serif"><%=strMilage %><%= FormatNumber(adoRS.Fields("rt_amt").Value, 2) %>&nbsp;</td>
						  <%
						  
						ElseIf curRate < adoRS.Fields("rt_amt").Value Then
						
					      %>
					        <td noWrap align="right" bgColor="<%=strBgColor  %>" style="font-size: 0.8em; color: green; font-family: Verdana, Arial, sans-serif"><%=strMilage %><%= FormatNumber(adoRS.Fields("rt_amt").Value, 2) %>&nbsp;</td>
						  <%
						  
						Else
					      %>
					        <td noWrap align="right" bgColor="<%=strBgColor  %>" style="font-size: 0.8em; color: black; font-family: Verdana, Arial, sans-serif"><%=strMilage %><%= FormatNumber(adoRS.Fields("rt_amt").Value, 2) %>&nbsp;</td>
						  <%
						
						
						End If						
						

					Else
					
					
	      				If adoRS.Fields("shop_msg").Value = "Rate was unavailable when website attempted to confirm with GDS" Then
		      				Rem As in not searched - no rate should be displayed
					      %>
		    			    <td noWrap align="right" bgColor="<%=strBgColor  %>" title="<%=adoRS.Fields("shop_msg").Value %>" style="font-size: 0.8em; color: black; font-family: Verdana, Arial, sans-serif">GDS Link</td>
						  <%

						Else

					      %>
					        <td noWrap align="right" bgColor="<%=strBgColor  %>" title="<%=adoRS.Fields("shop_msg").Value %>" style="font-size: 0.8em; color: navy; font-family: Verdana, Arial, sans-serif">
			                <sup> </sup>Closed</td>
						  <%

						End If

					End If


					Rem we need the details from the current date, prior to moving on to the next
					Rem date batch, so donn't advance after the last rate, but keep the ado row around
					Rem until the last coloumn can be created and then advance.
					If  intVendorIndex <> UBound(varVendors) - 1 Then
					
				  		If adoRS.EOF = False Then
						    adoRS.MoveNext
						    blnMoved = True
						End If
						
					End If
					
	

			    Next

		  		'If adoRS.EOF = False AND blnMoved = False Then
				'    adoRS.MoveNext
				'    blnMoved = True
				'End If

				blnMoved = False

			  %>
<!--			        
		        <td noWrap bgColor="<%=strBgColor%>" style="font-size: 0.8em; color: navy; font-family: Verdana, Arial, sans-serif">Lowest</td>
-->
		      </tr>
		      <%
		      	Next 
		      %>
		      </tr>
      		</table>
      	</table>
			<%
			
		End If
		
		'adoRS.MoveNext
		
	Wend
	
	adoRS.MoveFirst
'
'	'While (adoRS.EOF = False) And (adoRS.Fields("rt_amt").Value = -1)
'		adoRS.MoveNext
'
'	Wend


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

'SUB ErrorADOReport(parm_msg,parm_conn)
'   HowManyErrs=parm_conn.errors.count
'   IF HowManyErrs=0 then
'      exit sub
'   END IF
'   pad="&nbsp;&nbsp;&nbsp;&nbsp;"
'   response.write "<b>ADO Reports these Database Error(s) executing:<br>"
'          response.write SQLstmt & "</b><br>"
'   for counter= 0 to HowManyErrs-1
'      errornum=parm_conn.errors(counter).number
'      errordesc=parm_conn.errors(counter).description
'      response.write pad & "Error#=<b>" & errornum & "</b><br>"
'      response.write pad & "Error description=<b>"
'      response.write errordesc & "</b><p>"
'   next
'END SUB

	 	
 %>


<table cellSpacing="0" cellPadding="8" width="100%" border="0">
  <tr>
    <td vAlign="top" bgColor="#CFD7DB" style="font-size: 0.8em; color: navy; font-family: Verdana, Arial, sans-serif">
    &nbsp;</td>
  </tr>
</table>
<p>
<br>


</p>


<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" id="AutoNumber1" width="642">
  <tr>
    <td width="164" bgcolor="#92393A" bordercolor="#92393A">
    <font face="Verdana" color="#EAECF9" size="4">&nbsp;Report Legend</font></td>
    <td width="478" bordercolor="#92393A" bgcolor="#92393A">&nbsp;</td>
  </tr>
  <tr>
    <td width="164" class="report_detail_dark">Locations:</td>
    <td width="478" class="report_detail_light"><%=Left(strCityList, Len(strCityList) - 2) %></td>
  </tr>
  <tr>
    <td width="164" class="report_detail_dark">Companies: </td>
    <td width="478" class="report_detail_light"><%=Left(strVendList, Len(strVendList) - 2)%></td>
  </tr>
  <tr>
    <td width="164" class="report_detail_dark">Data Source:</td>
    <td width="478" class="report_detail_light"><%=Left(strSourceList, Len(strSourceList) - 2)%></td>
  </tr>
  <tr>
    <td width="164" class="report_detail_dark" valign="top">Dates:</td>
    <td width="478" class="report_detail_light"><%=FormatDateTime(varDates(LBound(varDates)+ 1),1) %> to <%=FormatDateTime(varDates(UBound(varDates)),1) %>, 
    <% Dim strDow
       Dim intDowIndex
       strDow = Split(strDowString, ",")
	   strDowString = ""
       For intDowIndex = LBound(strDow) To UBound(strDow)
       	 Select Case strDow(intDowIndex)
       	 
       	 	Case 1
	        	strDowString = strDowString & "Sun, "
       	 	Case 2
	        	strDowString = strDowString & "Mon, "
       	 	Case 3
	        	strDowString = strDowString & "Tue, "
       	 	Case 4
	        	strDowString = strDowString & "Wed, "
       	 	Case 5
	        	strDowString = strDowString & "Thu, "
       	 	Case 6
	        	strDowString = strDowString & "Fri, "
       	 	Case 7
	        	strDowString = strDowString & "Sat, "
	        	
	     End Select
	        	
	        	
       Next
    %><%=Left(strDowString, Len(strDowString) - 2)%></td>
  </tr>
  <tr>
    <td width="164" class="report_detail_dark">Car Types: </td>
    <td width="478" class="report_detail_light"><%=Left(strCarList, Len(strCarList) - 2)%></td>
  </tr>
  <tr>
    <td width="164" class="report_detail_dark">Pickup/Drop-off:</td>
    <td width="478" class="report_detail_light">
    <% If strLOR = 1000  Then %>
    	<%="Daily Rates"          %>    
    <% ElseIf strLOR = 1001  Then %>
    	<%="Weekend Daily Rates"         %>    
    <% ElseIf strLOR = 1002  Then %>
    	<%="Weekly Rates"         %>    
    <% Else                   %>
    	<%="LOR" & strLOR         %>
    <% End If                 %>
    ; Pickup <%=FormatDateTime(adoRS.Fields("arv_tm").Value, 3) %>; Drop-off <%=FormatDateTime(adoRS.Fields("rtrn_tm").Value, 3) %></font></td>
  </tr>
  <tr>
    <td width="164" class="report_detail_dark" valign="top">Options:</td>
    <td width="478" class="report_detail_light">Rates Shown: Rate Amount - not Estimated Total 
    Charge; <br> Highlighted Vendor: <%=varVendors(1) %> </td>
  </tr>
  <tr>
    <td width="164" class="report_detail_dark">Search Web Details:</td>
    <td width="478" class="report_detail_light">POS: <%=adoRS.Fields("shop_pos_cd").Value %>; Air Ticket Required: <%=adoRS.Fields("airline_arv_ind").Value %></font> 
    Discount Code Used: [None]</td>
  </tr>
  </table>
<font color="#800000" size="2">
<p><font face="Tahoma">The highlighted vendor is on the far left of the report 
and their rates are always in black. All of the other vendor�s rates will either 
be displayed in red, green, or black. Occasionally you will see other messages 
that appear on reports. The following is a list of each of their meanings:</font></p>
<p><font face="Tahoma"><font color="#FF0000"><b>Red Rates </b></font>� Rates are 
less than the highlighted vendor's</font></p>
<p><font face="Tahoma"><font color="#008000"><b>Green Rates </b></font>� Rates 
are more than the highlighted vendor's</font></p>
</font>
<font size="2">
<b>Black Rates<font color="#800000" size="2">
</font> </b>
</font>
<font color="#800000" size="2">
� </font>
<font size="2">
Rates are the 
  same as the highlighted vendor's
<font color="#800000" size="2">
<p><font face="Tahoma"><b>Closed</b> � The car type is not available on a certain 
day for a certain website (i.e. it is sold out).</font></p>
</font></font>
<font color="#800000" size="2">
<p>&nbsp;</p>
<p>Default New User Report - For custom versions please contact your 
Rate-Highway representative at (877) RATE-HWY</p>

<% Set adoRS = Nothing %>
<% Set adoRS1 = Nothing %>
<!--#INCLUDE FILE="footer.asp"-->
</body>
</html>