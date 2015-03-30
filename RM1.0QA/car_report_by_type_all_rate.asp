<%@ Language=VBScript %>
<!-- #INCLUDE FILE="inc/adovbs.asp" -->
<% Response.Expires = -1
   Response.cachecontrol="private" 
   Response.AddHeader "pragma", "no-cache"

	on error resume next

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
	Dim strSelectedVendor
	Dim strBgColor
	Dim blnDarkRow 
	Dim curRate
	Dim strCarList
	Dim strVendList
	Dim strCarCodeListArray 
	Dim strCarCodeList
	Dim strDowString
	Dim blnRedoEnabled
	Dim strIPAddress
	Dim strLOR	
	
	strIPAddress = Request.Servervariables("REMOTE_ADDR") 

	If UCASE(Request("redoenabled")) = "TRUE" Then
		blnRedoEnabled = True
	Else
		blnRedoEnabled = False
	End If


	'strConn = "Provider=SQLOLEDB.1; Network Library=dbmssocn;Password=symAgent;User ID=symAgent;Initial Catalog=production;Data Source=65.161.185.103;" 
	'
	strConn = Session("pro_con")
	
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
		
	If Request("car_type_cd") = "" Then
		adoCmd.Parameters.Append adoCmd.CreateParameter("@shop_car_type_cd", 200, 1, 4, Null)	
	Else
		adoCmd.Parameters.Append adoCmd.CreateParameter("@shop_car_type_cd", 200, 1, 4, Request("car_type_cd"))
	End If
	



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
	

	While NOT adoRS.EOF 
		intDateIndex = adoRS.Fields("date_count").Value
		intDataSourceIndex = adoRS.Fields("data_source_count").Value
		intCarTypeIndex = adoRS.Fields("car_type_count").Value
		intVendorIndex = adoRS.Fields("vendor_count").Value
		strCarCodeList = adoRS.Fields("car_type_list").Value
		strDowString = adoRS.Fields("dow_list").Value	
		strLOR = adoRS.Fields("lor").Value	
				
		adoRS.MoveNext

	Wend

	strCarCodeListArray = Split(strCarCodeList, ",")


	ReDim varCarTypes(intCarTypeIndex)
	ReDim varDataSources(intDataSourceIndex)
	ReDim varDates(intDateIndex)
	ReDim varVendors(intVendorIndex)

	intDateIndex = 0
	intCarTypeIndex = 0
	intDataSourceIndex = 0
	intVendorIndex = 0


	
	Set adoRS= adoRS.NextRecordset

	
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


			
		adoRS.MoveNext
			
	Wend
	
	'adoRS.Close
	'Set adoRS = Nothing
	'Set adoCmd = Nothing
	adoRS.MoveFirst
	
	If adoRS.Fields("city_cd").Value <> adoRS.Fields("rtrn_city_cd").Value Then
		strCityCd = adoRS.Fields("city_cd").Value & " to " & adoRS.Fields("rtrn_city_cd").Value
	Else
		strCityCd = adoRS.Fields("city_cd").Value 
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
<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Rate-Monitor.com | View Report By Car Type</title>
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
    <td background="images/top_middle.jpg">
    <img src="images/top_left.JPG" width="424" height="91"></td>
    <td background="images/top_middle.jpg"></td>
    <td background="images/top_tile.gif"><img src="images/top_right.jpg" width="365" height="91"></td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td background="images/med_bar_tile.gif"><img src="images/med_bar.gif"></td>
  </tr>
</table>
<p align="right"><font class="copyright">Copyright (c) 2001-2004,
<a target="_blank" href="http://www.rate-highway.com">Rate-Highway, Inc.</a> (www.rate-highway.com) 
All Rights Reserved.<br>
Rate-Monitor is a product of Rate-Highway, Inc. - the leader in competitive 
market intelligence technology for the auto rental industry.</font></p>
<p>
<p><font size="-1"><br>
&nbsp;</font></p>


	
	<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#92393A" width="100%" id="AutoNumber2">
      <tr>
        <td width="100%" bgcolor="#92393A">&nbsp;<font color="#FFFFFF" size="4" face="Verdana">Expanded Rate Detail </font></td>
      </tr>
</table>


	
	<%



	While NOT adoRS.EOF 

		If strCarType <> adoRS.Fields("shop_car_type_cd").Value Then
			strCarType = adoRS.Fields("shop_car_type_cd").Value
			%>
			<table cellSpacing='0' cellPadding='8' width='100%' border='0'>
			  <tr>
			    <td vAlign="top" bgColor="#CFD7DB" style="font-size: 0.8em; color: navy; font-family: Verdana, Arial, sans-serif">
			    <form method="GET" name="rate_detail">
			      <% If Request("car_type_cd") = "ALLL" Then  %>
                  <p><font size="+1"><b><%=strCityCd %>, <%=strCarType %></b></font>
			      <% Else %>
                  <br><font size="+1"><b><%=strCityCd %>,</b></font>
                  <select size="1" name="car_type_cd">
				  <% For intIndex = LBound(strCarCodeListArray) To UBound(strCarCodeListArray)	%>
				  <% If Request("car_type_cd") = strCarCodeListArray(intIndex) Then             %>
				   				   <option selected><%=strCarCodeListArray(intIndex) %> </option>
				  <% Else             %>
				   				   <option ><%=strCarCodeListArray(intIndex) %> </option>
				  <% End If           %>
 
				  <% Next %>                  
                  
                  
                  </select>
                  <input type="submit" value="Display" name="display"></p>
                  <input type="hidden" name="reportrequestid" value="<%=Request("reportrequestid") %>">
                  <input type="hidden" name="redoenabled" value="<%=Request("redoenabled") %>">
                </form>
                <% End If %>
			    <table cellSpacing="0" cellPadding="2" border="1">
		      <tr>
			<!--
		        <th align="right" bgColor="#B2BEC4" style="font-size: 0.8em; color: navy; font-family: Verdana, Arial, sans-serif">
		        &nbsp;</th>
			-->
		      	<%
		      	For intVendorIndex = LBound(varVendors)  To UBound(varVendors) 
		      		blnMoved = False
		      	%>
		        <th align="middle" bgColor="#B2BEC4" style="font-size: 0.8em; color: navy; font-family: Verdana, Arial, sans-serif"><%=varVendors(intVendorIndex) %>&nbsp;</th>
		      	<%
		      	Next
		      	%>

				<% If blnRedoEnabled Then %> 
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
					
					
					If (intVendorIndex = UBound(varVendors)) Then
		
						If blnRedoEnabled Then
					      %>
		    			    <td noWrap align="center" bgColor="#B2BEC4" style="font-size: 0.8em; color: black; font-family: Verdana, Arial, sans-serif">
                  <!--          <a href="javascript:centerPopUp( 'row_rerun_car.asp?shop_request_id=<%=adoRS.Fields("shop_request_id").Value%>&city_cd=<%=adoRS.Fields("city_cd").Value%>&shop_car_type_cd=<%=adoRS.Fields("shop_car_type_cd").Value%>&data_source=<%=adoRS.Fields("data_source").Value%>&arv_dt=<%=Server.URLEncode(adoRS.Fields("arv_dt").Value)%>', 're-request', 400, 250, 1 )">
                  -->
                            <a href="javascript:centerPopUp( 'row_rerun_car.asp?shop_request_id=<%=adoRS.Fields("shop_request_id").Value%>&city_cd=<%=adoRS.Fields("city_cd").Value%>&shop_car_type_cd=<%=adoRS.Fields("shop_car_type_cd").Value%>&data_source=<%=adoRS.Fields("data_source").Value%>&arv_dt=<%=Server.URLEncode(adoRS.Fields("arv_dt").Value)%>', 'rerequest', 400, 250, 0 );" >
                            <img src="images/re_run.jpg" align="middle" width="18" height="18" border="0"></a></td>
						  <%
						End If
		      
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
		    			    <td noWrap align="right" bgColor="<%=strBgColor  %>" style="font-size: 0.8em; color: red; font-family: Verdana, Arial, sans-serif"><%=strMilage %><%= FormatNumber(adoRS.Fields("rt_amt").Value, 2) & " / " & FormatNumber(adoRS.Fields("total_rt_amt").Value, 2) & " / " & FormatNumber(adoRS.Fields("est_rental_chrg_amt").Value, 2)  %>&nbsp;</td>
						  <%
						  
						ElseIf curRate < adoRS.Fields("rt_amt").Value Then
						
					      %>
					        <td noWrap align="right" bgColor="<%=strBgColor  %>" style="font-size: 0.8em; color: green; font-family: Verdana, Arial, sans-serif"><%=strMilage %><%= FormatNumber(adoRS.Fields("rt_amt").Value, 2) & " / " & FormatNumber(adoRS.Fields("total_rt_amt").Value, 2) & " / " & FormatNumber(adoRS.Fields("est_rental_chrg_amt").Value, 2)  %>&nbsp;</td>
						  <%
						  
						Else
					      %>
					        <td noWrap align="right" bgColor="<%=strBgColor  %>" style="font-size: 0.8em; color: black; font-family: Verdana, Arial, sans-serif"><%=strMilage %><%= FormatNumber(adoRS.Fields("rt_amt").Value, 2) & " / " & FormatNumber(adoRS.Fields("total_rt_amt").Value, 2) & " / " & FormatNumber(adoRS.Fields("est_rental_chrg_amt").Value, 2)  %>&nbsp;</td>
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
    <td width="164" bgcolor="#92393A" bordercolor="#92393A" valign="top">
    <font face="Verdana" color="#EAECF9" size="4">&nbsp;Report Legend</font></td>
    <td width="478" bordercolor="#92393A" bgcolor="#92393A" valign="top">&nbsp;</td>
  </tr>
  <tr>
    <td width="164" class="report_detail_dark" valign="top">Locations:</td>
    <td width="478" class="report_detail_light" valign="top"><%=strCityCd %></td>
  </tr>
  <tr>
    <td width="164" class="report_detail_dark" valign="top">Companies: </td>
    <td width="478" class="report_detail_light" valign="top"><%=Left(strVendList, Len(strVendList) - 2)%></td>
  </tr>
  <tr>
    <td width="164" class="report_detail_dark" valign="top">Data Source:</td>
    <td width="478" class="report_detail_light" valign="top"><%=Left(strSourceList, Len(strSourceList) - 2)%></td>
  </tr>
  <tr>
    <td width="164" class="report_detail_dark" valign="top">Dates:</td>
    <td width="478" class="report_detail_light" valign="top"><%=FormatDateTime(varDates(LBound(varDates)+ 1),1) %> to <%=FormatDateTime(varDates(UBound(varDates)),1) %>, 
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
    <td width="164" class="report_detail_dark" valign="top">Car Types: </td>
    <td width="478" class="report_detail_light" valign="top"><%=Left(strCarList, Len(strCarList) - 2)%>
    </td>
  </tr>
  <tr>
    <td width="164" class="report_detail_dark" valign="top">Pickup/Drop-off:</td>
    <td width="478" class="report_detail_light">
    <% If strLOR = 1000  Then %>
    	<%="Daily Rates"          %>    
    <% ElseIf strLOR = 1001  Then %>
    	<%="Weekend Daily Rates"         %>    
    <% ElseIf strLOR = 1002  Then %>
    	<%="Weekly Rates"         %>    
    <% Else                   %>
    	<%="LOR " & strLOR         %>
    <% End If                 %>; Pickup <%=FormatDateTime(adoRS.Fields("arv_tm").Value, 3) %>; Drop-off <%=FormatDateTime(adoRS.Fields("rtrn_tm").Value, 3) %></td>
  </tr>
  <tr>
    <td width="164" class="report_detail_dark" valign="top">Options:</td>
    <td width="478" class="report_detail_light" valign="top">Rates Shown: Rate Amount, 
    Total Rate Amount, Estimated Total 
    Charge; <br> Highlighted Vendor: <%=varVendors(1) %> </td>
  </tr>
  <tr>
    <td width="164" class="report_detail_dark" valign="top">Search Web Details:</td>
    <td width="478" class="report_detail_light" valign="top">POS: <%=adoRS.Fields("shop_pos_cd").Value %>; Air Ticket Required: <%=adoRS.Fields("airline_arv_ind").Value %></font> 
    Discount Code Used: [None]</td>
  </tr>
  </table>
<p>Default New User Report - For custom versions please contact your 
Rate-Highway representative at (877) RATE-HWY</p>

<% Set adoRS = Nothing %>
<!--#INCLUDE FILE="footer.asp"-->
</body>
</html>