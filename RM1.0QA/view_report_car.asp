<%@ Language=VBScript %>
<!-- #INCLUDE FILE="inc/login_check.asp" --> 
<!-- #INCLUDE FILE="inc/adovbs.asp" --> 


<% Response.Expires = -1
   Response.cachecontrol="private" 
   Response.AddHeader "pragma", "no-cache"

   Server.ScriptTimeout = 0
 
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
	Dim strSelectedVendor
	Dim strBgColor
	Dim blnDarkRow 
	Dim curRate
	Dim strCarList
	Dim strVendList


	
	'intReportRequestID = Request("ReportRequestID")
	
	'strConn = Session("pro_con")
	
  	'Set adoCmd = CreateObject("ADODB.Command")

	'adoCmd.ActiveConnection = strConn
	'adoCmd.CommandText = "symSearchResultSelect30dayByHotel"
	'adoCmd.CommandType = 4

	'adoCmd.Parameters.Refresh 

	'adoCmd.Parameters("@ReportRequestID").Value = intReportRequestID 

	intReportRequestID = Request("ReportRequestID")
	
	strConn = Session("pro_con")
	
  	Set adoRS = CreateObject("ADODB.Recordset")
	adoRS.CursorLocation = adUseClient
			
	adoRS.Open "car_shopped_rate_select " & intReportRequestID & " , 1", strConn, adOpenStatic, adLockReadOnly  

	Dim intDateIndex
	Dim intCarTypeIndex
	Dim intDataSourceIndex

	While NOT adoRS.EOF 
		intDateIndex = adoRS.Fields("date_count").Value
		intDataSourceIndex = adoRS.Fields("data_source_count").Value
		intCarTypeIndex = adoRS.Fields("car_type_count").Value

		adoRS.MoveNext

	Wend

	ReDim varCarTypes(intCarTypeIndex)
	ReDim varDataSources(intDataSourceIndex)
	ReDim varDates(intDateIndex)


	intDateIndex = 0
	intCarTypeIndex = 0
	intDataSourceIndex = 0


	
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
					varDataSources(intDataSourceIndex ) = adoRS.Fields("data_source_name").Value
					strVendList = strVendList & adoRS.Fields("data_source_name").Value & ", " 
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

	%>
	
    
<html>


<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Rate-Highway, Inc. C.A.R.S. | View report</title>
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
    <img src="images/top_left.jpg" width="423" height="91"></td>
    <td background="images/top_middle.jpg"></td>
    <td background="images/top_tile.gif"><img src="images/top_right.jpg" width="365" height="91"></td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td background="images/med_bar_tile.gif"><img src="images/med_bar.gif"></td>
  </tr>
</table>
<p align="right"><font class="copyright">Copyright (c) 2001-2004, Rate-Highway, Inc. All Rights Reserved.</font></p>
<p>
<p><font size="-1"><br>
&nbsp;</font></p>


	
	<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#92393A" width="100%" id="AutoNumber2">
      <tr>
        <td width="100%" bgcolor="#92393A">&nbsp;<font color="#FFFFFF" size="4" face="Verdana">Rate Detail</font></td>
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
			    <font size="+1"><b><%=strCityCd %>, <%=strCarType %></b></font>
			    <table cellSpacing="0" cellPadding="2" border="1">
		      <tr>
			<!--
		        <th align="right" bgColor="#B2BEC4" style="font-size: 0.8em; color: navy; font-family: Verdana, Arial, sans-serif">
		        &nbsp;</th>
			-->
		      	<%
		      	For intDataSourceIndex = LBound(varDataSources)  To UBound(varDataSources) 
		      	%>
		        <th align="middle" bgColor="#B2BEC4" style="font-size: 0.8em; color: navy; font-family: Verdana, Arial, sans-serif"><%=varDataSources(intDataSourceIndex) %>&nbsp;</th>
		      	<%
		      	Next
		      	%>
				
		        <th noWrap align="middle" bgColor="#B2BEC4" style="font-size: 0.8em; color: navy; font-family: Verdana, Arial, sans-serif">
		        Details</th>
		      </tr>
		      <%
		      
		      %>
		      <tr>
		      <% 
		      	blnDarkRow = True
		      	
				For intIndex = LBound(varDates) To UBound(varDates)	- 1	      
				
					'Dim strSelectedVendor
					
					blnDarkRow = Not blnDarkRow
					curRate = adoRS.Fields("rt_amt").Value
				

		      %>
		      <th noWrap align="right" bgColor="#B2BEC4" style="font-size: 0.8em; color: navy; font-family: Verdana, Arial, sans-serif"><%=FormatDateTime(adoRS.Fields("arv_dt").Value ,1) %>&nbsp;</th>
			  
		      <% 
				For intDataSourceIndex = LBound(varDataSources) To UBound(varDataSources) - 1	
				
					'If adoRS.Fields("data_source_name").Value = "Hertz" Then
					'	strBgColor = "#FFFFFF"
					'Else
						If blnDarkRow Then
							strBgColor = "#B2BEC4"
						Else
							strBgColor = "#CFD7DB"
						End If					
					'End If	      
		      
		      		If IsNumeric(adoRS.Fields("rt_amt").Value) Then
		      		
			      		If adoRS.Fields("rent_mi_alwnc_cd").Value = "Y" Then
			      			strMilage = "<sup>M </sup>"
			      		Else
			      			strMilage = ""
			      		End If
		      			
		      		
		      		
		      			If curRate > adoRS.Fields("rt_amt").Value Then
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
		      %>
		        <td noWrap align="right" bgColor="<%=strBgColor  %>" style="font-size: 0.8em; color: navy; font-family: Verdana, Arial, sans-serif">
                <sup>C </sup>Closed</td>
			  <%
					End If
					
			  		If adoRS.EOF = False Then
					    adoRS.MoveNext
					End If

			    Next
			  %>
			        
		        <td noWrap bgColor="<%=strBgColor%>" style="font-size: 0.8em; color: navy; font-family: Verdana, Arial, sans-serif">Lowest</td>
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
    <td width="478" class="report_detail_light"><%=strCityCd %>; All</td>
  </tr>
  <tr>
    <td width="164" class="report_detail_dark">Companies: </td>
    <td width="478" class="report_detail_light"><%=strVendList%></td>
  </tr>
  <tr>
    <td width="164" class="report_detail_dark">Data Source:</td>
    <td width="478" class="report_detail_light">Car Company Web Rates</font></td>
  </tr>
  <tr>
    <td width="164" class="report_detail_dark">Dates:</td>
    <td width="478" class="report_detail_light"><%=FormatDateTime(varDates(LBound(varDates)+ 1),1) %> to <%=FormatDateTime(varDates(UBound(varDates)),1) %>, 
    Mon Tue Wed Thur Fri Sat Sun</td>
  </tr>
  <tr>
    <td width="164" class="report_detail_dark">Car Types: </td>
    <td width="478" class="report_detail_light"><%=strCarList%></font></td>
  </tr>
  <tr>
    <td width="164" class="report_detail_dark">Pickup/Drop-off:</td>
    <td width="478" class="report_detail_light">LOR <%=adoRS.Fields("lor").Value %> 
    ; Pickup <%=adoRS.Fields("arv_tm").Value %>; Drop-off <%=adoRS.Fields("rtrn_tm").Value %></font></td>
  </tr>
  <tr>
    <td width="164" class="report_detail_dark">Options:</td>
    <td width="478" class="report_detail_light">Rates Shown: Estimated Total 
    Charge; Highlighted Vendor: Alamo; </td>
  </tr>
  <tr>
    <td width="164" class="report_detail_dark">Search Web Details:</td>
    <td width="478" class="report_detail_light">POS: <%=adoRS.Fields("shop_pos_cd").Value %>; Air Ticket Required: <%=adoRS.Fields("airline_arv_ind").Value %></font> 
    Discount Code Used: [None]</td>
  </tr>
  <tr>
    <td width="164" class="report_detail_dark">User/Profile:</td>
    <td width="478" class="report_detail_light">mmeyer / default [unlimited credit]</font></td>
  </tr>
</table>
<p>&nbsp;</p>

<% Set adoRS = Nothing %>
<!--#INCLUDE FILE="footer.asp"-->
</body>

</html>