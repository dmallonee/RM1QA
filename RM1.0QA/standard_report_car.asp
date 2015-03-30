<%@ Language=VBScript %>
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

	'If Request.QueryString("redo_off") = "true" Then
		bolRedoOff = True
	'Else
	'	bolRedoOff = False
	'End If

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
<title>Rate-Highway, Inc. | Rate-Monitor Standard Report</title>
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
 <table width="1108" border="0" cellpadding="2" cellspacing="0" background="images/alt_color.gif" style="border-collapse: collapse" bordercolor="#111111">
    <tr valign="bottom">
      <td width="169">&nbsp;<font size="2" face="Verdana, Arial, Helvetica, sans-serif"><a href="http://orion.mysymmetry.net/CARS/search_profiles.asp">|&lt;</a>
      <a href="http://orion.mysymmetry.net/CARS/search_profiles.asp">&lt;</a> Page 
      1 of 1 <a href="http://orion.mysymmetry.net/CARS/search_profiles.asp">&gt;</a>
      <a href="http://orion.mysymmetry.net/CARS/search_profiles.asp">&gt;|</a></font></td>
    </tr>
  </table>
  <table border="0" style="border-collapse: collapse" bordercolor="#FFFFFF" width="1110" cellspacing="0" height="4">
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
     <form name="maint" method="POST" action="car_rate_rule_maint.asp">
  <table border="1" style="border-collapse: collapse" bordercolor="#FFFFFF" width="1110" id="profiles">
    <tr>
      <td align="left" valign="bottom" bgcolor="#879AA2" height="45" width="30">&nbsp;</td>
      <td class="profile_header" width="58" style="background-color: #E07D1A" height="45">
      <font size="2">Selected</font></td>
      <td class="profile_header" width="291" height="45"><font size="2">
		Description</font></td>
      <td class="profile_header" width="46" height="45"><font size="2">Rate Code</font></td>
      <td class="profile_header" width="94" height="45"><font size="2">Proposed 
      Rate</font></td>
      <td class="profile_header" width="63" height="45"><font size="2">City</font></td>
      <td class="profile_header" width="153" height="45"><font size="2">Car Type</font></td>
      <td class="profile_header" width="349" height="45"><font size="2">Pick-up</font></td>
    </tr>
    
  <%
        
        Dim strClass
        Dim strOrange
        Dim intCount

		While adoRS4.EOF = False
		
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
	<%=adoRS4.Fields("car_rate_rule_change_id").Value%></td>
    <td bgcolor="#FDC677" align="center" height="20">
    <input type="checkbox" value="<%=adoRS4.Fields("car_rate_rule_change_id").Value %>" name="rate_rule_id"></td>
    <td class="<%=strClass %>" height="20" width="291">
    <a target="_self" title="<%=adoRS4.Fields("alert_desc").Value %>" href="alerts_rate_management_car.asp?rateruleid=<%=adoRS4.Fields("rate_rule_id").Value %>">
    <%=adoRS4.Fields("alert_desc").Value %></a></td>
    <td class="<%=strClass %>" height="20" width="46">
    <font color="#080000">
	<%=adoRS4.Fields("client_sys_rate_cd").Value %></font></td>
    <td class="<%=strClass %>" height="20" width="94">
	<%=adoRS4.Fields("new_rt_amt").Value %></td>
    <td class="<%=strClass %>" height="20" width="63">
	<%=adoRS4.Fields("city_cd").Value %>
	</td>
    <td class="<%=strClass %>" height="20"><font size="-1">
	<%=adoRS4.Fields("shop_car_type_cd").Value %> </font></td>
    <td class="<%=strClass %>" height="20">
    <%FormatDateTimel(adoRS4.Fields("arv_dt").Value, 2) %></td>
   
  </tr>
  <%
        
        	adoRS4.MoveNext
        	
        Wend
        
   		adoRS4.Close
		Set adoRS4 = Nothing
		Set adoCmd4 = Nothing

  %>
    
    
    </table>
 
  <table border="0" style="border-collapse: collapse" bordercolor="#FFFFFF" width="1110" cellspacing="0" height="4">
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
  <p><font size="2">&nbsp;| <a href="javascript:maint_action(1);">Delete</a> 
  | <!-- <a href="javascript:maint_action(2)">Copy</a> | -->
  <a href="javascript:maint_action(3)">Enable</a> |
  <a href="javascript:maint_action(4)">Disable</font></a> |</font></p>
  <input type="hidden" name="refresh_from" value="search">
  <input type="hidden" name="action" value="1">
</form>


	
	<font color="#800000" size="2">
<p>Default New User Report - For custom versions please contact your 
Rate-Highway representative at (877) RATE-HWY</p>

<% Set adoRS = Nothing %>
<% Set adoRS1 = Nothing %>
<!--#INCLUDE FILE="footer.asp"-->
</body>

</html>