<%@ Language=VBScript %>
<!-- #INCLUDE FILE="inc/adovbs.asp" -->
<%  Response.Expires = -1
    Response.cachecontrol="private" 
    Response.AddHeader "pragma", "no-cache"
    Response.Buffer = True

    Server.ScriptTimeout = 360
    
    Rem temporary
   	Rem Server.Transfer "maint_default.asp"

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
	Dim varVendors()
	Dim strSelectedVendor
	Dim strBgColor
	Dim blnDarkRow 
	Dim curRate
	Dim strCarList
	Dim strVendList
	Dim strCarCodeListArray 
	Dim strCarCodeList
	Dim blnTotalPrice

	bolOverrideOption = False
	intReportRequestID = 0

	intUserId = Request.Cookies("rate-monitor.com")("user_id")

	If IsNumeric(Request("reportrequestid")) Then
	
		intReportRequestID = Request("reportrequestid")

		strConn = Session("pro_con")
	
	  	Set adoRS = CreateObject("ADODB.Recordset")
	  	Set adoCmd = CreateObject("ADODB.Command")
	
		adoCmd.ActiveConnection = strConn
		adoCmd.CommandText = "car_shop_request_org_detail"
		adoCmd.CommandType = 4
	
		adoCmd.Parameters.Append adoCmd.CreateParameter("@shop_request_id", 3, 1, 0, intReportRequestID)

		Set adoRS = adoCmd.Execute
		
		'Catch the ABG users and re-route them
		If (adoRS.Fields("org_id").Value = 45) Then
		  Server.Transfer "rate_change_report_filtered.asp"
		End If

	
		
		Rem Located in user_access table
		If Request.Cookies("rate-monitor.com")("rate_override") = "True" Then
			bolOverrideOption = True
		ElseIf (adoRS.Fields("parent_id").Value = 6) Then 
			'Payless only
			bolOverrideOption = True
		Else
			bolOverrideOption = False
		End If

	
	  	Set adoRS = CreateObject("ADODB.Recordset")
	  	Set adoCmd = CreateObject("ADODB.Command")
	
		adoCmd.ActiveConnection = strConn
		adoCmd.CommandText = "car_rate_rule_change_select"
		adoCmd.CommandType = 4
	
		adoCmd.Parameters.Append adoCmd.CreateParameter("@shop_request_id", 3, 1, 0, intReportRequestID)

		Set adoRS = adoCmd.Execute
	
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
<meta name="GENERATOR" content="Microsoft FrontPage 12.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Rate-Monitor by Rate-Highway, Inc. | Rate-Change Report</title>
<link rel="stylesheet" type="text/css" href="inc/rh_standard.css">
<link rel="stylesheet" type="text/css" href="inc/rh_report.css">
<link rel="stylesheet" type="text/css" href="inc/sitewide.css">
<script language="javascript" type="text/javascript"  src="inc/sitewide.js" ></script>
<script language="javascript" type="text/javascript"  src="inc/header_menu_support.js"></script>
<script language='Javascript' type="text/javascript" > 
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

<script type="text/javascript" language="javascript"  >
<!-- Begin
function checkAll(field)
{
	//alert(field.name);

	for (i = 0; i < field.length; i++)
		field[i].checked = true ;
}

function uncheckAll(field)
{

	for (i = 0; i < field.length; i++)
		field[i].checked = false ;
}
//  End -->
</script>

<style type="text/css" >
<!--
P {
	COLOR: navy; FONT-FAMILY: Verdana, Arial, sans-serif
}
.copyright {
	FONT-SIZE: 0.7em; TEXT-ALIGN: right
}
.style2 {
	font-size: xx-small;
}
-->
</style>
</head>
<body topmargin="0">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td background="images/top_middle.jpg">
    <img src="images/top_left.jpg" width="423" height="91" alt="Rate-Highway, Inc."></td>
    <td background="images/top_middle.jpg"></td>
    <td background="images/top_tile.gif"><img src="images/top_right.jpg" width="365" height="91" alt="Rate-Highway, Inc."></td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td background="images/med_bar_tile.gif"><img src="images/med_bar.gif"></td>
  </tr>
</table>
<p>
<%

	Rem Default value
	blnTotalPrice = False

	If (Not adoRS Is Nothing) Then
	
		If (adoRS.State = adStateOpen) Then
		
			If (adoRS.EOF = False) Then
	
				If IsNull(adoRS.Fields("calc_total_price_amt").Value) Then
					blnTotalPrice = False
                   	//blnTotalPrice = True
				Else
					blnTotalPrice = True
				End If


			End If
			
		End If
		
	End If	

%> 
<p>&nbsp;<p><font size="-1">
 <table width="1500" border="0" cellpadding="2" background="images/alt_color.gif" style="border-collapse: collapse" bordercolor="#111111" id="table1">
    <tr valign="bottom">
      <td >&nbsp;Rate Rule Suggestion Report</td>
    </tr>
  </table>
  <table border="0" style="border-collapse: collapse" bordercolor="#FFFFFF" width="1500" height="4" id="table2">
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
    <table width="1500" border="0" cellpadding="2" background="images/alt_color.gif" style="border-collapse: collapse" bordercolor="#111111">
  <tr valign="bottom">
    <td ><font size="2">| <a href="javascript:checkAll(document.rate_change.car_rate_rule_change_id)" name="top_of_grid">Select All | 
</a> <a  href="javascript:uncheckAll(document.rate_change.car_rate_rule_change_id)">Unselect All</a> |
<a href="#bottom_of_grid">Bottom of Grid</a> | 
<% If (bolOverrideOption = True) And (blnTotalPrice = False) Then %>
<a target="_self" href="rate_change_report_override.asp?reportrequestid=<%=intReportRequestID %>&security_code=<%=Escape(Request("security_code")) %>">Override Report</a> |
<% End If %>
</font></td>
    </tr>
   </table>
     <table border="0" style="border-collapse: collapse" bordercolor="#FFFFFF" width="1500" height="4" id="table2">
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

<form name="rate_change" method="POST" action="rate_change_request.asp">
  <table border="1" bordercolor="#FFFFFF" id="suggestions" width="1500" cellspacing="0" cellpadding="0" >
    <tr>
      <th align="left" valign="bottom" bgcolor="#879AA2"><font size="2">Id</font></th>
      <th class="profile_header" style="background-color: #E07D1A"><font size="2">Update</font></th>
      <th class="profile_header"><font size="2">Alert Description</font></th>
      <th class="profile_header"><font size="2">Rate Code</font></th>
 	  	 
	  <% If blnTotalPrice Then %>
     	<th class="profile_header" width="75"><font size="2">Suggested Total</font></th>
        <th class="profile_header" width="75"><font size="2">Suggested Base</font></th>
        <th class="profile_header" width="75"><font size="2">Diff. in Totals</font></th>
      <% Else %>
        <th class="profile_header" width="75"><font size="2">Suggested Rate</font></th>
        <th class="profile_header" width="75"><font size="2">Rate Diff.</font></th>
      <% End If %>
      
      <th class="profile_header"><font size="2">City</font></th>
      <th class="profile_header"><font size="2">Car Type</font></th>
      <th class="profile_header"><font size="2">Pick-up</font></th>
      <th class="profile_header"><font size="2">LOR</font></th>
	  <% If blnTotalPrice Then %>
        <th class="profile_header"><font size="2">Current Total</font></th>
      <% Else %>
        <th class="profile_header"><font size="2">Current Rate</font></th>
      <% End If %>
      
      <th class="profile_header"><font size="2">Comp. Set Max</font></th>
      <th class="profile_header"><font size="2">Comp. Set Min</font></th>
      <th class="profile_header" height="45"><font size="2">Min Rate Vendor</font></th>
      <th class="profile_header" height="45"><font size="2">Util. level</font></th>
      
    </tr>
    
 <%
        
        Dim strClass
        Dim strOrange
        Dim intCount
        
        If adoRS Is Nothing Then

		ElseIf (adoRS.State = adStateOpen) Then
		
		While adoRS.EOF = False
		
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
    <td class="<%=strClass %>" width="29">
    <% If IsNull(adoRS.Fields("car_rate_change_id").Value) Then %>
		<%=adoRS.Fields("car_rate_rule_change_id").Value%>
    <% Else %>
    	(updated)
    <% End If %>

	</td>
    <td bgcolor="#FDC677" align="center" width="57">
    <% If IsNull(adoRS.Fields("car_rate_change_id").Value) Then %>
    	<input type="checkbox" value="<%=adoRS.Fields("car_rate_rule_change_id").Value %>" name="car_rate_rule_change_id" >
    <% Else %>
    	<input type="checkbox" value="0" name="car_rate_rule_change_id" disabled="disabled"  >
    <% End If %>
    </td>
    <td class="<%=strClass %>" width="340" >
    <%=adoRS.Fields("alert_desc").Value %></td>
    <td class="<%=strClass %>" width="73" >
    <font color="#080000">
    <% If adoRS.Fields("new_rt_amt").Value = 20000 Then %>
		drop chrg
	<% Else %>
		<%=adoRS.Fields("rate_cd").Value %>
    <% End If %>
	</font>
	</td>
    <td class="<%=strClass %>_right" align="right" width="74">
    <% If adoRS.Fields("new_rt_amt").Value = -1000 Then %>
		<font color='red'><%="<too low>" %></font>
    <% ElseIf adoRS.Fields("new_rt_amt").Value = 10000 Then %>
		<font color='blue'>close</font>
    <% ElseIf adoRS.Fields("new_rt_amt").Value = 10001 Then %>
		<font color='green'>open</font>
    <% ElseIf adoRS.Fields("new_rt_amt").Value = 20000 Then %>
		<%=FormatCurrency(adoRS.Fields("drop_chrg_amt").Value) %>
	<% Else %>
		<%=FormatCurrency(adoRS.Fields("suggested_new_rt_amt").Value) %>
    <% End If %>
	</td>
	
	<% If blnTotalPrice = True Then %>
    <td class="<%=strClass %>_right" >
		<font size="-1">
		<%=FormatCurrency(adoRS.Fields("calc_total_price_amt").Value)	%>
		</font>
	</td>
	<% End If %>
	
    <td class="<%=strClass %>_right" >
	
    <% If adoRS.Fields("new_rt_amt").Value >= 10000 Then %>
		<font size="-1">n/a</font> 
	<% ElseIf (adoRS.Fields("suggested_new_rt_amt").Value > adoRS.Fields("rt_amt").Value) Then %>
		<font size="-1">
		<%=FormatCurrency(adoRS.Fields("suggested_new_rt_amt").Value - adoRS.Fields("rt_amt").Value)	%>
		</font>
	<% Else	%>
		<font  size="-1" color="#FF0000">
		<%=FormatCurrency(adoRS.Fields("suggested_new_rt_amt").Value - adoRS.Fields("rt_amt").Value)	%>
		</font>
	<% End If %>
	</td>
    <td class="<%=strClass %>" width="50" >
	<%=adoRS.Fields("city_cd").Value %>
	</td>
    <td class="<%=strClass %>" width="73">
    <font size="-1">
    <% If IsNull(adoRS.Fields("cross_map_car_type_cd").Value) Then %>
    	<%=adoRS.Fields("shop_car_type_cd").Value %>
    <% ElseIf adoRS.Fields("shop_car_type_cd").Value <> adoRS.Fields("cross_map_car_type_cd").Value Then %>
    	<%=adoRS.Fields("shop_car_type_cd").Value & "|" & adoRS.Fields("cross_map_car_type_cd").Value %>
    <% Else %>
    	<%=adoRS.Fields("shop_car_type_cd").Value %>
    <% End If %>
    </font>
	</td>
    <td class="<%=strClass %>_ctr" width="74">
    <%=FormatDateTime(adoRS.Fields("arv_dt").Value, 2) %></td>
   
    <td class="<%=strClass %>_ctr" width="45">
    <font size="-1">
    <%=adoRS.Fields("lor").Value%></font></td>
   
    <td class="<%=strClass %>_right" align="right" width="75">
    <font size="-1" >
	<%=FormatCurrency(adoRS.Fields("rt_amt").Value) %></font></td>
   
    <td class="<%=strClass %>_right" align="right" width="74">
    <font size="-1">
	<%=FormatCurrency(adoRS.Fields("max_rt_amt").Value) %></font></td>
   
    <td class="<%=strClass %>_right" align="right" width="74">
    <font size="-1">
	<%=FormatCurrency(adoRS.Fields("min_rt_amt").Value) %></font></td>
   
    <td class="<%=strClass %>" height="24" width="74">
    <font size="-1">
	<%=adoRS.Fields("min_vend_cd").Value%></font></td>

    <td class="<%=strClass %>" height="24" width="74">
    <font size="-1">
	<%=FormatNumber(adoRS.Fields("current_util").Value) & "%" %></font></td>
   
    </tr>

<%	
	adoRS.MoveNext
	Response.Flush
	Wend
	
	End If
	
%>    
    
    </table>
 
  <table border="0" style="border-collapse: collapse" bordercolor="#FFFFFF" width="1500" height="4" id="table4">
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
  <p>&nbsp;| 
  <a href="javascript:checkAll(document.rate_change.car_rate_rule_change_id)" name="bottom_of_grid">Select All</a> 
  | <a  href="javascript:uncheckAll(document.rate_change.car_rate_rule_change_id)">Unselect All</a> | 
  <a href="#top_of_grid">Top of Grid</a> |</p>
  <p><input type="submit" value="Perform Update" name="update" ></p>
  <input type="hidden" name="refresh_from" value="search">
  <input type="hidden" name="action" value="1">
  <input type="hidden" name="reportrequestid" value="<%=intReportRequestID %>">
  <input type="hidden" name="security_code" value="<%=intReportRequestID %>">
</form>


	
<% Set adoRS = Nothing %>

<p>
<a target="_blank" href="car_report_by_type.asp?reportrequestid=<%=intReportRequestID %>&security_code=<%=Request("security_code")%>">
click to view source report</a></p>
<table border="1" cellpadding="2" style="border-collapse: collapse" bordercolor="#CFD7DB" id="table5" width="700">
  <tr>
    <td width="214" background="http://www.rate-monitor.com/images/alt_color.gif">
    <font size="2">Report Legend</font>&nbsp;&nbsp; </td>
    <td background="http://www.rate-monitor.com/images/alt_color.gif">&nbsp;</td>
  </tr>
  <tr>
    <td width="214"><font size="2">Description&nbsp;&nbsp; </font></td>
    <td><font size="2">Rate rule description</font></td>
  </tr>
  <tr>
    <td width="214" valign="top" style="font-size: 10pt">Rate Code</td>
    <td valign="top" style="font-size: 10pt">The code your CRS or counter system 
    will use to match your rate.</td>
  </tr>
  <tr>
    <td width="214" valign="top" style="font-size: 10pt">Suggested Rate</td>
    <td valign="top" style="font-size: 10pt">The new rate as calculated by the 
    rate rule</td>
  </tr>
  <tr>
    <td width="214" valign="top" style="font-size: 10pt">Diff.</td>
    <td valign="top" style="font-size: 10pt">The dollar amount difference 
    between your current rate and the proposed rate as generated by chosen rule 
    response</td>
  </tr>
  <tr>
    <td width="214" valign="top" style="font-size: 10pt">City</td>
    <td valign="top" style="font-size: 10pt">The city code that this rate 
    applies to.</td>
  </tr>
  <tr>
    <td width="214" valign="top" style="font-size: 10pt">Car Type</td>
    <td valign="top" style="font-size: 10pt">The car code (SIPP) that rate 
    applies to.</td>
  </tr>
  <tr>
    <td width="214" valign="top" style="font-size: 10pt">Pick-up</td>
    <td valign="top" style="font-size: 10pt">The rental pick-up date.</td>
  </tr>
  <tr>
    <td width="214" valign="top" style="font-size: 10pt">LOR</td>
    <td valign="top" style="font-size: 10pt">The rental length of rent.</td>
  </tr>
  <tr>
    <td width="214" valign="top" style="font-size: 10pt">Current Rate</td>
    <td valign="top" style="font-size: 10pt">The rate that this date/car 
    type/LOR/city combination is currently available at.</td>
  </tr>
  <tr>
    <td width="214" valign="top" style="font-size: 10pt">Comp. Set Max</td>
    <td valign="top" style="font-size: 10pt">The highest rate from your 
    competitive set.</td>
  </tr>
  <tr>
    <td width="214" valign="top" style="font-size: 10pt">Comp. Set Min</td>
    <td valign="top" style="font-size: 10pt">The lowest rate from your 
    competitive set.</td>
  </tr>
  <tr>
    <td width="214" valign="top" style="font-size: 10pt">Min Rate Vendor</td>
    <td valign="top" style="font-size: 10pt">The vendor code of the vendor (or 
    first vendor if there are multiple) that is is offering the Comp. Set Min 
    rate. </td>
  </tr>
  <tr>
    <td width="214" valign="top" style="font-size: 10pt">Util. level</td>
    <td valign="top" style="font-size: 10pt">The utilization level your 
    car/city/date combination is currently at. (Requires CRS or counter system 
    connectivity)</td>
  </tr>
</table>
<p>&nbsp;</p>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" id="export" width="642">
  <tr>
    <td width="222" bgcolor="#92393A" bordercolor="#92393A" valign="top">
    <font color="#EAECF9" size="4">&nbsp;Report Utilities</font></td>
    <td width="420" bordercolor="#92393A" bgcolor="#92393A" valign="top">&nbsp;</td>
  </tr>
  <tr>
    <td width="222" class="report_detail_dark" valign="top">Download to CSV 
    format:</td>
    <td width="420" class="report_detail_light" valign="top">&nbsp;<a href="rate_change_report_export.asp?reportrequestid=<%=intReportRequestId %>&reportformat=1&security_code=<%=Escape(strSecurityCode) %>">download</a> 
	(please choose an appropriate file name)</td>
  </tr>
  <tr>
    <td width="222" class="report_detail_dark" valign="top">Download to XLS 
    format:</td>
    <td width="420" class="report_detail_light" valign="top">&nbsp;<a href="rate_change_report_export.asp?reportrequestid=<%=intReportRequestId %>&reportformat=0&security_code=<%=Escape(strSecurityCode)%>">download</a> </td>
  </tr>
  <tr>
    <td width="222" class="report_detail_dark" valign="top">Download to XML 
    format:</td>
    <td width="420" class="report_detail_light" valign="top">&nbsp;<a href="rate_change_report_export.asp?reportrequestid=<%=intReportRequestId %>&reportformat=6&security_code=<%=Escape(strSecurityCode)%>">download</a> </td>
  </tr>
  </table>

<p>&nbsp;</p>

</font>

<p class="style2">
Debug Section<br>
Total Price = <%=blnTotalPrice %><br>
<a href="rate_change_report_filtered.asp?reportrequestid=<%=Request("reportrequestid")%>&security_code=<%=Request("security_code")%>">filter test</a>
</p>
<!--#INCLUDE FILE="footer.asp"-->
</body>
</html>