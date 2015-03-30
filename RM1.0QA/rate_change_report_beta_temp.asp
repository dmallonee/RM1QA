<%@ Language=VBScript %>
<!-- #INCLUDE FILE="inc/adovbs.asp" -->
<%  Response.Expires = -1
    Response.cachecontrol="private" 
    Response.AddHeader "pragma", "no-cache"

    Server.ScriptTimeout = 360

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

	strIPAddress = Request.Servervariables("REMOTE_ADDR") 

	If IsNumeric(Request("reportrequestid")) Then

		strConn = Session("pro_con")
	
	  	Set adoRS = CreateObject("ADODB.Recordset")
	  	Set adoCmd = CreateObject("ADODB.Command")
	
		adoCmd.ActiveConnection = strConn
		adoCmd.CommandText = "car_rate_rule_change_select"
		adoCmd.CommandType = 4
	
		adoCmd.Parameters.Append adoCmd.CreateParameter("@shop_request_id", 3, 1, 0, Request("reportrequestid"))

		Set adoRS = adoCmd.Execute
		
	 	Set adoRS1 = CreateObject("ADODB.Recordset")
	  	Set adoCmd = CreateObject("ADODB.Command")
		adoRS1.CursorLocation = adUseClient
	
		adoCmd.ActiveConnection = strConn
		adoCmd.CommandText = "car_shopped_rate_select_rpt2"
		adoCmd.CommandType = 4
	
		adoCmd.Parameters.Append adoCmd.CreateParameter("@shop_request_id", 3, 1, 0, Request("reportrequestid")) 
		adoCmd.Parameters.Append adoCmd.CreateParameter("@report", 3, 1, 0, 2)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@vend_override", 200, 1, 2, Null)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@ipaddress", 200, 1, 20, strIPAddress)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@shop_car_type_cd", 200, 1, 4, "ALLL")	
		'adoCmd.Parameters.Append adoCmd.CreateParameter("@city_cd", 200, 1, 6, "ALL")

		'adoRS1.Open adoCmd, , adOpenStatic, adLockReadOnly, adCmdStoredProc 
		
		

	
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
<title>Rate-Monitor by Rate-Highway, Inc. | Rate-Change Report</title>
<link rel="stylesheet" type="text/css" href="inc/rh_standard.css">
<link rel="stylesheet" type="text/css" href="inc/rh_report.css">
<link rel="stylesheet" type="text/css" href="inc/sitewide.css">
<script language="javascript" src="inc/sitewide.js" ></script>
<script language="javascript" src="inc/header_menu_support.js"></script>
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

<SCRIPT LANGUAGE="JavaScript"  >
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

<SCRIPT LANGUAGE="JavaScript">
<!-- Begin
function formatCurrency(num) {
	num = num.toString().replace(/\$|\,/g,'');
	if(isNaN(num))
		num = "0";
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
//  End -->
</script>
<style>
<!--
P {
	COLOR: navy; FONT-FAMILY: Verdana, Arial, sans-serif
}
.copyright {
	FONT-SIZE: 0.7em; TEXT-ALIGN: right
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
<p align="right"><font class="copyright">Copyright (c) 2001-<%=Year(Now)%>, Rate-Highway, Inc. All Rights Reserved.</font></p>
<p>
<p><font size="-1"><br>
&nbsp;</font></p>
<p><font size="-1">
 <table width="1100" border="0" cellpadding="2" background="images/alt_color.gif" style="border-collapse: collapse" bordercolor="#111111" id="table1">
    <tr valign="bottom">
      <td >&nbsp;Advanced Rate Rule Alert Report - Updateable&nbsp;&nbsp;&nbsp;
		<font color="#FF0000">BETA</font></td>
    </tr>
  </table>
<%  		 'Response.Write adoRS.Fields("shop_car_type_cd").Value & "/" & adoRS.Fields("alert_desc").Value & "/" & adoRS.Fields("car_rate_rule_change_id").Value   	%>

  <table border="0" style="border-collapse: collapse" bordercolor="#FFFFFF" width="1100" height="4" id="table2">
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
<%  		 'Response.Write adoRS.Fields("shop_car_type_cd").Value & "/" & adoRS.Fields("alert_desc").Value & "/" & adoRS.Fields("car_rate_rule_change_id").Value   	%>


  <form name="rate_change" method="POST" action="rate_change_request_beta.asp">
  <table border="1" bordercolor="#FFFFFF" id="rate_changes" width="1100" cellspacing="0" cellpadding="0" >
    <tr>
      <th class="profile_header"><font size="2">Car Type</font></th>
      <th class="profile_header"><font size="2">Rate Code</font></th>
      <th class="profile_header"><font size="2">Pick-up</font></th>
      <th class="profile_header"><font size="2">Current Rate</th>
      <th class="profile_header"><font size="2">Proposed Rate</font></th>
      <th class="profile_header" style="background-color: #E07D1A"><font size="2">Update</font></th>
      <th class="profile_header"><font size="2">Alert Description</font></th>
      <th class="profile_header" width="75"><font size="2">Diff</font>.</th>
      <th class="profile_header"><font size="2">City</font></th>
      <th class="profile_header"><font size="2">LOR</th>
      <th class="profile_header"><font size="2">Comp. Set Max</th>
      <th class="profile_header"><font size="2">Comp. Set Min</th>
      <th class="profile_header" height="45"><font size="2">Min Rate Vendor</th>
      <th class="profile_header" height="45"><font size="2">Util. level</th>      
          
    </tr>
    
 <%
        
        Dim strClass
        Dim strOrange
        Dim intCount
        
        If adoRS Is Nothing Then

		ElseIf (adoRS.State = adStateOpen) Then

		While adoRS.EOF = False
  		 'Response.Write "3/" & adoRS.Fields("shop_car_type_cd").Value & "/" & adoRS.Fields("alert_desc").Value & "/" & adoRS.Fields("car_rate_rule_change_id").Value   	

		
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
    <td class="<%=strClass %>" width="73"><font size="-1">
	<%=adoRS.Fields("shop_car_type_cd").Value %> </font></td>

    <td class="<%=strClass %>" width="73" >
    <font color="#080000">
	<%=adoRS.Fields("rate_cd").Value %></font></td>

    <td class="<%=strClass %>_ctr" width="74">
    <%=FormatDateTime(adoRS.Fields("arv_dt").Value, 2) %></td>

    <td class="<%=strClass %>_right" align="right" width="75">
    <font size="-1" >
    
    <% 
    	curComparisonRate =	adoRS.Fields("rt_amt").Value
    
    %>
	<%=FormatCurrency(curComparisonRate) %>
	</font></td>


    <td class="<%=strClass %>_right" align="right" width="74">
    <% If adoRS.Fields("new_rt_amt").Value = -1000 Then %>
	<font color='red'><%="<too low>" %></font>
    <% Else %>
	<input align="right"  maxlength="9" name="new_rate_amt" value="<%=FormatCurrency(adoRS.Fields("new_rt_amt").Value) %>" size="11" style="text-align: right" onBlur="this.value=formatCurrency(this.value);">
	<input type="hidden" name="new_rate_amt_id" value="<%=adoRS.Fields("car_rate_rule_change_id").Value %>"  >
    <% End If %>
	</td>

    <td bgcolor="#FDC677" align="center" width="57">
    <input type="checkbox" value="<%=adoRS.Fields("car_rate_rule_change_id").Value %>" name="car_rate_rule_change_id" ></td>

    <td class="<%=strClass %>" width="340" >
    <%=adoRS.Fields("alert_desc").Value %></td>
    
    <td class="<%=strClass %>_right" >
	<font size="-1">
	<% If adoRS.Fields("new_rt_amt").Value > adoRS.Fields("rt_amt").Value Then %>
	<%=FormatCurrency( adoRS.Fields("new_rt_amt").Value - adoRS.Fields("rt_amt").Value)	%>
	<% Else	%>
	<font color="#FF0000">
	<%=FormatCurrency(adoRS.Fields("new_rt_amt").Value - adoRS.Fields("rt_amt").Value)	%>
	</font>
	<% End If %>
	</font></td>
    
    <td class="<%=strClass %>" width="50" >
	<%=adoRS.Fields("city_cd").Value %>
	</td>
   
    <td class="<%=strClass %>_ctr" width="45">
    <font size="-1">
    <%=adoRS.Fields("lor").Value%></font></td>
   
   
    <td class="<%=strClass %>_right" align="right" width="74">
    <font size="-1">
	<%=FormatCurrency(adoRS.Fields("max_rt_amt").Value) %></font></td>
   
    <td class="<%=strClass %>_right" align="right" width="74">
    <font size="-1">
	<%=FormatCurrency(adoRS.Fields("min_rt_amt").Value) %></font></td>
   
    <td class="<%=strClass %>" height="24" width="74">
    <font size="-1">
	<%=adoRS.Fields("min_vend_cd").Value %></font></td>

    <td class="<%=strClass %>" height="24" width="74">
    <font size="-1">
	<%=FormatNumber(adoRS.Fields("current_util").Value) & "%" %></font></td>

 
   
    </tr>

<%	
 	    adoRS.MoveNext

	  Wend
	
	End If
	
%>    
    
    </table>
 
  <table border="0" style="border-collapse: collapse" bordercolor="#FFFFFF" width="1100" height="4" id="table4">
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
  <p>&nbsp;| <a href="javascript:checkAll(document.rate_change.car_rate_rule_change_id)">Select All</a> 
  | <a  href="javascript:uncheckAll(document.rate_change.car_rate_rule_change_id)">Unselect All</a> |</p>
  <p><input type="submit" value="Perform Update" name="update" ></font></p>
  <input type="hidden" name="refresh_from" value="search">
  <input type="hidden" name="action" value="1">
  <input type="hidden" name="reportrequestid" value="<%=Request("reportrequestid") %>">
  <input type="hidden" name="security_code" value="<%=Request("security_code") %>">
</form>


	
<% Set adoRS = Nothing %>

<p>
<a target="_blank" href="car_report_by_type.asp?reportrequestid=<%=Request("reportrequestid")%>&security_code=<%=Request("security_code")%>">
click to view source report</a></p>
<table border="1" cellpadding="2" style="border-collapse: collapse" bordercolor="#CFD7DB" id="table5" width="700">
  <tr>
    <td width="214" background="http://www.rate-monitor.com/images/alt_color.gif">
    <font size="2">Rate Alert Report Legend</font>&nbsp;&nbsp; </td>
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
    <td width="214" valign="top" style="font-size: 10pt">Proposed Rate</td>
    <td valign="top" style="font-size: 10pt">The new rate as calculated by the 
    chosen rate rule response.</td>
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
<!--#INCLUDE FILE="footer.asp"-->
</body>
</html>