<%@ Language=VBScript %>
<!-- #INCLUDE FILE="inc/adovbs.asp" -->
<%  Response.Expires = -1
    Response.cachecontrol="private" 
    Response.AddHeader "pragma", "no-cache"

    Server.ScriptTimeout = 360

	'On Error Resume Next

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


	If IsNumeric(Request("reportrequestid")) Then

		strConn = Session("pro_con")
	
	  	Set adoRS = CreateObject("ADODB.Recordset")
	  	Set adoCmd = CreateObject("ADODB.Command")
	
		adoCmd.ActiveConnection = strConn
		adoCmd.CommandText = "car_rate_rule_change_select"
		adoCmd.CommandType = 4
	
		adoCmd.Parameters.Append adoCmd.CreateParameter("@shop_request_id", 3, 1, 0, Request.Form("reportrequestid"))

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
	alert(field.name);

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
<p align="center"><font size="5" color="#384F5B">Please Confirm Your Rate Changes<br>
</font><font size="2" color="#384F5B"><b>TIP</b>: Use your browser's back button 
to go back and make any necessary changes to your list of selected rate changes<br>
</font></p>
<p><font size="-1">
 <table width="1300" border="0" cellpadding="2" background="images/alt_color.gif" style="border-collapse: collapse" bordercolor="#111111" id="table1">
    <tr valign="bottom">
      <td >&nbsp;Rate Change Report - Please Review and Confirm<font size="-1">&nbsp;&nbsp;</font></td>
    </tr>
  </table>
  <table border="0" style="border-collapse: collapse" bordercolor="#FFFFFF" width="1300" height="4" id="table2">
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
  <form name="rate_change" method="POST" action="rate_change_request_insert.asp">
  <table border="1" bordercolor="#FFFFFF" id="table3" width="1300" cellspacing="0" cellpadding="0" >
    <tr>
      <th align="left" valign="bottom" bgcolor="#879AA2"><font size="2">Id</font></th>
      <th class="profile_header"><font size="2">Alert Description</font></th>
      <th class="profile_header"><font size="2">Rate Code</font></th>
      <th class="profile_header" style="background-color: #E07D1A"><font size="2">Proposed Rate</font></th>
      <th class="profile_header" width="75"><font size="2">Diff.</th>
      <th class="profile_header"><font size="2">City</font></th>
      <th class="profile_header"><font size="2">Car Type</font></th>
      <th class="profile_header"><font size="2">Pick-up</font></th>
      <th class="profile_header"><font size="2">LOR</th>
      <th class="profile_header"><font size="2">Current Rate</th>
      <th class="profile_header"><font size="2">Comp. Set Max</th>
      <th class="profile_header"><font size="2">Comp. Set Min</th>
      <th class="profile_header" height="45"><font size="2">Min Rate Vendor</th>
      <th class="profile_header" height="45"><font size="2">Util. level</th>      
    </tr>
    
 <%
        
        Dim strClass
        Dim strOrange
        Dim intCount
        Dim strIDArray
        Dim strRateArray
        Dim intArrayHolder
		Dim curRateAmt         
        Dim strCompare
        Dim bolOverride
        
        intCount = 0
        
        strCompare = Request.Form("car_rate_rule_change_id")
        strRateArray = Split(Request.Form("new_rate_amt"), ", $", -1, 1)
		strIDArray = Split(Request.Form("new_rate_amt_id"), ",", -1, 1)
        
        If adoRS Is Nothing Then

		ElseIf (adoRS.State = adStateOpen) Then

		While adoRS.EOF = False
		

			If InStr(1, strCompare, adoRS.Fields("car_rate_rule_change_id").Value) Then

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
		<%=adoRS.Fields("car_rate_rule_change_id").Value%></td>
	    <td class="<%=strClass %>" width="340" >
	    <%=adoRS.Fields("alert_desc").Value %></td>
	    <td class="<%=strClass %>" width="73" >
	    <font color="#080000">
		<%=adoRS.Fields("rate_cd").Value %></font></td>
	    <td bgcolor="#FDC677" width="74" style="text-align: right; font-family:Verdana; font-size:10pt; vertical-align:bottom">
		<%

		curRateAmt = FormatCurrency(adoRS.Fields("new_rt_amt").Value)
		bolOverride = False 
		
		For intArrayHolder = LBound(strIDArray) to UBound(strIDArray)

			If Clng(strIDArray(intArrayHolder)) = CLng(adoRS.Fields("car_rate_rule_change_id").Value) Then
			
				If FormatCurrency(adoRS.Fields("new_rt_amt").Value) = FormatCurrency(strRateArray(intArrayHolder)) Then
					curRateAmt = FormatCurrency(strRateArray(intArrayHolder))
					Response.Write curRateAmt			
				Else
					curRateAmt = FormatCurrency(strRateArray(intArrayHolder))
					Response.Write "<span style='font-family: Wingdings' ><font size='4'>@</font></span> " & FormatCurrency(strRateArray(intArrayHolder))  
					bolOverride = True
				End If

				Exit For
			End If
		Next
		
		
		

		%>
		<input type="hidden" name="car_rate_rule_change_id" value="<%=strIDArray(intArrayHolder) %>" >
		<% If bolOverride = True Then %>
		<input type='hidden' name='new_rate_amt' value='<%="@" & Replace(curRateAmt, ",", "") %>' >
		<% Else %>
		<input type='hidden' name='new_rate_amt' value='<%=curRateAmt %>' >
		<% End If %>

		</td>
	    <td class="<%=strClass %>_right" >
		<font size="-1">
		<% If curRateAmt > adoRS.Fields("rt_amt").Value Then %>
		<%=FormatCurrency( curRateAmt - adoRS.Fields("rt_amt").Value)	%>
		<% Else	%>
		<font color="#FF0000">
		<%=FormatCurrency(curRateAmt - adoRS.Fields("rt_amt").Value)	%>
		</font>
		<% End If %>
		</font></td>
	    <td class="<%=strClass %>" width="50" >
		<%=adoRS.Fields("city_cd").Value %>
		</td>
	    <td class="<%=strClass %>" width="73"><font size="-1">
		<%=adoRS.Fields("shop_car_type_cd").Value %> </font></td>
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

		End If
	
	adoRS.MoveNext
	Wend
	
	End If
	
%>    
    
    </table>
 
  <table border="0" style="border-collapse: collapse" bordercolor="#FFFFFF" width="1300" height="4" id="table4">
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
  </table> <p>
  <!--
 <input type="checkbox" name="debug" value="1"></font><font size="1">test mode 
	- do not commit updates</p>
  </font><font size="-1">
  -->
  <p>
  <% If intCount = 0 Then %>
  <input type="submit" value="Commit Updates" name="update" disabled>
  TIP:
  You must select one or more updates in order to submit changes
  <% Else %>
  <input type="submit" value="Commit Updates" name="update" >
  <% End If %>
  </font></p>
  <input type="hidden" name="action" value="1">
  <input type="hidden" name="reportrequestid" value="<%=Request.Form("reportrequestid") %>">
  <input type="hidden" name="security_code" value="<%=Request.Form("security_code") %>">
  <!--
  <input type="hidden" name="car_rate_rule_change_id" value="<%=strCompare %>">
  -->
	<input type="hidden" name="debug" value="0">
</form>


	
<% Set adoRS = Nothing %>

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
    <td valign="top" style="font-size: 10pt">The rate at which this date/car 
    type/LOR/city combination is currently available.</td>
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
<font size="1">debug: <%=Request.Form("new_rate_amt") %>
</font>
<!--#INCLUDE FILE="footer.asp"-->
</body>
</html>