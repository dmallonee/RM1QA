<%@ Language=VBScript %>
<!-- #INCLUDE FILE="inc/adovbs.asp" -->
<%  Response.Expires = -1
    Response.cachecontrol="private" 
    Response.AddHeader "pragma", "no-cache"
    Response.Buffer = False  


    Server.ScriptTimeout = 120

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
	Dim strAlertDesc
	Dim intTableCount
	
	
	intTableCount = 0


	If IsNumeric(Request("reportrequestid")) Then

		strConn = Session("pro_con")
	
	  	Set adoRS = CreateObject("ADODB.Recordset")
	  	Set adoCmd = CreateObject("ADODB.Command")
	
		adoCmd.ActiveConnection = strConn
		adoCmd.CommandText = "car_rate_rule_change_select"
		adoCmd.CommandType = 4
	
		adoCmd.Parameters.Append adoCmd.CreateParameter("@shop_request_id", 3, 1, 0, Request("reportrequestid"))

		Set adoRS = adoCmd.Execute
	
	Else
	
		Server.Transfer "default.asp"
	
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
<title>Rate-Monitor by Rate-Highway, Inc. | Rate Change Receipt for Report <%=Request("reportrequestid")%> </title>
<link rel="stylesheet" type="text/css" href="inc/rh_standard.css">
<link rel="stylesheet" type="text/css" href="inc/rh_report.css">
<link rel="stylesheet" type="text/css" href="inc/sitewide.css">
<script type="text/javascript" language="javascript" src="inc/sitewide.js" ></script>
<script type="text/javascript" language="javascript" src="inc/header_menu_support.js"></script>
<script type="text/javascript" language="javascript" >
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
<script type="text/javascript" language="javascript" >
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
<style type="text/css" >
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
    <img src="images/top_left.jpg" width="423" height="91" alt=""></td>
    <td background="images/top_middle.jpg"></td>
    <td background="images/top_tile.gif"><img src="images/top_right.jpg" width="365" height="91" alt=""></td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td background="images/med_bar_tile.gif"><img src="images/med_bar.gif" alt=""></td>
  </tr>
</table>
<p align="right"><font class="copyright">Copyright (c) 2001-<%=Year(Now)%>, Rate-Highway, Inc. All Rights Reserved.</font></p>
<font size="-1">
 <table width="1000" border="0" cellpadding="2" style="border-collapse: collapse" id="table0">
    <tr valign="bottom">
      <td >
      <p align="center"><font size="5" color="#384F5B">The Following Changes Have Been 
Made</font></td>
    </tr>
    <tr valign="bottom">
      <td >&nbsp;</td>
    </tr>
  </table>

    
 <%
        
        Dim strClass
        Dim strOrange
        Dim intCount
        
        strAlertDesc = ""
        
        If adoRS Is Nothing Then

		ElseIf (adoRS.State = adStateOpen) Then

		While adoRS.EOF = False
		

%>
		
		<% If (strAlertDesc <> adoRS.Fields("alert_desc").Value)Then 
	
			If  (strAlertDesc <> "")  Then
		%>
	  	</table>
 
		<table border="0" style="border-collapse: collapse" bordercolor="#FFFFFF" width="1000" height="4" id="footerbar<%=intTableCount %>">
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
		<br>
		<br>	

		<%
		   End If
	     
	       strAlertDesc = adoRS.Fields("alert_desc").Value
		   strClass = "profile_dark"
		   intTableCount = intTableCount + 1
		     
		%>


		<table width="1000" border="0" cellpadding="2" background="images/alt_color.gif" style="border-collapse: collapse" bordercolor="#111111" id="desc<%=intTableCount %>">
		    <tr valign="bottom">
		      <td >&nbsp;<b>Rate Rule Description:</b> <%=adoRS.Fields("alert_desc").Value %></td>
		    </tr>
		</table>
		<table border="0" style="border-collapse: collapse" bordercolor="#FFFFFF" width="1000" height="4" id="headerbar<%=intTableCount %>">
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
 
		<table border="1" bordercolor="#FFFFFF" id="table3" width="1000" cellspacing="0" cellpadding="0" id="rate_changes<%=intTableCount %>">
    		<tr>
      			<th bgcolor="#879AA2">&nbsp;</th>
      			<th class="profile_header" style="background-color: #E07D1A"><font size="2">New Rate</font></th>
      			<th class="profile_header"><font size="2">Rate Code</font></th>
      			<th class="profile_header" width="75"><font size="2">Diff.</font></th>
      			<th class="profile_header"><font size="2">City</font></th>
      			<th class="profile_header"><font size="2">Car Type</font></th>
      			<th class="profile_header"><font size="2">Pick-up</font></th>
      			<th class="profile_header"><font size="2">LOR</font></th>
      			<th class="profile_header"><font size="2">Current Rate</font></th>
      			<th class="profile_header"><font size="2">Comp. Set Max</font></th>
      			<th class="profile_header"><font size="2">Comp. Set Min</font></th>
      			<th class="profile_header" height="45"><font size="2">Min Rate Vendor</font></th>
      			<th class="profile_header" height="45"><font size="2">Util. level</font></th>      
  
 			</tr>
		

	    <% End If %>
	    
		<%


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
	    <td class="<%=strClass %>_ctr" width="29" align="center" >
	    <% If IsNull(adoRS.Fields("car_rate_change_id").Value) Then %>
		<img border="0" src="images/failed-sm.gif" width="14" height="14" alt="" >
		<% Else %>
		<img border="0" src="images/success-sm.gif" width="14" height="14" alt="<%=adoRS.Fields("car_rate_change_id").Value %>">
		<% End If %>
		</td>
	    <td bgcolor="#FDC677" width="74" style="text-align: right; font-family:Verdana; font-size:10pt; vertical-align:bottom">
	    <% If adoRS.Fields("new_rt_amt").Value = -1000 Then %>
		<font color='red'><%="<too low>" %></font>
	    <% Else %>
		    <% If adoRS.Fields("new_rt_amt").Value <> adoRS.Fields("suggested_new_rt_amt").Value Then %>
				<span style='font-family: Wingdings' ><font size='4'>@</font></span><%=FormatCurrency(adoRS.Fields("new_rt_amt").Value) %>
			<% Else %>
				<%=FormatCurrency(adoRS.Fields("new_rt_amt").Value) %>
			<% End If %>
	    <% End If %>
		</td>
	    <td class="<%=strClass %>_ctr" width="73" >
	    <font color="#080000">
		<%=adoRS.Fields("rate_cd").Value %></font></td>
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
	    <td class="<%=strClass %>_ctr" width="50" >
		<%=adoRS.Fields("city_cd").Value %>
		</td>
	    <td class="<%=strClass %>_ctr" width="73"><font size="-1">
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
   
	    <td class="<%=strClass %>_ctr" height="24" width="74">
	    <font size="-1">
		<%=adoRS.Fields("min_vend_cd").Value%></font></td>

	    <td class="<%=strClass %>_right" height="24" width="74">
	    <font size="-1">
		<%=FormatNumber(adoRS.Fields("current_util").Value) & "%" %></font></td>
   
	    </tr>


<%	

	
	adoRS.MoveNext
	Wend
	
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
    
  </table>
 
  <table border="0" style="border-collapse: collapse" bordercolor="#FFFFFF" width="1000" height="4" id="footerbar<%=intTableCount %>">
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
    
  <input type="hidden" name="action" value="1">
  <input type="hidden" name="reportrequestid" value="<%=Request("reportrequestid") %>">
  <input type="hidden" name="security_code" value="<%=Request("security_code") %>">
  <input type="hidden" name="car_rate_rule_change_id" value="<%=strCompare %>">


	


	
<% Set adoRS = Nothing %>

<font size="2">

      &nbsp;</font><font size="3" color="#000000"> </font>
<table border="0" style="border-collapse: collapse" width="1000" id="table6">
  <tr>
    <td>&nbsp;</td>
    <td width="25">

<font size="2">

      <img border="0" src="images/success-lg.gif" width="25" height="25"></font></td>
    <td width="160">

<font size="2" color="#000000">

      New Rate Accepted</font></td>
    <td width="34">

<font size="2">

      <img border="0" src="images/failed-lg.gif" width="25" height="25"></font></td>
    <td width="151"><font color="#000000" size="2">New Rate Declined</font></td>
    <td width="35">
    <img border="0" src="images/remaining-lg.gif" width="25" height="25"></td>
    <td width="303"><font size="2">Rate Change Conflict</font></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td width="25">&nbsp;</td>
    <td width="160">&nbsp;</td>
    <td width="34">&nbsp;</td>
    <td width="151">&nbsp;</td>
    <td width="340" colspan="2">&nbsp;</td>
  </tr>
</table>

<table border="1" cellpadding="2" style="border-collapse: collapse" bordercolor="#CFD7DB" id="table5" width="700">
  <tr>
    <td background="images/alt_color.gif" colspan="2">
    <font size="2">Report Legend</font>&nbsp;&nbsp; </td>
  </tr>
  <tr>
    <td width="214"><font size="2">Description&nbsp;&nbsp; </font></td>
    <td><font size="2">Rate rule description that caused the alert</font></td>
  </tr>
  <tr>
    <td width="214" valign="top" style="font-size: 10pt">New Rate</td>
    <td valign="top" style="font-size: 10pt">The new rate as calculated by the 
    chosen rate rule response. The&nbsp;&quot;<span style='font-family: Wingdings' ><font size='4'>@</font></span>&quot; 
	symbol means the rate suggestion was overridden. </td>
  </tr>
  <tr>
    <td width="214" valign="top" style="font-size: 10pt">Rate Code</td>
    <td valign="top" style="font-size: 10pt">The code your CRS or counter system 
    will use to match your rate.</td>
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
<p>&nbsp;</p>
<p>Questions about what you are seeing on this report? Please
<a href="mailto:support@rate-highway.com?subject=Rate Change Receipt:<%=Request("reportrequestid") %> - Question">
click here</a> to send a question to Support.</p>

<p><font color="#000000" size="1">Report Number: <%=Request("reportrequestid") %></font></p>
<!--#INCLUDE FILE="footer.asp"-->
</body>
</html>