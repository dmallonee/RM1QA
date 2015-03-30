<%@ Language=VBScript %>
<!-- #INCLUDE FILE="inc/login_check.asp" -->
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

	strConn = Session("pro_con")
	
  	Set adoRS = CreateObject("ADODB.Recordset")
  	Set adoCmd = CreateObject("ADODB.Command")
	
	adoCmd.ActiveConnection = strConn
	adoCmd.CommandText = "car_rate_rule_change_select"
	adoCmd.CommandType = 4
	
	adoCmd.Parameters.Append adoCmd.CreateParameter("@shop_request_id", 3, 1, 0, 4107) 'Request("reportrequestid"))

	Set adoRS = adoCmd.Execute
	
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
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>New Page 0</title>
<meta name="VI60_defaultClientScript" content="JavaScript">
<link rel="stylesheet" type="text/css" href="inc/rh_standard.css">
<link rel="stylesheet" type="text/css" href="inc/rh_report.css">
<link rel="stylesheet" type="text/css" href="inc/sitewide.css">
<script src="inc/sitewide.js" language="javascript"></script>
<script language="JavaScript" type="text/JavaScript">
function SetSelectedRate(this_radio_value, array_index) {

	if (this_radio_value == 'override'){
		document.rate_list.override[array_index].value = 'ABCD'
	}

	if (this_radio_value == 'accept'){
		document.rate_list.override[array_index].value = document.rate_list.current_td[array_index].innerText
	}


	if (this_radio_value == 'reject'){
		document.rate_list.override[array_index].value = document.rate_list.suggested_td[array_index].innerText
	}



}

function SetSelectedRate2(this_radio_value, array_index, rate_amt) {

	if (this_radio_value == 'override'){
		document.rate_list.override[array_index].value = rate_amt
	}

	if (this_radio_value == 'accept'){
		document.rate_list.override[array_index].value = rate_amt	
	}


	if (this_radio_value == 'reject'){
		document.rate_list.override[array_index].value = rate_amt
	}



}



</script>
</head>

<body style="font-family: Arial; font-size: 8pt; background-color:#CED7DB" >

<form method="POST" action="rate_change_report_car_asp" name="rate_list" class='login'>
  <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse; font-size:10pt" width="1000" id="table1">
    <tr class="profile_header">
      <td width="23" align="center"><span style="font-family: Wingdings">
      <font size="4" color="#008000">ü</font></span></td>
      <td width="20" align="center"><span style="font-family: Wingdings">
      <font size="4" color="#800000">û</font></span></td>
      <td width="20" align="center" >
      <span style="font-family: Wingdings" ><font size="4">@</font></span></td>
      <td width="57" align="center">Date</td>
      <td width="74" align="center">Product</td>
      <td width="59" align="center">Car</td>
      <td width="63" align="center">Current</td>
      <td width="80" align="center">Suggested</td>
      <td width="79" align="center">Override</td>
      <td>Alert</td>
     
    </tr>
    
 <%
        
        Dim strClass
        Dim strOrange
        Dim intCount

		If (adoRS.State = adStateOpen) Then

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
    
    <tr class="<%=strClass%>">
      <td width="23" >
      <input type="radio" value="override" name="rate[<%=adoRS.Fields("car_rate_rule_change_id").Value%>]" checked onclick="SetSelectedRate2(this.value, 0,'<%=FormatCurrency(adoRS.Fields("new_rt_amt").Value)%>' )" class='radio'></td>
      <td width="20">
      <input type="radio" value="accept" name="rate[<%=adoRS.Fields("car_rate_rule_change_id").Value%>]" onclick="SetSelectedRate2(this.value, 0, '<%=FormatCurrency(adoRS.Fields("rt_amt").Value)%>')" class="radio"></td>
      <td width="20">
      <input type="radio" value="reject" name="rate[<%=adoRS.Fields("car_rate_rule_change_id").Value%>]" onclick="SetSelectedRate2(this.value, 0, '(enter value)')" class="radio"></td>
      <td width="57" align="center">
      <%=FormatDateTime(adoRS.Fields("arv_dt").Value, 2) %></td>
      <td width="74">&nbsp;</td>
      <td width="59" align="center"><%=adoRS.Fields("shop_car_type_cd").Value %></td>
      <td width="63" align="right" id='current_td'><%=FormatCurrency(adoRS.Fields("rt_amt").Value)%></td>
      <td width="80" align="right" id='suggested_td'><%=FormatCurrency(adoRS.Fields("new_rt_amt").Value)%></td>
      <td width="79">
      <input type="text" name="override[<%=adoRS.Fields("car_rate_rule_change_id").Value%>]" size="15" style="font-family: Vendana, Arial, Helvetica, sans-serif; font-size: 10pt; text-align:right" ></td>
      <td>
      <%=adoRS.Fields("alert_desc").Value %></td>
     
    </tr>
    
<% 	adoRS.MoveNext
	Wend
	
	End If
	
%>

    
  </table>
</form>
<!--#INCLUDE FILE="footer.asp"-->
</body>
</html>