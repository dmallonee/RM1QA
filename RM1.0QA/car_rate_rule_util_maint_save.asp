<%@ Language=VBScript %>
<!-- #INCLUDE FILE="inc/adovbs.asp" -->
<% Response.Expires = -1
   Response.cachecontrol="private" 
   Response.AddHeader "pragma", "no-cache"

   Server.ScriptTimeout = 180

	
	On Error Resume Next
	

	If err.number <> 0 Then
	   pad="&nbsp;&nbsp;&nbsp;&nbsp;"
	   response.write "<b>An error cccured while collecting data source information</b><br>"
	   response.write pad & "Error Number= #<b>" & err.number & "</b><br>"
	   response.write pad & "Error Desc.= <b>" & err.description & "</b><br>"
	   response.write pad & "Help Context= <b>" & err.HelpContext & "</b><br>"
	   response.write pad & "Help File Path=<b>" & err.helpfile & "</b><br>"
	   response.write pad & "Error Source= <b>" & err.source & "</b><br><hr>"

	End If

%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

<head>
<meta content="en-us" http-equiv="Content-Language" />
<meta content="text/html; charset=windows-1252" http-equiv="Content-Type" />
<title>Rule Name</title>
<style type="text/css">body {
	font-family: Arial, Helvetica, sans-serif;
	font-size: small;
	font-weight: normal;
	font-style: normal;
	font-variant: normal;
	text-transform: none;
	color: #000000;
}
select {
	font-family: Arial, Helvetica, sans-serif;
	font-size: small;
}
.td_label_rt {
	text-align: right;
}
.input_min {
	font-size: small;
	font-family: Arial, Helvetica, sans-serif;
	text-align: right;
	width: 75px;
	background-color: #666666;
	color: #FFFFFF;
}
.input_max {
	font-size: small;
	font-family: Arial, Helvetica, sans-serif;
	text-align: right;
	width: 75px;
	background-color: #CCCCCC;
	color: #000000;
}
.input_off {
	font-size: small;
	font-family: Arial, Helvetica, sans-serif;
	text-align: right;
	width: 75px;
	background-color: #FFFFFF;
	color: #FFFFFF;
}

.style1 {
	border-bottom-style: solid;
	border-bottom-width: 1px;
}
.style2 {
	border-top-style: solid;
	border-top-width: 1px;
}

</style>
</head>

<body>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td background="images/top_tile.gif">
    <img src="images/top.jpg" width="770" height="91"></td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td background="images/b_tile.gif">
    <!-- #INCLUDE FILE="inc/page_header_buttons.htm" -->
    </td>
  </tr>
</table>

<p>&nbsp;</p>
<form name="rules" id="rules">
<table cellpadding="0" cellspacing="0" style="width: 1600px" id="rules">
	<tr>
		<th colspan="2">&nbsp;</th>
		<td>&nbsp;</td>
		<th colspan="2">Same Day</th>
		<th style="width: 15px">&nbsp;</th>
		<th colspan="2">Next Day</th>
		<th style="width: 15px">&nbsp;</th>
		<th colspan="2">2 to 4 Days Out</th>
		<th style="width: 15px">&nbsp;</th>
		<th colspan="2">5 to 7 Days Out</th>
		<th style="width: 15px">&nbsp;</th>
		<th colspan="2">8 to 14 Days Out</th>
		<th style="width: 15px">&nbsp;</th>
		<th colspan="2">15 to 30 Days Out</th>
		<th style="width: 15px">&nbsp;</th>
		<th colspan="2">31 To 50 Days Out</th>
		<th style="width: 15px">&nbsp;</th>
		<th colspan="2">51+ Days Out</th>
	</tr>
	<tr>
		<th colspan="2" class="style1">Rule Name</th>
		<td>&nbsp;</td>
		<th style="width: 75px" class="style1">Min</th>
		<th style="width: 75px" class="style1">Max</th>
		<th style="width: 15px">&nbsp;</th>
		<th style="width: 75px">Min</th>
		<th style="width: 75px">Max</th>
		<th style="width: 15px">&nbsp;</th>
		<th style="width: 75px">Min</th>
		<th style="width: 75px">Max</th>
		<th style="width: 15px">&nbsp;</th>
		<th style="width: 75px">Min</th>
		<th style="width: 75px">Max</th>
		<th style="width: 15px">&nbsp;</th>
		<th style="width: 75px">Min</th>
		<th style="width: 75px">Max</th>
		<th style="width: 15px">&nbsp;</th>
		<th style="width: 75px">Min</th>
		<th style="width: 75px">Max</th>
		<th style="width: 15px">&nbsp;</th>
		<th style="width: 75px">Min</th>
		<th style="width: 75px">Max</th>
		<th style="width: 15px">&nbsp;</th>
		<th style="width: 75px">Min</th>
		<th style="width: 75px">Max</th>
	</tr>
	<tr>
		<td colspan="2" class="style2">
		<select name="parent_rule_1" id="parent_rule_1" style="width: 275px">
		<option value="0">No rule has been selected</option>
		</select></td>
		<td>&nbsp;</td>
		<td style="width: 75px" class="style2">
		<input name="Text2" class="input_min" type="text" /></td>
		<td class="style2">
		<input name="Text1" type="text" class="input_max" style="width: 75px" /></td>
		<td>&nbsp;</td>
		<td>
		<input name="Text3" class="input_min" type="text" /></td>
		<td>
		<input name="Text7" type="text" style="width: 75px" class="input_max" /></td>
		<td>&nbsp;</td>
		<td>
		<input name="Text4" class="input_min" type="text" /></td>
		<td>
		<input name="Text8" type="text" style="width: 75px" class="input_max" /></td>
		<td>&nbsp;</td>
		<td>
		<input name="Text5" class="input_min" type="text" /></td>
		<td>
		<input name="Text9" type="text" style="width: 75px" class="input_max" /></td>
		<td>&nbsp;</td>
		<td>
		<input name="Text6" class="input_min" type="text" /></td>
		<td>
		<input name="Text10" type="text" style="width: 75px" class="input_max" /></td>
		<td>&nbsp;</td>
		<td>
		<input name="Text11" class="input_min" type="text" /></td>
		<td>
		<input name="Text14" type="text" style="width: 75px" class="input_max" /></td>
		<td>&nbsp;</td>
		<td>
		<input name="Text12" class="input_min" type="text" /></td>
		<td>
		<input name="Text15" type="text" style="width: 75px" class="input_max" /></td>
		<td>&nbsp;</td>
		<td>
		<input name="Text13" class="input_min" type="text" /></td>
		<td>
		<input name="Text16" type="text" style="width: 75px" class="input_max" /></td>
	</tr>
	<tr>
		<td class="td_label_rt" style="width: 120px">If true:</td>
		<td colspan="2"><select name="true_rule_1" style="width: 275px">
		<option value="0">No rule has been selected</option>

		</select></td>
		<td style="width: 75px">
		<input name="Text17" class="input_off" type="text" disabled="disabled" /></td>
		<td>
		<input name="Text18" class="input_off" type="text" disabled="disabled" /></td>
		<td>&nbsp;</td>
		<td>
		<input name="Text21" class="input_off" type="text" disabled="disabled" /></td>
		<td>
		<input name="Text22" class="input_off" type="text" disabled="disabled" /></td>
		<td>&nbsp;</td>
		<td>
		<input name="Text25" class="input_off" type="text" disabled="disabled" /></td>
		<td>
		<input name="Text28" class="input_off" type="text" disabled="disabled" /></td>
		<td>&nbsp;</td>
		<td>
		<input name="Text29" class="input_off" type="text" disabled="disabled" /></td>
		<td>
		<input name="Text30" class="input_off" type="text" disabled="disabled" /></td>
		<td>&nbsp;</td>
		<td>
		<input name="Text33" class="input_off" type="text" disabled="disabled" /></td>
		<td>
		<input name="Text34" class="input_off" type="text" disabled="disabled" /></td>
		<td>&nbsp;</td>
		<td>
		<input name="Text37" class="input_off" type="text" disabled="disabled" /></td>
		<td>
		<input name="Text38" class="input_off" type="text" disabled="disabled" /></td>
		<td>&nbsp;</td>
		<td>
		<input name="Text41" class="input_off" type="text" disabled="disabled" /></td>
		<td>
		<input name="Text43" class="input_off" type="text" disabled="disabled" /></td>
		<td>&nbsp;</td>
		<td>
		<input name="Text45" class="input_off" type="text" disabled="disabled" /></td>
		<td>
		<input name="Text47" class="input_off" type="text" disabled="disabled" /></td>
	</tr>
	<tr>
		<td class="td_label_rt" style="width: 120px">If false:</td>
		<td colspan="2"><select name="false_rule_1" style="width: 275px">
		<option value="0">No rule has been selected</option>
		</select></td>
		<td style="width: 75px">
		<input name="Text20" class="input_off" type="text" disabled="disabled" /></td>
		<td>
		<input name="Text19" class="input_off" type="text" disabled="disabled" /></td>
		<td>&nbsp;</td>
		<td>
		<input name="Text24" class="input_off" type="text" disabled="disabled" /></td>
		<td>
		<input name="Text23" class="input_off" type="text" disabled="disabled" /></td>
		<td>&nbsp;</td>
		<td>
		<input name="Text26" class="input_off" type="text" disabled="disabled" /></td>
		<td>
		<input name="Text27" class="input_off" type="text" disabled="disabled" /></td>
		<td>&nbsp;</td>
		<td>
		<input name="Text32" class="input_off" type="text" disabled="disabled" /></td>
		<td>
		<input name="Text31" class="input_off" type="text" disabled="disabled" /></td>
		<td>&nbsp;</td>
		<td>
		<input name="Text36" class="input_off" type="text" disabled="disabled" /></td>
		<td>
		<input name="Text35" class="input_off" type="text" disabled="disabled" /></td>
		<td>&nbsp;</td>
		<td>
		<input name="Text40" class="input_off" type="text" disabled="disabled" /></td>
		<td>
		<input name="Text39" class="input_off" type="text" disabled="disabled" /></td>
		<td>&nbsp;</td>
		<td>
		<input name="Text42" class="input_off" type="text" disabled="disabled" /></td>
		<td>
		<input name="Text44" class="input_off" type="text" disabled="disabled" /></td>
		<td>&nbsp;</td>
		<td>
		<input name="Text46" class="input_off" type="text" disabled="disabled" /></td>
		<td>
		<input name="Text48" class="input_off" type="text" disabled="disabled" /></td>
	</tr>
	
<%
	On Error Resume Next
	

	If err.number = 0 Then
	   pad="&nbsp;&nbsp;&nbsp;&nbsp;"
	   response.write "<b>An error cccured while collecting data source information</b><br>"
	   response.write pad & "Error Number= #<b>" & err.number & "</b><br>"
	   response.write pad & "Error Desc.= <b>" & err.description & "</b><br>"
	   response.write pad & "Help Context= <b>" & err.HelpContext & "</b><br>"
	   response.write pad & "Help File Path=<b>" & err.helpfile & "</b><br>"
	   response.write pad & "Error Source= <b>" & err.source & "</b><br><hr>"

	End If
%>	
	<tr>
		<td style="width: 120px">&nbsp;</td>
		<td style="width: 170px">&nbsp;</td>
		<td>&nbsp;</td>
		<td style="width: 75px">&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	</tr>
	<tr>
		<td style="width: 120px">&nbsp;</td>
		<td style="width: 170px">&nbsp;</td>
		<td>&nbsp;</td>
		<td style="width: 75px">&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	</tr>
	<tr>
		<td style="width: 120px">&nbsp;</td>
		<td style="width: 170px">&nbsp;</td>
		<td>&nbsp;</td>
		<td style="width: 75px">&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	</tr>
	<tr>
		<td style="width: 120px">&nbsp;</td>
		<td style="width: 170px">&nbsp;</td>
		<td>&nbsp;</td>
		<td style="width: 75px">&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	</tr>
	<tr>
		<td style="width: 120px">&nbsp;</td>
		<td style="width: 170px">&nbsp;</td>
		<td>&nbsp;</td>
		<td style="width: 75px">&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	</tr>
</table>
</form>
<%
	On Error Resume Next
	

	If err.number = 0 Then
	   pad="&nbsp;&nbsp;&nbsp;&nbsp;"
	   response.write "<b>An error cccured while collecting data source information</b><br>"
	   response.write pad & "Error Number= #<b>" & err.number & "</b><br>"
	   response.write pad & "Error Desc.= <b>" & err.description & "</b><br>"
	   response.write pad & "Help Context= <b>" & err.HelpContext & "</b><br>"
	   response.write pad & "Help File Path=<b>" & err.helpfile & "</b><br>"
	   response.write pad & "Error Source= <b>" & err.source & "</b><br><hr>"

	End If
%>	


</body>

</html>
