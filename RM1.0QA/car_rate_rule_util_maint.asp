<%@ Language=VBScript %>
<!-- #INCLUDE FILE="inc/adovbs.asp" -->
<%  Response.Expires = -1
    Response.cachecontrol="private" 
    Response.AddHeader "pragma", "no-cache"

    Server.ScriptTimeout = 30

	Dim strUserId
	Dim strProfileId
	
	On Error Resume Next
	
	strUserId = Request.Cookies("rate-monitor.com")("user_id")
	
	strProfileId = CLng(Request("profile_id"))
		
	strConn = Session("pro_con")
	
	If strProfileId > 0 Then
	
		Rem Get the data sources
		Set adoCmd = CreateObject("ADODB.Command")
	
		adoCmd.ActiveConnection =  strConn
		adoCmd.CommandText = "car_rate_rule_select_util_maint"
		adoCmd.CommandType = adCmdStoredProc
		'adoCmd.Parameters.Append adoCmd.CreateParameter("@user_id",    3, 1, 0, strUserId)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@profile_id", 3, 1, 0, strProfileId)
	
		Set adoRS = adoCmd.Execute
		
		If adoRS.EOF = True Then
			blnNoRules = True
		Else
			varParents = adoRS.GetRows()
		
			Set adoRS = adoRS.NextRecordset
			
			If adoRS.EOF = True Then
				blnNoChildRules = True
			Else
				varFirstChildren = adoRS.GetRows()
				blnNoChildRules = False
			End If
			
			blnNoRules = False
			
		End If
		
		Set adoRS = Nothing
		Set adoCmd = Nothing
		
	End If	

	If err.number <> 0 Then
	   pad="&nbsp;&nbsp;&nbsp;&nbsp;"
	   response.write "<b>An error cccured while collecting rule information</b><br>"
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
<meta content="en-us" http-equiv="Content-Language" >
<meta content="text/html; charset=windows-1252" http-equiv="Content-Type" >
<title>Rate Monitor | Rule Utilization Maintenance</title>
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
.input_dark {
	font-size: small;
	font-family: Arial, Helvetica, sans-serif;
	text-align: right;
	width: 75px;
	background-color: #FFCC66;
	color: #000000;
}
.input_light {
	font-size: small;
	font-family: Arial, Helvetica, sans-serif;
	text-align: right;
	width: 75px;
	background-color: #FFCC66;
	color: #000000;
}
.input_off {
	font-size: small;
	font-family: Arial, Helvetica, sans-serif;
	text-align: right;
	width: 75px;
	background-color: #CCCCCC;
	color: #000000;
}

.input_closed {
	font-size: small;
	font-family: Arial, Helvetica, sans-serif;
	text-align: right;
	width: 75px;
	background-color: #000000;
	color: #000000;
}

.input_open {
	font-size: small;
	font-family: Arial, Helvetica, sans-serif;
	text-align: right;
	width: 75px;
	background-color: #FFFFFF;
	color: #000000;
}


.input_err_gap {
	font-size: small;
	font-family: Arial, Helvetica, sans-serif;
	text-align: right;
	width: 75px;
	background-color: #000080;
	color: #000000;
}
.input_err_overlap {
	font-size: small;
	font-family: Arial, Helvetica, sans-serif;
	text-align: right;
	width: 75px;
	background-color: #FF0000;
	color: #000000;
}


.style1 {
	border-bottom-style: solid;
	border-bottom-width: 1px;
}
.style2 {
	border-top-style: solid;
	border-top-width: 0px;
}

.style4 {
	border-color: #FFFFFF;
}

.style5 {
	text-align: center;
}

.input_no_rule {
	background-color: #CCFFFF;
}
.parent_rule {
	background-color: #FF9933;
}
.first_child_rule {
	background-color: #FFCC66;
}
.instructions {
	font-size: small;
}
</style>
<script type="text/javascript" language="javascript" >

function checkForInt(evt) {
	var charCode = ( evt.which ) ? evt.which : event.keyCode;
	return ( charCode >= 48 && charCode <= 57 );
}

function validateRange(sValue)
   {
      var s = sValue;
      var A = 0;
      var B = 999;
      
      alert(s);

      switch (isIntegerInRange(s, A, B))
      {
         case true:
            alert(s + " is in range from " + A + " to " + B)
            return true;
            break;
         case false:
            alert(s + " is not in range from " + A + " to " + B)
            return false;
      }
   }

// isIntegerInRange (STRING s, INTEGER a, INTEGER b)
function isIntegerInRange (s, a, b)
   {   
   
      alert(s);


   	  if (isEmpty(s))
         if (isIntegerInRange.arguments.length == 1) return false;
         else return (isIntegerInRange.arguments[1] == true);

      // Catch non-integer strings to avoid creating a NaN below,
      // which isn't available on JavaScript 1.0 for Windows.
      if (!isInteger(s, false)) return false;

      // Now, explicitly change the type to integer via parseInt
      // so that the comparison code below will work both on
      // JavaScript 1.2 (which typechecks in equality comparisons)
      // and JavaScript 1.1 and before (which doesn't).
      var num = parseInt (s);
      alert(s);

 
     return ((num >= a) && (num <= b));
   }

function isInteger (s)
   {
      var i;

      if (isEmpty(s))
      if (isInteger.arguments.length == 1) return 0;
      else return (isInteger.arguments[1] == true);

      for (i = 0; i < s.length; i++)
      {
         var c = s.charAt(i);

         if (!isDigit(c)) return false;
      }

      return true;
   }

function isEmpty(s)
   {
      return ((s == null) || (s.length == 0))
   }

</script>
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
 
    </td>
  </tr>
</table>

<p class="instructions">Directions: [DRAFT] This page is used to maintain the 
utilization levels for all of the rules for a single profile at the same time. 
To make changes, simply enter the utilization level for the rule and date range 
that requires updating. Please enter only numbers, as a percentage sign is not 
required. Once you have updated all the utilization levels that you want to 
change, click the &quot;Update&quot; button and the rules will be updated and the profile 
and it's rules re-displayed so that you can confirm that your changes were 
successful.&nbsp; </p>
<p class="instructions">WARNING: The Update button is active and will modify 
rules - do not use on production rules only test rules.</p>
<p><strong>Profile Name: <i><%=varParents(5, 1) %></i></strong></p>
<form name="rule_utilization" id="rules" action="car_rate_rule_util_maint_update.asp" method="post">
<table cellpadding="0" cellspacing="0" style="width: 1820" id="tbl_rules" class="style4">
	<tr>
		<th >&nbsp;</th>
		<th >&nbsp;</th>
		<th colspan="2">Same Day</th>
		<th >&nbsp;</th>
		<th colspan="2">Next Day</th>
		<th >&nbsp;</th>
		<th colspan="2">2 to 4 Days Out</th>
		<th >&nbsp;</th>
		<th colspan="2">5 to 7 Days Out</th>
		<th >&nbsp;</th>
		<th colspan="2">8 to 14 Days Out</th>
		<th >&nbsp;</th>
		<th colspan="2">15 to 30 Days Out</th>
		<th >&nbsp;</th>
		<th colspan="2">31 To 50 Days Out</th>
		<th >&nbsp;</th>
		<th colspan="2">51+ Days Out</th>
	</tr>
	<tr>
		<th style="width: 500px" class="style1">Rule Name</th>
		<th style="width: 15px">&nbsp;</th>
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

<%
	Dim intCount
	Dim strCellClass	
	Dim iRowLoop, iColLoop
	Dim iRowLoop2
	
	If blnNoRules = False Then

	For iRowLoop = 0 to UBound(varParents, 2)

%>
	<tr class="">
		<td class="style2" title="Rule Number <%=varParents(0, iRowLoop) %>" >
		<input type="text" name="parent_rule_<%=iRowLoop %>" id="parent_rule_<%=iRowLoop %>" value="<%=varParents(1, iRowLoop) %>" style="width: 275px" readonly="readonly" class="parent_rule"></td>
		<td>&nbsp;</td>

		
		<%	strCellValue = CStr(varParents(6, iRowLoop))	
			Select Case True	
				Case (strCellValue = "999")
					strCellClass = "input_closed"
				Case (strCellValue = "")
					strCellClass = "input_open"
				Case Else
					strCellClass = "input_dark"
			End Select
		%>
		<td>
			<input name="input_min_<%=iRowLoop %>" type="hidden" value="<%=varParents(0, iRowLoop) %>" >
			<input name="input_min_<%=iRowLoop %>" id="input_min_<%=iRowLoop %>" class="<%=strCellClass %>" type="text" value="<%=strCellValue %>" onkeypress="return checkForInt(event)">
		</td>
		<%	strCellValue = CStr(varParents(7, iRowLoop))	
			Select Case True	
				Case (strCellValue = "999")
					strCellClass = "input_closed"
				Case (strCellValue = "")
					strCellClass = "input_open"
				Case Else
					strCellClass = "input_dark"
			End Select
		%>
		<td>
			<input name="input_max_<%=iRowLoop %>" type="hidden" value="<%=varParents(0, iRowLoop) %>" >
			<input name="input_max_<%=iRowLoop %>" id="input_max_<%=iRowLoop %>" class="<%=strCellClass %>" type="text" value="<%=strCellValue %>" onkeypress="return checkForInt(event)" >
		</td>
		<td>&nbsp;</td>
		<% For intCount = 8 To 20 Step 2 %>
		<%	strCellValue = CStr(varParents(intCount, iRowLoop) & "")	
			Select Case True	
				Case (strCellValue = "999")
					strCellClass = "input_closed"
				Case (strCellValue = "")
					strCellClass = "input_open"
				Case Else
					strCellClass = "input_dark"
			End Select
		%>
		<td><input name="input_min_<%=iRowLoop %>" id="input_min_<%=iRowLoop %>" class="<%=strCellClass  %>" type="text" value="<%=strCellValue %>" onkeypress="return checkForInt(event)" ></td>
		<%	strCellValue = CStr(varParents(intCount + 1, iRowLoop) & "")	
			Select Case True	
				Case (strCellValue = "999")
					strCellClass = "input_closed"
				Case (strCellValue = "")
					strCellClass = "input_open"
				Case Else
					strCellClass = "input_dark"
			End Select
		%>
		<td><input name="input_max_<%=iRowLoop %>" id="input_max_<%=iRowLoop %>" class="<%=strCellClass  %>" type="text" value="<%=strCellValue %>" onkeypress="return checkForInt(event)" ></td>
		<td>&nbsp;</td>
		<% Next 'intCount                %>
	</tr>	

<%  

	blnTrueFound = False
	blnFalseFound = False

	If blnNoChildRules = False Then
	
		For iRowLoop2 = 0 to UBound(varFirstChildren, 2)
	
			Rem If the Rule = the childrens parent - use it
			If (varParents(0, iRowLoop) = varFirstChildren(2, iRowLoop2)) Then
				If varFirstChildren(3, iRowLoop2) Then 
					blnTrueFound = True
%>
	<tr>
		<td class="td_label_rt" >T:<input type="text" value="<%=varFirstChildren(1, iRowLoop2) %>" style="width: 250px" readonly="readonly" class="first_child_rule"></td>
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
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	</tr>
	
<%
			End If	
			
		End If
		
	Next 'iRowLoop2 	
%>	


<%	
	If blnTrueFound = False Then
%>

	<tr>
		<td class="td_label_rt" >T:<input type="text" value="none selected" style="width: 250px" readonly="readonly" class="input_no_rule"></td>
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
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	</tr>

<%
	End If
%>

			

<%
	blnFalseFound = False

	For iRowLoop2 = 0 to UBound(varFirstChildren, 2)

		Rem If the Rule = the childrens parent - use it
		If (varParents(0, iRowLoop) = varFirstChildren(2, iRowLoop2)) Then

			If Not varFirstChildren(3, iRowLoop2) Then 
				blnFalseFound = True
%>	
	<tr>
		<td class="td_label_rt" >F:<input type="text"  value="<%=varFirstChildren(1, iRowLoop2) %>" style="width: 250px" readonly="readonly" class="first_child_rule" ></td>
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
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	</tr>
<%
			End If	
			
		End If
		
	Next 'iRowLoop2 	

	If blnFalseFound = False Then
%>

	<tr>
		<td class="td_label_rt" >F:<input type="text" value="none selected" style="width: 250px" readonly="readonly" class="input_no_rule"></td>
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
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	</tr>

<%
	End If
	
	End If 

	Next 'iRowLoop

	End If

%>	
</table>
<br>
<br>
<div class="style5">
<input name="btn_update" type="submit" value="Update">&nbsp;
		<input name="btn_close" type="button" value="  Close  " onclick="window.close()">
</div>
	<input name="profile_id" type="hidden" value="<%=strProfileId %>">
</form>
<p>&nbsp;</p>
<table cellpadding="0" cellspacing="0" style="width: 392px">
	<tr>
		<td>&nbsp;</td>
		<td class="input_light">&nbsp;</td>
		<td>&nbsp;</td>
		<td class="instructions">Normal utilization</td>
	</tr>
	<tr>
		<td>&nbsp;</td>
		<td class="input_open">&nbsp;</td>
		<td>&nbsp;</td>
		<td class="instructions">Utilization is open</td>
	</tr>
	<tr>
		<td>&nbsp;</td>
		<td class="input_closed">&nbsp;</td>
		<td>&nbsp;</td>
		<td class="instructions">Utilization is closed</td>
	</tr>
	<tr>
		<td>&nbsp;</td>
		<td class="input_err_overlap">&nbsp;</td>
		<td>&nbsp;</td>
		<td class="instructions">Utilization overlaps with another rule</td>
	</tr>
	<tr>
		<td>&nbsp;</td>
		<td class="input_err_gap">&nbsp;</td>
		<td>&nbsp;</td>
		<td class="instructions">Utilization gap between rules</td>
	</tr>
	<tr>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	</tr>
	<tr>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	</tr>
	<tr>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	</tr>
</table>
</body>
</html>
