<%@ Language=VBScript %>
<!-- #INCLUDE FILE="inc/adovbs.asp" --> 
<!-- #INCLUDE FILE="inc/login_check.asp" -->
<% Response.Expires = -1
   Response.cachecontrol="private" 
   Response.AddHeader "pragma", "no-cache"

   Server.ScriptTimeout = 30

	Dim strConn	
	Dim adoCmd	
	Dim adoRS
	
	On Error Resume Next
	
	strUserId = Request.Cookies("rate-monitor.com")("user_id")
	
	
	strConn = Session("pro_con")	
	
	Rem Get the data sources
	Set adoCmd = CreateObject("ADODB.Command")

	adoCmd.ActiveConnection =  strConn
	adoCmd.CommandText = "user_city_select"
	adoCmd.CommandType = adCmdStoredProc

	adoCmd.Parameters.Append adoCmd.CreateParameter("@user_id",  3, 1,  0, strUserId)
		
	Set adoRS1 = adoCmd.Execute
	
	strCityCd    = Request("city_cd")
	strLOR       = Request("LOR")
	strStartDate = Request("start_date")
	strEndDate   = Request("end_date")
	strCarTypeCd = TRIM(Request("car_type_cd"))
	strHour      = Request("hour")
	
	
	If IsNumeric(strHour) = False Then
		strHour = 0
	End If
	
	If strCityCd <> "" Then	
	
		strConn = Session("pro_con")
		
		Set adoCmd = CreateObject("ADODB.Command")
	
		adoCmd.ActiveConnection = strConn
		adoCmd.CommandText = "car_rental_transaction_summary_select"
		adoCmd.CommandType = 4
		
		adoCmd.Parameters.Append adoCmd.CreateParameter("@city_cd",     200, 1, 6, strCityCd )
		adoCmd.Parameters.Append adoCmd.CreateParameter("@car_type_cd", 200, 1, 4, strCarTypeCd)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@lor",           3, 1, 0, strLOR)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@user_id",       3, 1, 0, strUserId)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@start_dt",    135, 1, 0, strStartDate)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@end_dt",      135, 1, 0, strEndDate)

		If strHour = 0 Then
			adoCmd.Parameters.Append adoCmd.CreateParameter("@hour",          17, 1, 0, Null)
		Else
			adoCmd.Parameters.Append adoCmd.CreateParameter("@hour",          17, 1, 0, CInt(strHour))
		End If
		
		Set adoRS = adoCmd.Execute
		
		If err.number <> 0 Then
		   pad="&nbsp;&nbsp;&nbsp;&nbsp;"
		   response.write "<b>An error cccured while collecting transaction information</b><br>"
		   response.write pad & "Error Number= #<b>" & err.number & "</b><br>"
		   response.write pad & "Error Desc.= <b>" & err.description & "</b><br>"
		   response.write pad & "Help Context= <b>" & err.HelpContext & "</b><br>"
		   response.write pad & "Help File Path=<b>" & err.helpfile & "</b><br>"
		   response.write pad & "Error Source= <b>" & err.source & "</b><br><hr>"
	
		End If
	
	Else
	
	  Set adoRS = CreateObject("ADODB.Recordset")


	End If
	
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<title>Utilization Report</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<style type="text/css">
.style1 {
	font-size: xx-small;
}
</style>
</head>

<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" style="font-family: Tahoma" >

<p>&nbsp;</p>
<p align="center">Utilization Report</p>
<div align="center">

<form method="GET" name="utilization_report">
	<table border="0" width="300" id="table2">
		<tr>
			<td><font size="2">City</font></td>
			<td>&nbsp;</td>
			<td><select size="1" name="city_cd">
					
					
				    <% While (adoRS1.EOF = False) 
		                If adoRS1.Fields("city_cd").Value = strCityCd Then %>
		                  <option selected ><%=adoRS1.Fields("city_cd").Value %></option>		           
		           <%   Else %>	 
		                  <option ><%=adoRS1.Fields("city_cd").Value %></option>
		           <%   End If %>
		 		   <%   adoRS1.MoveNext %>
		           <% Wend %>          
		          
		        </select></td>
		</tr>
		<tr>
			<td><font size="2">LOR</font></td>
			<td>&nbsp;</td>
			<td>
			<% If strLOR = "" Then %>
			<input type="text" name="LOR" size="3" value="2" style="text-align: right">
			<% Else %>
			<input type="text" name="LOR" size="3" value="<%=strLOR %>" style="text-align: right">
			<% End If %>
			</td>
		</tr>
		<tr>
			<td><font size="2">Start Date</font></td>
			<td>&nbsp;</td>
			<td>
			<% If strStartDate = "" Then %>
			<input type="text" name="start_date" size="10" value="MM/DD/YY">
			<% Else %>
			<input type="text" name="start_date" size="10" value="<%=strStartDate %>">
			<% End If %>
			
			</td>
		</tr>
		<tr>
			<td><font size="2">End Date</font></td>
			<td>&nbsp;</td>
			<td>
			<% If strEndDate = "" Then %>
			<input type="text" name="end_date" size="10" value="MM/DD/YY">
			<% Else %>
			<input type="text" name="end_date" size="10" value="<%=strEndDate %>">
			<% End If %>
			</td>
		</tr>
		<tr>
			<td><font size="2">Hour</font></td>
			<td>&nbsp;</td>
			<td>
			<select name="hour">
			<option selected="">Whole Day</option>
			<option value="24">Any hour</option>
			<option value="0">Midnight</option>
			<option value="1">1 am</option>
			<option value="2">2 am</option>
			<option value="3">3 am</option>
			<option value="4">4 am</option>
			<option value="5">5 am</option>
			<option value="6">6 am</option>
			<option value="7">7 am</option>
			<option value="8">8 am</option>
			<option value="9">9 am</option>
			<option value="10">10 am</option>
			<option value="11">12 am</option>
			<option value="12">Noon</option>
			<option value="13">1 pm</option>
			<option value="14">2 pm</option>
			<option value="15">3 pm</option>
			<option value="16">4 pm</option>
			<option value="17">5 pm</option>
			<option value="18">6 pm</option>
			<option value="19">7 pm</option>
			<option value="20">8 pm</option>
			<option value="21">9 pm</option>
			<option value="22">10 pm</option>
			<option value="23">11 pm</option>
			</select>
			</td>
		</tr>


		<tr>
			<td><font size="2">Car Type</font></td>
			<td>&nbsp;</td>
			<td><input type="text" name="car_type_cd" size="6" value="<%=strCarTypeCd %>"></td>
		</tr>


</table>
<p align="center"><input type="submit" value="Submit" name="B1"></p>
</form>
<p align="center">&nbsp;</p>
<table border="0" width="100" id="count">
	<tr>
		<td bgcolor="#000000" width="48"><font color="#FFFFFF">Status</font></td>
		<td bgcolor="#000000"><p align="center"><font color="#FFFFFF">Count</font></td>
	</tr>
		
		<% If strCityCd <> "" Then %>
			<% While (adoRS.EOF = False) %>
				<tr>
				<td width="48"><font size="2"><%=adoRS.Fields("status_code").Value %></font></td>
				<td align="center"><font size="2"><%=adoRS.Fields("count").Value %></font></td>
				</tr>
				<% adoRS.MoveNext %>
			<% Wend %>	
			<% Set adoRS = adoRS.NextRecordset %>
		<% End If %>
	</table>
<p align="center">&nbsp;</p>
<table border="0" width="150" id="detaily">
	<tr>
		<td bgcolor="#000000" width="48"><font color="#FFFFFF">Status</font></td>
		<td bgcolor="#000000"><p align="center"><font color="#FFFFFF">Res. No.</font></td>
	</tr>
		
		<% If strCityCd <> "" Then %>
			<% While (adoRS.EOF = False) %>
				<tr>
				<td width="48"><font size="2"><%=adoRS.Fields("status_code").Value %></font></td>
				<td align="left"><font size="2"><%=adoRS.Fields("res_number").Value %></font></td>
				</tr>
				<% adoRS.MoveNext %>
			<% Wend %>	
			<% Set adoRS = adoRS.NextRecordset %>
		<% End If %>
	</table>

<p align="center">&nbsp;</p>
<table border="0" width="300" id="detail">
	<tr>
		<td bgcolor="#000000" width="48"><font color="#FFFFFF">Status</font></td>
		<td bgcolor="#000000"><p align="center"><font color="#FFFFFF">Date</font></td>
		<td bgcolor="#000000"><p align="center"><font color="#FFFFFF">Time</font></td>
		<td bgcolor="#000000"><p align="center"><font color="#FFFFFF">LOR</font></td>
		<td bgcolor="#000000"><p align="center"><font color="#FFFFFF">Count</font></td>
	</tr>
		
		<% If strCityCd <> "" Then %>
			<% While (adoRS.EOF = False) %>
				<tr>
				<td width="48"><font size="2"><%=adoRS.Fields("status_code").Value %></font></td>
				<td align="center"><font size="2"><%=FormatDateTime(adoRS.Fields("arv_dt_day").Value & "/" & adoRS.Fields("arv_dt_month").Value & "/" & adoRS.Fields("arv_dt_year").Value, 2) %></font></td>
				<td align="center"><font size="2"><%=adoRS.Fields("arv_dt_hour").Value %></font></td>
				<td align="center"><font size="2"><%=adoRS.Fields("lor").Value %></font></td>
				<td align="center"><font size="2"><%=adoRS.Fields("count").Value %></font></td>
				</tr>
				<% adoRS.MoveNext %>
			<% Wend %>	
		<% End If %>
	</table>

<p>&nbsp;</p>
</div>
</body>
</html>
<%	
	Rem Clean up
	Set adoCmd = Nothing 
	Set adoRS = Nothing 
	
%>