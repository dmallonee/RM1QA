<%@ Language=VBScript %>
<!-- #INCLUDE FILE="inc/adovbs.asp" -->
<% Response.Expires = -1
   Response.cachecontrol="private" 
   Response.AddHeader "pragma", "no-cache"
   Response.Buffer = True
   
   Server.ScriptTimeout = 180


	'On Error Resume Next

	Dim strConn	
	Dim adoCmd	
	Dim adoRS
	Dim adoCmd1	
	Dim adoRS1
	Dim adoCmd2	
	Dim adoRS2
	Dim adoCmd3	
	Dim adoRS3
	Dim adoCmd4
	Dim adoRS4


	Dim adoPrices
	Dim strUserId
	Dim intRuleId
	Dim intSearchId
	Dim strAlertDesc
	Dim datBeginDate


	intSearchId = Request("txt_report")
	If CStr(intSearchId) = "" Then
		intSearchId = Request("recent_searches")
	End If

	strClassCode = TRIM(Request("classcode"))
	strStartDate = Request("startdate")
	strRateSystem = Request("ratesystem")

	strConn = Session("pro_con")
	
	If IsNumeric(intSearchId) Then
	
		If intSearchId  > 0 Then

			strConn = Session("pro_con")
	
			Set adoRS = CreateObject("ADODB.Recordset")
			Set adoCmd = CreateObject("ADODB.Command")
	
			adoCmd.ActiveConnection = strConn
			adoCmd.CommandText = "car_rate_change_queue_tsd_rezcentral_select_rpt"
			adoCmd.CommandType = 4

			adoCmd.CommandTimeout = 0
	
			adoCmd.Parameters.Append adoCmd.CreateParameter("@shop_request_id", 3, 1,  0, intSearchId) 'Request("reportrequestid"))
			
			If strClassCode = "" Then
				adoCmd.Parameters.Append adoCmd.CreateParameter("@ClassCode",     200, 1,  4, Null )
			Else
				adoCmd.Parameters.Append adoCmd.CreateParameter("@ClassCode",     200, 1,  4, strClassCode )
			End If
			
			If IsDate(strStartDate) Then
				adoCmd.Parameters.Append adoCmd.CreateParameter("@StartDate",     135, 1,  0, strStartDate )
			Else
				adoCmd.Parameters.Append adoCmd.CreateParameter("@StartDate",     135, 1,  0, Null )
			End If
			
			adoCmd.Parameters.Append adoCmd.CreateParameter("@RateSystem",    200, 1, 20, strRateSystem )			

			Set adoRS = adoCmd.Execute

			If err.number <> 0 Then
			   pad="&nbsp;&nbsp;&nbsp;&nbsp;"
			   response.write "<b>VBScript Errors during rule change select<br>"
			   response.write "</b><br>"
			   response.write pad & "Error Number= #<b>" & err.number & "</b><br>"
			   response.write pad & "Error Desc.= <b>" & err.description & "</b><br>"
			   response.write pad & "Help Context= <b>" & err.HelpContext & "</b><br>"
			   response.write pad & "Help File Path=<b>" & err.helpfile & "</b><br>"
			   response.write pad & "Error Source= <b>" & err.source & "</b><br><hr>"

			End If


		Else
	
			Set adoRS = CreateObject("ADODB.Recordset")

		End If

		
	End If

	
	If err.number <> 0 Then
	   pad="&nbsp;&nbsp;&nbsp;&nbsp;"
	   response.write "<b>VBScript Errors Occured!<br>"
	   response.write "</b><br>"
	   response.write pad & "Error Number= #<b>" & err.number & "</b><br>"
	   response.write pad & "Error Desc.= <b>" & err.description & "</b><br>"
	   response.write pad & "Help Context= <b>" & err.HelpContext & "</b><br>"
	   response.write pad & "Help File Path=<b>" & err.helpfile & "</b><br>"
	   response.write pad & "Error Source= <b>" & err.source & "</b><br><hr>"

	End If
	

%>



<!doctype HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>

<head>
<meta name="GENERATOR" content="Microsoft FrontPage 12.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Language" content="en-us">
<title>Rate-Monitor by Rate-Highway, Inc. | RezCentral Rate Update Status</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<style type="text/css" >
<!--
.profile_header {height="48" text-align:left; padding-left:3; padding-right:3; padding-top:0; background-color:#879AA2; font-family:Verdana, Arial, Helvetica, sans-serif; font-size:10pt; vertical-align:bottom; text-align:left }
.profile_light { height="68" text-align:left; padding-left:3; padding-right:3; padding-top:3; background-color:#CFD7DB; font-family:Verdana, Arial, Helvetica, sans-serif; font-size:10pt; vertical-align:bottom; text-align:left }
.profile_dark { height="48" text-align:left; padding-left:3; padding-right:3; padding-top:3; background-color:#B2BEC4; font-family:Verdana, Arial, Helvetica, sans-serif; font-size:10pt; vertical-align:bottom; text-align:left }
.style2 {
	text-align: center;
}
-->
</style>
<link rel="stylesheet" type="text/css" href="inc/rh_standard.css">
<link rel="stylesheet" type="text/css" href="inc/rh_report.css">
<link rel="stylesheet" type="text/css" href="inc/sitewide.css">
<script language="Javascript" type="text/javascript" src="inc/sitewide.js"></script>
<script language="Javascript" type="text/javascript" src="inc/header_menu_support.js" ></script>
</head>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" >
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td background="images/top_tile.gif">
    <img src="images/top.jpg" width="770" height="91"></td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td background="images/b_tile.gif">
    <table width="400" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td><img alt='' src="images/b_left.jpg"></td>
          <td>
          <img alt='' src="images/blanks/b_search_pro_of.gif" name="s1" border="0" id="s1" width="183" height="32"></td>
          <td>
          <img alt='' src="images/blanks/b_search_que_of.gif" name="s2" border="0" id="s2" width="97" height="32"></td>
          <td>
          <img alt='' src="images/blanks/b_search_cri_of.gif" name="s3" border="0" id="s3"></td>
          <td>
          <img alt='' src="images/blanks/b_rate_of.gif" name="ra" border="0" id="ra" width="88" height="32"></td>
          <td>
          <img alt='' src="images/blanks/b_alert_of.gif" name="al" border="0" id="al" width="53" height="32"></td>
          <td>
          <img alt='' src="images/blanks/b_user_of.gif" name="us" border="0" id="us" width="126" height="32"></td>
          <td>
          <img alt='' src="images/blanks/b_system_of.gif" name="sy" border="0" id="sy" width="58" height="32"></td>
        </tr>
      </table>

    </td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td background="images/med_bar_tile.gif">
    <img src="images/med_bar.gif" width="12" height="8"></td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td background="images/user_tile.gif">
    <table width="100" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td><img src="images/user_left.gif" width="580" height="31"></td>
        <td background="images/user_tile.gif">
        <table width="100" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td valign="bottom">
            <table width="100" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td align="right">
                <div align="right">
                  <font face="Vendana, Arial, Helvetica, sans-serif" size="2" color="#FFFFFF">
                  User: N/A</font></div>
                </td>
              </tr>
              <tr>
                <td><img src="images/separator.gif" width="183" height="6"></td>
              </tr>
            </table>
            </td>
            <td><img src="images/user_tile.gif" width="7" height="31"></td>
          </tr>
        </table>
        </td>
      </tr>
    </table>
    </td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td background="images/h_tile.gif">
    <table width="100" border="0" cellspacing="0" cellpadding="0" id="table1">
      <tr>
        <td><img src="images/h_system.gif" width="368" height="31"></td>
        <td><img src="images/h_right.gif" width="402" height="31"></td>
      </tr>
    </table>
    </td>
  </tr>
</table>

<p class="style2">&nbsp;&nbsp;&nbsp; <br>
&nbsp;<strong> 
</strong></p>

<div align="center">
<table cellpadding="0" cellspacing="0" border="0" width="1310" bgcolor="#FFFFFF">
<tr height="1">
<td >&nbsp;</td>
<td >&nbsp;</td>
<td >&nbsp;</td>
</tr>
<tr height="1">
<td bgcolor="#000000" colspan="1" height="1"><img src="images/pixel.gif" width="1" height="1"></td>
<td bgcolor="#000000" colspan="1" height="1"><img src="images/pixel.gif" width="1" height="1"></td>
<td bgcolor="#000000" colspan="1" height="1"><img src="images/pixel.gif" width="1" height="1"></td>
</tr>
</table>
</div>
<table cellpadding="0" cellspacing="0" border="0" ALIGN="CENTER" width="100%" bgcolor="#FFFFFF">
<tr >
<td  width="1" bgcolor="#000000"><img src="images/pixel.gif" width="1" height="1"></td>
<td colspan=3 bgcolor="#D9DEE1">
<table border="0" cellspacing="5" cellpadding="5">
<tr><td>
<font color="#080000">
&nbsp;
<form method="post" name="status" class="search">
<table border="0" style="border-collapse: collapse" bordercolor="#FFFFFF" width="100%" cellspacing="0" height="4">
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
 
  <table width="745" border="0" cellspacing="0" cellpadding="0" background="images/alt_color.gif">
    <tr>
      <td>
      <table width="1108" border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" background="images/alt_color.gif">
        <tr valign="bottom">
          <td width="10" height="51">&nbsp;</td>
          <td width="179" height="51">
          &nbsp;</td>
          <td width="583" colspan="3" height="51">
          <font size="2">Directions: To display the RezCentral 
			update status of a report, enter the reports number and enter the car class and/or start 
			date for limiting the results (Displays only the first 100 results).<br>
			</font></td>
          <td width="336" height="51">&nbsp;</td>
        </tr>
        <tr valign="bottom">
          <td width="10" height="26">&nbsp;</td>
          <td width="179" height="26">&nbsp;</td>
          <td width="177" height="26">
          <font size="2">Report # to use</font><font face="Verdana, Arial, Helvetica, sans-serif" size="2">:</font> </td>
          <td width="80" height="26">
<font color="#080000">
		  <input name="txt_report" type="text" value="<%=intSearchId %>" style="text-align:right"></font></td>
          <td width="662" colspan="2" height="26">
          <font size="2" face="Verdana, Arial, Helvetica, sans-serif">
          <input name="search" type="submit" id="Open2224" value="    Display   " class="rh_button"></font></td>
        </tr>
 <tr valign="bottom">
          <td width="10" height="22">&nbsp;</td>
          <td width="179" height="22">&nbsp;</td>
          <td width="177" height="22">
<font color="#080000">
          <font size="2">System</font><font size="2" face="Verdana, Arial, Helvetica, sans-serif">:</font></font></td>
          <td width="80" height="22">
<select name="ratesystem" style="width: 129px; border:1px solid #000000; background-color:#FF9933">
<option selected="">WSPAN</option>
<option>Sabre</option>
<option>Amadeus</option>
<option>CarRentals</option>
<option>Counter</option>
<option>Galileo</option>
<option>WebLink</option>
<option>RezCentral</option>
</select></td>
          <td width="662" colspan="2" height="22">&nbsp;</td>
        </tr>
        <tr valign="bottom">
          <td width="10" height="22">&nbsp;</td>
          <td width="179" height="22">&nbsp;</td>
          <td width="177" height="22">
          <font size="2">Optional - Car class</font><font size="2" face="Verdana, Arial, Helvetica, sans-serif">:</font></td>
          <td width="80" height="22">
          <input name="classcode" type="text" value="<%=strClassCode %>" ></td>
          <td width="662" colspan="2" height="22">&nbsp;</td>
        </tr>
        <tr valign="bottom">
          <td width="10" height="22">&nbsp;</td>
          <td width="179" height="22">&nbsp;</td>
          <td width="177" height="22"><font color="#080000">
          <font size="2">Optional - Start date</font><font size="2" face="Verdana, Arial, Helvetica, sans-serif">:</font></font></td>
          <td width="80" height="22"><input name="startdate" type="text" value="<%=strStartDate %>" ></td>
          <td width="662" colspan="2" height="22">&nbsp;</td>
        </tr>
        <tr valign="bottom">
          <td width="10" height="22">&nbsp;</td>
          <td width="179" height="22">&nbsp;</td>
          <td width="177" height="22">&nbsp;</td>
          <td width="80" height="22">
&nbsp;</td>
          <td width="662" colspan="2" height="22">&nbsp;</td>
        </tr>
      </table>
      </td>
    </tr>
  </table>
  </form>
 <table width="1300" border="0" cellpadding="2" background="images/alt_color.gif" style="border-collapse: collapse" bordercolor="#111111" id="table1">
    <tr valign="bottom">
      <td >&nbsp;RezCentral Queue Status</td>
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
  <input type="hidden" name="action" value="1">
  <input type="hidden" name="refresh_from" value="search">
  <table border="1" bordercolor="#FFFFFF" id="table3" width="1300" cellspacing="0" cellpadding="0" >
    <tr>
      <th align="left" valign="bottom" bgcolor="#879AA2"><font size="2">Id</font></th>
      <th class="profile_header"><font size="2">Branch</font></th>
      <th class="profile_header"><font size="2">Car</font></th>
      <th class="profile_header"><font size="2">Rate Code</font></th>
      <th class="profile_header"><font size="2">Date</font></th>
      <th class="profile_header"><font size="2">Rate Amt.</font></th>
      <th class="profile_header"><font size="2">xDay Amt.</font></th>
      <th class="profile_header"><font size="2">System</font></th>
      <th class="profile_header"><font size="2">Created (PST)</font></th>
      <th class="profile_header"><font size="2">Uploaded to TSD</font></th>
      <th class="profile_header"><font size="2">Error</font></th>
   </tr>
    
 <%
        
        Dim strClass
        Dim strOrange
        Dim intCount
        
        intCount = 0
        
        If adoRS Is Nothing Then
		%>
		
		Nothing
		
		<%

		ElseIf (adoRS.State <> adStateOpen) Then
		%>
		
		Closed State = <%=adoRS.State %>
		
		<%
		
		Else

		While (adoRS.EOF = False) And (intCount < 100)
		
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
	<%=adoRS.Fields("RemoteID").Value%></td>
    <td class="<%=strClass %>" width="50">
    <%=adoRS.Fields("branch").Value %>
    </td>
     <td class="<%=strClass %>" width="50" >
	<%=adoRS.Fields("classcode").Value %>
	</td>
     <td class="<%=strClass %>" width="50" >
	<%=adoRS.Fields("ratecode").Value %>
	</td>
     <td class="<%=strClass %>" width="70" >
	<%=adoRS.Fields("startdate").Value  %>
	</td>
    <td class="<%=strClass %>_right" align="right" width="74">
	<%=FormatCurrency(adoRS.Fields("rateamt").Value) %>
	</td>
    <td class="<%=strClass %>_right" align="right" width="74">
	<%=FormatCurrency(adoRS.Fields("extradayrate").Value) %>
	</td>
    <td class="<%=strClass %>" width="70" >
	<%=adoRS.Fields("ratesystem").Value %>
	</td>
     <td class="<%=strClass %>" width="160"  >
	<%=adoRS.Fields("inserted").Value  %>
	</td>
     <td class="<%=strClass %>" width="160"  >
	<%=adoRS.Fields("uploaded").Value & "" %>
	</td>
      <td class="<%=strClass %>"  >
	<%=adoRS.Fields("error").Value  & "" %>
	</td>
  
    </tr>

<%	
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
  </table>
  <table border="0" style="border-collapse: collapse" bordercolor="#FFFFFF" width="1210" height="4">
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
  <input type="hidden" name="refresh_from" value="create">
  <input type="hidden" name="rule_status" value="E">

<!-- Content goes before this comment -->
<!-- JUSTTABS BOTTOM OPEN -->
</font></td></tr></table>
</td>
<td  width="1" bgcolor="#000000"><img src="images/pixel.gif" width="1" height="1"></td>
</tr>
<tr bgcolor="#000000" height="1">
<td colspan=5><img src="images/pixel.gif" width="1" height="1"></td>
</tr>
</table>
<!-- JUSTTABS BOTTOM CLOSE -->
<font size="2" face="Verdana, Arial, Helvetica, sans-serif">
<p align="center">&nbsp;</p>
</font>
<!--#INCLUDE FILE="footer.asp"-->  
</body>
</html>
