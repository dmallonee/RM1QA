<%@ Language=VBScript %>
<!-- #INCLUDE FILE="inc/adovbs.asp" -->
<% Response.Expires = -1
   Response.cachecontrol="private" 
   Response.AddHeader "pragma", "no-cache"

   Server.ScriptTimeout = 60

	On Error Resume Next

	Dim strConn	
	Dim adoCmd	
	Dim adoRS
	Dim lngProfileId 

	lngProfileId = Request("profile_id")
	
	strConn = Session("pro_con")
	
	Rem Get the schedules

	Set adoCmd = CreateObject("ADODB.Command")

	adoCmd.ActiveConnection =  strConn
	adoCmd.CommandText = "car_shop_profile_schedule_select"
	adoCmd.CommandType = adCmdStoredProc
	
	adoCmd.Parameters.Append adoCmd.CreateParameter("@profile_id", 3, 1, 0, lngProfileId)
	
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
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Language" content="en-us">
<meta name="GENERATOR" content="Microsoft FrontPage 12.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Rate-Monitor by Rate-Highway, Inc. | Profile Search Schedule</title>
<link rel="stylesheet" type="text/css" href="inc/rh_standard.css">
<link rel="stylesheet" type="text/css" href="inc/rh_report.css">
<link rel="stylesheet" type="text/css" href="inc/sitewide.css">
<style type="text/css" >
<!--
P {
	COLOR: navy; FONT-FAMILY: Verdana, Arial, sans-serif
}
.data_cell   { width: 65; text-align: right; font-family: Tahoma; font-size: 10pt }
.header      { width: 65; text-align: center; background-color: #CFD7DB }
.copyright   { FONT-SIZE: 0.7em; TEXT-ALIGN: right }
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
<p align="right"><font class="copyright">Copyright (c) 2001-<%=Year(Now)%>, 
Rate-Highway, Inc. All Rights Reserved.</font></p>
<br>
&nbsp;[<a target="_self" href="profile_search_schedule_a.asp?profile_id=<%=lngProfileId %>">create or edit</a>] 
<strong>[view 
all schedules]</strong>&nbsp;
<form method="post" name="display_utilization" >
<p align="center"><font size="5" color="#384F5B">Search Schedules</font></p>
<div align="center">
        <table border="0" cellpadding="0" style="width: 600px;" bordercolor="#111111" class="style7">
          <tr>
           <td width="100%" class="boxtitle" colspan="7" style="height: 15px"><font size="2"><b>
           Directions:</b> You may review the search schedules listed below. To 
			delete a schedule please click the &quot;delete&quot; button to the right of 
			the search schedule you wish to delete. To add or change a schedule 
			please click the &quot;create or edit&quot; link on the upper left.</font><p>
           &nbsp;</p>
           </td>
           
          </tr>
          <tr>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           
          </tr>
          <tr>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           
          </tr>
          <tr>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           
          </tr>
          <tr>
            <td  class="boxtitle" style="height: 11px"><font size="2"><u>Item 
			No.</u></font></td>
            <td  class="boxtitle" style="height: 11px"><u><font size="2">
			Description</font></u></td>
            <td  class="boxtitle" style="height: 11px">&nbsp;</td>
            <td  class="boxtitle" style="height: 11px"><u><font size="2">Days of 
			Wk</font></u></td>
            <td  class="boxtitle" style="height: 11px"><u><font size="2">Time</font></u></td>
            <td  class="boxtitle" style="height: 11px"><u><font size="2">Status</font></u></td>
            <td  class="boxtitle" style="height: 11px">&nbsp;</td>
          </tr>
          <% 	Dim intCount 	%>
          <% 	intCount = 0	%>
          
          <% If adoRS.State = adStateOpen Then %>
          <% While (adoRS.EOF = False) %>
          <tr>
            <td class="boxtitle" style="width: 14%"><font size="2"><%=intCount + 1 %></font></td>
            <td colspan="2">
            <input type="text" name="description" readonly value="<%=adoRS.Fields("schedule_desc").Value %>" style="text-align: left; width: 160px;" size="40"></td>
            <td class="boxtitle" style="width: 14%">
            <input type="text" name="on_rent" size="20" readonly value="<%=adoRS.Fields("schedule_dow_list").Value %>" style="text-align: right; width: 80px;"></td>
            <td class="boxtitle" style="width: 14%">
            <input type="text" name="canceled" size="20" readonly value="<%=FormatDateTime(adoRS.Fields("schedule_dttm").Value, 4) %>" style="text-align: right; width: 80px;">            </td>
            <td class="boxtitle" style="width: 14%">
            <input type="text" name="returned" size="20" readonly value="<%=FormatDateTime(adoRS.Fields("updated").Value, 2) %>" style="text-align: right; width: 80px;"></td>
            <td class="boxtitle" style="width: 14%">
            <input type="button" name="delete" size="20" value="delete" ></td>
          </tr>
          <% 	intCount = intCount + 1	%>
          <%   adoRS.MoveNext         	%>
          <% Wend                     	%>
          <% End If 					%>
          </table>
        </div>
<p class="style11">&nbsp;&nbsp; 
          </p>
  </FORM>
<!--#INCLUDE FILE="footer.asp"-->
</body>
</html>