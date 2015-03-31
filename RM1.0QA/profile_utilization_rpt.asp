<%@ Language=VBScript %>
<!-- #INCLUDE FILE="inc/login_check.asp" -->
<!-- #INCLUDE FILE="inc/adovbs.asp" -->
<% Response.Expires = -1
   Response.cachecontrol="private" 
   Response.AddHeader "pragma", "no-cache"
   Response.Buffer = True
   
   Server.ScriptTimeout = 180


	On Error Resume Next

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


	strUserId = Request.Cookies("rate-monitor.com")("user_id")
	
	strProfileId = CLng(Request("profile_id"))
		
	strConn = Session("pro_con")
	
	If strProfileId > 0 Then
	
		Rem Get the data sources
		Set adoCmd = CreateObject("ADODB.Command")
	
		adoCmd.ActiveConnection =  strConn
		adoCmd.CommandText = "car_rate_rule_util_rpt"
		adoCmd.CommandType = adCmdStoredProc
		adoCmd.Parameters.Append adoCmd.CreateParameter("@user_id",    3, 1, 0, strUserId)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@profile_id", 3, 1, 0, strProfileId)
	
		Set adoRS = adoCmd.Execute

	
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
<title>Rate-Monitor by Rate-Highway, Inc. | Alerts! | Bulk Update</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<script language="Javascript" type="text/javascript" src="inc/sitewide.js"></script>
<script language="Javascript" type="text/javascript" src="inc/header_menu_support.js" ></script>
<script language="Javascript" type="text/javascript" src="inc/js_calendar_v2.js"></script>
<script language="Javascript" type="text/javascript" src="inc/validate2.js"></script>
<script language="JavaScript" type="text/javascript" src="inc/multiple_select_support.js"></script>
<script language="JavaScript" type="text/javascript" src="inc/multiple_select_support2.js"></script>
<style type="text/css" >
<!--
.profile_header { height="48" text-align:left; padding-left:3; padding-right:3; padding-top:0; background-color:#879AA2; font-family:Verdana, Arial, Helvetica, sans-serif; font-size:10pt; vertical-align:bottom; text-align:left }
.profile_light { height="68" text-align:left; padding-left:3; padding-right:3; padding-top:3; background-color:#CFD7DB; font-family:Verdana, Arial, Helvetica, sans-serif; font-size:10pt; vertical-align:bottom; text-align:left }
.profile_dark { height="48" text-align:left; padding-left:3; padding-right:3; padding-top:3; background-color:#B2BEC4; font-family:Verdana, Arial, Helvetica, sans-serif; font-size:10pt; vertical-align:bottom; text-align:left }
.style1 {
	text-align: center;
	color: black;
	font-size: small;
}
.style7 {
	border-collapse: collapse;
}
.style8 {
	border: 1px solid #C0C0C0;
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
                  User: <%=Request.Cookies("rate-monitor.com")("user_name")%></font></div>
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
        <td><img src="images/h_search_profiles.gif" width="368" height="31"></td>
        <td><img src="images/h_right.gif" width="402" height="31"></td>
      </tr>
    </table>
    </td>
  </tr>
</table>


<br>                
  <table border="0" style="width: 800px;" bordercolor="#FFFFFF" height="4" align="center" class="style7">
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
<p>&nbsp;<!-- 
	<p align="center">
		Peter - use this report for right now please =&gt;
	<a href="system_utilization_report.asp">utilization report</a></p>
	--><form method="post" name="display_utilization" >
<p class="style1">&nbsp; Rule Utilization Report - by Profile<p class="style1">
 <% If strProfileId > 0 Then %>
			<%=adoRS.Fields("Profile").Value %>
          <% End If 					%>
<div align="center">
        <table border="0" cellpadding="0" style="width: 1000px;" bordercolor="#111111" class="style7">
          <tr>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px"> 
          &nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           
          </tr>
          <tr>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px"><u><font size="2">Same</font></u></td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px"><u><font size="2">Next</font></u></td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px"><u><font size="2">2-4</font></u></td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px"><u><font size="2">5-7</font></u></td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px"><u><font size="2">8-14</font></u></td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px"><u><font size="2">15-30</font></u></td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px"><u><font size="2">31-50</font></u></td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px"><u><font size="2">51+</font></u></td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           
          </tr>
          <tr>
            <td  class="boxtitle" style="height: 11px"><u><font size="2">Rule</font></u></td>
            <td  class="boxtitle" style="height: 11px"><u><font size="2">Min</font></u></td>
            <td  class="boxtitle" style="height: 11px"><u><font size="2">Max</font></u></td>
            <td  class="boxtitle" style="height: 11px"><u><font size="2">Min</font></u></td>
            <td  class="boxtitle" style="height: 11px"><u><font size="2">Max</font></u></td>
            <td  class="boxtitle" style="height: 11px"><u><font size="2">Min</font></u></td>
            <td  class="boxtitle" style="height: 11px"><u><font size="2">Max</font></u></td>
            <td  class="boxtitle" style="height: 11px"><u><font size="2">Min</font></u></td>
            <td  class="boxtitle" style="height: 11px"><u><font size="2">Max</font></u></td>
            <td  class="boxtitle" style="height: 11px"><u><font size="2">Min</font></u></td>
            <td  class="boxtitle" style="height: 11px"><u><font size="2">Max</font></u></td>
            <td  class="boxtitle" style="height: 11px"><u><font size="2">Min</font></u></td>
            <td  class="boxtitle" style="height: 11px"><u><font size="2">Max</font></u></td>
            <td  class="boxtitle" style="height: 11px"><u><font size="2">Min</font></u></td>
            <td  class="boxtitle" style="height: 11px"><u><font size="2">Max</font></u></td>
            <td  class="boxtitle" style="height: 11px"><u><font size="2">Min</font></u></td>
            <td  class="boxtitle" style="height: 11px"><u><font size="2">Max</font></u></td>
 
          </tr>
          <% 	Dim intCount 	%>
          <% 	intCount = 0	%>
          
          <% If strProfileId > 0 Then %>

          <% If adoRS.State = adStateOpen Then %>
          <% While (adoRS.EOF = False) %>
          <tr>
            <td class="boxtitle" ><font size="2"><%=adoRS.Fields("Rule").Value %></font></td>
            <td class="style8" ><font size="2"><%=FormatPercent(adoRS.Fields("Same min").Value/100,0) %></font></td>
            <td class="style8" ><font size="2"><%=FormatPercent(adoRS.Fields("Same max").Value/100,0) %></font></td>
            <td class="style8" ><font size="2"><%=FormatPercent(adoRS.Fields("Next min").Value/100,0) %></font></td>
            <td class="style8" ><font size="2"><%=FormatPercent(adoRS.Fields("Next max").Value/100,0) %></font></td>
            <td class="style8" ><font size="2"><%=FormatPercent(adoRS.Fields("2-4 min").Value/100,0) %></font></td>
            <td class="style8" ><font size="2"><%=FormatPercent(adoRS.Fields("2-4 max").Value/100,0) %></font></td>
            <td class="style8" ><font size="2"><%=FormatPercent(adoRS.Fields("5-7 min").Value/100,0) %></font></td>
            <td class="style8" ><font size="2"><%=FormatPercent(adoRS.Fields("5-7 max").Value/100,0) %></font></td>
            <td class="style8" ><font size="2"><%=FormatPercent(adoRS.Fields("8-14 min").Value/100,0) %></font></td>
            <td class="style8" ><font size="2"><%=FormatPercent(adoRS.Fields("8-14 max").Value/100,0) %></font></td>
            <td class="style8" ><font size="2"><%=FormatPercent(adoRS.Fields("15-30 min").Value/100,0) %></font></td>
            <td class="style8" ><font size="2"><%=FormatPercent(adoRS.Fields("15-30 max").Value/100,0) %></font></td>
            <td class="style8" ><font size="2"><%=FormatPercent(adoRS.Fields("31-50 min").Value/100,0) %></font></td>
            <td class="style8" ><font size="2"><%=FormatPercent(adoRS.Fields("31-50 max").Value/100,0) %></font></td>
            <td class="style8" ><font size="2"><%=FormatPercent(adoRS.Fields("51+ min").Value/100,0) %></font></td>
            <td class="style8" ><font size="2"><%=FormatPercent(adoRS.Fields("51+ max").Value/100,0) %></font></td>


          </tr>
          <% 	intCount = intCount + 1	%>
          <%   adoRS.MoveNext         	%>
          <% Wend                     	%>
          <% End If 					%>
          <% End If 					%>
          
          </table>
		Total: <%=intCount %>
        </div>
        <p class="style11">&nbsp;&nbsp;&nbsp; 
          </p>
        <p align="center">
		&nbsp;</p>
  </FORM>

  <table border="0" style="width: 800px;" bordercolor="#FFFFFF" height="4" id="table2" align="center" class="style7">
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
<!--#INCLUDE FILE="footer.asp"-->
</body>
</html>
