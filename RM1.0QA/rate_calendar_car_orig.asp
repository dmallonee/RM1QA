<%@ Language=VBScript %>
<!-- #INCLUDE FILE="inc/adovbs.asp" --> 
<!-- #INCLUDE FILE="inc/login_check.asp" -->
<% 	Response.Expires = -1  
   	Response.cachecontrol="private" 
   	Response.AddHeader "pragma", "no-cache" 
   
   	'On Error Resume Next

   	Server.ScriptTimeout = 180

	Dim strSelected 

    strUserId =    Request.Cookies("rate-monitor.com")("user_id")
	strCityCd =    Request.Form("city_cd")
	strCarTypeCd = Request.Form("car_type_cd")
	
	strConn = Session("pro_con")
	
	Rem Get the cities for this user
	Set adoCmd = CreateObject("ADODB.Command")

	adoCmd.ActiveConnection =  strConn
	adoCmd.CommandText = "user_city_select"
	adoCmd.CommandType = adCmdStoredProc

	adoCmd.Parameters.Append adoCmd.CreateParameter("@user_id",  3, 1,  0, strUserId)
		
	Set adoRS = adoCmd.Execute

	Rem Re-use the command and get the cars for this user
	adoCmd.CommandText = "car_type_select"
		
	Set adoRS1 = adoCmd.Execute
	
	'adoCmd.Close 8005463672 x21202
		
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
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Language" content="en-us">
<title>Rate-Monitor by Rate-Highway, Inc. | Rate Reports | Rate Calendar</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="VI60_defaultClientScript" content="JavaScript">
<link rel="stylesheet" type="text/css" href="inc/rh_standard.css">
<link rel="stylesheet" type="text/css" href="inc/rh_report.css">
<link rel="stylesheet" type="text/css" href="inc/sitewide.css">
<script src="inc/sitewide.js" language="javascript"></script>
<script src="inc/header_menu_support.js" language="javascript"></script>
<script language="javascript">
	function NewWindow(mypage, myname, w, h, scroll) {
		var winl = (screen.width - w) / 2;
		var wint = (screen.height - h) / 2;
		winprops = 'height='+h+',width='+w+',top='+wint+',left='+winl+',scrollbars='+scroll+',resizable=0'
		win = window.open(mypage, myname, winprops);
		win.window.focus();
	}
</script>
<style>
<!--
.off_day     { background-color: #879AA2 }
-->
</style>
</head>

<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onload="MM_preloadImages('images/b_search_pro_on.gif','images/b_search_que_on.gif','images/b_search_cri_on.gif','images/b_rate_on.gif','images/b_alert_on.gif','images/b_user_on.gif','images/b_system_on.gif')" style="font-family: Tahoma; font-size: 10pt">

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
                  User: <%=Request.Cookies("rate-monitor.com")("user_name")%> </font></div>
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
    <table width="100" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td><img src="images/h_rates_reports.gif" width="368" height="31"></td>
        <td><img src="images/h_right.gif" width="402" height="31"></td>
      </tr>
    </table>
    </td>
  </tr>
</table>
<!-- JUSTTABS TOP OPEN -->
<p>&nbsp;</p>
<table cellpadding="0" cellspacing="0" border="0" ALIGN="CENTER" width="100%" bgcolor="#FFFFFF">
<tr height="1">
<td colspan="1" width="10">&nbsp;</td>
<td rowspan="2" width="335"><a href="javascript:not_enabled()">
<img src="images/ratecalendar0_a.GIF" width="100" height="25" hspace="0" vspace="0" border="0" alt="Rate Calendar" description="Rate Calendar"></a><a href="rate_wizard_car.asp"><img src="images/ratewizards1_ia.GIF" width="97" height="25" hspace="0" vspace="0" border="0" alt="Rate Wizards" description="Rate Calendar"></a><a href="rate_mgmt_car.asp"><img src="images/ratemananagement2_ia.GIF" width="138" height="25" hspace="0" vspace="0" border="0" alt="New ALT Text" description="Rate Calendar"></a></td>
<td colspan="1" >&nbsp;</td>
</tr>
<tr height="1">
<td bgcolor="#000000" colspan="1" height="1"><img src=pixel.gif width="1" height="1"></td>
<td bgcolor="#000000" colspan="1" height="1"><img src=pixel.gif width="1" height="1"></td>
</tr>
</table>
<table cellpadding="0" cellspacing="0" border="0" ALIGN="CENTER" width="100%" bgcolor="#FFFFFF">
<tr >
<td  width="1" bgcolor="#000000"><img src=pixel.gif width="1" height="1"></td>
<td colspan=3 bgcolor="#CED7DB">
<table border="0" cellspacing="5" cellpadding="5">
<tr><td>
<font color="#080000">
<!-- JUSTTABS TOP CLOSE -->
    <form method="POST" action>
      <table border="1" cellpadding="0" style="border-collapse: collapse; border-top-width:0px" width="300" id="table3" bordercolor="#384F5B">
        <tr>
          <td colspan="2" bgcolor="#384F5B" style="border-left:1px solid #384F5B; border-right:1px solid #384F5B; border-top:1px solid #384F5B; "><font color="#FFFFFF">
          &nbsp;<span style="font-size: 11pt">Report Properties</span></font></td>
        </tr>
        <tr>
          <td width="96" style="border-left-style: solid; border-left-width: 1px; border-right-style:none; border-right-width:medium; border-top-style:none; border-top-width:medium; border-bottom-style:none; border-bottom-width:medium" bgcolor="#FFFFFF">
          &nbsp;</td>
          <td width="201" style="border-right:1px solid #384F5B; border-left-style:none; border-left-width:medium; border-top-style:none; border-top-width:medium; border-bottom-style:none; border-bottom-width:medium" bgcolor="#FFFFFF">
          &nbsp;</td>
        </tr>
        <tr>
          <td width="96" style="border-left-style: solid; border-left-width: 1px; border-right-style:none; border-right-width:medium; border-top-style:none; border-top-width:medium; border-bottom-style:none; border-bottom-width:medium" bgcolor="#FFFFFF">
          <font size="2">&nbsp;Location:</font></td>
          <td width="201" style="border-right:1px solid #384F5B; border-left-style:none; border-left-width:medium; border-top-style:none; border-top-width:medium; border-bottom-style:none; border-bottom-width:medium" bgcolor="#FFFFFF">
		  <select size="1" name="city_cd" id="grp_city_cd" style="width:75;" tabindex="1" >
		  <% While (adoRS.EOF = False) 
		     If adoRS.Fields("city_cd").Value = strCityCd Then %>
		       <option selected ><%=adoRS.Fields("city_cd").Value %></option>		           
	      <% Else %>	 
		       <option ><%=adoRS.Fields("city_cd").Value %></option>
          <% End If %>
  	      <% adoRS.MoveNext %>
          <% Wend %>
		  </select>
          </td>
        </tr>
        <tr>
          <td width="96" style="border-left-style: solid; border-left-width: 1px; border-right-style:none; border-right-width:medium; border-top-style:none; border-top-width:medium; border-bottom-style:none; border-bottom-width:medium" bgcolor="#FFFFFF"><font size="2">
          &nbsp;Car Type:</font></td>
          <td width="201" style="border-right:1px solid #384F5B; border-top-style: none; border-top-width: medium; border-left-style:none; border-left-width:medium; border-bottom-style:none; border-bottom-width:medium" bgcolor="#FFFFFF">
          <select size="1" name="car_type_cd" id="grp_car_class" style="width:75;" tabindex="2" onchange="update_grp_car_type();">
		  <% While (adoRS1.EOF = False) 
		       If adoRS1.Fields("car_type_cd").Value = strCarTypeCd Then %>
		          <option selected ><%=adoRS1.Fields("car_type_cd").Value %></option>		           
		  <%   Else %>	 
		          <option ><%=adoRS1.Fields("car_type_cd").Value %></option>
		  <%   End If %>
		  <%   adoRS1.MoveNext %>
		  <% Wend %>
          </select>
          </td>
        </tr>
        <tr>
          <td width="96" style="border-left-style: solid; border-left-width: 1px; border-right-style:none; border-right-width:medium; border-top-style:none; border-top-width:medium; border-bottom-style:none; border-bottom-width:medium" bgcolor="#FFFFFF"><font size="2">
          &nbsp;Length of Rent:</font></td>
          <td width="201" style="border-right:1px solid #384F5B; border-left-style:none; border-left-width:medium; border-top-style:none; border-top-width:medium; border-bottom-style:none; border-bottom-width:medium" bgcolor="#FFFFFF">
			<select name="lor" style="width: 60px">
			<option selected="">1</option>
			<option>2</option>
			<option>3</option>
			<option>4</option>
			<option>5</option>
			<option>6</option>
			<option>7</option>
			<option>8</option>
			<option>9</option>
			<option>10</option>
			</select></td>
        </tr>
        <tr>
          <td width="96" style="border-left-style: solid; border-left-width: 1px; border-right-style:none; border-right-width:medium; border-top-style:none; border-top-width:medium; border-bottom-style:none; border-bottom-width:medium" bgcolor="#FFFFFF">
          <font size="2">&nbsp;Compare To:</font></td>
          <td width="201" style="border-right:1px solid #384F5B; border-left-style:none; border-left-width:medium; border-top-style:none; border-top-width:medium; border-bottom-style:none; border-bottom-width:medium" bgcolor="#FFFFFF">
          <select size="1" name="vendor_cd">
          <option selected>(none)</option>
          <option value="AL">Alamo</option>
          <option>Avis</option>
          <option>Budget</option>
          <option>Dollar</option>
          <option>National</option>
          <option>Payless</option>
          <option>Thrifty</option>
          </select></td>
        </tr>
        <tr>
          <td width="96" style="border-left-style: solid; border-left-width: 1px; border-right-style:none; border-right-width:medium; border-top-style:none; border-top-width:medium; border-bottom-style:none; border-bottom-width:medium" bgcolor="#FFFFFF">&nbsp;</td>
          <td width="201" style="border-right:1px solid #384F5B; border-left-style:none; border-left-width:medium; border-top-style:none; border-top-width:medium; border-bottom-style:none; border-bottom-width:medium" bgcolor="#FFFFFF">&nbsp;</td>
        </tr>
        <tr>
          <td width="96" style="border-left-style: solid; border-left-width: 1px; border-right-style:none; border-right-width:medium; border-top-style:none; border-top-width:medium; border-bottom-style:none; border-bottom-width:medium" bgcolor="#FFFFFF">&nbsp;</td>
          <td width="201" style="border-right:1px solid #384F5B; border-left-style:none; border-left-width:medium; border-top-style:none; border-top-width:medium; border-bottom-style:none; border-bottom-width:medium" bgcolor="#FFFFFF">
          &nbsp;<button name="display" type="submit">Display</button></td>
        </tr>
        <tr>
          <td width="96" style="border-bottom:1px solid #384F5B; border-left-style: solid; border-left-width: 1px; border-right-style:none; border-right-width:medium; border-top-style:none; border-top-width:medium" bgcolor="#FFFFFF">&nbsp;</td>
          <td width="201" style="border-right:1px solid #384F5B; border-bottom-style: solid; border-bottom-width: 1px; border-left-style:none; border-left-width:medium; border-top-style:none; border-top-width:medium" bgcolor="#FFFFFF">&nbsp;</td>
        </tr>
      </table>
      <p>&nbsp;</p>
    </form>
    <form action="javascript:not_enabled()" name="calendar">
      <p></p>
      <p align="center">&nbsp;| <a href="javascript:not_enabled()">&lt;&lt;</a> |
      <a href="javascript:not_enabled()">&lt;</a> | <%=MonthName(Month(Now), False) & " " & Year(Now) %> |
      <a href="javascript:not_enabled()">&gt;</a> | <a href="javascript:not_enabled()">&gt;&gt;</a> 
      |</p>
      <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="100%" id="table1">
          <tr>
          <td bgcolor="#384F5B" align="center" width="14%"><font color="#FFFFFF">
          Sunday</font></b></td>
          <td bgcolor="#384F5B" align="center" width="14%"><font color="#FFFFFF">
          Monday</font></b></td>
          <td bgcolor="#384F5B" align="center" width="14%"><font color="#FFFFFF">
          Tuesday</font></b></td>
          <td bgcolor="#384F5B" align="center" width="14%"><font color="#FFFFFF">
          Wednesday</font></b></td>
          <td bgcolor="#384F5B" align="center" width="14%"><font color="#FFFFFF">
          Thursday</font></b></td>
          <td bgcolor="#384F5B" align="center" width="14%"><font color="#FFFFFF">
          Friday</font></b></td>
          <td bgcolor="#384F5B" align="center" width="14%"><font color="#FFFFFF">
          Saturday</font></b></td>
          </tr>
        <tr>
          <td class="off_day" >&nbsp;</td>
          <td>
      <table border="0" cellpadding="0" style="border-collapse: collapse" width="100%" id="table6">
            <tr>
              <td bgcolor="#CFD7DB">&nbsp;</td>
              <td align="right" bgcolor="#CFD7DB">1&nbsp;&nbsp; </td>
            </tr>
            <tr>
              <td align="right" bgcolor="#FFFFFF"><font size="2">D:</font></td>
              <td align="right" bgcolor="#FFFFFF"><font size="2"><a href="javascript:BlahBlah()">
              $123.45</a></font></td>
            </tr>
            <tr>
              <td align="right" bgcolor="#FFFFFF"><font size="2">WD:</font></td>
              <td align="right" bgcolor="#FFFFFF">&nbsp;</td>
            </tr>
            <tr>
              <td align="right" bgcolor="#FFFFFF"><font size="2">WKLY:</font></td>
              <td align="right" bgcolor="#FFFFFF"><font size="2">$456.67</font></td>
            </tr>
          </table>
          </td>
          <td>
          <table border="0" cellpadding="0" style="border-collapse: collapse" width="100%" id="table5">
            <tr>
              <td bgcolor="#CFD7DB">&nbsp;</td>
              <td align="right" bgcolor="#CFD7DB">2&nbsp;&nbsp; </td>
            </tr>
            <tr>
              <td align="right" bgcolor="#FFFFFF"><font size="2">D:</font></td>
              <td align="right" bgcolor="#FFFFFF"><font size="2"><a href="javascript:BlahBlah()">
              $123.45</a></font></td>
            </tr>
            <tr>
              <td align="right" bgcolor="#FFFFFF"><font size="2">WD:</font></td>
              <td align="right" bgcolor="#FFFFFF">&nbsp;</td>
            </tr>
            <tr>
              <td align="right" bgcolor="#FFFFFF"><font size="2">WKLY:</font></td>
              <td align="right" bgcolor="#FFFFFF"><font size="2">$456.67</font></td>
            </tr>
          </table>
          </td>
          <td>
          <table border="0" cellpadding="0" style="border-collapse: collapse" width="100%" id="table6">
            <tr>
              <td bgcolor="#CFD7DB">&nbsp;</td>
              <td align="right" bgcolor="#CFD7DB">3 </td>
            </tr>
            <tr>
              <td align="right" bgcolor="#FFFFFF"><font size="2">D:</font></td>
              <td align="right" bgcolor="#FFFFFF"><font size="2"><a href="javascript:BlahBlah()">
              $123.45</a></font></td>
            </tr>
            <tr>
              <td align="right" bgcolor="#FFFFFF"><font size="2">WD:</font></td>
              <td align="right" bgcolor="#FFFFFF">&nbsp;</td>
            </tr>
            <tr>
              <td align="right" bgcolor="#FFFFFF"><font size="2">WKLY:</font></td>
              <td align="right" bgcolor="#FFFFFF"><font size="2">$456.67</font></td>
            </tr>
          </table>
          </td>
          <td>
          <table border="0" cellpadding="0" style="border-collapse: collapse" width="100%" id="table32">
            <tr>
              <td bgcolor="#CFD7DB">&nbsp;</td>
              <td align="right" bgcolor="#CFD7DB">4</td>
            </tr>
            <tr>
              <td align="right" bgcolor="#FFFFFF"><font size="2">D:</font></td>
              <td align="right" bgcolor="#FFFFFF"><font size="2"><a href="javascript:BlahBlah()">
              $123.45</a></font></td>
            </tr>
            <tr>
              <td align="right" bgcolor="#FFFFFF"><font size="2">WD:</font></td>
              <td align="right" bgcolor="#FFFFFF">&nbsp;</td>
            </tr>
            <tr>
              <td align="right" bgcolor="#FFFFFF"><font size="2">WKLY:</font></td>
              <td align="right" bgcolor="#FFFFFF"><font size="2">$456.67</font></td>
            </tr>
          </table>
          </td>
          <td>
          <table border="0" cellpadding="0" style="border-collapse: collapse" width="100%" id="table33">
            <tr>
              <td bgcolor="#CFD7DB">&nbsp;</td>
              <td align="right" bgcolor="#CFD7DB">5</td>
            </tr>
            <tr>
              <td align="right" bordercolor="#FFFFFF" bgcolor="#FFFFFF"><font size="2">
              D:</font></td>
              <td align="right" bordercolor="#FFFFFF" bgcolor="#FFFFFF"><font size="2"><a href="javascript:BlahBlah()">
              $123.45</a></font></td>
            </tr>
            <tr>
              <td align="right" bordercolor="#FFFFFF" bgcolor="#FFFFFF"><font size="2">
              WD:</font></td>
              <td align="right" bordercolor="#FFFFFF" bgcolor="#FFFFFF"><font size="2">
              $234.56</font></td>
            </tr>
            <tr>
              <td align="right" bordercolor="#FFFFFF" bgcolor="#FFFFFF"><font size="2">
              WKLY:</font></td>
              <td align="right" bordercolor="#FFFFFF" bgcolor="#FFFFFF"><font size="2">
              $456.67</font></td>
            </tr>
          </table>
          </td>
          <td>
          <table border="0" cellpadding="0" style="border-collapse: collapse" width="100%" id="table34">
            <tr>
              <td bgcolor="#CFD7DB">&nbsp;</td>
              <td align="right" bgcolor="#CFD7DB">6</td>
            </tr>
            <tr>
              <td align="right" bgcolor="#FFFFFF"><font size="2">D:</font></td>
              <td align="right" bgcolor="#FFFFFF">&nbsp;</td>
            </tr>
            <tr>
              <td align="right" bgcolor="#FFFFFF"><font size="2">WD:</font></td>
              <td align="right" bgcolor="#FFFFFF"><font size="2">$234.56</font></td>
            </tr>
            <tr>
              <td align="right" bgcolor="#FFFFFF"><font size="2">WKLY:</font></td>
              <td align="right" bgcolor="#FFFFFF"><font size="2">$456.67</font></td>
            </tr>
          </table>
          </td>
        </tr>
        <tr>
          <td width="14%">
          <table border="0" cellpadding="0" style="border-collapse: collapse" width="100%" id="table11">
            <tr>
              <td bgcolor="#CFD7DB">&nbsp;</td>
              <td align="right" bgcolor="#CFD7DB">7</td>
            </tr>
            <tr>
              <td align="right" bgcolor="#FFFFFF"><font size="2">D:</font></td>
              <td align="right" bgcolor="#FFFFFF"><font size="2"><a href="javascript:BlahBlah()">
              $123.45</a></font></td>
            </tr>
            <tr>
              <td align="right" bgcolor="#FFFFFF"><font size="2">WD:</font></td>
              <td align="right" bgcolor="#FFFFFF"><font size="2">$234.56</font></td>
            </tr>
            <tr>
              <td align="right" bgcolor="#FFFFFF"><font size="2">WKLY:</font></td>
              <td align="right" bgcolor="#FFFFFF"><font size="2">$456.67</font></td>
            </tr>
          </table>
          </td>
          <td>
          <table border="0" cellpadding="0" style="border-collapse: collapse" width="100%" id="table12">
            <tr>
              <td bgcolor="#CFD7DB">&nbsp;</td>
              <td align="right" bgcolor="#CFD7DB">8</td>
            </tr>
            <tr>
              <td align="right" bgcolor="#FFFFFF"><font size="2">D:</font></td>
              <td align="right" bgcolor="#FFFFFF"><font size="2"><a href="javascript:BlahBlah()">
              $123.45</a></font></td>
            </tr>
            <tr>
              <td align="right" bgcolor="#FFFFFF"><font size="2">WD:</font></td>
              <td align="right" bgcolor="#FFFFFF"><font size="2">$234.56</font></td>
            </tr>
            <tr>
              <td align="right" bgcolor="#FFFFFF"><font size="2">WKLY:</font></td>
              <td align="right" bgcolor="#FFFFFF"><font size="2">$456.67</font></td>
            </tr>
          </table>
          </td>
          <td>
          <table border="0" cellpadding="0" style="border-collapse: collapse" width="100%" id="table13">
            <tr>
              <td bgcolor="#CFD7DB">&nbsp;</td>
              <td align="right" bgcolor="#CFD7DB">9</td>
            </tr>
            <tr>
              <td align="right" bgcolor="#FFFFFF"><font size="2">D:</font></td>
              <td align="right" bgcolor="#FFFFFF"><font size="2"><a href="javascript:BlahBlah()">
              $123.45</a></font></td>
            </tr>
            <tr>
              <td align="right" bgcolor="#FFFFFF"><font size="2">WD:</font></td>
              <td align="right" bgcolor="#FFFFFF"><font size="2">$234.56</font></td>
            </tr>
            <tr>
              <td align="right" bgcolor="#FFFFFF"><font size="2">WKLY:</font></td>
              <td align="right" bgcolor="#FFFFFF"><font size="2">$456.67</font></td>
            </tr>
          </table>
          </td>
          <td>
          <table border="0" cellpadding="0" style="border-collapse: collapse" width="100%" id="table14">
            <tr>
              <td bgcolor="#CFD7DB">&nbsp;</td>
              <td align="right" bgcolor="#CFD7DB">10</td>
            </tr>
            <tr>
              <td align="right" bgcolor="#FFFFFF"><font size="2">D:</font></td>
              <td align="right" bgcolor="#FFFFFF"><font size="2"><a href="javascript:BlahBlah()">
              $123.45</a></font></td>
            </tr>
            <tr>
              <td align="right" bgcolor="#FFFFFF"><font size="2">WD:</font></td>
              <td align="right" bgcolor="#FFFFFF"><font size="2">$234.56</font></td>
            </tr>
            <tr>
              <td align="right" bgcolor="#FFFFFF"><font size="2">WKLY:</font></td>
              <td align="right" bgcolor="#FFFFFF"><font size="2">$456.67</font></td>
            </tr>
          </table>
          </td>
          <td>
          <table border="0" cellpadding="0" style="border-collapse: collapse" width="100%" id="table15">
            <tr>
              <td bgcolor="#CFD7DB">&nbsp;</td>
              <td align="right" bgcolor="#CFD7DB">11</td>
            </tr>
            <tr>
              <td align="right"><font size="2">D:</font></td>
              <td align="right"><font size="2"><a href="javascript:BlahBlah()">
              $123.45</a></font></td>
            </tr>
            <tr>
              <td align="right"><font size="2">WD:</font></td>
              <td align="right"><font size="2">$234.56</font></td>
            </tr>
            <tr>
              <td align="right"><font size="2">WKLY:</font></td>
              <td align="right"><font size="2">$456.67</font></td>
            </tr>
          </table>
          </td>
          <td>
          <table border="0" cellpadding="0" style="border-collapse: collapse" width="100%" id="table16">
            <tr>
              <td bgcolor="#CFD7DB">&nbsp;</td>
              <td align="right" bgcolor="#CFD7DB">12</td>
            </tr>
            <tr>
              <td align="right"><font size="2">D:</font></td>
              <td align="right"><font size="2"><a href="javascript:BlahBlah()">
              $123.45</a></font></td>
            </tr>
            <tr>
              <td align="right"><font size="2">WD:</font></td>
              <td align="right"><font size="2">$234.56</font></td>
            </tr>
            <tr>
              <td align="right"><font size="2">WKLY:</font></td>
              <td align="right"><font size="2">$456.67</font></td>
            </tr>
          </table>
          </td>
          <td>
          <table border="0" cellpadding="0" style="border-collapse: collapse" width="100%" id="table17">
            <tr>
              <td bgcolor="#CFD7DB">&nbsp;</td>
              <td align="right" bgcolor="#CFD7DB">13</td>
            </tr>
            <tr>
              <td align="right"><font size="2">D:</font></td>
              <td align="right"><font size="2"><a href="javascript:BlahBlah()">
              $123.45</a></font></td>
            </tr>
            <tr>
              <td align="right"><font size="2">WD:</font></td>
              <td align="right"><font size="2">$234.56</font></td>
            </tr>
            <tr>
              <td align="right"><font size="2">WKLY:</font></td>
              <td align="right"><font size="2">$456.67</font></td>
            </tr>
          </table>
          </td>
        </tr>
        <tr>
          <td width="14%">
          <table border="0" cellpadding="0" style="border-collapse: collapse" width="100%" id="table18">
            <tr>
              <td bgcolor="#CFD7DB">&nbsp;</td>
              <td align="right" bgcolor="#CFD7DB">14</td>
            </tr>
            <tr>
              <td align="right"><font size="2">D:</font></td>
              <td align="right"><font size="2"><a href="javascript:BlahBlah()">
              $123.45</a></font></td>
            </tr>
            <tr>
              <td align="right"><font size="2">WD:</font></td>
              <td align="right"><font size="2">$234.56</font></td>
            </tr>
            <tr>
              <td align="right"><font size="2">WKLY:</font></td>
              <td align="right"><font size="2">$456.67</font></td>
            </tr>
          </table>
          </td>
          <td>
          <table border="0" cellpadding="0" style="border-collapse: collapse" width="100%" id="table19">
            <tr>
              <td bgcolor="#CFD7DB">&nbsp;</td>
              <td align="right" bgcolor="#CFD7DB">15</td>
            </tr>
            <tr>
              <td align="right"><font size="2">D:</font></td>
              <td align="right"><font size="2"><a href="javascript:BlahBlah()">
              $123.45</a></font></td>
            </tr>
            <tr>
              <td align="right"><font size="2">WD:</font></td>
              <td align="right"><font size="2">$234.56</font></td>
            </tr>
            <tr>
              <td align="right"><font size="2">WKLY:</font></td>
              <td align="right"><font size="2">$456.67</font></td>
            </tr>
          </table>
          </td>
          <td>
          <table border="0" cellpadding="0" style="border-collapse: collapse" width="100%" id="table20">
            <tr>
              <td bgcolor="#CFD7DB">&nbsp;</td>
              <td align="right" bgcolor="#CFD7DB">16</td>
            </tr>
            <tr>
              <td align="right"><font size="2">D:</font></td>
              <td align="right"><font size="2"><a href="javascript:BlahBlah()">
              $123.45</a></font></td>
            </tr>
            <tr>
              <td align="right"><font size="2">WD:</font></td>
              <td align="right"><font size="2">$234.56</font></td>
            </tr>
            <tr>
              <td align="right"><font size="2">WKLY:</font></td>
              <td align="right"><font size="2">$456.67</font></td>
            </tr>
          </table>
          </td>
          <td>
          <table border="0" cellpadding="0" style="border-collapse: collapse" width="100%" id="table21">
            <tr>
              <td bgcolor="#CFD7DB">&nbsp;</td>
              <td align="right" bgcolor="#CFD7DB">17</td>
            </tr>
            <tr>
              <td align="right"><font size="2">D:</font></td>
              <td align="right"><font size="2"><a href="javascript:BlahBlah()">
              $123.45</a></font></td>
            </tr>
            <tr>
              <td align="right"><font size="2">WD:</font></td>
              <td align="right"><font size="2">$234.56</font></td>
            </tr>
            <tr>
              <td align="right"><font size="2">WKLY:</font></td>
              <td align="right"><font size="2">$456.67</font></td>
            </tr>
          </table>
          </td>
          <td>
          <table border="0" cellpadding="0" style="border-collapse: collapse" width="100%" id="table22">
            <tr>
              <td bgcolor="#CFD7DB">&nbsp;</td>
              <td align="right" bgcolor="#CFD7DB">18</td>
            </tr>
            <tr>
              <td align="right"><font size="2">D:</font></td>
              <td align="right"><font size="2"><a href="javascript:BlahBlah()">
              $123.45</a></font></td>
            </tr>
            <tr>
              <td align="right"><font size="2">WD:</font></td>
              <td align="right"><font size="2">$234.56</font></td>
            </tr>
            <tr>
              <td align="right"><font size="2">WKLY:</font></td>
              <td align="right"><font size="2">$456.67</font></td>
            </tr>
          </table>
          </td>
          <td>
          <table border="0" cellpadding="0" style="border-collapse: collapse" width="100%" id="table23">
            <tr>
              <td bgcolor="#CFD7DB">&nbsp;</td>
              <td align="right" bgcolor="#CFD7DB">19</td>
            </tr>
            <tr>
              <td align="right"><font size="2">D:</font></td>
              <td align="right"><font size="2"><a href="javascript:BlahBlah()">
              $123.45</a></font></td>
            </tr>
            <tr>
              <td align="right"><font size="2">WD:</font></td>
              <td align="right"><font size="2">$234.56</font></td>
            </tr>
            <tr>
              <td align="right"><font size="2">WKLY:</font></td>
              <td align="right"><font size="2">$456.67</font></td>
            </tr>
          </table>
          </td>
          <td>
          <table border="0" cellpadding="0" style="border-collapse: collapse" width="100%" id="table24">
            <tr>
              <td bgcolor="#CFD7DB">&nbsp;</td>
              <td align="right" bgcolor="#CFD7DB">20</td>
            </tr>
            <tr>
              <td align="right"><font size="2">D:</font></td>
              <td align="right"><font size="2"><a href="javascript:BlahBlah()">
              $123.45</a></font></td>
            </tr>
            <tr>
              <td align="right"><font size="2">WD:</font></td>
              <td align="right"><font size="2">$234.56</font></td>
            </tr>
            <tr>
              <td align="right"><font size="2">WKLY:</font></td>
              <td align="right"><font size="2">$456.67</font></td>
            </tr>
          </table>
          </td>
        </tr>
        <tr>
          <td width="14%">
          <table border="0" cellpadding="0" style="border-collapse: collapse" width="100%" id="table25">
            <tr>
              <td bgcolor="#CFD7DB">&nbsp;</td>
              <td align="right" bgcolor="#CFD7DB">21</td>
            </tr>
            <tr>
              <td align="right"><font size="2">D:</font></td>
              <td align="right"><font size="2"><a href="javascript:BlahBlah()">
              $123.45</a></font></td>
            </tr>
            <tr>
              <td align="right"><font size="2">WD:</font></td>
              <td align="right"><font size="2">$234.56</font></td>
            </tr>
            <tr>
              <td align="right"><font size="2">WKLY:</font></td>
              <td align="right"><font size="2">$456.67</font></td>
            </tr>
          </table>
          </td>
          <td>
          <table border="0" cellpadding="0" style="border-collapse: collapse" width="100%" id="table26">
            <tr>
              <td bgcolor="#CFD7DB">&nbsp;</td>
              <td align="right" bgcolor="#CFD7DB">22</td>
            </tr>
            <tr>
              <td align="right"><font size="2">D:</font></td>
              <td align="right"><font size="2"><a href="javascript:BlahBlah()">
              $123.45</a></font></td>
            </tr>
            <tr>
              <td align="right"><font size="2">WD:</font></td>
              <td align="right"><font size="2">$234.56</font></td>
            </tr>
            <tr>
              <td align="right"><font size="2">WKLY:</font></td>
              <td align="right"><font size="2">$456.67</font></td>
            </tr>
          </table>
          </td>
          <td>
          <table border="0" cellpadding="0" style="border-collapse: collapse" width="100%" id="table27">
            <tr>
              <td bgcolor="#CFD7DB">&nbsp;</td>
              <td align="right" bgcolor="#CFD7DB">23</td>
            </tr>
            <tr>
              <td align="right"><font size="2">D:</font></td>
              <td align="right"><font size="2"><a href="javascript:BlahBlah()">
              $123.45</a></font></td>
            </tr>
            <tr>
              <td align="right"><font size="2">WD:</font></td>
              <td align="right"><font size="2">$234.56</font></td>
            </tr>
            <tr>
              <td align="right"><font size="2">WKLY:</font></td>
              <td align="right"><font size="2">$456.67</font></td>
            </tr>
          </table>
          </td>
          <td>
          <table border="0" cellpadding="0" style="border-collapse: collapse" width="100%" id="table28">
            <tr>
              <td bgcolor="#CFD7DB">&nbsp;</td>
              <td align="right" bgcolor="#CFD7DB">24</td>
            </tr>
            <tr>
              <td align="right"><font size="2">D:</font></td>
              <td align="right"><font size="2"><a href="javascript:BlahBlah()">
              $123.45</a></font></td>
            </tr>
            <tr>
              <td align="right"><font size="2">WD:</font></td>
              <td align="right"><font size="2">$234.56</font></td>
            </tr>
            <tr>
              <td align="right"><font size="2">WKLY:</font></td>
              <td align="right"><font size="2">$456.67</font></td>
            </tr>
          </table>
          </td>
          <td>
          <table border="0" cellpadding="0" style="border-collapse: collapse" width="100%" id="table29">
            <tr>
              <td bgcolor="#CFD7DB">&nbsp;</td>
              <td align="right" bgcolor="#CFD7DB">25</td>
            </tr>
            <tr>
              <td align="right"><font size="2">D:</font></td>
              <td align="right"><font size="2"><a href="javascript:BlahBlah()">
              $123.45</a></font></td>
            </tr>
            <tr>
              <td align="right"><font size="2">WD:</font></td>
              <td align="right"><font size="2">$234.56</font></td>
            </tr>
            <tr>
              <td align="right"><font size="2">WKLY:</font></td>
              <td align="right"><font size="2">$456.67</font></td>
            </tr>
          </table>
          </td>
          <td>
          <table border="0" cellpadding="0" style="border-collapse: collapse" width="100%" id="table30">
            <tr>
              <td bgcolor="#CFD7DB">&nbsp;</td>
              <td align="right" bgcolor="#CFD7DB">26</td>
            </tr>
            <tr>
              <td align="right"><font size="2">D:</font></td>
              <td align="right"><font size="2"><a href="javascript:BlahBlah()">
              $123.45</a></font></td>
            </tr>
            <tr>
              <td align="right"><font size="2">WD:</font></td>
              <td align="right"><font size="2">$234.56</font></td>
            </tr>
            <tr>
              <td align="right"><font size="2">WKLY:</font></td>
              <td align="right"><font size="2">$456.67</font></td>
            </tr>
          </table>
          </td>
          <td>
          <table border="0" cellpadding="0" style="border-collapse: collapse" width="100%" id="table31">
            <tr>
              <td bgcolor="#CFD7DB">&nbsp;</td>
              <td align="right" bgcolor="#CFD7DB">27</td>
            </tr>
            <tr>
              <td align="right"><font size="2">D:</font></td>
              <td align="right"><font size="2"><a href="javascript:BlahBlah()">
              $123.45</a></font></td>
            </tr>
            <tr>
              <td align="right"><font size="2">WD:</font></td>
              <td align="right"><font size="2">$234.56</font></td>
            </tr>
            <tr>
              <td align="right"><font size="2">WKLY:</font></td>
              <td align="right"><font size="2">$456.67</font></td>
            </tr>
          </table>
          </td>
        </tr>
        <tr>
          <td width="14%">
          <table border="0" cellpadding="0" style="border-collapse: collapse" width="100%" id="table35">
            <tr>
              <td bgcolor="#CFD7DB">&nbsp;</td>
              <td align="right" bgcolor="#CFD7DB">28</td>
            </tr>
            <tr>
              <td align="right"><font size="2">D:</font></td>
              <td align="right"><font size="2"><a href="javascript:BlahBlah()">
              $123.45</a></font></td>
            </tr>
            <tr>
              <td align="right"><font size="2">WD:</font></td>
              <td align="right"><font size="2">$234.56</font></td>
            </tr>
            <tr>
              <td align="right"><font size="2">WKLY:</font></td>
              <td align="right"><font size="2">$456.67</font></td>
            </tr>
          </table>
          </td>
          <td>
          <table border="0" cellpadding="0" style="border-collapse: collapse" width="100%" id="table36">
            <tr>
              <td bgcolor="#CFD7DB">&nbsp;</td>
              <td align="right" bgcolor="#CFD7DB">29</td>
            </tr>
            <tr>
              <td align="right"><font size="2">D:</font></td>
              <td align="right"><font size="2"><a href="javascript:BlahBlah()">
              $123.45</a></font></td>
            </tr>
            <tr>
              <td align="right"><font size="2">WD:</font></td>
              <td align="right"><font size="2">$234.56</font></td>
            </tr>
            <tr>
              <td align="right"><font size="2">WKLY:</font></td>
              <td align="right"><font size="2">$456.67</font></td>
            </tr>
          </table>
          </td>
          <td>
          <table border="0" cellpadding="0" style="border-collapse: collapse" width="100%" id="table37">
            <tr>
              <td bgcolor="#CFD7DB">&nbsp;</td>
              <td align="right" bgcolor="#CFD7DB">30</td>
            </tr>
            <tr>
              <td align="right"><font size="2">D:</font></td>
              <td align="right"><font size="2"><a href="javascript:BlahBlah()">
              $123.45</a></font></td>
            </tr>
            <tr>
              <td align="right"><font size="2">WD:</font></td>
              <td align="right"><font size="2">$234.56</font></td>
            </tr>
            <tr>
              <td align="right"><font size="2">WKLY:</font></td>
              <td align="right"><font size="2">$456.67</font></td>
            </tr>
          </table>
          </td>
          <td>
          <table border="0" cellpadding="0" style="border-collapse: collapse" width="100%" id="table38">
            <tr>
              <td bgcolor="#CFD7DB">&nbsp;</td>
              <td align="right" bgcolor="#CFD7DB">31</td>
            </tr>
            <tr>
              <td align="right"><font size="2">D:</font></td>
              <td align="right"><font size="2"><a href="javascript:BlahBlah()">
              $123.45</a></font></td>
            </tr>
            <tr>
              <td align="right"><font size="2">WD:</font></td>
              <td align="right"><font size="2">$234.56</font></td>
            </tr>
            <tr>
              <td align="right"><font size="2">WKLY:</font></td>
              <td align="right"><font size="2">$456.67</font></td>
            </tr>
          </table>
          </td>
          <td class="off_day" >&nbsp;</td>
          <td bgcolor="#879AA2">&nbsp;</td>
          <td bgcolor="#879AA2">&nbsp;</td>
        </tr>
      </table>
    </form>
    <p>&nbsp;&nbsp;</p>
    <table width="745" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td><img src="images/ruler.gif" width="745" height="2"></td>
      </tr>
    </table>
    <p>&nbsp;</p>
    <p>&nbsp;</p>
    </td>
  </tr>
  </table>
<p>&nbsp;</p>
<p>
&nbsp;</p>
</td></tr></table>
<!-- JUSTTABS BOTTOM CLOSE -->
<!--#INCLUDE FILE="footer.asp"-->
</body>
</html>
