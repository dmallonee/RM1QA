<%@ Language=VBScript %>
<!-- #INCLUDE FILE="inc/login_check.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<title>Rate-Monitor by Rate-Highway, Inc. | Reports</title>
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
    <table width="400" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td><img src="images/b_left.jpg" width="62" height="32"></td>
        <td>
        <a href="search_profiles_car.asp" onmouseover="MM_swapImage('s1','','images/b_search_pro_on.gif',1)" onmouseout="MM_swapImgRestore()">
        <img src="images/b_search_pro_of.gif" name="s1" width="183" height="32" border="0" id="s1"></a></td>
        <td>
        <a href="search_queue_car.asp" onmouseover="MM_swapImage('s2','','images/b_search_que_on.gif',1)" onmouseout="MM_swapImgRestore()">
        <img src="images/b_search_que_of.gif" name="s2" width="97" height="32" border="0" id="s2"></a></td>
        <td>
        <a href="search_criteria_car.asp" onmouseover="MM_swapImage('s3','','images/b_search_cri_on.gif',1)" onmouseout="MM_swapImgRestore()">
        <img src="images/b_search_cri_of.gif" name="s3" width="103" height="32" border="0" id="s3"></a></td>
        <td>
        <a href="rate_wizard_car.asp" onmouseover="MM_swapImage('ra','','images/b_rate_on.gif',1)" onmouseout="MM_swapImgRestore()">
        <img src="images/b_rate_of.gif" name="ra" width="88" height="32" border="0" id="ra"></a></td>
        <td>
        <a href="javascript:not_enabled()" onmouseover="MM_swapImage('al','','images/b_alert_on.gif',1)" onmouseout="MM_swapImgRestore()">
        <img src="images/b_alert_of.gif" name="al" width="53" height="32" border="0" id="al"></a></td>
        <td>
        <a href="javascript:not_enabled()" onmouseover="MM_swapImage('us','','images/b_user_on.gif',1)" onmouseout="MM_swapImgRestore()">
        <img src="images/b_user_of.gif" name="us" width="126" height="32" border="0" id="us"></a></td>
        <td>
        <a href="javascript:not_enabled()" onmouseover="MM_swapImage('sy','','images/b_system_on.gif',1)" onmouseout="MM_swapImgRestore()">
        <img src="images/b_system_of.gif" name="sy" width="58" height="32" border="0" id="sy"></a></td>
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
    <table width="100" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td><img src="images/h_rates_reports.gif" width="368" height="31"></td>
		<td><map name="logout_map">
        <area alt="Click to log out of Rate-Monitor" href="http://www.rate-monitor.com/" shape="rect" coords="337, 10, 394, 25">
        </map>
        <img src="images/bottom_right.gif" width="402" height="31" usemap="#logout_map" border="0"></td>
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
<td rowspan="2" width="335">
<a title="Click to display the Rate Calendar" href="rate_calendar_car.asp">
<img src="images/ratecalendar0_ia.GIF" width="100" height="25" hspace="0" vspace="0" border="0" alt="Rate Calendar" description="Rate Calendar"></a><a href="rate_wizard_car.asp"><img src="images/ratewizards1_a.GIF" width="97" height="25" hspace="0" vspace="0" border="0" alt="Rate Wizards" description="Rate Calendar"></a><a href="rate_mgmt_car.asp"><img src="images/ratemananagement2_ia.GIF" width="138" height="25" hspace="0" vspace="0" border="0" alt="New ALT Text" description="Rate Calendar"></a></td>
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
<td bgcolor="#CED7DB">
<table border="0" cellspacing="5" cellpadding="5" width="860">
<tr>
<td width="316">
<font color="#080000">
<!-- JUSTTABS TOP CLOSE -->
    <form method="POST" action>
      <table border="1" cellpadding="0" style="border-collapse: collapse; border-top-width:0px" width="300" id="table3" bordercolor="#384F5B">
        <tr>
          <td colspan="2" bgcolor="#384F5B" style="border-left:1px solid #384F5B; border-right:1px solid #384F5B; border-top:1px solid #384F5B; "><font color="#FFFFFF">&nbsp;Alert Report 
          Properties</font></td>
        </tr>
        <tr>
          <td style="border-left-style: solid; border-left-width: 1px; border-right-style:none; border-right-width:medium; border-top-style:none; border-top-width:medium; border-bottom-style:none; border-bottom-width:medium; width: 122px;" bgcolor="#FFFFFF">
          &nbsp;</td>
          <td width="201" style="border-right:1px solid #384F5B; border-left-style:none; border-left-width:medium; border-top-style:none; border-top-width:medium; border-bottom-style:none; border-bottom-width:medium" bgcolor="#FFFFFF">
          &nbsp;</td>
        </tr>
        <tr>
          <td style="border-left-style: solid; border-left-width: 1px; border-right-style:none; border-right-width:medium; border-top-style:none; border-top-width:medium; border-bottom-style:none; border-bottom-width:medium; width: 122px;" bgcolor="#FFFFFF">
          <font size="2">&nbsp;Location:</font></td>
          <td width="201" style="border-right:1px solid #384F5B; border-left-style:none; border-left-width:medium; border-top-style:none; border-top-width:medium; border-bottom-style:none; border-bottom-width:medium" bgcolor="#FFFFFF">
          <select size="1" name="location">
          <option selected>LAX</option>
          <option>CCAR</option>
          <option>ICAR</option>
          <option>FCAR</option>
          </select></td>
        </tr>
        <tr>
          <td style="border-left-style: solid; border-left-width: 1px; border-right-style:none; border-right-width:medium; border-top-style:none; border-top-width:medium; border-bottom-style:none; border-bottom-width:medium; width: 122px;" bgcolor="#FFFFFF"><font size="2">&nbsp;Car Type:</font></td>
          <td width="201" style="border-right:1px solid #384F5B; border-top-style: none; border-top-width: medium; border-left-style:none; border-left-width:medium; border-bottom-style:none; border-bottom-width:medium" bgcolor="#FFFFFF">
          <p><select size="1" name="car_type">
          <option>ECAR</option>
          <option selected>CCAR</option>
          <option>ICAR</option>
          <option>FCAR</option>
          </select></p>
          </td>
        </tr>
        <tr>
<font color="#080000">
          <td style="border-left-style: solid; border-left-width: 1px; border-right-style:none; border-right-width:medium; border-top-style:none; border-top-width:medium; border-bottom-style:none; border-bottom-width:medium; width: 122px;" bgcolor="#FFFFFF">
			<font size="2" color="#080000">&nbsp;Length of Rent:</font></td>
          <td style="border-right:1px solid #384F5B; border-left-style:none; border-left-width:medium; border-top-style:none; border-top-width:medium; border-bottom-style:none; border-bottom-width:medium; width: 238px;" bgcolor="#FFFFFF">
			<select name="lor" style="width: 68px">
			<% For intIndex = 1 To 31 %>
			<% If CInt(intIndex) = CInt(strLOR) Then %>
				<option selected="selected" value="<%=intIndex %>"><%=intIndex %></option>
			<% Else                      %>
				<option value="<%=intIndex %>"><%=intIndex %></option>
			<% End If                    %>
			<% Next %>			
			<option value="0">All</option>
			</select></td>
    </font>
    	</tr>
        <tr>
          <td style="border-left-style: solid; border-left-width: 1px; border-right-style:none; border-right-width:medium; border-top-style:none; border-top-width:medium; border-bottom-style:none; border-bottom-width:medium; width: 122px;" bgcolor="#FFFFFF">
          <font size="2">&nbsp;Compare To:</font></td>
          <td width="201" style="border-right:1px solid #384F5B; border-left-style:none; border-left-width:medium; border-top-style:none; border-top-width:medium; border-bottom-style:none; border-bottom-width:medium" bgcolor="#FFFFFF">
          <select size="1" name="vendor">
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
          <td style="border-left-style: solid; border-left-width: 1px; border-right-style:none; border-right-width:medium; border-top-style:none; border-top-width:medium; border-bottom-style:none; border-bottom-width:medium; width: 122px;" bgcolor="#FFFFFF">&nbsp;<font size="2" color="#080000">From 
          Date:</font></td>
          <td width="201" style="border-right:1px solid #384F5B; border-left-style:none; border-left-width:medium; border-top-style:none; border-top-width:medium; border-bottom-style:none; border-bottom-width:medium" bgcolor="#FFFFFF">
          <input type="text" name="from_date" size="13" value="mm/dd/yyyy"></td>
        </tr>
        <tr>
          <td style="border-left-style: solid; border-left-width: 1px; border-right-style:none; border-right-width:medium; border-top-style:none; border-top-width:medium; border-bottom-style:none; border-bottom-width:medium; width: 122px;" bgcolor="#FFFFFF">&nbsp;<font size="2" color="#080000">To 
          Date:</font></td>
          <td width="201" style="border-right:1px solid #384F5B; border-left-style:none; border-left-width:medium; border-top-style:none; border-top-width:medium; border-bottom-style:none; border-bottom-width:medium" bgcolor="#FFFFFF">
<font color="#080000">
          <input type="text" name="from_date0" size="13" value="mm/dd/yyyy"></td>
        </tr>
        <tr>
          <td style="border-left-style: solid; border-left-width: 1px; border-right-style:none; border-right-width:medium; border-top-style:none; border-top-width:medium; border-bottom-style:none; border-bottom-width:medium; width: 122px;" bgcolor="#FFFFFF">
			&nbsp;</td>
          <td width="201" style="border-right:1px solid #384F5B; border-left-style:none; border-left-width:medium; border-top-style:none; border-top-width:medium; border-bottom-style:none; border-bottom-width:medium" bgcolor="#FFFFFF">
			&nbsp;</td>
        </tr>
        <tr>
          <td style="border-left-style: solid; border-left-width: 1px; border-right-style:none; border-right-width:medium; border-top-style:none; border-top-width:medium; border-bottom-style:none; border-bottom-width:medium; width: 122px;" bgcolor="#FFFFFF">
			&nbsp;</td>
          <td width="201" style="border-right:1px solid #384F5B; border-left-style:none; border-left-width:medium; border-top-style:none; border-top-width:medium; border-bottom-style:none; border-bottom-width:medium" bgcolor="#FFFFFF">
          &nbsp;<button name="display" type="submit">Display</button></td>
        </tr>
        <tr>
          <td style="border-bottom:1px solid #384F5B; border-left-style: solid; border-left-width: 1px; border-right-style:none; border-right-width:medium; border-top-style:none; border-top-width:medium; width: 122px;" bgcolor="#FFFFFF">
			&nbsp;</td>
          <td width="201" style="border-right:1px solid #384F5B; border-bottom-style: solid; border-bottom-width: 1px; border-left-style:none; border-left-width:medium; border-top-style:none; border-top-width:medium" bgcolor="#FFFFFF">
			&nbsp;</td>
        </tr>
      </table>
      <p>&nbsp;</p>
    </form>
    </td>
  <td width="316">
<font color="#080000">
<!-- JUSTTABS TOP CLOSE -->
    <form method="POST" action>
      <table border="1" cellpadding="0" style="border-collapse: collapse; border-top-width:0px" width="300" id="table3" bordercolor="#384F5B">
        <tr>
          <td colspan="2" bgcolor="#384F5B" style="border-left:1px solid #384F5B; border-right:1px solid #384F5B; border-top:1px solid #384F5B; "><font color="#FFFFFF">&nbsp;Ad-hoc 
          Rate Change</font></td>
        </tr>
        <tr>
          <td style="border-left-style: solid; border-left-width: 1px; border-right-style:none; border-right-width:medium; border-top-style:none; border-top-width:medium; border-bottom-style:none; border-bottom-width:medium; width: 123px;" bgcolor="#FFFFFF">
          &nbsp;</td>
          <td width="201" style="border-right:1px solid #384F5B; border-left-style:none; border-left-width:medium; border-top-style:none; border-top-width:medium; border-bottom-style:none; border-bottom-width:medium" bgcolor="#FFFFFF">
          &nbsp;</td>
        </tr>
        <tr>
          <td style="border-left-style: solid; border-left-width: 1px; border-right-style:none; border-right-width:medium; border-top-style:none; border-top-width:medium; border-bottom-style:none; border-bottom-width:medium; width: 123px;" bgcolor="#FFFFFF">
          <font size="2">&nbsp;Location:</font></td>
          <td width="201" style="border-right:1px solid #384F5B; border-left-style:none; border-left-width:medium; border-top-style:none; border-top-width:medium; border-bottom-style:none; border-bottom-width:medium" bgcolor="#FFFFFF">
          <select size="1" name="location">
          <option selected>LAX</option>
          <option>CCAR</option>
          <option>ICAR</option>
          <option>FCAR</option>
          </select></td>
        </tr>
        <tr>
          <td style="border-left-style: solid; border-left-width: 1px; border-right-style:none; border-right-width:medium; border-top-style:none; border-top-width:medium; border-bottom-style:none; border-bottom-width:medium; width: 123px;" bgcolor="#FFFFFF"><font size="2">&nbsp;Car Type:</font></td>
          <td width="201" style="border-right:1px solid #384F5B; border-top-style: none; border-top-width: medium; border-left-style:none; border-left-width:medium; border-bottom-style:none; border-bottom-width:medium" bgcolor="#FFFFFF">
          <p><select size="1" name="car_type">
          <option>ECAR</option>
          <option selected>CCAR</option>
          <option>ICAR</option>
          <option>FCAR</option>
          </select></p>
          </td>
        </tr>
        <tr>
<font color="#080000">
          <td style="border-left-style: solid; border-left-width: 1px; border-right-style:none; border-right-width:medium; border-top-style:none; border-top-width:medium; border-bottom-style:none; border-bottom-width:medium; width: 123px;" bgcolor="#FFFFFF">
			<font size="2" color="#080000">&nbsp;Length of Rent:</font></td>
          <td style="border-right:1px solid #384F5B; border-left-style:none; border-left-width:medium; border-top-style:none; border-top-width:medium; border-bottom-style:none; border-bottom-width:medium; width: 238px;" bgcolor="#FFFFFF">
			<select name="lor0" style="width: 68px">
			<% For intIndex = 1 To 31 %>
			<% If CInt(intIndex) = CInt(strLOR) Then %>
				<option selected="selected" value="<%=intIndex %>"><%=intIndex %></option>
			<% Else                      %>
				<option value="<%=intIndex %>"><%=intIndex %></option>
			<% End If                    %>
			<% Next %>			
			<option value="0">All</option>
			</select></td>
    </font>
    	</tr>
        <tr>
          <td style="border-left-style: solid; border-left-width: 1px; border-right-style:none; border-right-width:medium; border-top-style:none; border-top-width:medium; border-bottom-style:none; border-bottom-width:medium; width: 123px;" bgcolor="#FFFFFF">
          <font size="2">&nbsp;New Rate:</font></td>
          <td width="201" style="border-right:1px solid #384F5B; border-left-style:none; border-left-width:medium; border-top-style:none; border-top-width:medium; border-bottom-style:none; border-bottom-width:medium" bgcolor="#FFFFFF">
          <input type="text" name="T1" size="13"></td>
        </tr>
        <tr>
          <td style="border-left-style: solid; border-left-width: 1px; border-right-style:none; border-right-width:medium; border-top-style:none; border-top-width:medium; border-bottom-style:none; border-bottom-width:medium; width: 123px;" bgcolor="#FFFFFF">&nbsp;<font color="#080000"><font size="2">From 
          Date:</font></td>
          <td width="201" style="border-right:1px solid #384F5B; border-left-style:none; border-left-width:medium; border-top-style:none; border-top-width:medium; border-bottom-style:none; border-bottom-width:medium" bgcolor="#FFFFFF">
<font color="#080000">
          <input type="text" name="from_date1" size="13" value="mm/dd/yyyy"></td>
        </tr>
        <tr>
          <td style="border-left-style: solid; border-left-width: 1px; border-right-style:none; border-right-width:medium; border-top-style:none; border-top-width:medium; border-bottom-style:none; border-bottom-width:medium; width: 123px;" bgcolor="#FFFFFF">
<font color="#080000">
          <font size="2">&nbsp;To Date:</font></td>
          <td width="201" style="border-right:1px solid #384F5B; border-left-style:none; border-left-width:medium; border-top-style:none; border-top-width:medium; border-bottom-style:none; border-bottom-width:medium" bgcolor="#FFFFFF">
<font color="#080000">
          <input type="text" name="from_date2" size="13" value="mm/dd/yyyy"></td>
        </tr>
        <tr>
          <td style="border-left-style: solid; border-left-width: 1px; border-right-style:none; border-right-width:medium; border-top-style:none; border-top-width:medium; border-bottom-style:none; border-bottom-width:medium; width: 123px;" bgcolor="#FFFFFF">
			&nbsp;</td>
          <td width="201" style="border-right:1px solid #384F5B; border-left-style:none; border-left-width:medium; border-top-style:none; border-top-width:medium; border-bottom-style:none; border-bottom-width:medium" bgcolor="#FFFFFF">
			&nbsp;</td>
        </tr>
        <tr>
          <td style="border-left-style: solid; border-left-width: 1px; border-right-style:none; border-right-width:medium; border-top-style:none; border-top-width:medium; border-bottom-style:none; border-bottom-width:medium; width: 123px;" bgcolor="#FFFFFF">
			&nbsp;</td>
          <td width="201" style="border-right:1px solid #384F5B; border-left-style:none; border-left-width:medium; border-top-style:none; border-top-width:medium; border-bottom-style:none; border-bottom-width:medium" bgcolor="#FFFFFF">
          &nbsp;<button name="display" type="submit">Change</button></td>
        </tr>
        <tr>
          <td style="border-bottom:1px solid #384F5B; border-left-style: solid; border-left-width: 1px; border-right-style:none; border-right-width:medium; border-top-style:none; border-top-width:medium; width: 123px;" bgcolor="#FFFFFF">
			&nbsp;</td>
          <td width="201" style="border-right:1px solid #384F5B; border-bottom-style: solid; border-bottom-width: 1px; border-left-style:none; border-left-width:medium; border-top-style:none; border-top-width:medium" bgcolor="#FFFFFF">
			&nbsp;</td>
        </tr>
      </table>
      <p>&nbsp;</p>
    </form>
    </td>
  </tr>
  </table>
<p>&nbsp;</p>
<p>
<iframe name="rate_report" width="850" title="rate report" align="left" marginwidth="1" src="rate_change_report_car.asp" marginheight="0" height="200">
Your browser does not support inline frames or is currently configured not to display inline frames.
</iframe></p>
</td>
<td bgcolor="#CED7DB">
&nbsp;</td></tr></table>
&nbsp;<p>
<!-- JUSTTABS BOTTOM CLOSE -->
</p>
<!--#INCLUDE FILE="footer.asp"-->
</body>
</html>