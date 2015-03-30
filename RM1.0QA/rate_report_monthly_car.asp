<%@ Language=VBScript %>
<!-- #INCLUDE FILE="inc/login_check.asp" -->
<!-- #INCLUDE FILE="inc/login_check.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Rate-Monitor by Rate-Highway, Inc. | Reports</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" type="text/css" href="inc/sitewide.css">
<script src="inc/sitewide.js" language="javascript"></script>
<script src="inc/header_menu_support.js" language="javascript"></script>
<link rel="stylesheet" type="text/css" href="inc/rh_standard.css">
</head>

<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="MM_preloadImages('images/b_search_pro_on.gif','images/b_search_que_on.gif','images/b_search_cri_on.gif','images/b_rate_on.gif','images/b_alert_on.gif','images/b_user_on.gif','images/b_system_on.gif')">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td background="images/top_tile.gif"><img src="images/top.jpg" width="770" height="91"></td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td background="images/b_tile.gif">
<table width="400" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td><img src="images/b_left.jpg" width="62" height="32"></td>
          <td><a href="search_profiles_car.asp" onMouseOver="MM_swapImage('s1','','images/b_search_pro_on.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="images/b_search_pro_of.gif" name="s1" width="183" height="32" border="0" id="s1"></a></td>
          <td><a href="search_queue_car.asp" onMouseOver="MM_swapImage('s2','','images/b_search_que_on.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="images/b_search_que_of.gif" name="s2" width="97" height="32" border="0" id="s2"></a></td>
          <td><a href="search_criteria_car.asp" onMouseOver="MM_swapImage('s3','','images/b_search_cri_on.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="images/b_search_cri_of.gif" name="s3" width="103" height="32" border="0" id="s3"></a></td>
          <td>
          <a href="rate_report_monthly_car.asp" onMouseOver="MM_swapImage('ra','','images/b_rate_on.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="images/b_rate_of.gif" name="ra" width="88" height="32" border="0" id="ra"></a></td>
          <td><a href="javascript:not_enabled()" onMouseOver="MM_swapImage('al','','images/b_alert_on.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="images/b_alert_of.gif" name="al" width="53" height="32" border="0" id="al"></a></td>
          <td><a href="javascript:not_enabled()" onMouseOver="MM_swapImage('us','','images/b_user_on.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="images/b_user_of.gif" name="us" width="126" height="32" border="0" id="us"></a></td>
          <td><a href="javascript:not_enabled()" onMouseOver="MM_swapImage('sy','','images/b_system_on.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="images/b_system_of.gif" name="sy" width="58" height="32" border="0" id="sy"></a></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td background="images/med_bar_tile.gif"><img src="images/med_bar.gif" width="12" height="8"></td>
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
    <td background="images/h_tile.gif"><table width="100" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td><img src="images/h_rates_reports.gif" width="368" height="31"></td>
          <td><img src="images/h_right.gif" width="402" height="31"></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<table width="770" border="0" cellspacing="0" cellpadding="2">
  <tr> 
    <td width="20"><img src="images/separator.gif" width="20" height="20"></td>
    <td width="858" valign="top"> &nbsp;<p>&nbsp;Rate Report By Month</p>
    <p>&nbsp;<font face="Tahoma" size="2">Extra Report Types:</font></p>
    <p><font face="Tahoma" size="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
    <a target="_blank" href="Dollar_Rate_Change.doc">Dollar Rate Change Report</a></font></p>
    <p><font face="Tahoma" size="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Budget Rate 
    Change Form</font></p>
    <p><font face="Tahoma" size="2">&nbsp;&nbsp;&nbsp;&nbsp; </font></p>
      <table width="745" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td><img src="images/ruler.gif" width="745" height="2"></td>
        </tr>
      </table>
      <p>&nbsp;</p></td>
  </tr>
  <tr> 
    <td width="20">&nbsp;</td>
    <td width="858" valign="top"> &nbsp;</td>
  </tr>
</table>
<!--#INCLUDE FILE="footer.asp"-->
</body>
</html>