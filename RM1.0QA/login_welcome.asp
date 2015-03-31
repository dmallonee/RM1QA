<%@ Language=VBScript %>
<!-- #INCLUDE FILE="inc/login_check_ex.asp" -->
<!-- #INCLUDE FILE="inc/adovbs.asp" -->
<% Response.Expires = -1
   Response.cachecontrol="private" 
   Response.AddHeader "pragma", "no-cache"

   Server.ScriptTimeout = 180
   On Error Resume Next

		strConn = Session("pro_con")
	
  		Set adoCmd = CreateObject("ADODB.Command")

		adoCmd.ActiveConnection =  strConn
		adoCmd.CommandText = "login_welcome_select"
		adoCmd.CommandType = 4

		adoCmd.Parameters.Append adoCmd.CreateParameter("@user_id",      3, 1, 0)
		
		adoCmd.Parameters("@user_id").Value = Request.Cookies("rate-monitor.com")("user_id")
		
		'Create an ADO RecordSet object
		Set adoRS = Server.CreateObject("ADODB.Recordset")

		'Open the RecordSet
		adoRS.Open adoCmd, , adOpenStatic, adLockReadOnly

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
<title>Rate-Monitor by Rate-Highway, Inc. | Welcome</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<link rel="stylesheet" type="text/css" href="inc/rh_standard.css">
<link rel="stylesheet" type="text/css" href="inc/rh_report.css">
<link rel="stylesheet" type="text/css" href="inc/sitewide.css">
<script type='text/javascript' language='javascript' src="inc/sitewide.js" ></script>
<script type='text/javascript' language='javascript' src="inc/header_menu_support.js" ></script>
<style type="text/css">
div.centered 
{
text-align: center;
}
div.centered table 
{
margin: 0 auto; 
text-align: left;
}
.style1 {
	text-align: center;
}
.style2 {
	font-size: large;
}
.grid_value_black {
	vertical-align: top;
}
.style6 {
	font-size: xx-small;
    color:black;
}
.grid_value_red {
	vertical-align: top;
	color: #FF0000;
}
</style>
</head>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onload="MM_preloadImages('images/b_search_pro_on.gif','images/b_search_que_on.gif','images/b_search_cri_on.gif','images/b_rate_on.gif','images/b_alert_on.gif','images/b_user_on.gif','images/b_system_on.gif')">
<div class="style1">
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
                <td>
                <div align="right">
                  <a href="default.asp"><font face="Arial, Helvetica, sans-serif" size="2" color="#FFFFFF">User: <%=Request.Cookies("rate-monitor.com")("user_name")%></font></div>
                </td>
                <!--
                <td><a href="http://www.rate-monitor.com">
                <img src="images/logout.gif" width="54" height="19" align="middle" border="0" ></a>
                </td>
                -->
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
        <td><img src="images/h_blank.gif" width="368" height="31"></td>
        <td><map name="logout_map">
        <area alt="Click to log out of Rate-Monitor" href="http://www.rate-monitor.com/" shape="rect" coords="337, 10, 394, 25">
        </map>
        <img src="images/bottom_right.gif" width="402" height="31" usemap="#logout_map" border="0"></td>
      </tr>
    </table>
    </td>
  </tr>
</table>
&nbsp;
<br>
	<span class="style2">Welcome to Rate Monitor<br>
	</span>Today is <%=FormatDateTime(now, 1) %><br>
    <!--You have logged in at level <%=Session("user_level") %>-->
	<% If adoRS.EOF = False Then %>
	<% Dim intCounter
	   intCounter = 1
	%>
	<div class="centered">
	<table style="width: 600" >
		<tr>
			<td style="width: 64px" class="profile_header">&nbsp;</td>
			<td class="profile_header_center">Information Worth Noting</td>
		</tr>

	<% While adoRS.EOF = False  %>
	
		<% If Fix(intCounter/2) = (intCounter/2) Then %>
		<tr class="profile_light">
			<td style="width: 64px" class="grid_value_black"><%=intCounter %>)</td>
			<td class="grid_value_black"><%=adoRS.Fields("message_html").Value %></td>
		</tr>
		<% Else %>
		<tr class="profile_dark">
			<td style="width: 64px" class="grid_value_black"><%=intCounter %>)</td>
			<td class="grid_value_black"><%=adoRS.Fields("message_html").Value %></td>

		</tr>
		<% End If %>
		
	<% intCounter = intCounter + 1 %>
	<% adoRS.MoveNext %>
	<% Wend  %>
	</table>
	</div>
	<% End If %>

	<% Set adoRS = adoRS.NextRecordSet %>
	<% Set adoRS = adoRS.NextRecordSet %>
   	<% Set adoRS = adoRS.NextRecordSet %>
	<% Set adoRS = adoRS.NextRecordset  %>
	<% If adoRS.EOF = False Then 
	     intContracted = adoRS.Fields("contracted_monthly_amt").Value
	     intRates = adoRS.Fields("month_to_date_rates").Value
   	     intShops = adoRS.Fields("month_to_date_shops").Value
	     intCRSUpdates = adoRS.Fields("regular_crs_updates").Value
	     intCRSTethered = adoRS.Fields("tethered_crs_updates").Value
	     blnRates = adoRS.Fields("is_rate_customer").Value     
	   Else
	     intContracted = 0








         intRates = 0
   	     intShops = 0












	     intCRSUpdates = 0
	     intCRSTethered = 0
	     blnRates = 1     
	   End If	
	%>
<br />	
	<div class="centered">
<% If blnRates Then %>
	<table style="width:350px" id="rates" >
		<tr>
			<td class="profile_header_center" style="font-weight:bold;color:white;padding:5px;" colspan="2">Your month to date rates searched count</td>
		</tr>
	
		<tr class="profile_light">
			<td style="width:200px" class="grid_value_black">Total Rates:</td>
			<td class="grid_value_black" style="width:150px;text-align:right;"><%=FormatNumber(intRates, 0, 0, -1, -1) %></td>
		</tr>








            <%if intContracted > 0 then %>
		<tr class="profile_dark">
			<td style="width:200px" class="grid_value_black">Rate Budget:</td>
			<td class="grid_value_black" style="width:150px;text-align:right;"><%=FormatNumber(intContracted, 0, 0, -1, -1 ) %></td>
		</tr>
		<% If ((intRates + intGDSRates) > intContracted) Then %>
		<tr class="profile_light">
			<td style="width:200px" class="grid_value_black">Rate Overage:</td>
			<td class="grid_value_red" style="width:150px;text-align:right;">
			<%=FormatNumber(intRates - intContracted, 0, 0, -1, -1) %>
			</td>
		</tr>
            <%end if %>
		<% End If %>
        <tr class="profile_dark">
			<td style="width:200px" class="grid_value_black">Regular CRS Updates:</td>
			<td class="grid_value_black" style="width:150px;text-align:right;"><%=FormatNumber(intCRSUpdates, 0, 0, -1, -1) %></td>
		</tr>
		<tr class="profile_light">
			<td style="width:200px" class="grid_value_black">Tethered CRS Updates:</td>
			<td class="grid_value_black" style="width:150px;text-align:right;"><%=FormatNumber(intCRSTethered, 0, 0, -1, -1) %></td>
		</tr>

	</table>
    <p class="style6">
	<!--
    * Special rates are GDS or deep searches and<br>&nbsp;are counted and billed as two rates.

    * Special rates are GDS (Worldspan, etc) and OTA (Expedia, etc.) searches.<br>&nbsp;
    </div>-->
<% Else %>
	<table style="width:350px" id="shops" >
		<tr>
			<td class="profile_header_center" style="font-weight:bold;color:white;padding:5px;" colspan="2">Your month to date shops count</td>
		</tr>

	 	<tr class="profile_light">
			<td style="width:200px" class="grid_value_black">Total Shops:</td>
			<td class="grid_value_black" style="width:150px;text-align:right;"><%=FormatNumber(intShops, 0, 0, -1, -1) %></td>
		</tr>








            <% if intContracted > 0 then %>
		<tr class="profile_dark">
			<td style="width:200px" class="grid_value_black">Shop Budget:</td>
			<td class="grid_value_black" style="width:150px;text-align:right;"><%=FormatNumber(intContracted, 0, 0, -1, -1 ) %></td>
		</tr>
		<% If ((intShops + intGDSShops) > intContracted) Then %>
		<tr class="profile_light">
			<td style="width:200px" class="grid_value_black">Shop Overage:</td>
			<td class="grid_value_red" style="width:150px;text-align:right;">
			<%=FormatNumber(intShops - intContracted, 0, 0, -1, -1) %>
			</td>
		</tr>
            <%end if %>
		<% End If %>
        <tr class="profile_light">
			<td style="width:200px" class="grid_value_black">Regular CRS Updates:</td>
			<td class="grid_value_black" style="width:150px;text-align:right;"><%=FormatNumber(intCRSUpdates, 0, 0, -1, -1) %></td>
		</tr>
		<tr class="profile_dark">
			<td style="width:200px" class="grid_value_black">Tethered CRS Updates*:</td>
			<td class="grid_value_black" style="width:150px;text-align:right;"><%=FormatNumber(intCRSTethered, 0, 0, -1, -1) %></td>
		</tr>

	</table>
	<p class="style6">
    * Special shops are GDS (Worldspan, etc) and OTA (Expedia, etc.) searches.<br>&nbsp;
	</p>
<% End If %>
<br />






	<button class="rh_button" id="graphs" type="submit">Display Graphs</button>



<br />
<div id="graphdisplay" style="display:none">
<select id="displayformat">
    <option value="clientusage7">7-Day Usage</option>
    <option value="clientusage12">Last 12(+) Hours</option>
    <option value="clientusage1">Hourly Usage by User</option>
</select>
<div id="clientusagediv">
  <center><div id="chart_div" style="width: 900px; height: 500px; margin:0 auto;text-align:center;"></div></center>
  <script type="text/javascript" src="https://www.google.com/jsapi"></script>
    <script type="text/javascript">
        google.load("visualization", "1", { packages: ["corechart"] });
    </script>
</div>
</div>
</div>
</div>
<script type="text/javascript" src="http://code.jquery.com/jquery-1.10.2.js" ></script>
<script type="text/javascript">
    $('#displayformat').change(function (event) {
        $('#clientusagediv').load($('#displayformat').val() + '.asp');
    });
    $('#queue').click(function () {
        window.location("search_queue_car.asp");
    });
    $('#graphs').click(function () {
        if ($('#graphs').text() == "Display Graphs") {
            $('#graphdisplay').css('display', 'block');
            $('#clientusagediv').load('clientusage7.asp');
            $('#graphs').text("Hide Graphs");
        } else {
            $('#graphdisplay').css('display', 'none');
            $('#graphs').text("Display Graphs");
        }
    });
</script>	
<% Set adoRS = Nothing %>
<!--#INCLUDE FILE="footer.asp"-->
</body>
</html>