<%@ Language=VBScript %>
<!-- #INCLUDE FILE="inc/login_check.asp" -->
<!-- #INCLUDE FILE="inc/adovbs.asp" --><% Response.Expires = -1
   Response.cachecontrol="private" 
   Response.AddHeader "pragma", "no-cache"

   Server.ScriptTimeout = 180


	Dim strConn	
	Dim adoCmd	
	Dim adoRS
	Dim adoCmd1	
	Dim adoRS1
	Dim adoCmd2	
	Dim adoRS2

	Dim adoPrices
	Dim strUserId

	strUserId = Request.Cookies("rate-monitor.com")("user_id")
		
	strConn = Session("pro_con")
	
	Rem Get the cities
	Set adoCmd = CreateObject("ADODB.Command")

	adoCmd.ActiveConnection =  strConn
	adoCmd.CommandText = "city_select"
	adoCmd.CommandType = 4
	
	'adoCmd.Parameters.Append adoCmd3.CreateParameter("@user_id", 3, 1, 0, strUserId)
		
	Set adoRS = adoCmd.Execute


%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Language" content="en-us">

<title>Rate-Monitor by Rate-Highway, Inc. | User Configuration</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<link rel="stylesheet" type="text/css" href="inc/rh_standard.css">
<link rel="stylesheet" type="text/css" href="inc/sitewide.css">
<script src="inc/sitewide.js" language="javascript"></script>
<script src="inc/header_menu_support.js" language="javascript"></script>
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
          <td><a href="search_profiles_car.asp" onMouseOver="MM_swapImage('s1','','images/b_search_pro_on.gif',1)" onMouseOut="MM_swapImgRestore()">
          <img src="images/b_search_pro_of.gif" name="s1" border="0" id="s1" width="183" height="32"></a></td>
          <td><a href="search_queue_car.asp" onMouseOver="MM_swapImage('s2','','images/b_search_que_on.gif',1)" onMouseOut="MM_swapImgRestore()">
          <img src="images/b_search_que_of.gif" name="s2" border="0" id="s2" width="97" height="32"></a></td>
          <td><a href="search_criteria_car.asp" onMouseOver="MM_swapImage('s3','','images/b_search_cri_on.gif',1)" onMouseOut="MM_swapImgRestore()">
          <img src="images/b_search_cri_of.gif" name="s3" border="0" id="s3" width="103" height="32"></a></td>
          <td><a href="javascript:not_enabled()" onMouseOver="MM_swapImage('ra','','images/b_rate_on.gif',1)" onMouseOut="MM_swapImgRestore()">
          <img src="images/b_rate_of.gif" name="ra" border="0" id="ra" width="88" height="32"></a></td>
          <td><a href="javascript:not_enabled()" onMouseOver="MM_swapImage('al','','images/b_alert_on.gif',1)" onMouseOut="MM_swapImgRestore()">
          <img src="images/b_alert_of.gif" name="al" border="0" id="al" width="53" height="32"></a></td>
          <td><a href="javascript:not_enabled()" onMouseOver="MM_swapImage('us','','images/b_user_on.gif',1)" onMouseOut="MM_swapImgRestore()">
          <img src="images/b_user_of.gif" name="us" border="0" id="us" width="126" height="32"></a></td>
          <td>
          <a href="javascript:not_enabled()" onMouseOver="MM_swapImage('sy','','images/b_system_on.gif',1)" onMouseOut="MM_swapImgRestore()">
          <img src="images/b_system_of.gif" name="sy" border="0" id="sy" width="58" height="32"></a></td>
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
                      <td><div align="right">
                        <font face="Vendana, Arial, Helvetica, sans-serif" size="2" color="#FFFFFF">
                        User: <%=Request.Cookies("rate-monitor.com")("user_name")%></font></div></td>
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
    <td background="images/h_tile.gif"><table width="100" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td><img src="images/h_user_configuration.gif" width="368" height="31"></td>
          <td><img src="images/h_right.gif" width="402" height="31"></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<table width="770" border="0" cellspacing="0" cellpadding="2">
  <tr> 
    <td width="20">&nbsp;</td>
    <td width="858" valign="top"><br>
    <font face="Vendana, Arial, Helvetica, sans-serif" size="2" > <b>
    Instructions</b><br>
    You may create, alter or delete user 
    configurations. Please note that if a user configuration is deleted, all 
    Search Profiles, Alerts and searches that user created will be deleted along 
    with the user. </td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td>
<table width="804" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td width="25" valign="top"><img src="images/separator.gif" width="25" height="15"></td>
          <td width="788">
          	&nbsp;</td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<!--
</form>
-->
<form method="POST" action="javascript:not_enabled()" name="search_alerts">
  <p>&nbsp;</p>
  <input type="hidden" name="refresh_from" value="search">
</form>
<form name="user_config" method="POST" action="javascript:not_enabled()">
  <input type="hidden" value=" " name="_D:countryDropBox"> 
          <input type="hidden" value=" " name="_D:stateDropBox"> <input type="hidden" value=" " name="_D:city">
  <table border="0" style="border-collapse: collapse" bordercolor="#FFFFFF" width="1110" cellspacing="0" height="4">
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
  <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="1110" id="AutoNumber1" background="images/alt_color.gif" height="561">
    <tr>
      <td width="11" height="19">&nbsp;</td>
      <td width="247" height="19">&nbsp;</td>
      <td width="200" height="19">&nbsp;</td>
      <td width="672" height="19">&nbsp;</td>
      <td width="169" height="19">&nbsp;</td>
    </tr>
    <tr>
      <td width="11" height="25">&nbsp;</td>
      <td width="247" height="25">
      <img border="0" src="images/user_mgt.gif" width="162" height="25"></td>
      <td width="200" height="25"><font face="Verdana" size="2">1. User Name</font></td>
      <td width="672" height="25">
      <input type="text" name="username" size="20" style="width:200; font-family:Verdana; font-size:10pt">
      </td>
      <td width="169" height="25">&nbsp;</td>
    </tr>
    <tr>
      <td width="11" height="19">&nbsp;</td>
      <td width="247" height="19">&nbsp;</td>
      <td width="200" height="19">&nbsp;</td>
      <td width="672" height="19"><font face="Verdana" size="2">(Users full 
      name, first than last)</font></td>
      <td width="169" height="19">&nbsp;</td>
    </tr>
    <tr>
      <td width="11" height="23">&nbsp;</td>
      <td width="247" height="23">&nbsp;</td>
      <td width="200" height="23"><font face="Verdana" size="2">2a. User id</font></td>
      <td width="672" height="23">
      <input type="text" name="userid" size="20" style="width:200; font-family:Verdana; font-size:10pt">
      </td>
      <td width="169" height="23">&nbsp;</td>
    </tr>
    <tr>
      <td width="11" height="19">&nbsp;</td>
      <td width="247" height="19">&nbsp;</td>
      <td width="200" height="19">&nbsp;</td>
      <td width="672" height="19"><font face="Verdana" size="2">(email name that 
      precedes the @, i.e. jsmith@company.com)</font></td>
      <td width="169" height="19">&nbsp;</td>
    </tr>
    <tr>
      <td width="11" height="19">&nbsp;</td>
      <td width="247" height="19">&nbsp;</td>
      <td width="200" height="19"><font face="Verdana" size="2">2b. Client User id</font></td>
      <td width="672" height="19">
      <input type="text" name="userid0" size="20" style="width:200; font-family:Verdana; font-size:10pt"></td>
      <td width="169" height="19">&nbsp;</td>
    </tr>
    <tr>
      <td width="11" height="19">&nbsp;</td>
      <td width="247" height="19">&nbsp;</td>
      <td width="200" height="19">&nbsp;</td>
      <td width="672" height="19"><font face="Verdana" size="2">(user id used 
      within your organization, i.e. jsmith)</font></td>
      <td width="169" height="19">&nbsp;</td>
    </tr>
    <tr>
      <td width="11" height="22">&nbsp;</td>
      <td width="247" height="22">&nbsp;</td>
      <td width="200" height="22"><font face="Verdana" size="2">3. User role</font></td>
      <td width="672" height="22">
      <select size="1" name="role" style="width:200; font-family:Verdana; font-size:10pt">
      <option selected value="1">Standard User</option>
      <option value="2">Manager</option>
      <option value="3">Administrator</option>
      </select></td>
      <td width="169" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="11" height="23">&nbsp;</td>
      <td width="247" height="23">&nbsp;</td>
      <td width="200" height="23">&nbsp;</td>
      <td width="672" height="23">
      &nbsp;</td>
      <td width="169" height="23">&nbsp;</td>
    </tr>
    <tr>
      <td width="11" height="22">&nbsp;</td>
      <td width="247" height="22">&nbsp;</td>
      <td width="200" height="22"><font face="Verdana" size="2">4. Address</font></td>
      <td width="672" height="22">
      <input type="text" name="company2" size="20" style="width:200; font-family:Verdana; font-size:10pt"></td>
      <td width="169" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="11" height="19">&nbsp;</td>
      <td width="247" height="19">&nbsp;</td>
      <td width="200" height="19">&nbsp;</td>
      <td width="672" height="19">
      <input type="text" name="company3" size="20" style="width:200; font-family:Verdana; font-size:10pt"></td>
      <td width="169" height="19">&nbsp;</td>
    </tr>
    <tr>
      <td width="11" height="22">&nbsp;</td>
      <td width="247" height="22">&nbsp;</td>
      <td width="200" height="22"><font face="Verdana" size="2">5. City</font></td>
      <td width="672" height="22">
      <input type="text" name="car_type2" size="20" style="width:200; font-family:Verdana; font-size:10pt"></td>
      <td width="169" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="11" height="19">&nbsp;</td>
      <td width="247" height="19">&nbsp;</td>
      <td width="200" height="19">&nbsp;</td>
      <td width="672" height="19">
      &nbsp;</td>
      <td width="169" height="19">&nbsp;</td>
    </tr>
    <tr>
      <td width="11" height="22">&nbsp;</td>
      <td width="247" height="22">&nbsp;</td>
      <td width="200" height="22"><font face="Verdana" size="2">6. 
      State/Territory:</font></td>
      <td width="672" height="22">
      <select size="1" name="state" style="width:200; font-family:Verdana; font-size:10pt">
          <option value selected> </option>
          <option value="AL">Alabama</option>
          <option value="AK">Alaska</option>
          <option value="AZ">Arizona</option>
          <option value="AR">Arkansas</option>
          <option value="CA">California</option>
          <option value="CO">Colorado</option>
          <option value="CT">Connecticut</option>
          <option value="DE">Delaware</option>
          <option value="DC">District of Columbia</option>
          <option value="FL">Florida</option>
          <option value="GA">Georgia</option>
          <option value="HI">Hawaii</option>
          <option value="ID">Idaho</option>
          <option value="IL">Illinois</option>
          <option value="IN">Indiana</option>
          <option value="IA">Iowa</option>
          <option value="KS">Kansas</option>
          <option value="KY">Kentucky</option>
          <option value="LA">Louisiana</option>
          <option value="ME">Maine</option>
          <option value="MD">Maryland</option>
          <option value="MA">Massachusetts</option>
          <option value="MI">Michigan</option>
          <option value="MN">Minnesota</option>
          <option value="MS">Mississippi</option>
          <option value="MO">Missouri</option>
          <option value="MT">Montana</option>
          <option value="NE">Nebraska</option>
          <option value="NV">Nevada</option>
          <option value="NH">New Hampshire</option>
          <option value="NJ">New Jersey</option>
          <option value="NM">New Mexico</option>
          <option value="NY">New York</option>
          <option value="NC">North Carolina</option>
          <option value="ND">North Dakota</option>
          <option value="OH">Ohio</option>
          <option value="OK">Oklahoma</option>
          <option value="OR">Oregon</option>
          <option value="PA">Pennsylvania</option>
          <option value="PR">Puerto Rico</option>
          <option value="RI">Rhode Island</option>
          <option value="SC">South Carolina</option>
          <option value="SD">South Dakota</option>
          <option value="TN">Tennessee</option>
          <option value="TX">Texas</option>
          <option value="UT">Utah</option>
          <option value="VT">Vermont</option>
          <option value="VA">Virginia</option>
          <option value="WA">Washington</option>
          <option value="WV">West Virginia</option>
          <option value="WI">Wisconsin</option>
          <option value="WY">Wyoming</option>
          <option value="AB">Alberta</option>
          <option value="BC">British Columbia</option>
          <option value="MB">Manitoba</option>
          <option value="NB">New Brunswick</option>
          <option value="NL">Newfoundland/Labrador</option>
          <option value="NT">Northwest Territories</option>
          <option value="NS">Nova Scotia</option>
          <option value="NU">Nunavut</option>
          <option value="ON">Ontario</option>
          <option value="PE">Prince Edward Island</option>
          <option value="QC">Quebec</option>
          <option value="SK">Saskatchewan</option>
          <option value="YT">Yukon</option>
      </select></td>
      <td width="169" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="11" height="19">&nbsp;</td>
      <td width="247" height="19">&nbsp;</td>
      <td width="200" height="19">&nbsp;</td>
      <td width="672" height="19">&nbsp;</td>
      <td width="169" height="19">&nbsp;</td>
    </tr>
    <tr>
      <td width="11" height="22">&nbsp;</td>
      <td width="247" height="22">&nbsp;</td>
      <td width="200" height="22"><font face="Verdana" size="2">7. Country</font></td>
      <td width="672" height="22">
      <select size="1" name="country" style="width:200; font-family:Verdana; font-size:10pt">
          <option value="US" selected>United States</option>
          <option value="CA">Canada</option>
          <option value="DZ">Algeria</option>
          <option value="AR">Argentina</option>
          <option value="AU">Australia</option>
          <option value="AT">Austria</option>
          <option value="BS">Bahamas</option>
          <option value="BH">Bahrain</option>
          <option value="BE">Belgium</option>
          <option value="BR">Brazil</option>
          <option value="BG">Bulgaria</option>
          <option value="CM">Cameroon</option>
          <option value="CA">Canada</option>
          <option value="CN">China</option>
          <option value="CO">Colombia</option>
          <option value="CR">Costa Rica</option>
          <option value="AN">Curacao</option>
          <option value="CY">Cyprus</option>
          <option value="CZ">Czech Republic</option>
          <option value="DK">Denmark</option>
          <option value="DO">Dominican Republic</option>
          <option value="EC">Ecuador</option>
          <option value="EG">Egypt</option>
          <option value="ET">Ethiopia</option>
          <option value="FI">Finland</option>
          <option value="FR">France</option>
          <option value="DE">Germany</option>
          <option value="GR">Greece</option>
          <option value="GU">Guam</option>
          <option value="HK">Hong Kong</option>
          <option value="HU">Hungary</option>
          <option value="IN">India</option>
          <option value="IE">Ireland</option>
          <option value="IL">Israel</option>
          <option value="IT">Italy</option>
          <option value="JM">Jamaica</option>
          <option value="JP">Japan</option>
          <option value="KE">Kenya</option>
          <option value="KW">Kuwait</option>
          <option value="LU">Luxembourg</option>
          <option value="MG">Madagascar</option>
          <option value="MY">Malaysia</option>
          <option value="MV">Maldives</option>
          <option value="MT">Malta</option>
          <option value="MU">Mauritius</option>
          <option value="MX">Mexico</option>
          <option value="MA">Morocco</option>
          <option value="NL">Netherlands</option>
          <option value="NZ">New Zealand</option>
          <option value="NG">Nigeria</option>
          <option value="OM">Oman</option>
          <option value="PE">Peru</option>
          <option value="PR">Puerto Rico</option>
          <option value="RO">Romania</option>
          <option value="LC">Saint Lucia</option>
          <option value="SA">Saudi Arabia</option>
          <option value="SG">Singapore</option>
          <option value="ZA">South Africa</option>
          <option value="KR">South Korea</option>
          <option value="ES">Spain</option>
          <option value="LK">Sri Lanka</option>
          <option value="SE">Sweden</option>
          <option value="CH">Switzerland</option>
          <option value="TH">Thailand</option>
          <option value="TT">Trinidad/Tobago</option>
          <option value="TR">Turkey</option>
          <option value="AE">U.A.E</option>
          <option value="GB">United Kingdom</option>
          <option value="US">United States</option>
          <option value="UY">Uruguay</option>
          <option value="VE">Venezuela</option>
          <option value="VN">Vietnam</option>
      </select></td>
      <td width="169" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="11" height="19">&nbsp;</td>
      <td width="247" height="19">&nbsp;</td>
      <td width="200" height="19">&nbsp;</td>
      <td width="672" height="19">&nbsp;</td>
      <td width="169" height="19">&nbsp;</td>
    </tr>
    <tr>
      <td width="11" height="22">&nbsp;</td>
      <td width="247" height="22">&nbsp;</td>
      <td width="200" height="22"><font face="Verdana" size="2">8. Allowable 
      Search<br>
&nbsp;&nbsp;&nbsp; Locations</font></td>
      <td width="672" height="22">
      <select size="5" name="city_codes" multiple style="width:200; font-family:Verdana; font-size:10pt">
                    <% While adoRS.EOF = False %>
              <option value="<%=adoRS.Fields("city_cd").Value %>" ><%=adoRS.Fields("city_name").Value & " (" & adoRS.Fields("city_cd").Value & ")" %>
              </option>
              <%	adoRS.MoveNext
					   Wend
					   
					   Set adoRS  = Nothing
					   Set adoCmd = Nothing
					   
					%>
      </select></td>
      <td width="169" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="11" height="19">&nbsp;</td>
      <td width="247" height="19">&nbsp;</td>
      <td width="200" height="19">&nbsp;</td>
      <td width="672" height="19"><font face="Verdana" size="2">(enter 
      airport/city codes &quot;any&quot; for any location) </font></td>
      <td width="169" height="19">&nbsp;</td>
    </tr>
    <tr>
      <td width="11" height="22">&nbsp;</td>
      <td width="247" height="22">&nbsp;</td>
      <td width="200" height="22"><font face="Verdana" size="2">9. Email address</font>:</td>
      <td width="672" height="22">
      <input type="text" name="recipient" size="20" style="width:200; font-family:Verdana; font-size:10pt"></td>
      <td width="169" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="11" height="19">&nbsp;</td>
      <td width="247" height="19">&nbsp;</td>
      <td width="200" height="19">&nbsp;</td>
      <td width="672" height="19">&nbsp;</td>
      <td width="169" height="19">&nbsp;</td>
    </tr>
    <tr>
      <td width="11" height="22">&nbsp;</td>
      <td width="247" height="22">&nbsp;</td>
      <td width="200" height="22"><font face="Verdana" size="2">10. Search Type</font></td>
      <td width="672" height="22">
      <select size="1" name="search_type" style="width:200; font-family:Verdana; font-size:10pt">
      <option selected value="1">As searched (all searches)</option>
      <option value="2">Link to Profile</option>
      </select></td>
      <td width="169" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="11" height="61">&nbsp;</td>
      <td width="247" height="61">&nbsp;</td>
      <td width="872" colspan="2" height="61">
      <font size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;<input name="submit_alert" type="submit" id="Open2225" value="    Create   " class="rh_button"></font></td>
      <td width="169" height="61">&nbsp;</td>
    </tr>
    <tr>
      <td width="11" height="19">&nbsp;</td>
      <td width="247" height="19">&nbsp;</td>
      <td width="872" colspan="2" height="19">&nbsp;</td>
      <td width="169" height="19">&nbsp;</td>
    </tr>
  </table>
  <table border="0" style="border-collapse: collapse" bordercolor="#FFFFFF" width="1110" cellspacing="0" height="4">
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
</form>
<!--#INCLUDE FILE="footer.asp"-->
</body>
</html>