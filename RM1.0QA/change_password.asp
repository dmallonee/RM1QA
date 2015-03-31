<%@ Language=VBScript %>
<!-- #INCLUDE FILE="inc/adovbs.asp" -->
<% Response.Expires = -1
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
	Dim strMsg

	Dim adoPrices
	Dim strUserId
	
	strMsg = ""
	
	if Request.Form("login") <> "" Then
	If Request.Form("email_address") <> "" And Request.Form("password") <> "" And Request.Form("new_password") <> "" AND Request.Form("new_password") = Request.Form("rpt_password") Then

		strConn = Session("pro_con")
	
		Set adoCmd = CreateObject("ADODB.Command")

		adoCmd.ActiveConnection =  strConn
		adoCmd.CommandText = "change_password"
		adoCmd.CommandType = 4
	
		adoCmd.Parameters.Append adoCmd.CreateParameter("RETURN_VALUE",   adInteger ,adParamReturnValue)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@email_address", 200, 1, 50, Request.Form("email_address"))
		adoCmd.Parameters.Append adoCmd.CreateParameter("@old_password",  200, 1, 50, Request.Form("password"))
		adoCmd.Parameters.Append adoCmd.CreateParameter("@new_password",  200, 1, 50, Request.Form("new_password"))
		
		Set adoRS = adoCmd.Execute
	
	
		If adoRS.EOF = False Then
			Server.Transfer "change_password_success.asp"
	
		Else
			Server.Transfer "change_password_fail.asp"
	
		End If
	Else
		Server.Transfer "change_password_fail.asp"
	End If
	End If
	
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<title>Rate-Monitor.com | Welcome and please login</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<link rel="stylesheet" type="text/css" href="inc/sitewide.css">
<script src="inc/sitewide.js" language="javascript"></script>
<link rel="stylesheet" type="text/css" href="inc/rh_standard.css">
<script language="JavaScript" type="text/JavaScript">
<!--
    function UpdateParent(UserId) {

        myCreator = self.opener; // window that opened me
        myCreator.document.login_form.email_address.value = UserId;
    }

    function checkform() {
        var why = "";
        why += checkUsername(document.login_form.email_address.value);
        why += checkPassword(document.login_form.password.value);
        why += checkNewPassword(document.login_form.new_password.value);
        why += checkRptPassword(document.login_form.rpt_password.value, document.login_form.new_password.value);

        if (why != "") {
            alert(why);
            return false;
        }
        return true;
    }

    function checkNewPassword(strng) {
        var error = "";
        if (strng == "") {
            error = "Please enter a new password.\n";
            login_form.new_password.style.background = 'rgba(162,0,0,0.5)';
        }
        else {
            login_form.new_password.style.background = 'White';
        }
        return error;
    }

    function checkRptPassword(strng, strng2) {
        var error = "";
        if (strng != strng2 || strng == "") {
            error = "Please repeat your new password.\n";
            login_form.rpt_password.style.background = 'rgba(162,0,0,0.5)';
        }
        else {
            login_form.rpt_password.style.background = 'White';
        }
        return error;
    }

    function checkPassword(strng) {
        var error = "";
        if (strng == "") {
            error = "Please enter your old password.\n";
            login_form.password.style.background = 'rgba(162,0,0,0.5)';
        }
        else {
            login_form.password.style.background = 'White';
        }
        return error;
    }
    function checkUsername(strng) {
        var error = "";
        if (strng == "") {
            error = "Please enter your user name.\n";
            login_form.email_address.style.background = 'rgba(162,0,0,0.5)';
        }
        else {
            login_form.email_address.style.background = 'White';
        }
        return error;
    }
-->
</script>
</head>

<body>

<form method="POST" action="change_password.asp" name="login_form" onsubmit="return checkform();" class="login">
  <table width="590" border="0" cellspacing="0" cellpadding="0" id="table1">
    <tr>
      <td width="100%" valign="top">&nbsp;<table width="100%" border="0" cellspacing="0" cellpadding="0" id="table2">
        <tr>
          <td><img src="images/ruler.gif" width="600" height="2"></td>
        </tr>
      </table>
      <p><font color="#800000"><b>Change Password</b></font><br>
      <font size="2" face="Arial, Helvetica, sans-serif">To change your password, 
      simply enter your current password and your new password below and <br>
      click
      the change button. </font></p>
      <table width="590" border="0" cellspacing="0" cellpadding="0" id="table3" background="images/alt_color.gif">
        <tr>
          <td><img src="images/ruler.gif" width="590" height="2"></td>

        </tr>
        </table> 
      <table width="590" border="0" cellspacing="0" cellpadding="0" id="table4" background="images/alt_color.gif">
        
        <tr>
          <td></td>
          <td></td>
          <td></td>
          <td></td>
        </tr>
        <tr valign="bottom">
          <td width="200"><font size="2" face="Arial, Helvetica, sans-serif">
          <img src="images/ti_log_in.gif"></font></td>
          <td width="128"><font size="2" face="Arial, Helvetica, sans-serif">User Name:</font> </td>
          <td width="166"><font size="2" face="Arial, Helvetica, sans-serif">
          <input type="text" name="email_address" size="20" tabindex="1" onfocus="this.className='focus';cl(this,'email');" onblur="this.className='';fl(this,'email');" style="width: 150"></font></td>
          <td width="95"><font size="2" face="Arial, Helvetica, sans-serif">
          <input name="login" type="submit" id="Open2224" name="Open2224" value="Change" tabindex="3" class="rh_button"></font></td>
        </tr>
        <tr valign="bottom">
          <td width="200">&nbsp;</td>
          <td width="128">
          <p align="left"><font size="2" face="Arial, Helvetica, sans-serif">Current Password:</font> </p>
          </td>
          <td width="166"><font size="2" face="Arial, Helvetica, sans-serif">
          <input type="password" name="password" size="20" tabindex="2" onfocus="this.className='focus';cl(this,'');" onblur="this.className='';fl(this,'');" style="width: 150"></font></td>
          <td width="95"><font size="2" face="Arial, Helvetica, sans-serif">&nbsp;</font></td>
        </tr>
        <tr valign="bottom">
          <td width="200"></td>
          <td width="128"><font size="2" face="Arial, Helvetica, sans-serif">New Password:</font> </td>
          <td width="166"><font size="2" face="Arial, Helvetica, sans-serif">
          <input type="password" name="new_password" size="20" tabindex="2" onfocus="this.className='focus';cl(this,'');" onblur="this.className='';fl(this,'');" style="width: 150"></font></td>
          <td width="95">&nbsp;</td>
        </tr>
        <tr valign="bottom">
          <td width="200">&nbsp;</td>
          <td width="128"><font size="2" face="Arial, Helvetica, sans-serif">Repeat New Pwd:</font> </td>
          <td width="166"><font size="2" face="Arial, Helvetica, sans-serif">
          <input type="password" name="rpt_password" size="20" tabindex="2" onfocus="this.className='focus';cl(this,'');" onblur="this.className='';fl(this,'');" style="width: 150"></font></td>
          <td width="95">&nbsp;</td>
        </tr>
        <tr valign="bottom">
          <td width="200">&nbsp;</td>
          <td width="128">&nbsp;</td>
          <td width="166">&nbsp;</td>
          <td width="95">&nbsp;</td>
        </tr>
      </table>
      <table width="590" border="0" cellspacing="0" cellpadding="0" id="table9">
        <tr>
          <td><img src="images/ruler.gif" width="590" height="2"></td>
        </tr>
      </table>
      <p align="center"><font size="3" face="Arial, Helvetica, sans-serif"><br>
      &nbsp;</font><font size="2" face="Arial, Helvetica, sans-serif">If you 
      don't want to change your password click here to
      <a href="javascript:setTimeout( self.close(), 1 );">cancel</a></font></p>
      <table width="590" border="0" cellspacing="0" cellpadding="0" id="table10">
        <tr>
          <td><img src="images/ruler.gif" width="590" height="2"></td>
        </tr>
      </table>
      <p>&nbsp;</p>
      </td>
    </tr>
  </table>
</form>
<!--#INCLUDE FILE="footer.asp"-->
</body>
</html>