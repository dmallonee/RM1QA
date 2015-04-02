<%@ Language=VBScript %>
<!-- #INCLUDE FILE="inc/adovbs.asp" -->
<% Response.Expires = -1
   Response.cachecontrol="private" 
   Response.AddHeader "pragma", "no-cache"

   Server.ScriptTimeout = 180


	Dim	strUser 
	Dim strPassword 
	
	strUser = Request.Form("email_address")
	strPassword = Request.Form("new_password")
	
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
    function jumpTo() {

        var NewAction
        var SelIndex

        SelIndex = document.login_form.system.selectedIndex
        NewAction = document.login_form.system.options[selectedIndex].text

        document.login_form.action = NewAction

    }

    function CheckLoginAlerts() {
        if ("ANY" != "ANY") {
            //alert("No alerts at this time.");  
            return true;
        }
        return true;
    }

    function UpdateParent(UserId, Password) {

        myCreator = self.opener; // window that opened me
        myCreator.document.login_form.email_address.value = UserId;
        myCreator.document.login_form.password.value = Password;

    }
    //-->
</script>
</head>

<body>

<form method="POST" action="change_password_success.asp" name="login_form" onsubmit="CheckLoginAlerts();return true" class="login">
  <table width="590" border="0" cellspacing="0" cellpadding="0" id="table1">
    <tr>
      <td width="100%" valign="top">&nbsp;<table width="100%" border="0" cellspacing="0" cellpadding="0" id="table2">
        <tr>
          <td><img src="images/ruler.gif" width="600" height="2"></td>
        </tr>
      </table>
      <p><font color="#800000"><b>Your Password has been Changed</b></font><br>
      <font size="2" face="Arial, Helvetica, sans-serif">Click the okay button 
      to close this window and automatically enter your new login information<br>
      in the login screen.</font></p>
      <p align="center"><font face="Arial, Helvetica, sans-serif" size="2">
      <a href='javascript:UpdateParent("<%=strUser %>", "<%=strPassword %>");setTimeout( self.close(), 1 );'>
      Return to login screen</a></font></p>
      <p>&nbsp;</p>
      <p>&nbsp;</p>
      <table width="590" border="0" cellspacing="0" cellpadding="0" id="table10">
        <tr>
          <td><img src="images/ruler.gif" width="590" height="2"></td>
        </tr>
      </table>
      <p>&nbsp;</p>
      </td>
    </tr>
  </table>
  <input type="hidden" name="email_address" value="<%=strUser %>">
  <input type="hidden" name="password" value="<%=strPassword %>">
</form>
<!--#INCLUDE FILE="footer.asp"-->
</body>
</html>