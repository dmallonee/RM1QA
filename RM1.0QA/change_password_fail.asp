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

	If Request.Form("email_address") <> "" And Request.Form("password") <> "" And Request.Form("new_password") <> "" Then

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
	
	
	If adoCmd.Parameters("RETURN_VALUE").Value = 1 Then
		strMsg = "Change successful"
	
	Else
		strMsg = "change failed"
	
	
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
        myCreator.document.login_form.email_address.value = Password;

    }
    //-->
</script>
</head>

<body>

<form method="POST" action="change_password_fail.asp" name="login_form" onsubmit="CheckLoginAlerts();return true" class="login">
  <table width="590" border="0" cellspacing="0" cellpadding="0" id="table1">
    <tr>
      <td width="100%" valign="top">&nbsp;<table width="100%" border="0" cellspacing="0" cellpadding="0" id="table2">
        <tr>
          <td><img src="images/ruler.gif" width="600" height="2"></td>
        </tr>
      </table>
      <p><font color="#800000"><b>Changing Password Failed</b></font><br>
      &nbsp;</p>
      <p><font size="2" face="Arial, Helvetica, sans-serif">To try again
      <a href="change_password.asp">click here</a></font></p>
      <p><font face="Arial, Helvetica, sans-serif" size="2">To cancel and go 
      back to the login page <a href="javascript:setTimeout( self.close(), 1 );">click here</a><br>
&nbsp;</font></p>
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