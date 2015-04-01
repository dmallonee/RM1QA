<%@ language=vbscript%>
<HTML>
<HEAD>
<TITLE>ServerInfo.asp</TITLE>
<STYLE>
  BODY 
  { font-family:arial, helvetica, "sans-serif"; font-size:12px; }
  
  TABLE 
  { font-family:arial, helvetica, "sans-serif"; 
    font-size:12px; 
    border-width:1px;
    border-collapse:collapse;
  }

  TD
  { font-family:arial, helvetica, "sans-serif";
    font-size:12px;
    border-width:thin;
    border-style:groove;
    border-collapse:collapse;
    padding-left:4px;
    padding-right:4px;
    padding-top:3px;
    vertical-align:top;
  }


</STYLE>
</head>
<body bgcolor="#FFFFFF">

<TABLE>
<SCRIPT>
  document.write("<TR><TD>Cookies Enabled</TD><TD><b>" + window.navigator.cookieEnabled + "</b></TD></TR>");
</SCRIPT>
</TABLE>

<script language=jscript runat=server>
response.write("<TABLE>");
response.write("<TR><TD>JScript Server Scripting Engine</TD><TD><b>" + ScriptEngine() + "</b></TD></TR>");
response.write("<TR><TD>JScript buildversion</TD><TD><b>" + ScriptEngineBuildVersion() + "</b></TD></TR>");
response.write("<TR><TD>JScript majorversion</TD><TD><b>" + ScriptEngineMajorVersion() + "</b></TD></TR>");
response.write("<TR><TD>JScript minorversion</TD><TD><b>" + ScriptEngineMinorVersion() + "</b></TD></TR>");
response.write("</TABLE>");
</script>

<%
Function OutputServerVariable(varName)
  Dim varResult
  response.write("<TR><TD>")
  response.write(varName)
  response.write("</TD><TD>")
  varResult = request.servervariables(varName)
  response.write(varResult)

  'if(varResult = nul) then
  '  response.write("(NULL)")
  'end if

  response.write("</TD></TR>" & vbcrlf)
End Function

response.write("<TABLE>")
response.write "<TR><TD>ASP Server Script buildversion</TD><TD><b>" & scriptenginebuildversion() & "</b></TD></TR>" & vbcrlf
response.write "<TR><TD>ASP Server Script majorversion</TD><TD><b>" & scriptenginemajorversion() & "</b></TD></TR>" & vbcrlf
response.write "<TR><TD>ASP Server Script minorversion</TD><TD><b>" & scriptengineminorversion() & "</b></TD></TR>" & vbcrlf

set tempconn=server.createobject("adodb.connection")
response.write "<TR><TD>ado version</TD><TD><b>"  & vbcrlf
response.write tempconn.version & "</b></TD></TR>" & vbcrlf
set tempconn=nothing

serversoftware=request.servervariables("server_software")
response.write "<TR><TD>server software</TD><TD><b>"  & vbcrlf
response.write serversoftware & "</b></TD></TR>" & vbcrlf

Response.Write "<TR><TD>Script Timeout</TD><TD><b>" & Server.ScriptTimeout & " seconds</b></TD></TR>" & vbcrlf
Response.Write "<TR><TD>Session Timeout</TD><TD><b>" & Session.Timeout & " minutes</b></TD></TR>" & vbcrlf


OutputServerVariable("APPL_MD_PATH")
OutputServerVariable("APPL_PHYSICAL_PATH")

OutputServerVariable("AUTH_TYPE")
OutputServerVariable("AUTH_PASSWORD")
OutputServerVariable("AUTH_USER")

OutputServerVariable("CERT_FLAGS")
OutputServerVariable("CERT_ISSUER")
OutputServerVariable("CERT_KEYSIZE")
OutputServerVariable("CERT_SECRETKEYSIZE")
OutputServerVariable("CERT_SERIALNUMBER")
OutputServerVariable("CERT_SERVER_ISSUER")
OutputServerVariable("CERT_SERVER_SUBJECT")
OutputServerVariable("CERT_SUBJECT")

OutputServerVariable("CONTENT_LENGTH")
OutputServerVariable("CONTENT_TYPE")
OutputServerVariable("DATE_GMT")
OutputServerVariable("DATE_LOCAL")

OutputServerVariable("DOCUMENT_NAME")
OutputServerVariable("DOCUMENT_URI")
OutputServerVariable("GATEWAY_INTERFACE")

OutputServerVariable("HTTP_ACCEPT")
OutputServerVariable("HTTP_ACCEPT_LANGUAGE")
OutputServerVariable("HTTP_ACCEPT_ENCODING")
OutputServerVariable("HTTP_CONNECTION")
OutputServerVariable("HTTP_COOKIE")
OutputServerVariable("HTTP_HOST")
OutputServerVariable("HTTP_REFERER")
OutputServerVariable("HTTP_USER_AGENT")

OutputServerVariable("HTTPS")
OutputServerVariable("HTTPS_KEYSIZE")
OutputServerVariable("HTTPS_SECRETKEYSIZE")
OutputServerVariable("HTTPS_SERVER_ISSUER")
OutputServerVariable("HTTPS_SERVER_SUBJECT")

OutputServerVariable("INSTANCE_ID")
OutputServerVariable("INSTANCE_META_PATH")

OutputServerVariable("LOCAL_ADDR")
OutputServerVariable("LAST_MODIFIED")
OutputServerVariable("LOGON_USER")

OutputServerVariable("PATH_INFO")
OutputServerVariable("PATH_TRANSLATED")
OutputServerVariable("QUERY_STRING")
OutputServerVariable("QUERY_STRING_UNESCAPED")
OutputServerVariable("REMOTE_ADDR")
OutputServerVariable("REMOTE_HOST")
OutputServerVariable("REMOTE_PORT")
OutputServerVariable("REMOTE_IDENT")
OutputServerVariable("REMOTE_USER")
OutputServerVariable("REQUEST_METHOD")


OutputServerVariable("SCRIPT_NAME")
OutputServerVariable("SERVER_NAME")
OutputServerVariable("SERVER_PORT")
OutputServerVariable("SERVER_PORT_SECURE")
OutputServerVariable("SERVER_PROTOCOL")
OutputServerVariable("SERVER_SOFTWARE")

OutputServerVariable("UNMAPPED_REMOTE_USER")


OutputServerVariable("URL")

OutputServerVariable("ALL_HTTP")
OutputServerVariable("ALL_RAW")

Response.Write ("<TR><TD>ASP Session ID</TD><TD>" & Session.SessionID & "</TD></TD>")


response.write("</TABLE>") & vbcrlf
%>
Load <A href="ServerInfo.asp">ServerInfo.asp</A> again.<BR>

<COMMENT>The text below may not be removed</COMMENT>
Copyright © 2004, <A href="http://www.spectrum-research.com">Spectrum Research Inc.</A>, All Rights Reserved.<BR>
Written by Alfred J. Heyman. May not be sold or included in a commercial product without written permission.
</BODY>
</HTML>

