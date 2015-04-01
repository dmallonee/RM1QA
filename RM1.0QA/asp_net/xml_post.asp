<%@ language=vbscript %>
<%
Option Explicit
'-----------------------------------------------------------------------
' Define Variables
'-----------------------------------------------------------------------
Dim xmlExample, XMLhttp, xmlDoc2, Response_Doc
'-----------------------------------------------------------------------
' HTML
'-----------------------------------------------------------------------
%>
<html>
<head>
<title>XML Post</title>
</head>
<body>
<pre>
<%
'-----------------------------------------------------------------------
' Open XML document
'-----------------------------------------------------------------------
Set xmlExample = CreateObject("MSXML2.DOMDocument")
xmlExample.SetProperty "ServerHTTPRequest", False
xmlExample.Async = False
xmlExample.Load "E:\aspsms\xml\request1.xml"
'-----------------------------------------------------------------------
' Check XML document
'-----------------------------------------------------------------------
If Not xmlExample.ParseError = 0 Then
  Response.Write "<b>Error Code:</b> " & xmlExample.ParseError & "<br>"
  Response.Write "<b>Error Description:</b> " & xmlExample.ParseError.reason & "<br>"
  Response.Write "<b>Error File Position:</b> " & xmlExample.ParseError.filepos & "<br>"
  Response.Write "<b>Error Line:</b> " & xmlExample.ParseError.line & "<br>"
  Response.Write "<b>Error Line Position:</b> " & xmlExample.ParseError.linepos & "<br>"
  Response.Write "<b>Error Source Text:</b> " & xmlExample.ParseError.srcText & "<br>"
Else
  '-----------------------------------------------------------------------
  ' Send XML request
  '-----------------------------------------------------------------------
  Set XMLhttp = CreateObject("MSXML2.ServerXMLHTTP")
  XMLhttp.Open "POST", "http://xml1.aspsms.com:5061/xmlsvr.asp", False
  XMLhttp.Send xmlExample.xml
  '-----------------------------------------------------------------------
  ' Get server status
  '-----------------------------------------------------------------------
  Response.Write "<br>"
  Response.Write "<b>xmlSvr Server Status:</b><br>"
  Response.Write "-----------------------------------------<br>"
  Response.Write "<b>Status (Value must be 200): </b>" & XMLhttp.status & "<br>"
  Response.Write "<b>ReadyState (Value must be 4): </b>" & XMLhttp.ReadyState & "<br>"
  Response.Write "<b>StatusText (Value must be OK): </b>" & XMLhttp.StatusText & "<br>"
  Response.Write "<b>AllResponseHeaders:</b><br>" & XMLhttp.GetAllResponseHeaders & "<br>"
  '-----------------------------------------------------------------------
  ' Get XML response from xmlSvr
  '-----------------------------------------------------------------------
  Set xmlDoc2 = CreateObject("MSXML2.DOMDocument")
  xmlDoc2.setProperty "ServerHTTPRequest", True
  xmlDoc2.async = False
  xmlDoc2.LoadXML XMLhttp.ResponseXML.xml
  Response.Write "<br>"
  Response.Write "<b>xmlSvr XML Response:</b><br>"
  Response.Write "-----------------------------------------<br>"
  Response_Doc = xmlhttp.responseXML.xml
  Response_Doc = Replace (Response_Doc,"<","&lt;")
  Response_Doc = Replace (Response_Doc,">","&gt;")
  Response.Write Response_Doc & "<br>"
End If
%>
</pre>
</body>
</html>

