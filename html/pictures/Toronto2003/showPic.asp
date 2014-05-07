<%@LANGUAGE=JSCRIPT%>
<HTML>
<HEAD>
	<TITLE></TITLE>
</HEAD>
<BODY TOPMARGIN=0 LEFTMARGIN=0 ONFOCUSOUT="self.focus();">
<%
	strImage = Request.QueryString("img");
%>
<IMG SRC="<%=strImage%>">
<BR>
<CENTER><INPUT TYPE=BUTTON ONCLICK="window.close();" VALUE=Close></CENTER>
</BODY>
</HTML>