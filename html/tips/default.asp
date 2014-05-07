<%
	Option Explicit
	'On Error Resume Next

	Function FormatText(strText)
		strText = Replace(strText, Chr(13), "<P>")
		FormatText = strText
	
	End Function
	
	Dim dbConn, rs, strSQL
	
	Set dbConn = Server.CreateObject("ADODB.Connection")
	'dbConn.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("/invaders/database/hockey.mdb")
	dbConn.Open "DSN=hockey"
	
	strSQL = "SELECT TipTitle, TipText, TipDate FROM tbl_Tips ORDER BY TipDate ASC"
	Set rs = Server.CreateObject("ADODB.Recordset")
	
	rs.Open strSQL, dbConn
%>
<!-- #include virtual="/include/adovbs.inc"-->
<HTML>
<HEAD>
<TITLE>Hockey Tips</TITLE>
<LINK REL="stylesheet" HREF="/styles/hockey.css" TYPE="text/css">
</HEAD>
<BODY BACKGROUND="/images/stars.gif">
<TABLE WIDTH=660 CELLPADDING=0 CELLSPACING=0 BORDER=0>
<TR>
	<TD><CENTER><IMG SRC="/images/invader.gif" BORDER=0>
		<H2><FONT COLOR="#FFFF66">Tips<SUP>*</SUP></FONT></H2></CENTER>

		
		<table>
<%
	While not rs.EOF 
%>
		<tr><td>
		<B STYLE="color: #FF0000">(<%=rs("TipDate") %>) <%=rs("TipTitle") %></B><BR>
		<TABLE WIDTH=100% CELLPADDING=0 CELLSPACING=0 BORDER=0>
		<TR>
			<TD WIDTH=15>&nbsp;</TD>
			<TD><%=FormatText(rs("TipText")) %>
			</TD>
		</TR>
		</TABLE>
		</td></tr>
<%
		rs.MoveNext
	Wend
%>
		</table>
		<BR><BR>
		<P><SMALL>*Tips are provided by users at <A HREF="http://www.eteamz.com/">eteamz</A> bulliten boards.</SMALL></P>
		<BR><BR><BR>
	</TD>
</TR>
</TABLE>
<%
	rs.Close
	Set rs = nothing
	
	dbConn.Close
	Set dbConn=nothing
%>
</BODY>
</HTML>