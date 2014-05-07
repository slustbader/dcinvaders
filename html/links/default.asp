<%
	Option Explicit
	'On Error Resume Next
	
	Dim dbConn, rs, strSQL
	Dim strLinkName, strLinkAddress, strLinkDescription, strLinkImage
	
	Set dbConn = Server.CreateObject("ADODB.Connection")
	'dbConn.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("/invaders/database/hockey.mdb")
	dbConn.Open "DSN=hockey"
	
	Set rs = Server.CreateObject("ADODB.Recordset")
	strSQL = "SELECT * FROM tbl_Links"
	rs.Open strSQL, dbConn
%>
<!--#include virtual="/include/adovbs.inc"-->
<HTML>
<HEAD>
<TITLE>DC Invaders</TITLE>
<LINK REL="stylesheet" HREF="/styles/hockey.css" TYPE="text/css">
</HEAD>
<BODY BACKGROUND="/images/stars.gif">
<TABLE WIDTH=660 CELLPADDING=0 CELLSPACING=0 BORDER=0>
<TR>
	<TD><CENTER><IMG SRC="/images/invader.gif" BORDER=0>
		<H2><FONT COLOR="#FFFF66">Links</FONT></H2></CENTER>

<%

		If Not rs.BOF AND Not rs.EOF Then
			rs.MoveFirst
			
			Do While Not rs.EOF
				strLinkName = rs("LinkName")
				strLinkAddress = rs("LinkAddress")
				strLinkDescription = rs("LinkDescription")
%>				
				
				

		<P>
		<A HREF="<%=strLinkAddress %>" TARGET="New"><%=strLinkName %> -</A>
		<TABLE WIDTH=100% CELLPADDING=0 CELLSPACING=0 BORDER=0>
		<TR>
			<TD WIDTH=20>&nbsp;</TD>
			<TD><%= strLinkDescription %></TD>
		</TR>
		</TABLE>
		</P>
<%
				rs.MoveNext
			Loop
		End If
		
		rs.Close
		Set rs = nothing
		
		dbConn.Close
		Set dbConn = nothing
%>

		<BR><BR>
	</TD>
</TR>
</TABLE>
</BODY>
</HTML>