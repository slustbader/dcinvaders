<%
	Option Explicit
	On Error Resume Next
	
	Function FormatPhone(strNumber)
		strNumber = "(" & Left(strNumber,3) & ") " & Mid(strNumber,4,3) & "-" & Right(strNumber,4)
		FormatPhone = strNumber	
	End Function

	Dim dbConn, rs, strSQL
	
	Set dbConn = Server.CreateObject("ADODB.Connection")
	'dbConn.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("/invaders/database/hockey.mdb")
	dbConn.Open "DSN=hockey"
	
	Set rs = Server.CreateObject("ADODB.Recordset")
	
	strSQL = "SELECT * FROM tbl_Locations"
	
	If Not IsEmpty(Request.QueryString("ID")) Then
		strSQL = strSQL & " WHERE LocationID=" & Request.QueryString("ID")
	End If
	
	rs.Open strSQL, dbConn
	
%>
<!-- #include virtual="/include/adovbs.inc"-->
<HTML>
<HEAD>
<TITLE>DC Invaders</TITLE>
<LINK REL="stylesheet" HREF="/styles/hockey.css" TYPE="text/css">
</HEAD>
<BODY BACKGROUND="/images/stars.gif">
<TABLE WIDTH=660 CELLPADDING=0 CELLSPACING=0 BORDER=0>
<TR>
	<TD><CENTER><IMG SRC="../images/invader.gif" BORDER=0>
		<H2><FONT COLOR="#FFFF66">Rinks</FONT></H2></CENTER>

<%
		rs.MoveFirst
		
		Do While Not rs.EOF
%>
			<P><% If Not IsNull(rs("Website")) Then response.write "<A HREF=""" & rs("Website") & """ TARGET=""NEW"""">" %>
			<%=rs("LocationName") %> -
			<% If Not IsNull(rs("Website")) Then response.write "</A>" %>
			<TABLE WIDTH=100% CELLPADDING=0 CELLSPACING=0 BORDER=0>
			<TR>
				<TD WIDTH=20>&nbsp;</TD>
				<TD><%=rs("Address1") %><BR>
					<%	If Not IsNull(rs("Address2")) Then response.write rs("Address2") & "<BR>" %>
					<%=rs("City") %>, <%=rs("State") & " " %><%=rs("Zip") %><BR>
					<%=FormatPhone(rs("Phone")) %><BR>
					<BR>
					<%=rs("Comments") %><BR>
					<% If Not IsNull(rs("MapLink")) And rs("MapLink")<>"" Then response.write "<A HREF=""" & rs("MapLink") & """>Map</A>" %>
				</TD>
			</TR>
			</TABLE>
			</P>

			<HR COLOR="#08479E">
<%			
			rs.MoveNext
		Loop
		
		rs.Close
		Set rs = nothing
		
		dbConn.Close
		Set dbConn = nothing
%>
	</TD>
</TR>
</TABLE>
</BODY>
</HTML>