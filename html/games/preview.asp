<%
	Option Explicit
	'On Error Resume Next
	
	Function FormatText(strText)
		
		Dim strFormattedText
		
		strFormattedText = Replace(strText, Chr(13), "<P>")
		FormatText = strFormattedText
	
	End Function

	Dim strGameID
	
	strGameID = Request.QueryString("ID")
	
	Dim dbConn, rs, strSQL
	
	Set dbConn = Server.CreateObject("ADODB.Connection")
	'dbConn.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("/invaders/database/hockey.mdb")
	dbConn.Open "DSN=hockey"
	
	strSQL = "SELECT tbl_Games.GameID, tbl_Games.Recap, tbl_Games.PreviewTitle, tbl_Games.Preview, tbl_Games.LocationID, tbl_HomeTeams.HomeTeamName, tbl_AwayTeams.AwayTeamName, tbl_Locations.LocationName, tbl_Games.GameDate, tbl_Games.GameTime FROM tbl_Locations INNER JOIN (tbl_HomeTeams INNER JOIN (tbl_AwayTeams INNER JOIN tbl_Games ON tbl_AwayTeams.AwayTeamID = tbl_Games.AwayTeamID) ON tbl_HomeTeams.HomeTeamID = tbl_Games.HomeTeamID) ON tbl_Locations.LocationID = tbl_Games.LocationID WHERE GameID=" & strGameID
	rs = dbConn.Execute(strSQL)
	
%>
<!-- #include virtual="/include/adovbs.inc"-->
<HTML>
<HEAD>
<LINK REL="stylesheet" HREF="/styles/hockey.css" TYPE="text/css">
</HEAD>
<BODY BACKGROUND="/images/stars.gif">
<TABLE WIDTH=660 CELLPADDING=0 CELLSPACING=0 BORDER=0>
<TR>
	<TD><CENTER><IMG SRC="/images/invader.gif" BORDER=0>
		<H2><FONT COLOR="#FFFF66">Game Preview</FONT></H2></CENTER>

		<CENTER>
		<TABLE BORDER=0 WIDTH=100% CELLPADDING=2 CELLSPACING=0>
		<TR BGCOLOR="#FFFF33">
			<TD ALIGN=RIGHT NOWRAP CLASS="black"><SMALL><B>Preview</B> - 
<%
			If Now() > rs("GameDate")+1 Then
				response.write "<A HREF=""summary.asp?ID=" & rs("GameID") & """>Box Score</A> - "
			Else
				response.write "<FONT COLOR=""#C0C0C0"">Box Score</FONT> - "
			End If
			
			If IsNull(rs("Recap")) Or rs("Recap")="" Then
				response.write "<FONT COLOR=""#C0C0C0"">Recap</FONT>"
			Else
				response.write "<A HREF=""recap.asp?ID=" & rs("GameID") & """>Recap</A>"
			End If
%>					
			</SMALL></TD>
		</TR>
		</TABLE>
		</CENTER>
		
		<BR>
		<FONT COLOR="#FF0000" SIZE=+1><%=rs("AwayTeamName") %> at <%=rs("HomeTeamName") %></FONT><BR>
		<% =FormatDateTime(rs("GameDate"),1) %><BR>
		<% =FormatDateTime(rs("GameTime"),3) %><BR>
		<A HREF="/rinks/default.asp?ID=<%=rs("LocationID") %>"><% =rs("LocationName") %></A>
		<H3><%= rs("PreviewTitle")%></H3>
		<%= FormatText(rs("Preview")) %>
		<BR><BR><BR>
<%
		dbConn.Close
		Set dbConn = nothing
%>
</BODY>
</HTML>