<%
	Option Explicit
	'On Error Resume Next
	
	Function FormatText(strText)
		
		Dim strFormattedText
		
		strFormattedText = Replace(strText, Chr(13), "<P>")
		FormatText = strFormattedText
	
	End Function

	Dim dbConn, rs, strSQL
	Dim strGameID
	Dim strHomeTeam, strAwayTeam
	Dim strRegHomeGoals, strRegAwayGoals, strTotalHomeGoals, strTotalAwayGoals
	Dim strOT
	
	strGameID = Request.QueryString("ID")
	
	Set dbConn = Server.CreateObject("ADODB.Connection")
	'dbConn.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("/invaders/database/hockey.mdb")
	dbConn.Open "DSN=hockey"

	strSQL = "SELECT tbl_Games.*, tbl_HomeTeams.HomeTeamName, tbl_AwayTeams.AwayTeamName FROM tbl_AwayTeams INNER JOIN (tbl_HomeTeams INNER JOIN tbl_Games ON tbl_HomeTeams.HomeTeamID = tbl_Games.HomeTeamID) ON tbl_AwayTeams.AwayTeamID = tbl_Games.AwayTeamID WHERE GameID=" & strGameID	
	rs = dbConn.Execute(strSQL)

	'Team Info	
	strHomeTeam = rs("HomeTeamName")
	strAwayTeam = rs("AwayTeamName")

	'Goals
	strRegHomeGoals = rs("FirstPeriodHomeGoals") + rs("SecondPeriodHomeGoals") + rs("ThirdPeriodHomeGoals")
	strRegAwayGoals = rs("FirstPeriodAwayGoals") + rs("SecondPeriodAwayGoals") + rs("ThirdPeriodAwayGoals")
	strTotalHomeGoals = strRegHomeGoals + rs("OvertimeHomeGoals")
	strTotalAwayGoals = strRegAwayGoals + rs("OvertimeAwayGoals")
	
	If strRegHomeGoals = strRegAwayGoals Then
		strOT = 1
	End If
	
%>
<!-- #include virtual="/include/adovbs.inc"-->
<HTML>
<HEAD>
<LINK REL="stylesheet" HREF="/styles/hockey.css" TYPE="text/css">
</HEAD>
<BODY BACKGROUND="/images/stars.gif">
<% if Request.QueryString("ID") = 100 Then Response.Write " <BGSOUND SRC=""/media/mforbidd.wav"" LOOP=1>" %>
<TABLE WIDTH=660 CELLPADDING=0 CELLSPACING=0 BORDER=0>
<TR>
	<TD><CENTER><IMG SRC="/images/invader.gif" BORDER=0>
		<H2><FONT COLOR="#FFFF66">Game Recap</FONT></H2></CENTER>

		<CENTER>
		<TABLE BORDER=0 WIDTH=100% CELLPADDING=2 CELLSPACING=0>
		<TR BGCOLOR="#FFFF33">
			<TD ALIGN=RIGHT NOWRAP CLASS="black"><SMALL>
<%
					If IsNull(rs("Preview")) or rs("Preview") = "" Then
						response.write "<FONT COLOR=""#9999CC"">Preview - </FONT>"
					Else
						response.write "<A HREF=""preview.asp?ID=" & rs("GameID") & """>Preview</A> - "
					End If
%>
					<A HREF="summary.asp?ID=<%=rs("GameID") %>">Box Score</A> - <B>Recap</B>

					</SMALL></TD>
				</TR>
				</TABLE>
		</CENTER>
		
		<H3><FONT COLOR="#FF0000">
<%
		If strTotalHomeGoals > strTotalAwayGoals then
			response.write strHomeTeam & " " & strTotalHomeGoals & ", " & strAwayTeam & " " & strTotalAwayGoals
		Else
			response.write strAwayTeam & " " & strTotalAwayGoals & ", " & strHomeTeam & " " & strTotalHomeGoals
		End If
		
		If strOT = 1 Then
			response.write " (OT) "
		End If
%>
		</FONT>
		</H3>
			

		<H3><%= rs("RecapTitle")%></H3>
		<%= FormatText(rs("Recap")) %>
		<P><A HREF="summary.asp?ID=<%=strGameID %>">Go to box score.</A></P>
		<BR><BR><BR>
<%
		dbConn.Close
		Set dbConn = nothing
%>
</BODY>
</HTML>