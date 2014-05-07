<%
	Option Explicit
	'On Error Resume Next

	Function FormatText(strText)
		If Instr(strText, Chr(13)) Then
			strText = Replace(strText, Chr(13), "<P>")
		End If
		FormatText = strText
	End Function

	Dim dbConn, rs, strSQL, strScoreSQL, rsScore

	Set dbConn = Server.CreateObject("ADODB.Connection")
	'dbConn.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("/invaders/database/hockey.mdb")
	dbConn.Open "DSN=hockey"

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
	<TD><CENTER><IMG SRC="/images/invader.gif" BORDER=0>
		<H2><FONT COLOR="#FFFF66">News</FONT></H2></CENTER>
<%
	Dim strNewsText, i

	strSQL = "SELECT NewsTitle, NewsText, Duration FROM tbl_News WHERE ((([PostDate]+[Duration])>=Now())) ORDER BY PostDate DESC"
	Set rs = dbConn.Execute(strSQL)

	'rs.Filter = "#" & Now() & "# < ([PostDate] + [Duration])"

	If Not rs.BOF and Not rs.EOF Then

		rs.MoveFirst

		response.write "<H3><FONT COLOR=""red"">Top Stories</FONT></H3>"

		Do While Not rs.EOF
			strNewsText = rs("NewsText")

			response.write "<B><FONT COLOR=""yellow"">" & rs("NewsTitle") & "</FONT></B><BR>"
			response.write FormatText(strNewsText) & "<BR>"
			response.write "<HR>"

			rs.MoveNext
		Loop

	End If

	Dim strPreview, strRecap

	strSQL = "SELECT GameID, GameDate, GameTime, HomeTeamName, AwayTeamName, tbl_Locations.LocationID, LocationName, tbl_HomeTeams.HomeTeamID, tbl_AwayTeams.AwayTeamID, Preview, Recap, Scored FROM tbl_Locations INNER JOIN (tbl_HomeTeams INNER JOIN (tbl_AwayTeams INNER JOIN tbl_Games ON tbl_AwayTeams.AwayTeamID = tbl_Games.AwayTeamID) ON tbl_HomeTeams.HomeTeamID = tbl_Games.HomeTeamID) ON tbl_Locations.LocationID = tbl_Games.LocationID ORDER BY GameDate"
	Set rs = dbConn.Execute(strSQL)
	rs.Filter = "GameDate > #" & Now()-1 & "# And GameDate < #" & Now()+7 & "#"
	If Not rs.BOF AND Not rs.EOF Then
		rs.MoveFirst

		response.write "<TABLE WIDTH=100% CELLPADDING=2 CELLSPACING=0 BORDER=1 BORDERCOLOR=""#C0C0C0""><TR><TD BGCOLOR=""#0000FF""><FONT SIZE=+1><EM>Upcoming Games</EM></FONT></TD></TR></TABLE>"

		i = 0

		response.write "<TABLE WIDTH=100% CELLPADDING=2 CELLSPACING=0 BORDER=0>"
		response.write "<TR>"
		response.write "<TD WIDTH=20 BGCOLOR=""#999999"">&nbsp;</TD>"
		response.write "<TD BGCOLOR=""#999999"" CLASS=""black"">"

		Do While Not rs.EOF
			strPreview = rs("Preview")
			If i > 0 Then
				response.write "<BR><HR COLOR=""#08479E""><BR>"
			End If
			response.write "<B>" & rs("AwayTeamName") & " vs. " & rs("HomeTeamName") & "</B><BR>"
			response.write "<A HREF=""../rinks/default.asp?ID=" & rs("LocationID") & """>" & rs("LocationName") & "</A><BR>"
			response.write FormatDateTime(rs("GameDate"),1) & " - "
			If DatePart("h", rs("GameTime")) > 12 Then
				response.write DatePart("h",rs("GameTime"))-12 & ":"
				If DatePart("n",rs("GameTime")) < 10 Then
					response.write "0"
				End If
				response.write DatePart("n", rs("GameTime"))
				response.write " PM<BR>"
			Else
				response.write DatePart("h", rs("GameTime")) & ":"
				If DatePart("n", rs("GameTime")) < 10 Then
					response.write "0"
				End If
				response.write DatePart("n", rs("GameTime")) & " AM<BR>"
			End If
			If Not IsNull(strPreview) AND strPreview <> "" Then
				response.write "<A HREF=""/games/preview.asp?ID=" & rs("GameID") & """>Preview</A>"
			End If

			i = i + 1
			rs.MoveNext
		Loop
		response.write "</TD>"
		response.write "</TR>"
		response.write "</TABLE>"
	End If

		response.write "<BR><BR>"
	rs.Filter = "GameDate < #" & Now()-1 & "# AND GameDate > #" & Now()-7 & "# AND Scored=TRUE"
	If Not rs.BOF AND Not rs.EOF Then
		rs.MoveFirst

		response.write "<TABLE WIDTH=100% CELLPADDING=2 CELLSPACING=0 BORDER=1 BORDERCOLOR=""#C0C0C0""><TR><TD BGCOLOR=""#FF0000""><FONT SIZE=+1><EM>Recent Games</EM></FONT></TD></TR></TABLE>"

		response.write "<TABLE WIDTH=100% CELLPADDING=2 CELLSPACING=0 BORDER=0>"
		response.write "<TR>"
		response.write "<TD WIDTH=20 BGCOLOR=""#999999"">&nbsp;</TD>"
		response.write "<TD BGCOLOR=""#999999"" CLASS=""black"">"

		i = 0
		Do While Not rs.EOF
			strRecap = rs("recap")
			strPreview = rs("Preview")

		 	If i > 0 Then
		 		response.write "<BR><HR COLOR=""#08479E""><BR>"
		 	End If

		 	strScoreSQL = "SELECT tbl_HomeTeams.HomeTeamName, tbl_AwayTeams.AwayTeamName, [FirstPeriodHomeGoals]+[SecondPeriodHomeGoals]+[ThirdPeriodHomeGoals]+[OvertimeHomeGoals] AS HomeScore, [FirstPeriodAwayGoals]+[SecondPeriodAwayGoals]+[ThirdPeriodAwayGoals]+[OvertimeAwayGoals] AS AwayScore FROM tbl_HomeTeams INNER JOIN (tbl_AwayTeams INNER JOIN tbl_Games ON tbl_AwayTeams.AwayTeamID = tbl_Games.AwayTeamID) ON tbl_HomeTeams.HomeTeamID = tbl_Games.HomeTeamID WHERE (((tbl_Games.GameID)=" & rs("GameID") & "))"
		 	Set rsScore = dbConn.Execute(strScoreSQL)

			If rsScore("AwayScore") > rsScore("HomeScore") Then
				response.write "<B>" & rs("AwayTeamName") & " <FONT COLOR=""#0000FF"">" & rsScore("AwayScore") & "</FONT>, " & rs("HomeTeamName") & " <FONT COLOR=""#0000FF"">" & rsScore("HomeScore") & "</FONT></B><BR>"
			Else
				response.write "<B>" & rs("HomeTeamName") & " <FONT COLOR=""#0000FF"">" & rsScore("HomeScore") & "</FONT>, " & rs("AwayTeamName") & " <FONT COLOR=""#0000FF"">" & rsScore("AwayScore") & "</FONT></B><BR>"
			End If

			response.write "<A HREF=""../rinks/default.asp?ID=" & rs("LocationID") & """>" & rs("LocationName") & "</A><BR>"
			response.write FormatDateTime(rs("GameDate"),1) & " - "
			If DatePart("h", rs("GameTime")) > 12 Then
				response.write DatePart("h",rs("GameTime"))-12 & ":"
				If DatePart("n",rs("GameTime")) < 10 Then
					response.write "0"
				End If
				response.write DatePart("n", rs("GameTime"))
				response.write " PM<BR>"
			Else
				response.write DatePart("h", rs("GameTime")) & ":"
				If DatePart("n", rs("GameTime")) < 10 Then
					response.write "0"
				End If
				response.write DatePart("n", rs("GameTime")) & " AM<BR>"
			End If
			If Not IsNull(strPreview) AND strPreview <> "" Then
				response.write "<A HREF=""/dcinvaders/games/preview.asp?ID=" & rs("GameID") & """>Preview</A> - "
			End If

			If Not IsNull(strRecap) AND strRecap <> "" Then
				response.write "<A HREF=""../games/recap.asp?ID=" & rs("GameID") & """>Recap</A> - "
			End If
			response.write "<A HREF=""../games/summary.asp?ID=" & rs("GameID") & """>Box Score</A><BR>"

			i = i + 1
			rs.MoveNext
		Loop

		response.write "</TD>"
		response.write "</TR>"
		response.write "</TABLE>"

	End If

	response.write "<BR><BR>"

	strSQL = "SELECT PracticeDate, LocationName, tbl_Locations.LocationID, StartTime, EndTime, tbl_Practices.Comments FROM tbl_Practices INNER JOIN tbl_Locations ON tbl_Practices.PracticeLocationID = tbl_Locations.LocationID"
	Set rs = dbConn.Execute(strSQL)

	Dim strComments

	rs.Filter = "PracticeDate > #" & Now()-1 & "# And PracticeDate < #" & Now()+7 & "#"

	If Not rs.BOF AND Not rs.EOF Then

		response.write "<TABLE WIDTH=100% CELLPADDING=2 CELLSPACING=0 BORDER=1 BORDERCOLOR=""#C0C0C0""><TR><TD BGCOLOR=""#FFFF66"" CLASS=""black""><FONT SIZE=+1><EM>Upcoming Practices</EM></FONT></TD></TR></TABLE>"

		rs.MoveFirst

		response.write "<TABLE WIDTH=100% CELLPADDING=2 CELLSPACING=0 BORDER=0>"
		response.write "<TR>"
		response.write "<TD WIDTH=20 BGCOLOR=""#999999"">&nbsp;</TD>"
		response.write "<TD BGCOLOR=""#999999"" CLASS=""black"">"

		i = 0

		Do While Not rs.EOF

			strComments = rs("Comments")

		 	If i > 0 Then
		 		response.write "<BR><HR COLOR=""#08479E""><BR>"
		 	End If

			response.write "<B>"& FormatDateTime(rs("PracticeDate"),1) & "</B><BR>"
			response.write "Time: " & FormatDateTime(rs("StartTime"),3) & " - " & FormatDateTime(rs("EndTime"),3) & "<BR>"
			response.write "Location: <A HREF=""../rinks/default.asp?ID=" & rs("LocationID") & """>" & rs("LocationName") & "</A><BR>"
			If Not IsNull(strComments) AND strComments <> "" Then
				response.write "<B>Comments:</B> " & FormatText(strComments)
			End If

			rs.MoveNext

			i = i + 1
		Loop

		response.write "</TD>"
		response.write "</TR>"
		response.write "</TABLE>"
	End If

	'rs.Filter = "PracticeDate < #" & Now()-1 & "# And PracticeDate > #" & Now()-7 & "#"

	'If Not rs.BOF AND Not rs.EOF Then

	'	response.write "<H3>Recent Practices</H3>"

	'	rs.MoveFirst

	'	Do While Not rs.EOF
	'		strComments = rs("Comments")

	'		response.write "<B STYLE=""color: #FF0000"">"& FormatDateTime(rs("PracticeDate"),1) & "</B><BR>"
	'		response.write "Time: " & FormatDateTime(rs("StartTime"),3) & " - " & FormatDateTime(rs("EndTime"),3) & "<BR>"
	'		response.write "Location: <A HREF=""../rinks/default.asp?ID=" & rs("LocationID") & """>" & rs("LocationName") & "<BR>"
	'		If Not IsNull(strComments) AND strComments <> "" Then
	'			response.write "<B>Comments:</B> " & FormatText(strComments)
	'		End If
	'		rs.MoveNext
	'	Loop
	'	response.write "<HR>"
	'End If

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