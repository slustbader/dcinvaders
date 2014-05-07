<%
	Option Explicit
	On Error Resume Next

	Dim fmSeasonID
	Dim dbConn, strCon, rs, strSQL
	Dim intTable
	Dim blnAway

	fmSeasonID = Request.QueryString("SeasonID")

	Set dbConn = Server.CreateObject("ADODB.Connection")
	'strCon = "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("/invaders/database/hockey.mdb")
	'dbConn.Open strCon
	dbConn.Open "DSN=hockey"
	
	Set rs = Server.CreateObject("ADODB.Recordset")

	If fmSeasonID = "" Or IsNull(fmSeasonID) Then
		strSQL = "SELECT tbl_Games.GameID, tbl_Locations.LocationName, tbl_Locations.LocationID, tbl_Games.HomeTeamID, tbl_Games.AwayTeamID, tbl_Games.GameDate, tbl_Games.GameTime, ([FirstPeriodHomeGoals]+[SecondPeriodHomeGoals]+[ThirdPeriodHomeGoals]+[OvertimeHomeGoals]) AS HomeScore, ([FirstPeriodAwayGoals]+[SecondPeriodAwayGoals]+[ThirdPeriodAwayGoals]+[OvertimeAwayGoals]) AS AwayScore, Recap, Scored, tbl_Seasons.SeasonName, tbl_Seasons.StartDate AS SeasonStart, tbl_Seasons.EndDate AS SeasonEnd FROM tbl_Seasons INNER JOIN (tbl_Locations INNER JOIN tbl_Games ON tbl_Locations.LocationID = tbl_Games.LocationID) ON tbl_Seasons.SeasonID = tbl_Games.SeasonID WHERE (((tbl_Seasons.StartDate)<Now()) AND ((tbl_Seasons.EndDate)>Now())) ORDER BY GameDate"
	Else
		strSQL = "SELECT tbl_Games.GameID, tbl_Locations.LocationName, tbl_Locations.LocationID, tbl_Games.HomeTeamID, tbl_Games.AwayTeamID, tbl_Games.GameDate, tbl_Games.GameTime, ([FirstPeriodHomeGoals]+[SecondPeriodHomeGoals]+[ThirdPeriodHomeGoals]+[OvertimeHomeGoals]) AS HomeScore, ([FirstPeriodAwayGoals]+[SecondPeriodAwayGoals]+[ThirdPeriodAwayGoals]+[OvertimeAwayGoals]) AS AwayScore, Recap, Scored, tbl_Seasons.SeasonName, tbl_Seasons.StartDate AS SeasonStart, tbl_Seasons.EndDate AS SeasonEnd FROM tbl_Seasons INNER JOIN (tbl_Locations INNER JOIN tbl_Games ON tbl_Locations.LocationID = tbl_Games.LocationID) ON tbl_Seasons.SeasonID = tbl_Games.SeasonID WHERE tbl_Games.SeasonID=" & fmSeasonID & " ORDER BY GameDate"
	End If

	rs.Open strSQL, dbConn, adOpenStatic

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<!--#include virtual="/include/adovbs.inc"-->
<HTML>
	<HEAD>
		<TITLE>DC Invaders</TITLE>
		<LINK REL="stylesheet" HREF="/styles/hockey.css" TYPE="text/css">
	</HEAD>
	<BODY BACKGROUND="/images/stars.gif">
		<TABLE CELLPADDING="0" CELLSPACING="0" BORDER="0">
			<TR>
				<TD><CENTER><IMG SRC="/images/invader.gif" BORDER="0">
						<H2><FONT COLOR="#FFFF66">Invaders Game Schedule</FONT></H2>
						<BR>
						<H3><% If Not rs.BOF AND Not rs.EOF Then Response.write rs("SeasonName") Else Response.write "The Invaders are currently relaxing in the offseason." %></H3>
					</CENTER>
					<%
	If Not rs.BOF AND Not rs.EOF Then
		Response.write "<CENTER>"
		Response.write "<TABLE CELLPADDING=2 CELLSPACING=0 BORDER=1 BORDERCOLOR=""black"" BGCOLOR=""#EEEEEE"" rules=""all"">"
		Response.write "<TR BGCOLOR=""#08479E"">"
			Response.write "<TD><B>Date</B></TD>"
			Response.write "<TD><B>Opponent</B></TD>"
			Response.write "<TD><B>Location</B></TD>"
			Response.write "<TD><B>Home/Away</B></TD>"
			Response.write "<TD><B>Time/Result</B></TD>"
		Response.write "</TR>"

		intTable = 0
		Do While Not rs.EOF
				If intTable Mod 2 = 1 Then
					response.write "<TR BGCOLOR=""#DCDCDC"">"
				Else
					response.write "<TR BGCOLOR=""#C0C0C0"">"
				End If

				Response.Write "<TD CLASS=""black"">" & FormatDateTime(rs("GameDate"),1) & "</TD>"
				Response.Write "<TD CLASS=""black"">"
				If rs("HomeTeamID") <> 1 Then
					strSQL = "SELECT TeamName FROM tbl_Teams WHERE TeamID=" & rs("HomeTeamID")
					blnAway = "TRUE"
				End If

				If rs("AwayTeamID") <> 1 Then
					strSQL = "SELECT TeamName FROM tbl_Teams WHERE TeamID=" & rs("AwayTeamID")
					blnAway = "FALSE"
				End If

				Dim rs2
				Set rs2 = dbConn.Execute(strSQL)

				response.write rs2("TeamName")

				Set rs2 = nothing
				Response.write "</TD>"
				Response.write "<TD CLASS=""black""><A HREF=""../rinks/default.asp?ID=" & rs("LocationID") & """>" & rs("LocationName") & "</A></TD>"

				If blnAway = "TRUE" Then
					Response.write "<TD ALIGN=CENTER><IMG SRC=""../images/tinyawayjersey.gif"" WIDTH=20 HEIGHT=17 BORDER=0 ALT=""Away""></TD>"
				Else
					Response.write "<TD ALIGN=CENTER><IMG SRC=""../images/tinyhomejersey.gif"" WIDTH=20 HEIGHT=17 BORDER=0 ALT=""Home""></TD>"
				End If

				Dim strResult

					strResult=""

					If rs("GameDate")<(Now()) AND rs("Scored")=TRUE Then
						Dim strHomeScore
						Dim strAwayScore

						strHomeScore = rs("HomeScore")
						strAwayScore = rs("AwayScore")
						If strHomeScore = strAwayScore then
							strResult = "T"
						End If

						If (rs("HomeTeamID")=1 AND strHomeScore>strAwayScore) OR (rs("AwayTeamID")=1 AND strAwayScore>strHomeScore) Then
							strResult = "W"
						End If

						If (rs("HomeTeamID")=1 AND strHomeScore<strAwayScore) OR (rs("AwayTeamID")=1  AND strAwayScore<strHomeScore) Then
							strResult = "L"
						End If

						Response.Write "<TD CLASS=""black"">"
							If IsNull(rs("Recap")) or rs("Recap") = "" Then
								response.write "<A HREF=""summary.asp?ID=" & rs("GameID") & """>" & strResult & " "
							Else
								response.write "<A HREF=""recap.asp?ID=" & rs("GameID") & """>" & strResult & " "
							End If

							If strAwayScore > strHomeScore Then
								response.write strAwayScore & "-" & strHomeScore
							Else
								response.write strHomeScore & "-" & strAwayScore
							End If
						Response.Write "</A></TD>"

					Else
						response.write "<TD CLASS=""black"">" & rs("GameTime") & "</TD>"
					End If

			Response.write "</TR>"

			intTable = intTable + 1
			rs.MoveNext
		Loop

		Response.Write "</TABLE>"
		Response.Write "</CENTER>"
	End If
%>
	<p>View this schedule in Google Calendar: <a target="_blank" href="http://www.google.com/calendar/render?cid=dcinvaders%40gmail.com"><img src="http://www.google.com/calendar/images/ext/gc_button1_en.gif" border=0></a>

	<p>&nbsp;</p>
<%

	strSQL = "Select * FROM tbl_Seasons WHERE EndDate<Now()"

	If fmSeasonID <> "" Then
		strSQL = strSQL &  " AND SeasonID<>" & fmSeasonID
	End If

	Set rs = dbConn.Execute(strSQL)
	response.write "<P><B>"
	If fmSeasonID <> "" Then
		response.write "Other "
	End If
	response.write "Past Season Schedules</B>:"
	response.write "<UL>"

	Do While Not rs.EOF
		response.write "<LI><A HREF=""default.asp?SeasonID=" & rs("SeasonID") & """>" & rs("SeasonName") & "</A></LI>"
		rs.MoveNext
	Loop

	response.write "</UL>"

	If fmSeasonID <> "" Then
		response.write "<TABLE CELLPADDING=0 CELLSPACING=0 BORDER=0><TR><TD ALIGN=RIGHT>"
		response.write "<A HREF=""default.asp"">Return to current schedule</A>"
		response.write "</TD></TR></TABLE>"
	End If


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
