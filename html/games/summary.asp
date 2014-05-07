<%
	Option Explicit
	On Error Resume Next

	Sub Scoring(strGameID, strPeriod)
		If strPeriod = "OT" then
			strPeriod=4
		End If
		strSQL = "SELECT Time, Assist1PlayerID, Assist2PlayerID, PPG, SH, EN, GW, GT, LastName, Number, TeamName, tbl_Players.PlayerID FROM (tbl_Goals INNER JOIN tbl_Players ON tbl_Goals.PlayerID = tbl_Players.PlayerID) INNER JOIN tbl_Teams ON tbl_Players.TeamID = tbl_Teams.TeamID WHERE GameID=" & strGameID & " AND Period=" & strPeriod & " ORDER BY Time DESC"
		Set rs = dbConn.Execute(strSQL)
		If Not rs.BOF AND Not rs.EOF Then
			rs.MoveFirst
			i = 0
			Do While Not rs.EOF
				If i > 0 Then
					response.write "<BR>"
				End If
				response.write "<B>" & rs("TeamName") & "</B> - " & FormatDateTime(rs("Time"),4) & ", "

				response.write "<A HREF=""../roster/default.asp?ID=" & rs("PlayerID") & """>"
				If Not IsNull(rs("LastName")) Then
					response.write rs("LastName")
				Else
					response.write rs("Number")
				End If
				response.write "</A>"
				If rs("PPG") = TRUE Then response.write " (Power Play)"
				If rs("SH") = TRUE Then response.write " (Short Handed)"
				If rs("EN") = TRUE Then response.write " (Empty Net)"
				If rs("Assist1PlayerID") <> 0 Then
					strSQL = "SELECT LastName, Number FROM tbl_Players WHERE PlayerID=" & rs("Assist1PlayerID")
					Set rs2 = dbConn.Execute(strSQL)
					response.write " (<A HREF=""../roster/default.asp?ID=" & rs("Assist1PlayerID") & """>"
					If Not IsNull(rs2("LastName")) Then
						response.write rs2("LastName")
					Else
						response.write rs2("Number")
					End If
					response.write "</A>"

					If rs("Assist2PlayerID") <> 0 Then
						strSQL = "SELECT LastName, Number FROM tbl_Players WHERE PlayerID=" & rs("Assist2PlayerID")
						Set rs2 = dbConn.Execute(strSQL)
						response.write ", <A HREF=""../roster/default.asp?ID=" & rs("Assist2PlayerID") & """>"

						If Not IsNull(rs2("LastName")) Then
							response.write rs2("LastName")
						Else
							response.write rs2("Number")
						End If
						response.write "</A>"

					End If
					response.write ")"
				Else
					response.write " (unassisted)"
				End If
				i = i + 1
				rs.MoveNext
			Loop
		Else
			response.write "None"
		End If
	End Sub

	Sub Penalties(strGameID, strPeriod)
		strSQL = "SELECT tbl_Teams.TeamName, tbl_Penalties.Start, tbl_Players.LastName, tbl_Penalties.PlayerID, tbl_Players.Number, pick_Penalties.PenaltyName, pick_Periods.Period FROM pick_Periods INNER JOIN (tbl_Teams INNER JOIN (tbl_Players INNER JOIN (pick_Penalties INNER JOIN tbl_Penalties ON pick_Penalties.PenaltyID = tbl_Penalties.PenaltyCharge) ON tbl_Players.PlayerID = tbl_Penalties.PlayerID) ON tbl_Teams.TeamID = tbl_Players.TeamID) ON pick_Periods.ID = tbl_Penalties.Period WHERE (((tbl_Penalties.GameID)=" & strGameID & ") AND ((pick_Periods.Period)='" & strPeriod & "')) ORDER BY tbl_Penalties.Start DESC;"
		Set rs = dbConn.Execute(strSQL)
		If Not rs.BOF AND Not rs.EOF Then
			rs.MoveFirst
			i = 0
			Do While Not rs.EOF
				If i > 0 Then
					response.write "<BR>"
				End If
				response.write "<B>" & rs("TeamName") & "</B> - " & FormatDateTime(rs("Start"),4) & ", "
				response.write "<A HREF=""../roster/default.asp?ID=" & rs("PlayerID") & """>"
				If Not IsEmpty(rs("LastName")) Then
					response.write rs("LastName")
				Else
					response.write rs("Number")
				End If
				response.write "</A>"

				response.write " (" & rs("PenaltyName") & ")"
				i = i + 1
				rs.MoveNext
			Loop
		Else
			response.write "None"
		End If
	End Sub

	Dim dbConn, rs, rs2, strSQL, strGameID
	Dim strHomeTeamID, strAwayTeamID, strHomeTeam, strAwayTeam
	Dim strFirstPeriodHomeShots, strFirstPeriodAwayShots, strSecondPeriodHomeShots, strSecondPeriodAwayShots, strThirdPeriodHomeShots, strThirdPeriodAwayShots, strOvertimeHomeShots, strOvertimeAwayShots
	Dim strFirstPeriodHomeGoals, strFirstPeriodAwayGoals, strSecondPeriodHomeGoals, strSecondPeriodAwayGoals, strThirdPeriodHomeGoals, strThirdPeriodAwayGoals, strRegTotalHomeGoals, strRegTotalAwayGoals, strOvertimeHomeGoals, strOvertimeAwayGoals, strTotalHomeGoals, strTotalAwayGoals
	Dim strOT
	Dim strAttendance, strReferee1, strReferee2, strScorekeeper, strNotes
	Dim strTable
	Dim i

	strGameID=Request.QueryString("ID")

	Set dbConn = Server.CreateObject("ADODB.Connection")
	'dbConn.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("/invaders/database/hockey.mdb")
	dbConn.Open "DSN=hockey"

	Set rs = Server.CreateObject("ADODB.Recordset")
	Set rs2 = Server.CreateObject("ADODB.Recordset")

	strSQL = "SELECT tbl_Games.*, tbl_HomeTeams.HomeTeamName, tbl_AwayTeams.AwayTeamName FROM tbl_AwayTeams INNER JOIN (tbl_HomeTeams INNER JOIN tbl_Games ON tbl_HomeTeams.HomeTeamID = tbl_Games.HomeTeamID) ON tbl_AwayTeams.AwayTeamID = tbl_Games.AwayTeamID WHERE GameID=" & strGameID
	Set rs = dbConn.Execute(strSQL)

	'Team Info
	strHomeTeamID = rs("HomeTeamID")
	strAwayTeamID = rs("AwayTeamID")
	strHomeTeam = rs("HomeTeamName")
	strAwayTeam = rs("AwayTeamName")

	'Shots
	strFirstPeriodHomeShots = rs("FirstPeriodHomeShots")
	strFirstPeriodAwayShots = rs("FirstPeriodAwayShots")

	strSecondPeriodHomeShots = rs("SecondPeriodHomeShots")
	strSecondPeriodAwayShots = rs("SecondPeriodAwayShots")

	strThirdPeriodHomeShots = rs("ThirdPeriodHomeShots")
	strThirdPeriodAwayShots = rs("ThirdPeriodAwayShots")

	strOvertimeHomeShots = rs("OvertimeHomeShots")
	strOvertimeAwayShots = rs("OvertimeAwayShots")

	'Goals
	strFirstPeriodHomeGoals = rs("FirstPeriodHomeGoals")
	strFirstPeriodAwayGoals = rs("FirstPeriodAwayGoals")

	strSecondPeriodHomeGoals = rs("SecondPeriodHomeGoals")
	strSecondPeriodAwayGoals = rs("SecondPeriodAwayGoals")

	strThirdPeriodHomeGoals = rs("ThirdPeriodHomeGoals")
	strThirdPeriodAwayGoals = rs("ThirdPeriodAwayGoals")

	strRegTotalHomeGoals = strFirstPeriodHomeGoals + strSecondPeriodHomeGoals + strThirdPeriodHomeGoals
	strRegTotalAwayGoals = strFirstPeriodAwayGoals + strSecondPeriodAwayGoals + strThirdPeriodAwayGoals

	strOvertimeHomeGoals = rs("OvertimeHomeGoals")
	strOvertimeAwayGoals = rs("OvertimeAwayGoals")

	strTotalHomeGoals = strRegTotalHomeGoals + strOvertimeHomeGoals
	strTotalAwayGoals = strRegTotalAwayGoals + strOvertimeAwayGoals

	'Notes Info
	strAttendance = rs("Attendance")
	strReferee1 = rs("Referee1")
	strReferee2 = rs("Referee2")
	strScorekeeper = rs("Scorekeeper")
	strNotes = rs("Notes")

%>
<!--#include virtual="/include/adovbs.inc"-->
<HTML>
<HEAD>
<TITLE>DC Invaders Game Summary</TITLE>
<LINK REL="stylesheet" HREF="/styles/hockey.css" TYPE="text/css">
</HEAD>
<BODY BACKGROUND="/images/stars.gif">
<TABLE WIDTH=660 CELLPADDING=0 CELLSPACING=0 BORDER=0>
<TR>
	<TD><CENTER><IMG SRC="/images/invader.gif" BORDER=0>
		<H2><FONT COLOR="#FFFF66">Game Recap</FONT></H2></CENTER>

		<CENTER>
		<TABLE BORDER=0 WIDTH=90% CELLPADDING=2 CELLSPACING=0>
		<TR BGCOLOR="#FFFF33">
			<TD ALIGN=RIGHT NOWRAP CLASS="black"><SMALL>
<%
					If IsNull(rs("Preview")) or rs("Preview") = "" Then
						response.write "<FONT COLOR=""#C0C0C0"">Preview - </FONT>"
					Else
						response.write "<A HREF=""preview.asp?ID=" & rs("GameID") & """>Preview</A> - "
					End If
%>
					<B>Box Score</B> -
<%
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
		<CENTER>
		<TABLE WIDTH=90% CELLPADDING=2 CELLSPACING=0 BORDER=1 BORDERCOLOR="#003399">
		<TR BGCOLOR="#003399">
			<TD><B>Final</B></TD>
			<TD ALIGN=CENTER><B>1st</B></TD>
			<TD ALIGN=CENTER><B>2nd</B></TD>
			<TD ALIGN=CENTER><B>3rd</B></TD>
<%
	If strRegTotalHomeGoals=strRegTotalAwayGoals Then
		strOT=1
%>
			<TD ALIGN=CENTER><B>OT</B></TD>
<%
	End If
%>
			<TD ALIGN=CENTER><B>Total</B></TD>
		</TR>

		<TR BGCOLOR="#FFFFFF">
			<TD CLASS="black" ><%=strAwayTeam %></TD>
			<TD CLASS="black" ALIGN=CENTER><%=strFirstPeriodAwayGoals %></TD>
			<TD CLASS="black" ALIGN=CENTER><%=strSecondPeriodAwayGoals %></TD>
			<TD CLASS="black" ALIGN=CENTER><%=strThirdPeriodAwayGoals %></TD>
			<% If strOT=1 Then %>
				<TD CLASS="black" ALIGN=CENTER><%=strOvertimeAwayGoals %></TD>
			<% End If %>
			<TD CLASS="black" ALIGN=CENTER><%=strTotalAwayGoals %></TD>
		</TR>

		<TR BGCOLOR="#FFFFFF">
			<TD CLASS="black"><%=strHomeTeam %></TD>
			<TD CLASS="black" ALIGN=CENTER><%=strFirstPeriodHomeGoals %></TD>
			<TD CLASS="black" ALIGN=CENTER><%=strSecondPeriodHomeGoals %></TD>
			<TD CLASS="black" ALIGN=CENTER><%=strThirdPeriodHomeGoals %></TD>
			<% If strOT=1 Then %>
				<TD CLASS="black" ALIGN=CENTER><%=strOvertimeHomeGoals %></TD>
			<% End If %>
			<TD CLASS="black" ALIGN=CENTER><%=strTotalHomeGoals %></TD>
		</TR>
		</TABLE>

		<BR>

		<TABLE WIDTH=90% CELLPADDING=2 CELLSPACING=0 BORDER=0 BGCOLOR="#EEEEEE">
		<TR BGCOLOR="#08479E">
			<TD><B><SMALL>Scoring Summary</SMALL></B></TD>
		</TR>

		<TR BGCOLOR="#DCDCDC">
			<TD CLASS="black"><B><SMALL>First Period</SMALL></B></TD>
		</TR>

		<TR>
			<TD CLASS="black"><SMALL><% Call Scoring(strGameID, 1) %></SMALL></TD>
		</TR>

		<TR BGCOLOR="#DCDCDC">
			<TD CLASS="black"><B><SMALL>Second Period</SMALL></B></TD>
		</TR>

		<TR>
			<TD CLASS="black"><SMALL><% Call Scoring(strGameID, 2) %></SMALL></TD>
		</TR>

		<TR BGCOLOR="#DCDCDC">
			<TD CLASS="black"><B><SMALL>Third Period</SMALL></B></TD>
		</TR>

		<TR>
			<TD CLASS="black"><SMALL><% Call Scoring(strGameID, 3) %></SMALL></TD>
		</TR>

		<%	If strOT=1 Then  %>
			<TR BGCOLOR="#DCDCDC">
				<TD CLASS="black"><B><SMALL>Ovetime</SMALL></B></TD>
			</TR>

			<TR>
				<TD CLASS="black"><SMALL><% Call Scoring(strGameID, "OT") %></SMALL></TD>
			</TR>
		<%	End If
	
			rs2.Close
			Set rs2 = nothing
		
		%>
		</TABLE>

		<BR>
		<BR>

<%
	Dim strAwayPPConv, strAwayPPTotal, strHomePPConv, strHomePPTotal

	strSQL = "SELECT COUNT(*) FROM tbl_Teams INNER JOIN (tbl_Penalties INNER JOIN tbl_Players ON tbl_Penalties.PlayerID = tbl_Players.PlayerID) ON tbl_Teams.TeamID = tbl_Players.TeamID WHERE tbl_Penalties.GameID=" & strGameID & " AND tbl_Players.TeamID=" & strHomeTeamID
	Set rs = dbConn.Execute(strSQL)
	strAwayPPTotal = rs.Fields(0).Value

	strSQL = "SELECT COUNT(*) FROM tbl_Teams INNER JOIN (tbl_Penalties INNER JOIN tbl_Players ON tbl_Penalties.PlayerID = tbl_Players.PlayerID) ON tbl_Teams.TeamID = tbl_Players.TeamID WHERE tbl_Penalties.GameID=" & strGameID & " AND tbl_Players.TeamID=" & strAwayTeamID
	Set rs = dbConn.Execute(strSQL)
	strHomePPTotal = rs.Fields(0).Value

	strSQL = "SELECT Count(*) FROM (tbl_Goals INNER JOIN tbl_Players ON tbl_Goals.PlayerID = tbl_Players.PlayerID) INNER JOIN tbl_Teams ON tbl_Players.TeamID = tbl_Teams.TeamID WHERE (((tbl_Goals.GameID)=" & strGameID & ") AND ((tbl_Goals.PPG)=True) AND ((tbl_Players.TeamID)=" & strAwayTeamID & "))"
	Set rs = dbConn.Execute(strSQL)
	strAwayPPConv = rs.Fields(0).Value

	strSQL = "SELECT Count(*) FROM (tbl_Goals INNER JOIN tbl_Players ON tbl_Goals.PlayerID = tbl_Players.PlayerID) INNER JOIN tbl_Teams ON tbl_Players.TeamID = tbl_Teams.TeamID WHERE (((tbl_Goals.GameID)=" & strGameID & ") AND ((tbl_Goals.PPG)=True) AND ((tbl_Players.TeamID)=" & strHomeTeamID & "))"
	Set rs = dbConn.Execute(strSQL)
	strHomePPConv = rs.Fields(0).Value

%>
		<TABLE WIDTH=90% CELLPADDING=2 CELLSPACING=0 BORDER=1 BORDERCOLOR="black" BGCOLOR="#FFFFFF">
		<TR BGCOLOR="#C0CCF8">
			<TD CLASS="black"><B><SMALL>Team</SMALL></B>
			<TD CLASS="black" COLSPAN=<% If strOT=1 Then %>5<% Else %>4<% End If %> ALIGN=CENTER><B><SMALL>Shots on Goal</SMALL></B></TD>
			<TD CLASS="black" ALIGN=CENTER COLSPAN=2><B><SMALL>Power Plays</SMALL></B>
		</TR>

		<TR BGCOLOR="#1048A0">
			<TD>&nbsp;</TD>
			<TD ALIGN=CENTER><SMALL><B>1st</B></SMALL></TD>
			<TD ALIGN=CENTER><SMALL><B>2nd</B></SMALL></TD>
			<TD ALIGN=CENTER><SMALL><B>3rd</B></SMALL></TD>
			<% If strOT=1 Then %><TD ALIGN=CENTER><SMALL><B>OT</B></SMALL></TD><% End If %>
			<TD ALIGN=CENTER><SMALL><B>TOTAL</B></SMALL></TD>
			<TD ALIGN=CENTER><SMALL><B>CONVERTED</B></SMALL></TD>
			<TD ALIGN=CENTER><SMALL><B>TOTAL</B></SMALL></TD>
		</TR>

		<TR>
			<TD CLASS="black"><SMALL><%=strAwayTeam %></SMALL></TD>
			<TD CLASS="black" ALIGN=CENTER><SMALL><%=strFirstPeriodAwayShots %></SMALL></TD>
			<TD CLASS="black" ALIGN=CENTER><SMALL><%=strSecondPeriodAwayShots%></SMALL></TD>
			<TD CLASS="black" ALIGN=CENTER><SMALL><%=strThirdPeriodAwayShots %></SMALL></TD>
			<% If strOT=1 Then %><TD CLASS="black" ALIGN=CENTER><SMALL><%=strOverTimeAwayShots %></SMALL></TD><% End If %>
			<TD CLASS="black" ALIGN=CENTER><SMALL><%=strFirstPeriodAwayShots+strSecondPeriodAwayShots+strThirdPeriodAwayShots+strOverTimeAwayShots %></SMALL></TD>
			<TD CLASS="black" ALIGN=CENTER><SMALL><%=strAwayPPConv %></SMALL></TD>
			<TD CLASS="black" ALIGN=CENTER><SMALL><%=strAwayPPTotal %></SMALL></TD>
		</TR>

		<TR>
			<TD CLASS="black"><SMALL><%=strHomeTeam %></SMALL></TD>
			<TD CLASS="black" ALIGN=CENTER><SMALL><%=strFirstPeriodHomeShots %></SMALL></TD>
			<TD CLASS="black" ALIGN=CENTER><SMALL><%=strSecondPeriodHomeShots%></SMALL></TD>
			<TD CLASS="black" ALIGN=CENTER><SMALL><%=strThirdPeriodHomeShots %></SMALL></TD>
			<% If strOT=1 Then %><TD CLASS="black" ALIGN=CENTER><SMALL><%=strOverTimeHomeShots %></SMALL></TD><% End If %>
			<TD CLASS="black" ALIGN=CENTER><SMALL><%=strFirstPeriodHomeShots+strSecondPeriodHomeShots+strThirdPeriodHomeShots+strOverTimeHomeShots %></SMALL></TD>
			<TD CLASS="black" ALIGN=CENTER><SMALL><%=strHomePPConv %></SMALL></TD>
			<TD CLASS="black" ALIGN=CENTER><SMALL><%=strHomePPTotal %></SMALL></TD>
		</TR>
		</TABLE>

		<BR><BR>
		<TABLE WIDTH=90% CELLPADDING=2 CELLSPACING=0 BORDER=0 BGCOLOR="#EEEEEE">
		<TR BGCOLOR="#08479E">
			<TD><B><SMALL>Penalty Summary</SMALL></B></TD>
		</TR>

		<TR BGCOLOR="#DCDCDC">
			<TD CLASS="black"><B><SMALL>First Period</SMALL></B></TD>
		</TR>

		<TR>
			<TD CLASS="black"><SMALL><% Call Penalties(strGameID, 1) %></SMALL></TD>
		</TR>

		<TR BGCOLOR="#DCDCDC">
			<TD CLASS="black"><B><SMALL>Second Period</SMALL></B></TD>
		</TR>

		<TR>
			<TD CLASS="black"><SMALL><% Call Penalties(strGameID, 2) %></SMALL></TD>
		</TR>

		<TR BGCOLOR="#DCDCDC">
			<TD CLASS="black"><B><SMALL>Third Period</SMALL></B></TD>
		</TR>

		<TR>
			<TD CLASS="black"><SMALL><% Call Penalties(strGameID, 3) %></SMALL></TD>
		</TR>
		<% If strOT=1 Then %>
			<TR BGCOLOR="#DCDCDC">
				<TD CLASS="black"><B><SMALL>Ovetime</SMALL></B></TD>
			</TR>

			<TR>
				<TD CLASS="black"><SMALL><% Call Penalties(strGameID, "OT") %></SMALL></TD>
			</TR>
		<% End If %>
		</TABLE>

		<BR><BR>
		<TABLE WIDTH=90% CELLPADDING=2 CELLSPACING=0 BORDER=0 BGCOLOR="#EEEEEE">
		<TR BGCOLOR="#08479E">
			<TD><B><SMALL>Game Notes</SMALL></B></TD>
		</TR>

		<TR BGCOLOR="#DCDCDC">
			<TD CLASS="black" HEIGHT=50 VALIGN=TOP><SMALL><B>Attendance</B><BR><%=strAttendance %></SMALL></TD>
		</TR>

		<TR BGCOLOR="#DCDCDC">
			<TD CLASS="black" HEIGHT=50 VALIGN=TOP><SMALL><B>Referees</B><BR><%=strReferee1 %><% If Not IsNull(strReferee2) Then response.write ", " & strReferee2 %></SMALL></TD>
		</TR>

		<TR BGCOLOR="#DCDCDC">
			<TD CLASS="black" HEIGHT=50 VALIGN=TOP><SMALL><B>Scorekeeper</B><BR><%=strScorekeeper %></SMALL></TD>
		</TR>

		<TR BGCOLOR="#DCDCDC">
			<TD CLASS="black" HEIGHT=50 VALIGN=TOP><SMALL><B>Notes</B><BR><%=strNotes %></SMALL></TD>
		</TR>
		</TABLE>

<%

	rs.Close
	Set rs = nothing

	dbConn.Close
	Set dbConn=nothing
%>
		</CENTER>
		<BR><BR><BR>

	</TD>
</TR>
</TABLE>
</BODY>
</HTML>