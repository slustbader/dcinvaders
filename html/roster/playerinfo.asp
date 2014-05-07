<%
	Option Explicit
	'On Error Resume Next
	
	Dim strPlayerID, strStatType, strLastName, strFirstName, strNumber

	Dim dbConn, rs, strSQL
	
	strPlayerID = Request.QueryString("ID")
	strStatType = Request.QueryString("stat")
	strLastName = Request.QueryString("LN")
	strFirstName = Request.QueryString("FN")
	strNumber = Request.QueryString("N")

	Set dbConn = Server.CreateObject("ADODB.Connection")
	'dbConn.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("/invaders/database/hockey.mdb")
	dbConn.Open "DSN=hockey"
	
	Set rs = Server.CreateObject("ADODB.Recordset")

%>
<!-- #include virtual="/include/adovbs.inc"-->
<HTML>
<HEAD>
<SCRIPT LANGUAGE="JavaScript">
<!--

	function SendToOpener(GameID){
		window.opener.document.location.href="/games/summary.asp?ID=" + GameID;
	}

//-->
</SCRIPT>
<LINK REL="stylesheet" HREF="/styles/hockey.css" TYPE="text/css">
</HEAD>
<BODY BACKGROUND="/images/stars.gif">
<%
	response.write "<TABLE WIDTH=100% CELLPADDING=0 CELLSPACING=0 BORDER=0><TR><TD><B>" & strNumber & " - " & strFirstName & " " & strLastName & "</B></TD><TD ALIGN=RIGHT><SMALL><A HREF=""javascript: window.close()"">Close Window</A></TD></TR></TABLE><BR><BR>"
	
	Select Case (strStatType)
		Case "GP"
			response.write "<CENTER><FONT COLOR=""#00CC00""><SMALL><B>Games Played</B></SMALL></FONT></CENTER><BR>"
			strSQL = "SELECT tbl_Games.GameID, tbl_PlayerGames.PlayerID, tbl_HomeTeams.HomeTeamName, tbl_AwayTeams.AwayTeamName, tbl_Games.GameDate FROM tbl_HomeTeams INNER JOIN (tbl_AwayTeams INNER JOIN (tbl_Games INNER JOIN tbl_PlayerGames ON tbl_Games.GameID = tbl_PlayerGames.GameID) ON tbl_AwayTeams.AwayTeamID = tbl_Games.AwayTeamID) ON tbl_HomeTeams.HomeTeamID = tbl_Games.HomeTeamID WHERE (((tbl_PlayerGames.PlayerID)=" & strPlayerID & "))"
			rs.Open strSQL, dbConn, adOpenStatic
			
			If Not rs.EOF AND Not rs.BOF Then
				rs.MoveFirst
				Do While Not rs.EOF
					response.write "<SMALL><A HREF=""javascript: SendToOpener('" & rs("GameID") & "')"">" & rs("GameDate") & " - " & rs("AwayTeamName") & " vs. " & rs("HomeTeamName") & "</A></SMALL><BR>"
					rs.MoveNext
				Loop
			Else
				response.write "None."
			End If
		Case "G"
			response.write "<CENTER><FONT COLOR=""#00CC00""><SMALL><B>Goals</B></SMALL></FONT></CENTER><BR>"
			strSQL = "SELECT tbl_Games.GameID, tbl_HomeTeams.HomeTeamName, tbl_AwayTeams.AwayTeamName, tbl_Games.GameDate, pick_Periods.Period, tbl_Goals.Time FROM pick_Periods INNER JOIN (tbl_HomeTeams INNER JOIN (tbl_AwayTeams INNER JOIN (tbl_Games INNER JOIN tbl_Goals ON tbl_Games.GameID = tbl_Goals.GameID) ON tbl_AwayTeams.AwayTeamID = tbl_Games.AwayTeamID) ON tbl_HomeTeams.HomeTeamID = tbl_Games.HomeTeamID) ON pick_Periods.ID = tbl_Goals.Period WHERE (((tbl_Goals.PlayerID)=" & strPlayerID & ")) ORDER BY tbl_Games.GameDate, pick_Periods.Period, tbl_Goals.Time DESC"

			rs.Open strSQL, dbConn, adOpenStatic

			If Not rs.EOF AND Not rs.BOF Then
				rs.MoveFirst

				response.write "<CENTER>"
				response.write "<TABLE WIDTH=90%  BGCOLOR=""#EEEEEE"" CELLPADDING=2 CELLSPACING=0 BORDER=1 BORDERCOLOR=black>"
				response.write "<TR BGCOLOR=""#08479E"">"
				response.write "<TD><B><SMALL>Game</SMALL></B></TD>"
				response.write "<TD><B><SMALL>Period</SMALL></B></TD>"
				response.write "<TD><B><SMALL>Time</SMALL></B></TD>"
				response.write "</TR>"
				Do While Not rs.EOF
					response.write "<TR BGCOLOR=""#DCDCDC"">"
					response.write "<TD CLASS=""black""><SMALL><A HREF=""javascript: SendToOpener('" & rs("GameID") & "')"">" & rs("GameDate") & " - " & rs("AwayTeamName") & " vs. " & rs("HomeTeamName") & "</A></SMALL></TD>"
					response.write "<TD CLASS=""black""><SMALL>" & rs("Period") & "</SMALL></TD>"
					response.write "<TD CLASS=""black""><SMALL>" & FormatDateTime(rs("Time"),4) & "</SMALL></TD>"
					response.write "</TR>"
					rs.MoveNext
				Loop
				response.write "</TABLE>"
				response.write "</CENTER>"
			Else
				response.write "None."
			End If

		Case "A"
			response.write "<CENTER><FONT COLOR=""#00CC00""><SMALL><B>Assists</B></SMALL></FONT></CENTER><BR>"
			strSQL = "SELECT tbl_Games.GameID, tbl_HomeTeams.HomeTeamName, tbl_AwayTeams.AwayTeamName, tbl_Games.GameDate, pick_Periods.Period, tbl_Goals.Time FROM tbl_HomeTeams INNER JOIN ((tbl_AwayTeams INNER JOIN tbl_Games ON tbl_AwayTeams.AwayTeamID = tbl_Games.AwayTeamID) INNER JOIN (pick_Periods INNER JOIN tbl_Goals ON pick_Periods.ID = tbl_Goals.Period) ON tbl_Games.GameID = tbl_Goals.GameID) ON tbl_HomeTeams.HomeTeamID = tbl_Games.HomeTeamID WHERE (((tbl_Goals.Assist1PlayerID)=" & strPlayerID & ")) OR (((tbl_Goals.Assist2PlayerID)=" & strPlayerID & ")) ORDER BY tbl_Games.GameDate, pick_Periods.Period, tbl_Goals.Time DESC"

			rs.Open strSQL, dbConn, adOpenStatic

			If Not rs.EOF AND Not rs.BOF Then
				rs.MoveFirst

				response.write "<CENTER>"
				response.write "<TABLE WIDTH=90%  BGCOLOR=""#EEEEEE"" CELLPADDING=2 CELLSPACING=0 BORDER=1 BORDERCOLOR=black>"
				response.write "<TR BGCOLOR=""#08479E"">"
				response.write "<TD><B><SMALL>Game</SMALL></B></TD>"
				response.write "<TD><B><SMALL>Period</SMALL></B></TD>"
				response.write "<TD><B><SMALL>Time</SMALL></B></TD>"
				response.write "</TR>"
				Do While Not rs.EOF
					response.write "<TR BGCOLOR=""#DCDCDC"">"
					response.write "<TD CLASS=""black""><SMALL><A HREF=""javascript: SendToOpener('" & rs("GameID") & "')"">" & rs("GameDate") & " - " & rs("AwayTeamName") & " vs. " & rs("HomeTeamName") & "</A></SMALL></TD>"
					response.write "<TD CLASS=""black""><SMALL>" & rs("Period") & "</SMALL></TD>"
					response.write "<TD CLASS=""black""><SMALL>" & FormatDateTime(rs("Time"),4) & "</SMALL></TD>"
					response.write "</TR>"
					rs.MoveNext
				Loop
				response.write "</TABLE>"
				response.write "</CENTER>"
			Else
				response.write "None."
			End If

		Case "PPG"
			response.write "<CENTER><FONT COLOR=""#00CC00""><SMALL><B>Power Play Goals</B></SMALL></FONT></CENTER><BR>"
			strSQL = "SELECT tbl_Games.GameID, tbl_HomeTeams.HomeTeamName, tbl_AwayTeams.AwayTeamName, tbl_Games.GameDate, pick_Periods.Period, tbl_Goals.Time FROM tbl_HomeTeams INNER JOIN ((tbl_AwayTeams INNER JOIN tbl_Games ON tbl_AwayTeams.AwayTeamID = tbl_Games.AwayTeamID) INNER JOIN (pick_Periods INNER JOIN tbl_Goals ON pick_Periods.ID = tbl_Goals.Period) ON tbl_Games.GameID = tbl_Goals.GameID) ON tbl_HomeTeams.HomeTeamID = tbl_Games.HomeTeamID WHERE (((tbl_Goals.PlayerID)=" & strPlayerID & ") AND ((tbl_Goals.PPG)=True)) ORDER BY tbl_Games.GameDate, pick_Periods.Period, tbl_Goals.Time DESC"

			rs.Open strSQL, dbConn, adOpenStatic

			If Not rs.EOF AND Not rs.BOF Then
				rs.MoveFirst

				response.write "<CENTER>"
				response.write "<TABLE WIDTH=90%  BGCOLOR=""#EEEEEE"" CELLPADDING=2 CELLSPACING=0 BORDER=1 BORDERCOLOR=black>"
				response.write "<TR BGCOLOR=""#08479E"">"
				response.write "<TD><B><SMALL>Game</SMALL></B></TD>"
				response.write "<TD><B><SMALL>Period</SMALL></B></TD>"
				response.write "<TD><B><SMALL>Time</SMALL></B></TD>"
				response.write "</TR>"
				Do While Not rs.EOF
					response.write "<TR BGCOLOR=""#DCDCDC"">"
					response.write "<TD CLASS=""black""><SMALL><A HREF=""javascript: SendToOpener('" & rs("GameID") & "')"">" & rs("GameDate") & " - " & rs("AwayTeamName") & " vs. " & rs("HomeTeamName") & "</A></SMALL></TD>"
					response.write "<TD CLASS=""black""><SMALL>" & rs("Period") & "</SMALL></TD>"
					response.write "<TD CLASS=""black""><SMALL>" & FormatDateTime(rs("Time"),4) & "</SMALL></TD>"
					response.write "</TR>"
					rs.MoveNext
				Loop
				response.write "</TABLE>"
				response.write "</CENTER>"
			Else
				response.write "None."
			End If

		Case "PPA"
			response.write "<CENTER><FONT COLOR=""#00CC00""><SMALL><B>Power Play Assists</B></SMALL></FONT></CENTER><BR>"
			strSQL = "SELECT tbl_Games.GameID, tbl_HomeTeams.HomeTeamName, tbl_AwayTeams.AwayTeamName, tbl_Games.GameDate, pick_Periods.Period, tbl_Goals.Time, tbl_Goals.Assist1PlayerID, tbl_Goals.Assist2PlayerID FROM tbl_HomeTeams INNER JOIN ((tbl_AwayTeams INNER JOIN tbl_Games ON tbl_AwayTeams.AwayTeamID = tbl_Games.AwayTeamID) INNER JOIN (pick_Periods INNER JOIN tbl_Goals ON pick_Periods.ID = tbl_Goals.Period) ON tbl_Games.GameID = tbl_Goals.GameID) ON tbl_HomeTeams.HomeTeamID = tbl_Games.HomeTeamID WHERE (((tbl_Goals.PPG)=True) AND ((tbl_Goals.Assist1PlayerID)=" & strPlayerID & ")) OR (((tbl_Goals.PPG)=True) AND ((tbl_Goals.Assist2PlayerID)=" & strPlayerID & ")) ORDER BY tbl_Games.GameDate, pick_Periods.Period, tbl_Goals.Time DESC"

			rs.Open strSQL, dbConn, adOpenStatic

			If Not rs.EOF AND Not rs.BOF Then
				rs.MoveFirst

				response.write "<CENTER>"
				response.write "<TABLE WIDTH=90%  BGCOLOR=""#EEEEEE"" CELLPADDING=2 CELLSPACING=0 BORDER=1 BORDERCOLOR=black>"
				response.write "<TR BGCOLOR=""#08479E"">"
				response.write "<TD><B><SMALL>Game</SMALL></B></TD>"
				response.write "<TD><B><SMALL>Period</SMALL></B></TD>"
				response.write "<TD><B><SMALL>Time</SMALL></B></TD>"
				response.write "</TR>"
				Do While Not rs.EOF
					response.write "<TR BGCOLOR=""#DCDCDC"">"
					response.write "<TD CLASS=""black""><SMALL><A HREF=""javascript: SendToOpener('" & rs("GameID") & "')"">" & rs("GameDate") & " - " & rs("AwayTeamName") & " vs. " & rs("HomeTeamName") & "</A></SMALL></TD>"
					response.write "<TD CLASS=""black""><SMALL>" & rs("Period") & "</SMALL></TD>"
					response.write "<TD CLASS=""black""><SMALL>" & FormatDateTime(rs("Time"),4) & "</SMALL></TD>"
					response.write "</TR>"
					rs.MoveNext
				Loop
				response.write "</TABLE>"
				response.write "</CENTER>"
			Else
				response.write "None."
			End If

		Case "SHG"
			response.write "<CENTER><FONT COLOR=""#00CC00""><SMALL><B>Short Handed Goals</B></SMALL></FONT></CENTER><BR>"
			strSQL = "SELECT tbl_Games.GameID, tbl_HomeTeams.HomeTeamName, tbl_AwayTeams.AwayTeamName, tbl_Games.GameDate, pick_Periods.Period, tbl_Goals.Time FROM tbl_HomeTeams INNER JOIN ((tbl_AwayTeams INNER JOIN tbl_Games ON tbl_AwayTeams.AwayTeamID = tbl_Games.AwayTeamID) INNER JOIN (pick_Periods INNER JOIN tbl_Goals ON pick_Periods.ID = tbl_Goals.Period) ON tbl_Games.GameID = tbl_Goals.GameID) ON tbl_HomeTeams.HomeTeamID = tbl_Games.HomeTeamID WHERE (((tbl_Goals.PlayerID)=" & strPlayerID & ") AND ((tbl_Goals.SH)=True)) ORDER BY tbl_Games.GameDate, pick_Periods.Period, tbl_Goals.Time DESC"

			rs.Open strSQL, dbConn, adOpenStatic

			If Not rs.EOF AND Not rs.BOF Then
				rs.MoveFirst

				response.write "<CENTER>"
				response.write "<TABLE WIDTH=90%  BGCOLOR=""#EEEEEE"" CELLPADDING=2 CELLSPACING=0 BORDER=1 BORDERCOLOR=black>"
				response.write "<TR BGCOLOR=""#08479E"">"
				response.write "<TD><B><SMALL>Game</SMALL></B></TD>"
				response.write "<TD><B><SMALL>Period</SMALL></B></TD>"
				response.write "<TD><B><SMALL>Time</SMALL></B></TD>"
				response.write "</TR>"
				Do While Not rs.EOF
					response.write "<TR BGCOLOR=""#DCDCDC"">"
					response.write "<TD CLASS=""black""><SMALL><A HREF=""javascript: SendToOpener('" & rs("GameID") & "')"">" & rs("GameDate") & " - " & rs("AwayTeamName") & " vs. " & rs("HomeTeamName") & "</A></SMALL></TD>"
					response.write "<TD CLASS=""black""><SMALL>" & rs("Period") & "</SMALL></TD>"
					response.write "<TD CLASS=""black""><SMALL>" & FormatDateTime(rs("Time"),4) & "</SMALL></TD>"
					response.write "</TR>"
					rs.MoveNext
				Loop
				response.write "</TABLE>"
				response.write "</CENTER>"
			Else
				response.write "None."
			End If

		Case "SHA"
			response.write "<CENTER><FONT COLOR=""#00CC00""><SMALL><B>Short Handed Assists</B></SMALL></FONT></CENTER><BR>"
			strSQL = "SELECT tbl_Games.GameID, tbl_HomeTeams.HomeTeamName, tbl_AwayTeams.AwayTeamName, tbl_Games.GameDate, pick_Periods.Period, tbl_Goals.Time FROM tbl_HomeTeams INNER JOIN ((tbl_AwayTeams INNER JOIN tbl_Games ON tbl_AwayTeams.AwayTeamID = tbl_Games.AwayTeamID) INNER JOIN (pick_Periods INNER JOIN tbl_Goals ON pick_Periods.ID = tbl_Goals.Period) ON tbl_Games.GameID = tbl_Goals.GameID) ON tbl_HomeTeams.HomeTeamID = tbl_Games.HomeTeamID WHERE (((tbl_Goals.SH)=True) AND ((tbl_Goals.Assist1PlayerID)=" & strPlayerID & ")) OR (((tbl_Goals.SH)=True) AND ((tbl_Goals.Assist2PlayerID)=" & strPlayerID & ")) ORDER BY tbl_Games.GameDate, pick_Periods.Period, tbl_Goals.Time DESC"

			rs.Open strSQL, dbConn, adOpenStatic

			If Not rs.EOF AND Not rs.BOF Then
				rs.MoveFirst

				response.write "<CENTER>"
				response.write "<TABLE WIDTH=90%  BGCOLOR=""#EEEEEE"" CELLPADDING=2 CELLSPACING=0 BORDER=1 BORDERCOLOR=black>"
				response.write "<TR BGCOLOR=""#08479E"">"
				response.write "<TD><B><SMALL>Game</SMALL></B></TD>"
				response.write "<TD><B><SMALL>Period</SMALL></B></TD>"
				response.write "<TD><B><SMALL>Time</SMALL></B></TD>"
				response.write "</TR>"
				Do While Not rs.EOF
					response.write "<TR BGCOLOR=""#DCDCDC"">"
					response.write "<TD CLASS=""black""><SMALL><A HREF=""javascript: SendToOpener('" & rs("GameID") & "')"">" & rs("GameDate") & " - " & rs("AwayTeamName") & " vs. " & rs("HomeTeamName") & "</A></SMALL></TD>"
					response.write "<TD CLASS=""black""><SMALL>" & rs("Period") & "</SMALL></TD>"
					response.write "<TD CLASS=""black""><SMALL>" & FormatDateTime(rs("Time"),4) & "</SMALL></TD>"
					response.write "</TR>"
					rs.MoveNext
				Loop
				response.write "</TABLE>"
				response.write "</CENTER>"
			Else
				response.write "None."
			End If

		Case "GW"
			response.write "<CENTER><FONT COLOR=""#00CC00""><SMALL><B>Game Winning Goals</B></SMALL></FONT></CENTER><BR>"
			strSQL = "SELECT tbl_Games.GameID, tbl_HomeTeams.HomeTeamName, tbl_AwayTeams.AwayTeamName, tbl_Games.GameDate, pick_Periods.Period, tbl_Goals.Time FROM tbl_HomeTeams INNER JOIN ((tbl_AwayTeams INNER JOIN tbl_Games ON tbl_AwayTeams.AwayTeamID = tbl_Games.AwayTeamID) INNER JOIN (pick_Periods INNER JOIN tbl_Goals ON pick_Periods.ID = tbl_Goals.Period) ON tbl_Games.GameID = tbl_Goals.GameID) ON tbl_HomeTeams.HomeTeamID = tbl_Games.HomeTeamID WHERE (((tbl_Goals.PlayerID)=" & strPlayerID & ") AND ((tbl_Goals.GW)=True)) ORDER BY tbl_Games.GameDate, pick_Periods.Period, tbl_Goals.Time DESC"

			rs.Open strSQL, dbConn, adOpenStatic

			If Not rs.EOF AND Not rs.BOF Then
				rs.MoveFirst

				response.write "<CENTER>"
				response.write "<TABLE WIDTH=90%  BGCOLOR=""#EEEEEE"" CELLPADDING=2 CELLSPACING=0 BORDER=1 BORDERCOLOR=black>"
				response.write "<TR BGCOLOR=""#08479E"">"
				response.write "<TD><B><SMALL>Game</SMALL></B></TD>"
				response.write "<TD><B><SMALL>Period</SMALL></B></TD>"
				response.write "<TD><B><SMALL>Time</SMALL></B></TD>"
				response.write "</TR>"
				Do While Not rs.EOF
					response.write "<TR BGCOLOR=""#DCDCDC"">"
					response.write "<TD CLASS=""black""><SMALL><A HREF=""javascript: SendToOpener('" & rs("GameID") & "')"">" & rs("GameDate") & " - " & rs("AwayTeamName") & " vs. " & rs("HomeTeamName") & "</A></SMALL></TD>"
					response.write "<TD CLASS=""black""><SMALL>" & rs("Period") & "</SMALL></TD>"
					response.write "<TD CLASS=""black""><SMALL>" & FormatDateTime(rs("Time"),4) & "</SMALL></TD>"
					response.write "</TR>"
					rs.MoveNext
				Loop
				response.write "</TABLE>"
				response.write "</CENTER>"
			Else
				response.write "None."
			End If

		Case "GT"
			response.write "<CENTER><FONT COLOR=""#00CC00""><SMALL><B>Game Tying</B></SMALL></FONT></CENTER><BR>"
			strSQL = "SELECT tbl_Games.GameID, tbl_HomeTeams.HomeTeamName, tbl_AwayTeams.AwayTeamName, tbl_Games.GameDate, pick_Periods.Period, tbl_Goals.Time FROM tbl_HomeTeams INNER JOIN ((tbl_AwayTeams INNER JOIN tbl_Games ON tbl_AwayTeams.AwayTeamID = tbl_Games.AwayTeamID) INNER JOIN (pick_Periods INNER JOIN tbl_Goals ON pick_Periods.ID = tbl_Goals.Period) ON tbl_Games.GameID = tbl_Goals.GameID) ON tbl_HomeTeams.HomeTeamID = tbl_Games.HomeTeamID WHERE (((tbl_Goals.PlayerID)=" & strPlayerID & ") AND ((tbl_Goals.GT)=True)) ORDER BY tbl_Games.GameDate, pick_Periods.Period, tbl_Goals.Time DESC"

			rs.Open strSQL, dbConn, adOpenStatic

			If Not rs.EOF AND Not rs.BOF Then
				rs.MoveFirst

				response.write "<CENTER>"
				response.write "<TABLE WIDTH=90%  BGCOLOR=""#EEEEEE"" CELLPADDING=2 CELLSPACING=0 BORDER=1 BORDERCOLOR=black>"
				response.write "<TR BGCOLOR=""#08479E"">"
				response.write "<TD><B><SMALL>Game</SMALL></B></TD>"
				response.write "<TD><B><SMALL>Period</SMALL></B></TD>"
				response.write "<TD><B><SMALL>Time</SMALL></B></TD>"
				response.write "</TR>"
				Do While Not rs.EOF
					response.write "<TR BGCOLOR=""#DCDCDC"">"
					response.write "<TD CLASS=""black""><SMALL><A HREF=""javascript: SendToOpener('" & rs("GameID") & "')"">" & rs("GameDate") & " - " & rs("AwayTeamName") & " vs. " & rs("HomeTeamName") & "</A></SMALL></TD>"
					response.write "<TD CLASS=""black""><SMALL>" & rs("Period") & "</SMALL></TD>"
					response.write "<TD CLASS=""black""><SMALL>" & FormatDateTime(rs("Time"),4) & "</SMALL></TD>"
					response.write "</TR>"
					rs.MoveNext
				Loop
				response.write "</TABLE>"
				response.write "</CENTER>"
			Else
				response.write "None."
			End If

		Case "PIM"
			response.write "<CENTER><FONT COLOR=""#00CC00""><SMALL><B>Penalty Minutes</B></SMALL></FONT></CENTER><BR>"
			strSQL = "SELECT tbl_AwayTeams.AwayTeamName, tbl_HomeTeams.HomeTeamName, tbl_Games.GameDate, pick_Penalties.PenaltyName, tbl_Penalties.Period, tbl_Penalties.Start, tbl_Games.GameID, tbl_Seasons.SeasonName, tbl_Penalties.Length FROM tbl_Seasons INNER JOIN (tbl_HomeTeams INNER JOIN ((tbl_AwayTeams INNER JOIN tbl_Games ON tbl_AwayTeams.AwayTeamID = tbl_Games.AwayTeamID) INNER JOIN (pick_Penalties INNER JOIN tbl_Penalties ON pick_Penalties.PenaltyID = tbl_Penalties.PenaltyCharge) ON tbl_Games.GameID = tbl_Penalties.GameID) ON tbl_HomeTeams.HomeTeamID = tbl_Games.HomeTeamID) ON tbl_Seasons.SeasonID = tbl_Games.SeasonID WHERE (((tbl_Penalties.PlayerID)=" & strPlayerID & ")) ORDER BY GameDate, Period, Start DESC"

			rs.Open strSQL, dbConn, adOpenStatic

			If Not rs.EOF AND Not rs.BOF Then
				rs.MoveFirst
				response.write "<CENTER>"
				response.write "<TABLE WIDTH=90%  BGCOLOR=""#EEEEEE"" CELLPADDING=2 CELLSPACING=0 BORDER=1 BORDERCOLOR=black>"
				response.write "<TR BGCOLOR=""#08479E"">"
				response.write "<TD><B><SMALL>Game</SMALL></B></TD>"
				response.write "<TD><B><SMALL>Penalty</SMALL></B></TD>"
				response.write "<TD ALIGN=CENTER><B><SMALL>Length</SMALL></B></TD>"
				response.write "<TD ALIGN=CENTER><B><SMALL>Period</SMALL></B></TD>"
				response.write "<TD ALIGN=CENTER><B><SMALL>Time</SMALL></B></TD>"
				response.write "</TR>"
				
				Do While Not rs.EOF
					response.write "<TR BGCOLOR=""#DCDCDC"">"
					response.write "<TD CLASS=""black""><SMALL><A HREF=""javascript: SendToOpener('" & rs("GameID") & "')"">" & rs("AwayTeamName") & " vs. " & rs("HomeTeamName") & " - " & rs("GameDate") & "</A></SMALL></TD>"
					response.write "<TD CLASS=""black""><SMALL>" & rs("PenaltyName") & "</SMALL></TD>"
					response.write "<TD CLASS=""black"" ALIGN=CENTER><SMALL>" & rs("Length") & "</SMALL></TD>"
					response.write "<TD CLASS=""black"" ALIGN=CENTER><SMALL>" & rs("Period") & "</SMALL></TD>"
					response.write "<TD CLASS=""black"" ALIGN=CENTER><SMALL>" & FormatDateTime(rs("Start"),4) & "</SMALL></TD>"
					response.write "</TR>"
					rs.MoveNext
				Loop
				response.write "</TABLE>"
				response.write "</CENTER>"
			Else
				response.write "None."
			End If
			
	End Select
			
	rs.Close
	Set rs = nothing
	
	dbConn.Close
	Set dbConn = nothing

%>
	
</BODY>
</HTML>