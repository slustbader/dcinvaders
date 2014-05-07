<!--#include virtual="/include/adovbs.inc"-->
<%
Class Tops
	Public playerId(), name(), totals(), startYear(), endYear(), season(), seasonId(), game(), gameId(), gameTeam()

	Public Sub processRS(num, rs, scope)
		ReDim playerId(num), name(num), totals(num), startYear(num), endYear(num), season(num), seasonId(num), game(num), gameId(num), gameTeam(num)

		Dim i
		i = 0
		While i < num and not rs.EOF
		    playerId(i) = rs("playerId")
			name(i) = rs("firstName") & " " & rs("lastName")
			totals(i) = rs("total")
			if scope = "career" then
				startYear(i) = rs("start")
				endYear(i) = rs("end")
			elseif scope = "season" then
				season(i) = rs("season")
				seasonId(i) = rs("seasonId")
			elseif scope = "game" then
				game(i) = rs("gameDate")
				gameId(i) = rs("gameId")
				gameTeam(i) = rs("team")
			end if
			rs.MoveNext
			i = i + 1
		Wend
	End Sub

End Class

Dim dbConn, strSQ, rs
Dim goalTops

Set dbConn = Server.CreateObject("ADODB.Connection")
'dbConn.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("/invaders/database/hockey.mdb")
dbConn.Open "DSN=hockey"

%>
<HTML>
<HEAD>
<TITLE>DC Invaders Records</TITLE>
<LINK REL="stylesheet" HREF="../styles/hockey.css" TYPE="text/css">
</HEAD>
<BODY BACKGROUND="../images/stars.gif">
<TABLE WIDTH=660 CELLPADDING=0 CELLSPACING=0 BORDER=0>
<TR>
	<TD><CENTER><IMG SRC="../images/invader.gif" BORDER=0></CENTER>
	<CENTER><H2>Invader Records</H2></CENTER>
		<TABLE WIDTH=100% CELLPADDING=0 CELLSPACING=0 BORDER=0>
<%
		call prepareOffenseQuery(null, "season", "points", null)
%>
		<TR>
			<TD COLSPAN=3>&nbsp;</TD>
		</TR>
		
<%
		call prepareOffenseQuery(null, "game", "points", null)
%>
		
		<TR>
			<TD COLSPAN=3>&nbsp;</TD>
		</TR>
		
<%		
		call prepareOffenseQuery(null, "career", "points", null)
%>
		<TR>
			<TD COLSPAN=3>&nbsp;</TD>
		</TR>

<%
		call prepareOffenseQuery(null, "season", "goals", null)
%>
		<TR>
			<TD COLSPAN=3>&nbsp;</TD>
		</TR>

<%
		call prepareOffenseQuery(null, "game", "goals", null)
%>
		<TR>
			<TD COLSPAN=3>&nbsp;</TD>
		</TR>

<%
		call prepareOffenseQuery(null, "career", "goals", null)
%>
		<TR>
			<TD COLSPAN=3>&nbsp;</TD>
		</TR>
		
<%
		call prepareOffenseQuery(null, "season", "assists", null)
%>
		<TR>
			<TD COLSPAN=3>&nbsp;</TD>
		</TR>

<%
		call prepareOffenseQuery(null, "game", "assists", null)
%>
		<TR>
			<TD COLSPAN=3>&nbsp;</TD>
		</TR>
		
<%
		call prepareOffenseQuery(null, "career", "assists", null)
%>
		<TR>
			<TD COLSPAN=3>&nbsp;</TD>
		</TR>

<%
		call prepareGoalieQuery(null, "career", "gaa")
%>
		<TR>
			<TD COLSPAN=3>&nbsp;</TD>
		</TR>

<%
		call prepareGoalieQuery(null, "season", "gaa")
%>
		<TR>
			<TD COLSPAN=3>&nbsp;</TD>
		</TR>

<%
		call prepareShutoutQuery(null, "career")
%>
		<TR>
			<TD COLSPAN=3>&nbsp;</TD>
		</TR>

<%
		call prepareShutoutQuery(null, "season")
%>
		<TR>
			<TD COLSPAN=3>&nbsp;</TD>
		</TR>

<%
		call prepareGoalieQuery(null, "career", "spct")
%>
		<TR>
			<TD COLSPAN=3>&nbsp;</TD>
		</TR>

<%
		call prepareGoalieQuery(null, "season", "spct")
%>
		<TR>
			<TD COLSPAN=3>&nbsp;</TD>
		</TR>

<%
		call preparePimQuery(null, "season")
%>
		<TR>
			<TD COLSPAN=3>&nbsp;</TD>
		</TR>

<%
		call preparePimQuery(null, "game")
%>
		<TR>
			<TD COLSPAN=3>&nbsp;</TD>
		</TR>		
		
<%
		call preparePimQuery(null, "career")
%>
		<TR>
			<TD COLSPAN=3>&nbsp;</TD>
		</TR>

<%
		call prepareOffenseQuery(null, "season", "goals", "PPG")
%>
		<TR>
			<TD COLSPAN=3>&nbsp;</TD>
		</TR>

<%
		call prepareOffenseQuery(null, "game", "goals", "PPG")
%>
		<TR>
			<TD COLSPAN=3>&nbsp;</TD>
		</TR>

<%
		call prepareOffenseQuery(null, "career", "goals", "PPG")
%>
		<TR>
			<TD COLSPAN=3>&nbsp;</TD>
		</TR>

<%
		call prepareOffenseQuery(null, "season", "assists", "PPG")
%>
		<TR>
			<TD COLSPAN=3>&nbsp;</TD>
		</TR>

<%
		call prepareOffenseQuery(null, "game", "assists", "PPG")
%>
		<TR>
			<TD COLSPAN=3>&nbsp;</TD>
		</TR>

<%
		call prepareOffenseQuery(null, "career", "assists", "PPG")
%>
		<TR>
			<TD COLSPAN=3>&nbsp;</TD>
		</TR>

<%
		call prepareOffenseQuery(null, "season", "goals", "SH")
%>
		<TR>
			<TD COLSPAN=3>&nbsp;</TD>
		</TR>

<%
		call prepareOffenseQuery(null, "game", "goals", "SH")
%>
		<TR>
			<TD COLSPAN=3>&nbsp;</TD>
		</TR>

<%
		call prepareOffenseQuery(null, "career", "goals", "SH")
%>
		<TR>
			<TD COLSPAN=3>&nbsp;</TD>
		</TR>

		
<%
		call prepareOffenseQuery(null, "season", "assists", "SH")
%>
		<TR>
			<TD COLSPAN=3>&nbsp;</TD>
		</TR>

<%
		call prepareOffenseQuery(null, "game", "assists", "SH")
%>
		<TR>
			<TD COLSPAN=3>&nbsp;</TD>
		</TR>

<%
		call prepareOffenseQuery(null, "career", "assists", "SH")
%>		
		<TR>
			<TD COLSPAN=3>&nbsp;</TD>
		</TR>

<%
		call prepareOffenseQuery(null, "season", "goals", "GW")
%>
		<TR>
			<TD COLSPAN=3>&nbsp;</TD>
		</TR>

<%
		call prepareOffenseQuery(null, "career", "goals", "GW")
%>
		<TR>
			<TD COLSPAN=3>&nbsp;</TD>
		</TR>

<%
		call prepareGamesQuery(10)
%>
		</TABLE>
			<a href="garibotto.html">Mark Garibotto's records</a>

	</TD>
</TR>
</TABLE>
<BR><BR><BR>
</BODY>
</HTML>

<%
Function prepareOffenseQuery(num, timeSpan, stat, modifier)
	dim newTops, strSQL, selectClause, statClause, join2, join3, whereClause, groupby, orderby
	
	if timeSpan = "career" then
		selectClause = "year(p.DateActivated) as start, year(p.DateDeactivated) as end, "
		join2 = ""
		join3 = ""
		groupby = ", p.DateActivated, p.DateDeactivated "
		orderby = ""
	elseif timespan = "season" then
		selectClause = "s.SeasonName as season, s.SeasonID as seasonId, "
		join2 = "inner join tbl_Games game on g.GameID = game.GameID"
		join3 = "inner join tbl_Seasons s on game.SeasonID = s.SeasonID"
		groupby = ", s.SeasonName, s.SeasonID"
		orderby = ", s.SeasonID"
	elseif timespan = "game" then
		selectClause = "game.GameDate as gameDate, game.GameID as gameId, t.TeamName as team, "
		join2 = "inner join tbl_Games game on g.GameID = game.GameID"
		join3 = "inner join tbl_Teams t on game.HomeTeamID = t.TeamID or game.AwayTeamID = t.TeamID"
		whereClause = " and t.TeamID <> 1"
		groupby = ", game.GameID, game.GameDate, t.TeamName"
		orderby = ", game.GameDate"
	end if
	
	if stat = "points" then
		statClause = "g.PlayerID = p.PlayerID or p.PlayerID = g.Assist1PlayerID or p.PlayerID = g.Assist2PlayerID"
	elseif stat = "goals" then
		statClause = "g.PlayerID = p.PlayerID"
	elseif stat = "assists" then 
		statClause = "p.PlayerID = g.Assist1PlayerID or p.PlayerID = g.Assist2PlayerID"
	end if
	
	'modifier can be GW, GT, PPG, SH, EN
	if modifier <> "" then
		whereClause = whereClause & " and g." & modifier & " = true"
	end if
	
	if isNull(num) Then
		num = 6
	End If
	
	strSQL = "select top " & num & " p.PlayerID, p.FirstName, p.LastName, " & selectClause & " count(*) as total"
	strSQL = strSQL & " from ((tbl_Goals g inner join tbl_Players p on " & statClause & ") "
	strSQL = strSQL & join2 & ") "
	strSQL = strSQL & join3
	strSQL = strSQL & " where p.TeamID = 1 and p.PlayerID <> 290 " & whereClause
	strSQL = strSQL & " group by p.PlayerID, p.LastName, p.FirstName" & groupby
	strSQL = strSQL & " order by count(*) desc" & orderby
	
	Set rs = dbConn.Execute(strSQL)
	Set newTops = New Tops
	newTops.processRS num, rs, timeSpan

	Dim title
	If modifier = "PPG" Then
		modifier = "PP"
	End If
	title = "Most " & modifier & " " & stat & " (" & timeSpan & ")"
	
	printRecords newTops, timeSpan, title, num, false
	
End Function

Function preparePimQuery(num, timeSpan)
	dim newTops, strSQL, selectClause, statClause, join2, join3, whereClause, groupby, orderby
	
	if timeSpan = "career" then
		selectClause = "year(pl.DateActivated) as start, year(pl.DateDeactivated) as end, "
		join2 = ""
		join3 = ""
		groupby = ", pl.DateActivated, pl.DateDeactivated "
		orderby = ""
	elseif timespan = "season" then
		selectClause = "s.SeasonName as season, s.SeasonID as seasonId, "
		join2 = "inner join tbl_Games game on p.GameID = game.GameID"
		join3 = "inner join tbl_Seasons s on game.SeasonID = s.SeasonID"
		groupby = ", s.SeasonName, s.SeasonID"
		orderby = ", s.SeasonID"
	elseif timespan = "game" then
		selectClause = "game.GameDate as gameDate, game.GameID as gameId,  t.TeamName as team, "
		join2 = "inner join tbl_Games game on p.GameID = game.GameID"
		join3 = "inner join tbl_Teams t on game.HomeTeamID = t.TeamID or game.AwayTeamID = t.TeamID"
		whereClause = " and t.TeamID <> 1"
		groupby = ", game.GameID, game.GameDate, t.TeamName"
		orderby = ", game.GameDate"
	end if
	
	if isNull(num) Then
		num = 6
	End If

	strSQL = "select top " & num & " pl.PlayerID, pl.FirstName, pl.LastName, " & selectClause & " sum(p.length) as total"
	strSQL = strSQL & " from ((tbl_Penalties p inner join  tbl_Players pl on p.PlayerID = pl.PlayerID)"
	strSQL = strSQL & join2 & ")"
	strSQL = strSQL & join3
	strSQL = strSQL & " where pl.TeamID = 1 " & whereClause
	strSQL = strSQL & " group by pl.PlayerID, pl.LastName, pl.FirstName" & groupby
	strSQL = strSQL & " order by sum(p.length) desc" & orderby
	
	Set rs = dbConn.Execute(strSQL)
	Set newTops = New Tops
	newTops.processRS num, rs, timeSpan

	Dim title
	title = "Most PIMs (" & timeSpan & ")"
	
	printRecords newTops, timeSpan, title, num, false

End Function

' Save % and GAA
Function prepareGoalieQuery(num, timeSpan, stat)
	dim newTops, strSQL, selectClause, subSelectClause, statClause, join2, where, groupby, orderby
	dim title

	if timeSpan = "career" then
		selectClause = "year(tmpView.DateActivated) as start, year(tmpView.DateDeactivated) as end, "
		subSelectClause = "p.DateActivated, p.DateDeactivated,"
		join2 = ""
		groupby = ", tmpView.DateActivated, tmpView.DateDeactivated "
		orderby = ""
	elseif timespan = "season" then
		selectClause = "s.SeasonName as season, s.SeasonID as seasonId, "
		subSelectClause = " s.SeasonID, s.SeasonName,"
		join2 = "inner join tbl_Seasons s on g.SeasonID = s.SeasonID"
		groupby = ", s.SeasonName, s.SeasonID"
		orderby = ", s.SeasonID"
	end if
	
	if stat = "spct" then
		statClause = "format((sum(shots) - sum(goals)) / sum(shots), '.000') AS total,"
    	title = "Best SV % (" & timeSpan & ")"
    	orderby = "(sum(shots) - sum(goals)) / sum(shots) DESC"
		where = "shots <> 0"
	elseif stat = "gaa" then
		statClause = "format(Avg(goals), '0.00') AS total,"	  
		orderby = "Avg(goals) ASC"  & orderby
	    title = "Best GAA (" & timeSpan & ")"
		where = "1 = 1"
	end if

	if isNull(num) Then
		num = 3
	End If
	
	strSQL = "select top " & num & " tmpView.PlayerID, tmpView.FirstName, tmpView.LastName, tmpView.PlayerID, " & selectClause & statClause & " Count(*) AS games"
	strSQL = strSQL & " FROM [select *, (shots - goals) / shots AS pct from ("
	strSQL = strSQL & " SELECT p.FirstName, p.LastName, p.PlayerID, " & subSelectClause
	strSQL = strSQL & " iif(t.TeamID <> g.HomeTeamID, "
	strSQL = strSQL & " g.FirstPeriodHomeShots +  g.SecondPeriodHomeShots +  g.ThirdPeriodHomeShots +  g.OvertimeHomeShots, "
	strSQL = strSQL & " g.FirstPeriodAwayShots +  g.SecondPeriodAwayShots +  g.ThirdPeriodAwayShots +  g.OvertimeAwayShots) AS shots, "
	strSQL = strSQL & " iif(t.TeamID <> g.HomeTeamID, "
	strSQL = strSQL & " g.FirstPeriodHomeGoals +  g.SecondPeriodHomeGoals +  g.ThirdPeriodHomeGoals +  g.OvertimeHomeGoals, "
	strSQL = strSQL & " g.FirstPeriodAwayGoals +  g.SecondPeriodAwayGoals +  g.ThirdPeriodAwayGoals +  g.OvertimeAwayGoals) AS goals"
	strSQL = strSQL & " FROM (((tbl_Players AS p INNER JOIN tbl_PlayerGames AS pg ON p.PlayerID = pg.PlayerID) "
	strSQL = strSQL & " INNER JOIN tbl_Games AS g ON g.GameID = pg.GameID) "
	strSQL = strSQL & " INNER JOIN tbl_Teams AS t ON t.TeamID = g.HomeTeamID or t.TeamID = g.AwayTeamID)"
	strSQL = strSQL & join2
	strSQL = strSQL & " WHERE p.TeamID = 1 and p.Position = 5 and t.TeamID = 1"
	strSQL = strSQL & " ) where " & where & "]. AS tmpView"
	strSQL = strSQL & " GROUP BY tmpView.PlayerID, tmpView.FirstName, tmpView.LastName, tmpView.PlayerID " & groupby
	strSQL = strSQL & " HAVING (((Count(*))>5))"
	strSQL = strSQL & " ORDER BY " & orderby

	Set rs = dbConn.Execute(strSQL)
	Set newTops = New Tops
	newTops.processRS num, rs, timeSpan
	
	printRecords newTops, timeSpan, title, num, true	
End Function

Function prepareShutoutQuery(num, timeSpan)
	dim newTops, strSQL
	dim selectClause, statClause, join2, whereClause, groupby, orderby
	
	if timeSpan = "career" then
		selectClause = "year(p.DateActivated) as start, year(p.DateDeactivated) as end, "
		join2 = ""
		groupby = ", p.DateActivated, p.DateDeactivated "
		orderby = ""
	elseif timespan = "season" then
		selectClause = "s.SeasonName as season, s.SeasonID as seasonId, "
    	join2 = "inner join tbl_Seasons s on g.SeasonID = s.SeasonID"
		groupby = ", s.SeasonName, s.SeasonID"
		orderby = ", s.SeasonID"
	end if
	
	if isNull(num) Then
		num = 3
	End If

    strSQL = "select top " & num & " p.PlayerID, p.lastname, p.firstname, " & selectClause & " count(*) as total"
    strSQL = strSQL & " from ((tbl_games g inner join  tbl_playergames pg on g.GameID = pg.GameID) inner join tbl_players p on p.PlayerID = pg.PlayerID) "
    strSQL = strSQL & join2
    strSQL = strSQL & " where p.teamid=1 and p.position =5  and "
    strSQL = strSQL & " ((g.hometeamid = 1 and g.firstperiodawaygoals + g.secondperiodawaygoals + g.thirdperiodawaygoals + g.overtimeawaygoals = 0) or (g.awayteamid = 1 and g.firstperiodhomegoals + g.secondperiodhomegoals + g.thirdperiodhomegoals + g.overtimehomegoals = 0))"
    strSQL = strSQL & " group by p.playerid, p.lastname, p.firstname" & groupBy & " order by count(*) desc" & orderBy
	
	Set rs = dbConn.Execute(strSQL)
	Set newTops = New Tops
	newTops.processRS num, rs, timeSpan

	Dim title
	title = "Most shutouts (" & timeSpan & ")"
	printRecords newTops, timeSpan, title, num, true
End Function

Function prepareGamesQuery(num)
	dim newTops, strSQL

	if isNull(num) Then
		num = 3
	End If

	strSQL = "select top " & num & " p.PlayerID, p.FirstName, p.LastName, year(p.DateActivated) as start, year(p.DateDeactivated) as end, count(*) as total"
	strSQL = strSQL & " from tbl_PlayerGames pg inner join tbl_Players p on p.PlayerID = pg.PlayerID"
	strSQL = strSQL & " where p.TeamID = 1"
	strSQL = strSQL & " group by p.PlayerID, p.LastName, p.FirstName, p.DateActivated, p.DateDeactivated"
	strSQL = strSQL & " order by count(*) desc, p.LastName asc"
	
	Set rs = dbConn.Execute(strSQL)
	Set newTops = New Tops
	newTops.processRS num, rs, "career"

	Dim title
	title = "Most Games Played (career)"
	
	printRecords newTops, "career", title, num, false

End Function


Sub printRecords(newTops, timeSpan, title, num, isGoalie)
	goalieLink = ""
	if isGoalie then
		goalieLink = "&G=1"
	end if
	Dim i
	if timeSpan = "career" then
		Response.write "<TR>"
			Response.write "<TD ROWSPAN=2 VALIGN=TOP><B>" & title & ":</B></TD>"
			Response.write "<TD COLSPAN=2><a href=""/roster/default.asp?ID=" & newTops.playerId(0) & goalieLink & """>" & newTops.name(0) & "</a>, " & newTops.totals(0) & " (" & newTops.startYear(0) & " - " & newTops.endYear(0) & ")</TD>"
		Response.write "</TR>"
		if newTops.name(1) <> "" then
		    Response.write "<TR>"
			    Response.write "<TD WIDTH=20>&nbsp;</TD>"
    			Response.write "<TD><SMALL>"
	    		For i = 1 to num - 1
            		if newTops.name(i) <> "" then
		        		Response.write "<a href=""/roster/default.asp?ID=" & newTops.playerId(i) & goalieLink & """>" & newTops.name(i) & "</a>, " & newTops.totals(i) & " (" & newTops.startYear(i) & " - " & newTops.endYear(i) & ")<BR>"
		        	end if
			    Next
			    Response.write "</SMALL></TD>"
    		Response.write "</TR>"
		end if
	elseif timespan = "season" then
		Response.write "<TR>"
			Response.write "<TD ROWSPAN=2 VALIGN=TOP><B>" & title & ":</B></TD>"
			Response.write "<TD COLSPAN=2><a href=""/roster/default.asp?ID=" & newTops.playerId(0) & goalieLink & """>" & newTops.name(0) & "</a>, " & newTops.totals(0) & " (" & newTops.season(0) & ")</TD>"
		Response.write "</TR><TR>"
			Response.write "<TD WIDTH=20>&nbsp;</TD>"
			Response.write "<TD><SMALL>"
			For i = 1 to num - 1
				Response.write "<a href=""/roster/default.asp?ID=" & newTops.playerId(i) & goalieLink & """>" & newTops.name(i) & "</a>, " & newTops.totals(i) & " (" & newTops.season(i) & ")<BR>"
			Next
			Response.write "</SMALL></TD>"
		Response.write "</TR>"
	elseif timespan = "game" then
		Response.write "<TR>"
			Response.write "<TD ROWSPAN=2 VALIGN=TOP><B>" & title & ":</B></TD>"
			Response.write "<TD COLSPAN=2><a href=""/roster/default.asp?ID=" & newTops.playerId(0) & goalieLink & """>" & newTops.name(0) & "</a>, " & newTops.totals(0) & " (" & newTops.game(0) & " vs. " & newTops.gameTeam(0) & ")</TD>"
		Response.write "</TR>"
		Response.write "<TR>"
			Response.write "<TD WIDTH=20>&nbsp;</TD>"
			Response.write "<TD><SMALL>"
			For i = 1 to num - 1
				Response.write "<a href=""/roster/default.asp?ID=" & newTops.playerId(i) & goalieLink & """>" & newTops.name(i) & "</a>, " & newTops.totals(i) & " (" & newTops.game(i) & " vs. " & newTops.gameTeam(i) & ")<BR>"
			Next
			Response.write "</SMALL></TD>"
		Response.write "</TR>"		
	end if
End Sub

%>
