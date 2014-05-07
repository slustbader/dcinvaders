<%
	Option Explicit
	'On Error Resume Next

	Function FormatPhone(strPhone)
		Dim strFormattedPhone

		strFormattedPhone = "(" & Left(strPhone,3) & ") " & Mid(strPhone,4,3) & "-" & Right(strPhone,4)

		FormatPhone=strFormattedPhone

	End Function

	Function ChangeNullToZero(strText)
		If IsNull(strText) Then
			strText = 0
		End If
		ChangeNullToZero = strText
	End Function

	Function ChangeNullToNbsp(strText)
		If IsNull(strText) or strText = "" Then
			strText = "&nbsp;"
		End If
		ChangeNullToNbsp = strText
	End Function

	Dim dbConn
	Dim rsRoster
	Dim strSQL
	Dim strPlayerID
	Dim strFilter, strSortBy

	strSortBy = Request.QueryString("sortby")
	strFilter = Request.QueryString("selfilter")

	strPlayerID = request.QueryString("ID")

	Set dbConn = Server.CreateObject("ADODB.Connection")
	'dbConn.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("/invaders/database/hockey.mdb")
	dbConn.Open "DSN=hockey"
	
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<!-- #include virtual="/include/adovbs.inc"-->
<HTML>
<HEAD>
<TITLE>DC Invaders</TITLE>
<LINK REL="stylesheet" HREF="/styles/hockey.css" TYPE="text/css">
<SCRIPT LANGUAGE="JavaScript">
<!--

	function SpawnWindow(PlayerID, StatType, LastName, FirstName, Number, winwidth, winheight){
		strQueryString = "playerinfo.asp?ID=" + PlayerID + "&Stat=" + StatType + "&LN=" + LastName + "&FN=" + FirstName + "&N=" + Number;
		eval("window.open(strQueryString,'playerinfo','scrollbars=yes,width=" + winwidth + ",height=" + winheight + "')");
	}

//-->
</SCRIPT>
</HEAD>
<BODY BACKGROUND="/images/stars.gif">

<TABLE WIDTH=660 CELLPADDING=0 CELLSPACING=0 BORDER=0>
<TR>
	<TD><CENTER><IMG SRC="../images/invader.gif" BORDER=0>
	<H2><FONT COLOR="#FFFF66">Roster</FONT></H2></CENTER>

<%
	If IsNull(strPlayerID) or strPlayerID="" Then
%>

		<P><EM>You can filter the Invader's roster according to the player's position.  To do so, select the position you would like to filter for from the drop-down box below.  You can also sort the roster by clicking any of the column headings. To view detailed information on a particular player, click on his last name.</EM></P>
		<FORM NAME="filter" ACTION="default.asp" METHOD=GET>
		<B>Filter:</B> <SELECT NAME="selfilter" ONCHANGE="submit()"><OPTION VALUE="ALL" <% If strFilter="ALL" OR strFilter="" Then response.write "DEFAULT SELECTED" %>>All<OPTION VALUE="f" <% If strFilter="f" Then response.write "DEFAULT SELECTED" %>>Forwards<OPTION VALUE="lw" <% If strFilter="lw" Then response.write "DEFAULT SELECTED" %>>Left Wingers<OPTION VALUE="c" <% If strFilter="c" Then response.write "DEFAULT SELECTED" %>>Centers<OPTION VALUE="rw" <% If strFilter="rw" Then response.write "DEFAULT SELECTED" %>>Right Wingers<OPTION VALUE="d" <% If strFilter="d" Then response.write "DEFAULT SELECTED" %>>Defensemen<OPTION VALUE="g" <% If strFilter="g" Then response.write "DEFAULT SELECTED" %>>Goalies</SELECT>
		</FORM>

		<TABLE WIDTH=600 CELLPADDING=2 CELLSPACING=0 BORDER=1 BORDERCOLOR="black" BGCOLOR="#EEEEEE" rules="all">
		<TR BGCOLOR="#08479E">
			<TD WIDTH=160><B><A HREF="default.asp?sortby=l" ONMOUSEOVER="self.status='Last Name'; return true" ONMOUSEOUT="self.status=''; return true">Last Name</A></B></TD>
			<TD WIDTH=160><B><A HREF="default.asp?sortby=f" ONMOUSEOVER="self.status='First Name'; return true" ONMOUSEOUT="self.status=''; return true">First Name</A></B></TD>
			<TD WIDTH=160><B><A HREF="default.asp?sortby=p" ONMOUSEOVER="self.status='Position'; return true" ONMOUSEOUT="self.status=''; return true">Position</A></B></TD>
			<TD WIDTH=40><B><A HREF="default.asp?sortby=n" ONMOUSEOVER="self.status='Number'; return true" ONMOUSEOUT="self.status=''; return true">Number</A></B></TD>
			<TD WIDTH=40><B><A HREF="default.asp?sortby=sn" ONMOUSEOVER="self.status='Summer Jersey Number'; return true" ONMOUSEOUT="self.status=''; return true">Summer Jersey Number</A></B></TD>
			<TD WIDTH=40><B><A HREF="default.asp?sortby=s" ONMOUSEOVER="self.status='Shoots'; return true" ONMOUSEOUT="self.status=''; return true">Shoots</A></B></TD>
		</TR>

<%

		Select Case strFilter
			Case "f"
				strFilter = " AND (tbl_Positions.Position='Left Wing' OR tbl_Positions.Position='Center' OR tbl_Positions.Position='Right Wing')"
			Case "lw"
				strFilter = " AND tbl_Positions.Position='Left Wing'"
			Case "c"
				strFilter = " AND tbl_Positions.Position='Center'"
			Case "rw"
				strFilter = " AND tbl_Positions.Position='Right Wing'"
			Case "d"
				strFilter = " AND tbl_Positions.Position='Defenseman'"
			Case "g"
				strFilter = " AND tbl_Positions.Position='Goalie'"
			Case else
				strFilter = ""
		End Select

		Select Case strSortBy
			Case "p"
				strSortBy = " ORDER BY tbl_Positions.Position"
			Case "l"
				strSortBy = " ORDER BY LastName"
			Case "f"
				strSortBy = " ORDER BY FirstName"
			Case "n"
				strSortBy = " ORDER BY Number"
			Case "s"
				strSortBy = " ORDER BY LeftRight"
			Case "sn"
				strSortBy = " ORDER BY AlternateJerseyNumber"
			Case Else
				strSortBy = " ORDER BY LastName"
		End Select

		strSQL = "SELECT PlayerID, LastName, FirstName, tbl_Positions.Position, Number, LeftRight, AlternateJerseyNumber FROM tbl_Players LEFT JOIN tbl_Positions ON tbl_Players.Position = tbl_Positions.PositionID WHERE TeamID=1 AND Alternate=FALSE AND Active=TRUE" & strFilter & strSortBy
		Set rsRoster = Server.CreateObject("ADODB.Recordset")
		'response.write strSQL
		rsRoster.Open strSQL, dbConn, adOpenStatic

		If Not rsRoster.BOF and Not rsRoster.EOF Then

			rsRoster.MoveFirst


			Do While Not rsRoster.EOF
				response.write "<TR BGCOLOR=""#DCDCDC"">"
				response.write 	"<TD CLASS=""black""><A HREF=""default.asp?ID=" & rsRoster("PlayerID") & "&G="
					If rsRoster("Position") = "Goalie" Then
						response.write "1"
					End If
				response.write """ TARGET=""main"">" & rsRoster("LastName") & "</A></TD>"
				response.write 	"<TD CLASS=""black"">" & rsRoster("FirstName") & "</TD>"
				response.write 	"<TD CLASS=""black"">" & rsRoster("Position") & "</TD>"
				response.write	"<TD CLASS=""black"">"
				If Not IsNull(rsRoster("Number")) Then
					response.write	rsRoster("Number")
				Else
					response.write	"&nbsp;"
				End If
				response.write	"</TD>"
				response.write	"<TD CLASS=""black"">"
				If Not IsNull(rsRoster("AlternateJerseyNumber")) Then
					response.write	rsRoster("AlternateJerseyNumber")
				Else
					response.write	"&nbsp;"
				End If
				response.write	"</TD>"
				response.write	"<TD CLASS=""black"">"
					If Not IsNull(rsRoster("LeftRight")) Then
						response.write rsRoster("LeftRight")
					Else
						response.write "&nbsp;"
					End If
				response.write "</TD>"
				response.write "</TR>"

				rsRoster.MoveNext
			Loop

		End If

		rsRoster.Close
	%>

		</TABLE>
	<%
	Else

		strSQL = "SELECT * FROM tbl_Players LEFT JOIN tbl_Positions ON tbl_Players.Position = tbl_Positions.PositionID WHERE PlayerID=" & strPlayerID
		Set rsRoster = dbConn.Execute(strSQL)

		Dim strHobbies
		strHobbies = rsRoster("Hobbies")

		response.write "<TABLE WIDTH=100% CELLPADDING=0 CELLSPACING=0 BORDER=0>"
		response.write "<TR>"
		Dim strNickName
		strNickName = " "
		if (rsRoster("NickName") <> null) or (rsRoster("NickName") <> "") Then strNickName = " (" & rsRoster("NickName") & ") "
		response.write "	<TD VALIGN=TOP COLSPAN=2><H3>" & rsRoster("FirstName") & strNickName & rsRoster("LastName") & "</H3></TD>"
		response.write "	<TD ALIGN=RIGHT VALIGN=TOP ROWSPAN=2>"
						If Not IsNull(rsRoster("Photo")) Then response.write "<IMG SRC=""" & rsRoster("Photo") & """>"
		response.write "<BR></TD>"
		response.write "</TR>"

		response.write "<TR>"
		response.write "	<TD WIDTH=35% VALIGN=TOP>"
						If Not IsNull(rsRoster("Number")) Then response.write "<FONT COLOR=""#FFFF66""><B>Number:</B></FONT> " & rsRoster("Number") & "<BR>"
						If Not IsNull(rsRoster("Position")) Then response.write "<FONT COLOR=""#FFFF66""><B>Position:</B></FONT> " & rsRoster("Position") & "<BR>"
						If Not IsNull(rsRoster("LeftRight")) Then response.write "<FONT COLOR=""#FFFF66""><B>Shoots:</B></FONT> " & rsRoster("LeftRight") & "<BR>"
						If Not IsNull(rsRoster("Height")) AND rsRoster("Height")<>0 Then response.write "<FONT COLOR=""#FFFF66""><B>Height:</B></FONT> " & Int(rsRoster("Height")/12) & "' " & rsRoster("Height") Mod 12 & """<BR>"
						If Not IsNull(rsRoster("Weight")) AND rsRoster("Weight")<>0 Then response.write "<FONT COLOR=""#FFFF66""><B>Weight:</B></FONT> " & rsRoster("Weight") & "lbs<BR>"
		response.write "	</TD>"
		response.write "	<TD VALIGN=TOP>"
						If Not IsNull(rsRoster("Birthdate")) Then response.write "<FONT COLOR=""#FFFF66""><B>Birth Date:</B></FONT> " & FormatDateTime(rsRoster("Birthdate"),2) & "<BR>"
						If Not IsNull(rsRoster("Birthplace")) Then response.write "<FONT COLOR=""#FFFF66""><B>Birth Place:</B></FONT> " & rsRoster("Birthplace") & "<BR>"
						If Not IsNull(strHobbies) Then response.write "<FONT COLOR=""#FFFF66""><B>Hobbies:</B></FONT> " & strHobbies & "<BR>"
						If Not IsNull(rsRoster("FavoriteTeam")) Then response.write "<FONT COLOR=""#FFFF66""><B>Favorite Team:</B></FONT> " & rsRoster("FavoriteTeam") & "<BR>"
						If Not IsNull(rsRoster("FavoriteHockeyPlayer")) Then response.write "<FONT COLOR=""#FFFF66""><B>Favorite Hockey Player:</B></FONT> " & rsRoster("FavoriteHockeyPlayer") & "<BR>"
						If Not IsNull(rsRoster("FavoriteFood")) Then response.write "<FONT COLOR=""#FFFF66""><B>Favorite Food:</B></FONT> " & rsRoster("FavoriteFood") & "<BR>"
		response.write "	</TD>"
		response.write "</TR>"
		response.write "</TABLE>"

		response.write "<P><CENTER><H3>Personal Stats</H3></CENTER></P>"
		response.write "<EM>Click on the column headings to see a breakdown of the player's corresponding statsistics.</EM><BR><BR>"

		Dim rsSeasons
		Dim fmG
		Dim rs, strGP, strPIM

		strSQL = "SELECT SeasonName FROM tbl_Seasons"
		Set rsSeasons = dbConn.Execute(strSQL)

		fmG = Request.QueryString("G")

		If fmG = "1" Then

%>
			<CENTER>
			<TABLE WIDTH=100% CELLPADDING=3 CELLSPACING=0 BORDER=1 BORDERCOLOR="black" BGCOLOR="#EEEEEE" rules="all">
			<TR BGCOLOR="#08479E">
				<TH WIDTH=30%>Season</TH>
				<TH WIDTH=7%>GP</TH>
				<TH WIDTH=7%>W</TH>
				<TH WIDTH=7%>L</TH>
				<TH WIDTH=7%>T</TH>
				<TH WIDTH=7%>SA</TH>
				<TH WIDTH=7%>GA</TH>
				<TH WIDTH=7%>GAA</TH>
				<TH WIDTH=7%>SO</TH>
				<TH WIDTH=7%>SV%</TH>
				<TH WIDTH=7%>PIM</TH>
			</TR>

<%
			rsSeasons.MoveFirst

			Do While Not rsSeasons.EOF
%>
				<TR BGCOLOR="#DCDCDC">
					<TD CLASS="black"><SMALL><%=rsSeasons.Fields(0) %></SMALL></TD>

<%
					strSQL = "SELECT tbl_Players.PlayerID, tbl_Players.FirstName, tbl_Players.LastName, (SELECT COUNT(*) FROM tbl_Seasons INNER JOIN (tbl_Games INNER JOIN tbl_PlayerGames ON tbl_Games.GameID = tbl_PlayerGames.GameID) ON tbl_Seasons.SeasonID = tbl_Games.SeasonID WHERE tbl_PlayerGames.PlayerID=tbl_Players.PlayerID AND tbl_Seasons.SeasonName='" & rsSeasons.Fields(0) & "') AS GP, (SELECT COUNT(*) FROM tbl_Seasons INNER JOIN (tbl_Games INNER JOIN tbl_PlayerGames ON tbl_Games.GameID = tbl_PlayerGames.GameID) ON tbl_Seasons.SeasonID = tbl_Games.SeasonID WHERE (((IIf(([HomeTeamID]=tbl_Players.TeamID And ([FirstPeriodHomeGoals]+[SecondPeriodHomeGoals]+[ThirdPeriodHomeGoals]+[OvertimeHomeGoals])>([FirstPeriodAwayGoals]+[SecondPeriodAwayGoals]+[ThirdPeriodAwayGoals]+[OvertimeAwayGoals])) Or ([AwayTeamID]=tbl_Players.TeamID And ([FirstPeriodAwayGoals]+[SecondPeriodAwayGoals]+[ThirdPeriodAwayGoals]+[OvertimeAwayGoals])>([FirstPeriodHomeGoals]+[SecondPeriodHomeGoals]+[ThirdPeriodHomeGoals]+[OvertimeHomeGoals])),'True','False'))='True') AND ((tbl_PlayerGames.PlayerID)=tbl_Players.PlayerID) AND ((tbl_Seasons.SeasonName)='" & rsSeasons.Fields(0) & "'))) AS W, (SELECT COUNT(*) FROM tbl_Seasons INNER JOIN (tbl_Games INNER JOIN tbl_PlayerGames ON tbl_Games.GameID = tbl_PlayerGames.GameID) ON tbl_Seasons.SeasonID = tbl_Games.SeasonID WHERE (((IIf(([HomeTeamID]=tbl_Players.TeamID And ([FirstPeriodHomeGoals]+[SecondPeriodHomeGoals]+[ThirdPeriodHomeGoals]+[OvertimeHomeGoals])<([FirstPeriodAwayGoals]+[SecondPeriodAwayGoals]+[ThirdPeriodAwayGoals]+[OvertimeAwayGoals])) Or ([AwayTeamID]=tbl_Players.TeamID And ([FirstPeriodAwayGoals]+[SecondPeriodAwayGoals]+[ThirdPeriodAwayGoals]+[OvertimeAwayGoals])<([FirstPeriodHomeGoals]+[SecondPeriodHomeGoals]+[ThirdPeriodHomeGoals]+[OvertimeHomeGoals])),'True','False'))='True') AND ((tbl_PlayerGames.PlayerID)=tbl_Players.PlayerID) AND ((tbl_Seasons.SeasonName)='" & rsSeasons.Fields(0) & "'))) AS L, (SELECT COUNT(*) FROM tbl_Seasons INNER JOIN (tbl_Games INNER JOIN tbl_PlayerGames ON tbl_Games.GameID = tbl_PlayerGames.GameID) ON tbl_Seasons.SeasonID = tbl_Games.SeasonID WHERE ((([FirstPeriodHomeGoals]+[SecondPeriodHomeGoals]+[ThirdPeriodHomeGoals]+[OvertimeHomeGoals])=([FirstPeriodAwayGoals]+[SecondPeriodAwayGoals]+[ThirdPeriodAwayGoals]+[OvertimeAwayGoals])) AND ((tbl_PlayerGames.PlayerID)=tbl_Players.PlayerID) AND ((tbl_Seasons.SeasonName)='" & rsSeasons.Fields(0) & "'))) AS T, (SELECT Sum(IIf([HomeTeamID]=tbl_Players.TeamID,[FirstPeriodAwayShots]+[SecondPeriodAwayShots]+[ThirdPeriodAwayShots]+[OvertimeAwayShots],[FirstPeriodHomeShots]+[SecondPeriodHomeShots]+[ThirdPeriodHomeShots]+[OvertimeHomeShots])) AS Expr1 FROM tbl_Seasons INNER JOIN (tbl_Games INNER JOIN tbl_PlayerGames ON tbl_Games.GameID = tbl_PlayerGames.GameID) ON tbl_Seasons.SeasonID = tbl_Games.SeasonID GROUP BY tbl_PlayerGames.PlayerID, tbl_Seasons.SeasonName HAVING (((tbl_PlayerGames.PlayerID)=tbl_Players.PlayerID) AND ((tbl_Seasons.SeasonName)='" & rsSeasons.Fields(0) & "'))) AS SA, (SELECT Sum(IIf([HomeTeamID]=tbl_Players.TeamID,[FirstPeriodAwayGoals]+[SecondPeriodAwayGoals]+[ThirdPeriodAwayGoals]+[OvertimeAwayGoals],[FirstPeriodHomeGoals]+[SecondPeriodHomeGoals]+[ThirdPeriodHomeGoals]+[OvertimeHomeGoals])) AS Expr1 FROM tbl_Seasons INNER JOIN (tbl_Games INNER JOIN tbl_PlayerGames ON tbl_Games.GameID = tbl_PlayerGames.GameID) ON tbl_Seasons.SeasonID = tbl_Games.SeasonID GROUP BY tbl_PlayerGames.PlayerID, tbl_Seasons.SeasonName HAVING (((tbl_PlayerGames.PlayerID)=tbl_Players.PlayerID) AND ((tbl_Seasons.SeasonName)='" & rsSeasons.Fields(0) & "'))) AS GA, Format([GA]/[GP],'0.00') AS GAA, (SELECT Count(*) FROM tbl_Seasons INNER JOIN (tbl_Games INNER JOIN tbl_PlayerGames ON tbl_Games.GameID = tbl_PlayerGames.GameID) ON tbl_Seasons.SeasonID = tbl_Games.SeasonID WHERE (((IIf([HomeTeamID]=tbl_Players.TeamID And [FirstPeriodAwayGoals]+[SecondPeriodAwayGoals]+[ThirdPeriodAwayGoals]+[OvertimeAwayGoals]=0,'True',IIf([AwayTeamID]=tbl_Players.TeamID And [FirstPeriodHomeGoals]+[SecondPeriodHomeGoals]+[ThirdPeriodHomeGoals]+[OvertimeHomeGoals]=0,'True')))='True') AND ((tbl_PlayerGames.PlayerID)=tbl_Players.PlayerID)) GROUP BY tbl_Seasons.SeasonName HAVING (((tbl_Seasons.SeasonName)='" & rsSeasons.Fields(0) & "'))) AS SO, Format(([SA]-[GA])/[SA],'0.000') AS SPCT, (SELECT Sum(Length) FROM tbl_Seasons INNER JOIN (tbl_Games INNER JOIN tbl_Penalties ON tbl_Games.GameID = tbl_Penalties.GameID) ON tbl_Seasons.SeasonID = tbl_Games.SeasonID WHERE (((tbl_Penalties.PlayerID)=tbl_Players.PlayerID) AND ((tbl_Seasons.SeasonName)='" & rsSeasons.Fields(0) & "'))) AS PIM, tbl_Players.Alternate FROM tbl_Players WHERE PlayerID=" & strPlayerID
					Set rs = dbConn.Execute(strSQL)
%>

					<TD CLASS="black" ALIGN=CENTER><%=rs("GP") %></TD>
					<TD CLASS="black" ALIGN=CENTER><%=rs("W") %></TD>
					<TD CLASS="black" ALIGN=CENTER><%=rs("L") %></TD>
					<TD CLASS="black" ALIGN=CENTER><%=rs("T") %></TD>
					<TD CLASS="black" ALIGN=CENTER><%=ChangeNullToZero(rs("SA")) %></TD>
					<TD CLASS="black" ALIGN=CENTER><%=ChangeNullToZero(rs("GA")) %></TD>
					<TD CLASS="black" ALIGN=CENTER><%=ChangeNullToNbsp(rs("GAA")) %></TD>
					<TD CLASS="black" ALIGN=CENTER><%=ChangeNullToZero(rs("SO")) %></TD>
					<TD CLASS="black" ALIGN=CENTER><%=ChangeNullToNbsp(rs("SPCT")) %></TD>
					<TD CLASS="black" ALIGN=CENTER><%=ChangeNullToZero(rs("PIM")) %></TD>
				</TR>
<%
				rsSeasons.MoveNext
			Loop
%>
			<TR BGCOLOR="#DCDCDC">
				<TD CLASS="black"><B><FONT COLOR="#3333FF">Career</FONT></B></TD>
<%
					strSQL = "SELECT tbl_Players.PlayerID, tbl_Players.FirstName, tbl_Players.LastName, (SELECT COUNT(*) FROM tbl_PlayerGames WHERE tbl_PlayerGames.PlayerID=tbl_Players.PlayerID) AS GP, (SELECT COUNT(*) FROM tbl_Games INNER JOIN tbl_PlayerGames ON tbl_Games.GameID = tbl_PlayerGames.GameID WHERE (IIf(([HomeTeamID]=tbl_Players.TeamID AND ([FirstPeriodHomeGoals]+[SecondPeriodHomeGoals]+[ThirdPeriodHomeGoals]+[OvertimeHomeGoals])>([FirstPeriodAwayGoals]+[SecondPeriodAwayGoals]+[ThirdPeriodAwayGoals]+[OvertimeAwayGoals])) Or ([AwayTeamID]=tbl_Players.TeamID AND ([FirstPeriodAwayGoals]+[SecondPeriodAwayGoals]+[ThirdPeriodAwayGoals]+[OvertimeAwayGoals])>([FirstPeriodHomeGoals]+[SecondPeriodHomeGoals]+[ThirdPeriodHomeGoals]+[OvertimeHomeGoals])), 'True', 'False')='True') AND tbl_PlayerGames.PlayerID=tbl_Players.PlayerID) AS W, (SELECT COUNT(*) FROM tbl_Games INNER JOIN tbl_PlayerGames ON tbl_Games.GameID = tbl_PlayerGames.GameID WHERE (IIf(([HomeTeamID]=tbl_Players.TeamID AND ([FirstPeriodHomeGoals]+[SecondPeriodHomeGoals]+[ThirdPeriodHomeGoals]+[OvertimeHomeGoals])<([FirstPeriodAwayGoals]+[SecondPeriodAwayGoals]+[ThirdPeriodAwayGoals]+[OvertimeAwayGoals])) Or ([AwayTeamID]=tbl_Players.TeamID AND ([FirstPeriodAwayGoals]+[SecondPeriodAwayGoals]+[ThirdPeriodAwayGoals]+[OvertimeAwayGoals])<([FirstPeriodHomeGoals]+[SecondPeriodHomeGoals]+[ThirdPeriodHomeGoals]+[OvertimeHomeGoals])), 'True', 'False')='True') AND tbl_PlayerGames.PlayerID=tbl_Players.PlayerID) AS L, (SELECT COUNT(*) FROM tbl_Games INNER JOIN tbl_PlayerGames ON tbl_Games.GameID = tbl_PlayerGames.GameID WHERE (([FirstPeriodHomeGoals]+[SecondPeriodHomeGoals]+[ThirdPeriodHomeGoals]+[OvertimeHomeGoals])=([FirstPeriodAwayGoals]+[SecondPeriodAwayGoals]+[ThirdPeriodAwayGoals]+[OvertimeAwayGoals])) AND tbl_PlayerGames.PlayerID=tbl_Players.PlayerID) AS T, (SELECT Sum(IIf([TeamID]=[HomeTeamID],[FirstPeriodAwayShots]+[SecondPeriodAwayShots]+[ThirdPeriodAwayShots]+[OvertimeAwayShots],[FirstPeriodHomeShots]+[SecondPeriodHomeShots]+[ThirdPeriodHomeShots]+[OvertimeHomeShots])) FROM tbl_Games INNER JOIN tbl_PlayerGames ON tbl_Games.GameID=tbl_PlayerGames.GameID GROUP BY tbl_PlayerGames.PlayerID HAVING tbl_PlayerGames.PlayerID=tbl_Players.PlayerID) AS SA, (SELECT Sum(IIf([TeamID]=[HomeTeamID],[FirstPeriodAwayGoals]+[SecondPeriodAwayGoals]+[ThirdPeriodAwayGoals]+[OvertimeAwayGoals],[FirstPeriodHomeGoals]+[SecondPeriodHomeGoals]+[ThirdPeriodHomeGoals]+[OvertimeHomeGoals])) FROM tbl_Games INNER JOIN tbl_PlayerGames ON tbl_Games.GameID=tbl_PlayerGames.GameID GROUP BY tbl_PlayerGames.PlayerID HAVING tbl_PlayerGames.PlayerID=tbl_Players.PlayerID) AS GA, Format([GA]/[GP],'0.00') AS GAA, (SELECT Count(*) FROM tbl_Games INNER JOIN tbl_PlayerGames ON tbl_Games.GameID = tbl_PlayerGames.GameID WHERE (((IIf([TeamID]=[HomeTeamID] And [FirstPeriodAwayGoals]+[SecondPeriodAwayGoals]+[ThirdPeriodAwayGoals]+[OvertimeAwayGoals]=0,'True',IIf([TeamID]=[AwayTeamID] And [FirstPeriodHomeGoals]+[SecondPeriodHomeGoals]+[ThirdPeriodHomeGoals]+[OvertimeHomeGoals]=0,'True')))='True') AND ((tbl_PlayerGames.PlayerID)=tbl_Players.PlayerID))) AS SO, Format(([SA]-[GA])/[SA],'0.000') AS SPCT, (SELECT Sum(Length) FROM tbl_Penalties GROUP BY PlayerID HAVING PlayerID=tbl_Players.PlayerID) AS PIM, tbl_Players.Alternate FROM tbl_Players WHERE PlayerID=" & strPlayerID
					Set rs = dbConn.Execute(strSQL)
%>

				<TD CLASS="black" ALIGN=CENTER><B><FONT COLOR="#3333FF"><%=rs("GP") %></FONT></B></TD>
				<TD CLASS="black" ALIGN=CENTER><B><FONT COLOR="#3333FF"><%=rs("W") %></FONT></B></TD>
				<TD CLASS="black" ALIGN=CENTER><B><FONT COLOR="#3333FF"><%=rs("L") %></FONT></B></TD>
				<TD CLASS="black" ALIGN=CENTER><B><FONT COLOR="#3333FF"><%=rs("T") %></FONT></B></TD>
				<TD CLASS="black" ALIGN=CENTER><B><FONT COLOR="#3333FF"><%=rs("SA") %></FONT></B></TD>
				<TD CLASS="black" ALIGN=CENTER><B><FONT COLOR="#3333FF"><%=rs("GA") %></FONT></B></TD>
				<TD CLASS="black" ALIGN=CENTER><B><FONT COLOR="#3333FF"><%=rs("GAA") %></FONT></B></TD>
				<TD CLASS="black" ALIGN=CENTER><B><FONT COLOR="#3333FF"><%=rs("SO") %></FONT></B></TD>
				<TD CLASS="black" ALIGN=CENTER><B><FONT COLOR="#3333FF"><%=rs("SPCT") %></FONT></B></TD>
				<TD CLASS="black" ALIGN=CENTER><B><FONT COLOR="#3333FF"><%=ChangeNullToZero(rs("PIM")) %></FONT></B></TD>
			</TR>
			</TABLE>
			</CENTER>

<%
			rs.Close
			Set rs = nothing

		Else

%>

			<TABLE WIDTH=100% CELLPADDING=3 CELLSPACING=0 BORDER=1 BORDERCOLOR="black" BGCOLOR="#EEEEEE" rules="all">
			<TR BGCOLOR="#08479E">
				<TH WIDTH=23%>Season</TH>
				<TH WIDTH=7%><A HREF="javascript: SpawnWindow('<%=strPlayerID%>','GP','<%=rsRoster("LastName")%>','<%=rsRoster("FirstName")%>','<%=rsRoster("Number")%>',300,200)" ONMOUSEOVER="self.status='Games Played'; return true" ONMOUSEOUT="self.status=''; return true">GP</A></TH>
				<TH WIDTH=7%><A HREF="javascript: SpawnWindow('<%=strPlayerID%>','G','<%=rsRoster("LastName")%>','<%=rsRoster("FirstName")%>','<%=rsRoster("Number")%>',400,200)" ONMOUSEOVER="self.status='Goals'; return true" ONMOUSEOUT="self.status=''; return true">G</A></TH>
				<TH WIDTH=7%><A HREF="javascript: SpawnWindow('<%=strPlayerID%>','A','<%=rsRoster("LastName")%>','<%=rsRoster("FirstName")%>','<%=rsRoster("Number")%>',400,200)" ONMOUSEOVER="self.status='Assists'; return true" ONMOUSEOUT="self.status=''; return true">A</A></TH>
				<TH WIDTH=7%>PTS</TH>
				<TH WIDTH=7%><A HREF="javascript: SpawnWindow('<%=strPlayerID%>','PIM','<%=rsRoster("LastName")%>','<%=rsRoster("FirstName")%>','<%=rsRoster("Number")%>',500,200)" ONMOUSEOVER="self.status='Games Played'; return true" ONMOUSEOUT="self.status=''; return true">PIM</A></TH>
				<TH WIDTH=7%><A HREF="javascript: SpawnWindow('<%=strPlayerID%>','PPG','<%=rsRoster("LastName")%>','<%=rsRoster("FirstName")%>','<%=rsRoster("Number")%>',400,200)" ONMOUSEOVER="self.status='Power Play Goals'; return true" ONMOUSEOUT="self.status=''; return true">PPG</A></TH>
				<TH WIDTH=7%><A HREF="javascript: SpawnWindow('<%=strPlayerID%>','PPA','<%=rsRoster("LastName")%>','<%=rsRoster("FirstName")%>','<%=rsRoster("Number")%>',400,200)" ONMOUSEOVER="self.status='Power Play Assists'; return true" ONMOUSEOUT="self.status=''; return true">PPA</A></TH>
				<TH WIDTH=7%><A HREF="javascript: SpawnWindow('<%=strPlayerID%>','SHG','<%=rsRoster("LastName")%>','<%=rsRoster("FirstName")%>','<%=rsRoster("Number")%>',400,200)" ONMOUSEOVER="self.status='Short Handed Goals'; return true" ONMOUSEOUT="self.status=''; return true">SHG</A></TH>
				<TH WIDTH=7%><A HREF="javascript: SpawnWindow('<%=strPlayerID%>','SHA','<%=rsRoster("LastName")%>','<%=rsRoster("FirstName")%>','<%=rsRoster("Number")%>',400,200)" ONMOUSEOVER="self.status='Short Handed Assists'; return true" ONMOUSEOUT="self.status=''; return true">SHA</A></TH>
				<TH WIDTH=7%><A HREF="javascript: SpawnWindow('<%=strPlayerID%>','GW','<%=rsRoster("LastName")%>','<%=rsRoster("FirstName")%>','<%=rsRoster("Number")%>',400,200)" ONMOUSEOVER="self.status='Game Winning Goals'; return true" ONMOUSEOUT="self.status=''; return true">GW</A></TH>
				<TH WIDTH=7%><A HREF="javascript: SpawnWindow('<%=strPlayerID%>','GT','<%=rsRoster("LastName")%>','<%=rsRoster("FirstName")%>','<%=rsRoster("Number")%>',400,200)" ONMOUSEOVER="self.status='Game Tying Goals'; return true" ONMOUSEOUT="self.status=''; return true">GT</A></TH>
			</TR>

<%
			rsSeasons.MoveFirst

			Do While Not rsSeasons.EOF
%>
				<TR BGCOLOR="#DCDCDC">
					<TD CLASS="black"><SMALL><%=rsSeasons.Fields(0) %></SMALL></TD>
<%
					strSQL = "SELECT tbl_Players.PlayerID, tbl_Players.FirstName, tbl_Players.LastName, (SELECT COUNT (*) FROM tbl_Seasons INNER JOIN (tbl_Games INNER JOIN tbl_PlayerGames ON tbl_Games.GameID = tbl_PlayerGames.GameID) ON tbl_Seasons.SeasonID = tbl_Games.SeasonID WHERE (((tbl_PlayerGames.PlayerID)=tbl_Players.PlayerID) AND ((tbl_Seasons.SeasonName)='" & rsSeasons.Fields(0) & "'))) AS GP, (SELECT COUNT (*) FROM (tbl_Seasons INNER JOIN tbl_Games ON tbl_Seasons.SeasonID = tbl_Games.SeasonID) INNER JOIN tbl_Goals ON tbl_Games.GameID = tbl_Goals.GameID WHERE tbl_Seasons.SeasonName='" & rsSeasons.Fields(0) & "' AND tbl_Goals.PlayerID=tbl_Players.PlayerID) AS G, (SELECT COUNT (*) FROM (tbl_Seasons INNER JOIN tbl_Games ON tbl_Seasons.SeasonID = tbl_Games.SeasonID) INNER JOIN tbl_Goals ON tbl_Games.GameID = tbl_Goals.GameID WHERE (Assist1PlayerID=tbl_Players.PlayerID OR Assist2PlayerID=tbl_Players.PlayerID) AND tbl_Seasons.SeasonName='" & rsSeasons.Fields(0) & "') AS A, [G]+[A] AS PTS, (SELECT Sum(Length) AS SumOfLength FROM tbl_Seasons INNER JOIN (tbl_Games INNER JOIN tbl_Penalties ON tbl_Games.GameID = tbl_Penalties.GameID) ON tbl_Seasons.SeasonID = tbl_Games.SeasonID GROUP BY tbl_Penalties.PlayerID, tbl_Seasons.SeasonName HAVING tbl_Penalties.PlayerID=tbl_Players.PlayerID AND tbl_Seasons.SeasonName='" & rsSeasons.Fields(0) & "') AS PIM, (SELECT COUNT (*) FROM tbl_Seasons INNER JOIN (tbl_Games INNER JOIN tbl_Goals ON tbl_Games.GameID = tbl_Goals.GameID) ON tbl_Seasons.SeasonID = tbl_Games.SeasonID WHERE tbl_Seasons.SeasonName='" & rsSeasons.Fields(0) & "' AND tbl_Goals.PPG=True AND tbl_Goals.PlayerID=tbl_Players.PlayerID) AS PPG, (SELECT COUNT (*) FROM tbl_Seasons INNER JOIN (tbl_Games INNER JOIN tbl_Goals ON tbl_Games.GameID = tbl_Goals.GameID) ON tbl_Seasons.SeasonID = tbl_Games.SeasonID WHERE (tbl_Seasons.SeasonName='" & rsSeasons.Fields(0) & "' AND tbl_Goals.Assist1PlayerID=tbl_Players.PlayerID AND tbl_Goals.PPG=True) OR (tbl_Seasons.SeasonName='" & rsSeasons.Fields(0) & "' AND tbl_Goals.PPG=True AND tbl_Goals.Assist2PlayerID=tbl_Players.PlayerID)) AS PPA, (SELECT COUNT (*) FROM tbl_Seasons INNER JOIN (tbl_Games INNER JOIN tbl_Goals ON tbl_Games.GameID = tbl_Goals.GameID) ON tbl_Seasons.SeasonID = tbl_Games.SeasonID WHERE tbl_Seasons.SeasonName='" & rsSeasons.Fields(0) & "' AND tbl_Goals.SH=True AND tbl_Goals.PlayerID=tbl_Players.PlayerID) AS SHG, (SELECT COUNT (*) FROM tbl_Seasons INNER JOIN (tbl_Games INNER JOIN tbl_Goals ON tbl_Games.GameID = tbl_Goals.GameID) ON tbl_Seasons.SeasonID = tbl_Games.SeasonID WHERE (tbl_Goals.Assist1PlayerID=tbl_Players.PlayerID AND tbl_Goals.SH=True AND tbl_Seasons.SeasonName='" & rsSeasons.Fields(0) & "') OR (tbl_Goals.SH=True AND tbl_Goals.Assist2PlayerID=tbl_Players.PlayerID AND tbl_Seasons.SeasonName='" & rsSeasons.Fields(0) & "')) AS SHA, (SELECT COUNT (*) FROM tbl_Seasons INNER JOIN (tbl_Games INNER JOIN tbl_Goals ON tbl_Games.GameID = tbl_Goals.GameID) ON tbl_Seasons.SeasonID = tbl_Games.SeasonID WHERE tbl_Goals.GW=True AND tbl_Goals.PlayerID=tbl_Players.PlayerID AND tbl_Seasons.SeasonName='" & rsSeasons.Fields(0) & "') AS GW, (SELECT COUNT (*) FROM tbl_Seasons INNER JOIN (tbl_Games INNER JOIN tbl_Goals ON tbl_Games.GameID = tbl_Goals.GameID) ON tbl_Seasons.SeasonID = tbl_Games.SeasonID WHERE tbl_Goals.GT=True AND tbl_Goals.PlayerID=tbl_Players.PlayerID AND tbl_Seasons.SeasonName='" & rsSeasons.Fields(0) & "') AS GT, tbl_Players.Alternate FROM tbl_Players WHERE PlayerID=" & strPlayerID
					Set rs = dbConn.Execute(strSQL)
%>
					<TD CLASS="black" ALIGN=CENTER><%=rs("GP") %></TD>
					<TD CLASS="black" ALIGN=CENTER><%=rs("G") %></TD>
					<TD CLASS="black" ALIGN=CENTER><%=rs("A") %></TD>
					<TD CLASS="black" ALIGN=CENTER><%=rs("PTS") %></TD>
					<TD CLASS="black" ALIGN=CENTER><%=ChangeNullToZero(rs("PIM")) %></TD>
					<TD CLASS="black" ALIGN=CENTER><%=rs("PPG") %></TD>
					<TD CLASS="black" ALIGN=CENTER><%=rs("PPA") %></TD>
					<TD CLASS="black" ALIGN=CENTER><%=rs("SHG") %></TD>
					<TD CLASS="black" ALIGN=CENTER><%=rs("SHA") %></TD>
					<TD CLASS="black" ALIGN=CENTER><%=rs("GW") %></TD>
					<TD CLASS="black" ALIGN=CENTER><%=rs("GT") %></TD>
				</TR>
<%
				rsSeasons.MoveNext
			Loop
%>
				<TR BGCOLOR="#DCDCDC">
					<TD CLASS="black"><B><FONT COLOR="#3333FF">Career</FONT></B></TD>
<%
					strSQL = "SELECT tbl_Players.PlayerID, tbl_Players.FirstName, tbl_Players.LastName, (SELECT COUNT (*) FROM tbl_PlayerGames WHERE PlayerID=tbl_Players.PlayerID) AS GP, (SELECT COUNT (*) FROM tbl_Goals WHERE PlayerID=tbl_Players.PlayerID) AS G, (SELECT COUNT (*) FROM tbl_Goals WHERE Assist1PlayerID=tbl_Players.PlayerID OR Assist2PlayerID=tbl_Players.PlayerID) AS A, [G]+[A] AS PTS, (SELECT Sum(Length) AS SumOfLength FROM tbl_Penalties GROUP BY PlayerID HAVING PlayerID=tbl_Players.PlayerID) AS PIM, (SELECT COUNT (*) FROM tbl_Goals WHERE PPG=TRUE AND PlayerID=tbl_Players.PlayerID) AS PPG, (SELECT COUNT (*) FROM tbl_Goals WHERE PPG=TRUE AND (Assist1PlayerID=tbl_Players.PlayerID OR Assist2PlayerID=tbl_Players.PlayerID)) AS PPA, (SELECT COUNT (*) FROM tbl_Goals WHERE SH=TRUE AND PlayerID=tbl_Players.PlayerID) AS SHG, (SELECT COUNT (*) FROM tbl_Goals WHERE SH=TRUE AND (Assist1PlayerID=tbl_Players.PlayerID OR Assist2PlayerID=tbl_Players.PlayerID)) AS SHA, (SELECT COUNT (*) FROM tbl_Goals WHERE GW=TRUE AND PlayerID=tbl_Players.PlayerID) AS GW, (SELECT COUNT (*) FROM tbl_Goals WHERE GT=TRUE AND PlayerID=tbl_Players.PlayerID) AS GT, tbl_Players.Alternate FROM tbl_Players WHERE PlayerID=" & strPlayerID
					Set rs = dbConn.Execute(strSQL)
%>
					<TD ALIGN=CENTER><B><FONT COLOR="#3333FF"><%=rs("GP") %></FONT></B></TD>
					<TD ALIGN=CENTER><B><FONT COLOR="#3333FF"><%=rs("G") %></FONT></B></TD>
					<TD ALIGN=CENTER><B><FONT COLOR="#3333FF"><%=rs("A") %></FONT></B></TD>
					<TD ALIGN=CENTER><B><FONT COLOR="#3333FF"><%=rs("PTS") %></FONT></B></TD>
					<TD ALIGN=CENTER><B><FONT COLOR="#3333FF"><%=ChangeNullToZero(rs("PIM")) %></FONT></B></TD>
					<TD ALIGN=CENTER><B><FONT COLOR="#3333FF"><%=rs("PPG") %></FONT></B></TD>
					<TD ALIGN=CENTER><B><FONT COLOR="#3333FF"><%=rs("PPA") %></FONT></B></TD>
					<TD ALIGN=CENTER><B><FONT COLOR="#3333FF"><%=rs("SHG") %></FONT></B></TD>
					<TD ALIGN=CENTER><B><FONT COLOR="#3333FF"><%=rs("SHA") %></FONT></B></TD>
					<TD ALIGN=CENTER><B><FONT COLOR="#3333FF"><%=rs("GW") %></FONT></B></TD>
					<TD ALIGN=CENTER><B><FONT COLOR="#3333FF"><%=rs("GT") %></FONT></B></TD>
				</TR>
			</TABLE>

<%

			rs.Close
			Set rs = nothing

		End If

	End If
%>
	<BR><BR>
	</TD>
</TR>
</TABLE>

<%

	Set rsRoster = nothing

	dbConn.Close
	Set dbConn = nothing
%>
</BODY>
</HTML>