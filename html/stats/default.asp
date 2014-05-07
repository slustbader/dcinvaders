<%
	Option Explicit
	'On Error Resume Next

	Dim strSortBy, fmSeasonID

	fmSeasonID = Request.QueryString("SeasonID")
	strSortBy = Request.QueryString("sortby")

	Function ChangeNullToZero(strText)
		If IsNull(strText) Then
			strText = 0
		End If
		ChangeNullToZero = strText
	End Function

%>
<!--#include virtual="/include/adovbs.inc"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<script src="/javascripts/prototype.js" type="text/javascript"></script>
<script src="/javascripts/scriptaculous.js" type="text/javascript"></script>
<SCRIPT LANGUAGE="JavaScript">
<!--

	function ShowLegend(Num){
		eval ("window.open('" + Num + ".html','legend','width=250,height=260,scrollbars=no,resizable=no')");
	}

	var req;
	var search;
	
	function autoComplete(keyCode) {
	    if (keyCode == 8 && document.getElementById("playerSearch").value != "") {
	        setTimeout("doAutoComplete()", 3000);
	        return;
	    } else if ((keyCode >= 97 && keyCode <= 122) || (keyCode >= 65 && keyCode <= 90)) {
	        doAutoComplete();
	        return;
	    }
	}
	
	function doAutoComplete() {
        search = document.getElementById("playerSearch").value;
	    if (search != "") {
    	    var url = "findPlayer.asp?lastName=" + escape(search);
    	    req = newXMLHttpRequest();
            req.open("GET", url, true);
            req.onreadystatechange = processRequest;
            req.send(null);
        }
	}

function newXMLHttpRequest() {
  var xmlreq = false;
  if (window.XMLHttpRequest) {
    // Create XMLHttpRequest object in non-Microsoft browsers
    xmlreq = new XMLHttpRequest();
  } else if (window.ActiveXObject) {
    // Create XMLHttpRequest via MS ActiveX
    try {
      // Try to create XMLHttpRequest in later versions
      // of Internet Explorer
      xmlreq = new ActiveXObject("Msxml2.XMLHTTP");
    } catch (e1) {
      // Failed to create required ActiveXObject
      try {
        // Try version supported by older versions
        // of Internet Explorer
        xmlreq = new ActiveXObject("Microsoft.XMLHTTP");
      } catch (e2) {
        // Unable to create an XMLHttpRequest with ActiveX
      }
    }
  }
  return xmlreq;
}

	function processRequest() {
        if (req.readyState == 4 && req.status == 200) {
            var name = req.responseXML.getElementsByTagName("lastName")[0].childNodes[0].nodeValue;
            setTimeout("updateText('" + name + "')", 0);
        }
    }
    
    function updateText(text) {
        var input = document.getElementById("playerSearch");
        input.value = text;
        if (input.setSelectionRange) {
            // mozilla
            input.setSelectionRange(search.length, input.value.length); 
        } else {
            // IE
            var range = input.createTextRange();
            range.moveStart("character", search.length);
            range.moveEnd("character", input.value.length);      
            range.select();        
        }
    }
//-->
</SCRIPT>
<TITLE>DC Invaders Stats</TITLE>
<LINK REL="stylesheet" HREF="../styles/hockey.css" TYPE="text/css">
</HEAD>
<BODY BACKGROUND="../images/stars.gif">
<TABLE WIDTH=660 CELLPADDING=0 CELLSPACING=0 BORDER=0>
<TR>
	<TD><CENTER><IMG SRC="../images/invader.gif" BORDER=0></CENTER>
<%
		Dim dbConn, strSQL, rsSeason, rsTeamStats, rsPlayers, rsGoalies
		Dim strSeason, strSeasonID, strGP, strG, strA, strPTS, strPIM, strPPG, strPPA, strSHG, strSHA, strGW, strGT
		Dim strPlayerID
		Dim strTemptable

		Set dbConn = Server.CreateObject("ADODB.Connection")
		'dbConn.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("/invaders/database/hockey.mdb")
		dbConn.Open "DSN=hockey"

		If fmSeasonID = "" Then
			strSQL = "SELECT SeasonName, SeasonID FROM tbl_Seasons WHERE StartDate < Now() AND EndDate > Now()"
			Set rsSeason = dbConn.Execute(strSQL)
		Else
			strSQL = "SELECT SeasonName, SeasonID FROM tbl_Seasons WHERE SeasonID=" & fmSeasonID
			Set rsSeason = dbConn.Execute(strSQL)
		End If

		If Not (rsSeason.BOF AND rsSeason.EOF) Then
			strSeason = rsSeason("SeasonName")
			strSeasonID = rsSeason("SeasonID")

			strSQL = "SELECT tbl_Teams.TeamID, tbl_Teams.TeamName, (SELECT Count(*) FROM tbl_Games WHERE (((tbl_Games.HomeTeamID)=1) AND ((tbl_Games.SeasonID)=" & strSeasonID & ") AND ((tbl_Games.Scored)=True)) OR (((tbl_Games.AwayTeamID)=1) AND ((tbl_Games.SeasonID)=" & strSeasonID & ") AND ((tbl_Games.Scored)=True))) AS GP, (SELECT Count(*) FROM tbl_Games WHERE (((IIf(([HomeTeamID]=1 And ([FirstPeriodHomeGoals]+[SecondPeriodHomeGoals]+[ThirdPeriodHomeGoals]+[OvertimeHomeGoals]>[FirstPeriodAwayGoals]+[SecondPeriodAwayGoals]+[ThirdPeriodAwayGoals]+[OvertimeAwayGoals]) And [SeasonID]=" & strSeasonID & ") Or ([AwayTeamID]=1 And ([FirstPeriodAwayGoals]+[SecondPeriodAwayGoals]+[ThirdPeriodAwayGoals]+[OvertimeAwayGoals]>[FirstPeriodHomeGoals]+[SecondPeriodHomeGoals]+[ThirdPeriodHomeGoals]+[OvertimeHomeGoals]) And [SeasonID]=" & strSeasonID & "),1,0))=1))) AS W, (SELECT Count(*) FROM tbl_Games WHERE (((IIf(([HomeTeamID]=1 And ([FirstPeriodHomeGoals]+[SecondPeriodHomeGoals]+[ThirdPeriodHomeGoals]+[OvertimeHomeGoals]<[FirstPeriodAwayGoals]+[SecondPeriodAwayGoals]+[ThirdPeriodAwayGoals]+[OvertimeAwayGoals]) And [SeasonID]=" & strSeasonID & ") Or ([AwayTeamID]=1 And ([FirstPeriodAwayGoals]+[SecondPeriodAwayGoals]+[ThirdPeriodAwayGoals]+[OvertimeAwayGoals]<[FirstPeriodHomeGoals]+[SecondPeriodHomeGoals]+[ThirdPeriodHomeGoals]+[OvertimeHomeGoals]) And [SeasonID]=" & strSeasonID & "),1,0))=1))) AS L, (SELECT Count(*) FROM tbl_Games WHERE (((IIf((([HomeTeamID]=1 Or [AwayTeamID]=1) And ([FirstPeriodHomeGoals]+[SecondPeriodHomeGoals]+[ThirdPeriodHomeGoals]+[OvertimeHomeGoals]=[FirstPeriodAwayGoals]+[SecondPeriodAwayGoals]+[ThirdPeriodAwayGoals]+[OvertimeAwayGoals]) And [SeasonID]=" & strSeasonID & " And [Scored]=TRUE),1,0))=1))) AS T, (SELECT Sum(IIf([HomeTeamID]=1,[FirstPeriodHomeShots]+[SecondPeriodHomeShots]+[ThirdPeriodHomeShots]+[OvertimeHomeShots],[FirstPeriodAwayShots]+[SecondPeriodAwayShots]+[ThirdPeriodAwayShots]+[OvertimeAwayShots])) AS S FROM tbl_Games HAVING (((tbl_Games.HomeTeamID)=1) AND ((tbl_Games.SeasonID)=" & strSeasonID & ")) OR (((tbl_Games.SeasonID)=" & strSeasonID & ") AND ((tbl_Games.AwayTeamID)=1))) AS S, (SELECT Sum(IIf([HomeTeamID]=1,[FirstPeriodAwayShots]+[SecondPeriodAwayShots]+[ThirdPeriodAwayShots]+[OvertimeAwayShots],[FirstPeriodHomeShots]+[SecondPeriodHomeShots]+[ThirdPeriodHomeShots]+[OvertimeHomeShots])) AS SA FROM tbl_Games HAVING (((tbl_Games.HomeTeamID)=1) AND ((tbl_Games.SeasonID)=" & strSeasonID & ")) OR (((tbl_Games.SeasonID)=" & strSeasonID & ") AND ((tbl_Games.AwayTeamID)=1))) AS SA, (SELECT Sum(IIf([HomeTeamID]=1,[FirstPeriodHomeGoals]+[SecondPeriodHomeGoals]+[ThirdPeriodHomeGoals]+[OvertimeHomeGoals],[FirstPeriodAwayGoals]+[SecondPeriodAwayGoals]+[ThirdPeriodAwayGoals]+[OvertimeAwayGoals])) AS G FROM tbl_Games HAVING ((([HomeTeamID]=1 Or [AwayTeamID]=1) And [SeasonID]=" & strSeasonID & "))) AS G, (SELECT Sum(IIf([HomeTeamID]=1,[FirstPeriodAwayGoals]+[SecondPeriodAwayGoals]+[ThirdPeriodAwayGoals]+[OvertimeAwayGoals],[FirstPeriodHomeGoals]+[SecondPeriodHomeGoals]+[ThirdPeriodHomeGoals]+[OvertimeHomeGoals])) AS G FROM tbl_Games HAVING (((tbl_Games.HomeTeamID)=1) AND ((tbl_Games.SeasonID)=" & strSeasonID & ")) OR (((tbl_Games.SeasonID)=" & strSeasonID & ") AND ((tbl_Games.AwayTeamID)=1))) AS GA, (SELECT Sum(tbl_Penalties.Length) AS SumOfLength FROM tbl_Players INNER JOIN (tbl_Games INNER JOIN tbl_Penalties ON tbl_Games.GameID = tbl_Penalties.GameID) ON tbl_Players.PlayerID = tbl_Penalties.PlayerID GROUP BY tbl_Players.TeamID, tbl_Games.SeasonID HAVING (((tbl_Players.TeamID)=1) AND ((tbl_Games.SeasonID)=" & strSeasonID & "))) AS PIM, [G]/[GP] AS GPG, [S]/[GP] AS SPG, [GA]/[GP] AS GAPG, [SA]/[GP] AS SAPG FROM tbl_Teams WHERE (((tbl_Teams.TeamID)=1))"

			Set rsTeamStats = dbConn.Execute(strSQL)
		End If

' This code breaks the stats page between seasons, since it references rsTeamStats("GP") = 0 even though the rsSeason clause is true  
'		If (rsSeason.BOF AND rsSeason.EOF) OR rsTeamStats("GP") = 0  Then
'			fmSeasonId = rsSeason("SeasonID") - 1
'			strSQL = "SELECT SeasonName, SeasonID FROM tbl_Seasons WHERE SeasonID=" & fmSeasonID
'			Set rsSeason = dbConn.Execute(strSQL)		
'			If Not (rsSeason.BOF AND rsSeason.EOF) Then
'				strSeason = rsSeason("SeasonName")
'				strSeasonID = rsSeason("SeasonID")
'				strSQL = "SELECT tbl_Teams.TeamID, tbl_Teams.TeamName, (SELECT Count(*) FROM tbl_Games WHERE (((tbl_Games.HomeTeamID)=1) AND ((tbl_Games.SeasonID)=" & strSeasonID & ") AND ((tbl_Games.Scored)=True)) OR (((tbl_Games.AwayTeamID)=1) AND ((tbl_Games.SeasonID)=" & strSeasonID & ") AND ((tbl_Games.Scored)=True))) AS GP, (SELECT Count(*) FROM tbl_Games WHERE (((IIf(([HomeTeamID]=1 And ([FirstPeriodHomeGoals]+[SecondPeriodHomeGoals]+[ThirdPeriodHomeGoals]+[OvertimeHomeGoals]>[FirstPeriodAwayGoals]+[SecondPeriodAwayGoals]+[ThirdPeriodAwayGoals]+[OvertimeAwayGoals]) And [SeasonID]=" & strSeasonID & ") Or ([AwayTeamID]=1 And ([FirstPeriodAwayGoals]+[SecondPeriodAwayGoals]+[ThirdPeriodAwayGoals]+[OvertimeAwayGoals]>[FirstPeriodHomeGoals]+[SecondPeriodHomeGoals]+[ThirdPeriodHomeGoals]+[OvertimeHomeGoals]) And [SeasonID]=" & strSeasonID & "),1,0))=1))) AS W, (SELECT Count(*) FROM tbl_Games WHERE (((IIf(([HomeTeamID]=1 And ([FirstPeriodHomeGoals]+[SecondPeriodHomeGoals]+[ThirdPeriodHomeGoals]+[OvertimeHomeGoals]<[FirstPeriodAwayGoals]+[SecondPeriodAwayGoals]+[ThirdPeriodAwayGoals]+[OvertimeAwayGoals]) And [SeasonID]=" & strSeasonID & ") Or ([AwayTeamID]=1 And ([FirstPeriodAwayGoals]+[SecondPeriodAwayGoals]+[ThirdPeriodAwayGoals]+[OvertimeAwayGoals]<[FirstPeriodHomeGoals]+[SecondPeriodHomeGoals]+[ThirdPeriodHomeGoals]+[OvertimeHomeGoals]) And [SeasonID]=" & strSeasonID & "),1,0))=1))) AS L, (SELECT Count(*) FROM tbl_Games WHERE (((IIf((([HomeTeamID]=1 Or [AwayTeamID]=1) And ([FirstPeriodHomeGoals]+[SecondPeriodHomeGoals]+[ThirdPeriodHomeGoals]+[OvertimeHomeGoals]=[FirstPeriodAwayGoals]+[SecondPeriodAwayGoals]+[ThirdPeriodAwayGoals]+[OvertimeAwayGoals]) And [SeasonID]=" & strSeasonID & " And [Scored]=TRUE),1,0))=1))) AS T, (SELECT Sum(IIf([HomeTeamID]=1,[FirstPeriodHomeShots]+[SecondPeriodHomeShots]+[ThirdPeriodHomeShots]+[OvertimeHomeShots],[FirstPeriodAwayShots]+[SecondPeriodAwayShots]+[ThirdPeriodAwayShots]+[OvertimeAwayShots])) AS S FROM tbl_Games HAVING (((tbl_Games.HomeTeamID)=1) AND ((tbl_Games.SeasonID)=" & strSeasonID & ")) OR (((tbl_Games.SeasonID)=" & strSeasonID & ") AND ((tbl_Games.AwayTeamID)=1))) AS S, (SELECT Sum(IIf([HomeTeamID]=1,[FirstPeriodAwayShots]+[SecondPeriodAwayShots]+[ThirdPeriodAwayShots]+[OvertimeAwayShots],[FirstPeriodHomeShots]+[SecondPeriodHomeShots]+[ThirdPeriodHomeShots]+[OvertimeHomeShots])) AS SA FROM tbl_Games HAVING (((tbl_Games.HomeTeamID)=1) AND ((tbl_Games.SeasonID)=" & strSeasonID & ")) OR (((tbl_Games.SeasonID)=" & strSeasonID & ") AND ((tbl_Games.AwayTeamID)=1))) AS SA, (SELECT Sum(IIf([HomeTeamID]=1,[FirstPeriodHomeGoals]+[SecondPeriodHomeGoals]+[ThirdPeriodHomeGoals]+[OvertimeHomeGoals],[FirstPeriodAwayGoals]+[SecondPeriodAwayGoals]+[ThirdPeriodAwayGoals]+[OvertimeAwayGoals])) AS G FROM tbl_Games HAVING ((([HomeTeamID]=1 Or [AwayTeamID]=1) And [SeasonID]=" & strSeasonID & "))) AS G, (SELECT Sum(IIf([HomeTeamID]=1,[FirstPeriodAwayGoals]+[SecondPeriodAwayGoals]+[ThirdPeriodAwayGoals]+[OvertimeAwayGoals],[FirstPeriodHomeGoals]+[SecondPeriodHomeGoals]+[ThirdPeriodHomeGoals]+[OvertimeHomeGoals])) AS G FROM tbl_Games HAVING (((tbl_Games.HomeTeamID)=1) AND ((tbl_Games.SeasonID)=" & strSeasonID & ")) OR (((tbl_Games.SeasonID)=" & strSeasonID & ") AND ((tbl_Games.AwayTeamID)=1))) AS GA, (SELECT Sum(tbl_Penalties.Length) AS SumOfLength FROM tbl_Players INNER JOIN (tbl_Games INNER JOIN tbl_Penalties ON tbl_Games.GameID = tbl_Penalties.GameID) ON tbl_Players.PlayerID = tbl_Penalties.PlayerID GROUP BY tbl_Players.TeamID, tbl_Games.SeasonID HAVING (((tbl_Players.TeamID)=1) AND ((tbl_Games.SeasonID)=" & strSeasonID & "))) AS PIM, [G]/[GP] AS GPG, [S]/[GP] AS SPG, [GA]/[GP] AS GAPG, [SA]/[GP] AS SAPG FROM tbl_Teams WHERE (((tbl_Teams.TeamID)=1))"
'				Set rsTeamStats = dbConn.Execute(strSQL)
'			End If
'		End If

        If IsEmpty(rsTeamStats) Or (rsSeason.BOF AND rsSeason.EOF) Then
			Response.write "<H3>The Invaders are currently relaxing in the offseason.</H3>"
		ElseIf rsTeamStats("GP") = 0 Or (rsSeason.BOF AND rsSeason.EOF) Then
		    Response.write "<H3>The Invaders are currently relaxing in the offseason.</H3>"
		Else
%>
			<CENTER><H2><FONT COLOR="#FFFF66"><% =strSeason %> Stats</FONT></H2></CENTER>
<div id="stats">
<div id="teamStats">
			<CENTER><H2>Team Stats</H2></CENTER>
			<HR>
			<BR>

			<CENTER><TABLE WIDTH=650 CELLPADDING=2 CELLSPACING=0 BORDER=1 BORDERCOLOR="black" rules="all">
			<tbody>
			<TR>
				<TD COLSPAN=13 ALIGN=RIGHT><A HREF="javascript: ShowLegend('teamlegend')">Legend</A></TD>
			</TR>

			<TR BGCOLOR="#08479E">
				<TD ALIGN=CENTER WIDTH=50><B>GP</B></TD>
				<TD ALIGN=CENTER WIDTH=50><B>W</B></TD>
				<TD ALIGN=CENTER WIDTH=50><B>L</B></TD>
				<TD ALIGN=CENTER WIDTH=50><B>T</B></TD>
				<TD ALIGN=CENTER WIDTH=50><B>S</B></TD>
				<TD ALIGN=CENTER WIDTH=50><B>SA</B></TD>
				<TD ALIGN=CENTER WIDTH=50><B>G</B></TD>
				<TD ALIGN=CENTER WIDTH=50><B>GA</B></TD>
				<TD ALIGN=CENTER WIDTH=50><B>SPG</B></TD>
				<TD ALIGN=CENTER WIDTH=50><B>SAPG</B></TD>
				<TD ALIGN=CENTER WIDTH=50><B>GPG</B></TD>
				<TD ALIGN=CENTER WIDTH=50><B>GAPG</B></TD>
				<TD ALIGN=CENTER WIDTH=50><B>PIM</B></TD>
			</TR>

			<TR BGCOLOR="#DCDCDC">
				<TD CLASS="black" ALIGN=CENTER><%=rsTeamStats("GP") %></TD>
				<TD CLASS="black" ALIGN=CENTER><%=rsTeamStats("W") %></TD>
				<TD CLASS="black" ALIGN=CENTER><%=rsTeamStats("L") %></TD>
				<TD CLASS="black" ALIGN=CENTER><%=rsTeamStats("T") %></TD>
				<TD CLASS="black" ALIGN=CENTER><%=rsTeamStats("S") %></TD>
				<TD CLASS="black" ALIGN=CENTER><%=rsTeamStats("SA") %></TD>
				<TD CLASS="black" ALIGN=CENTER><%=rsTeamStats("G") %></TD>
				<TD CLASS="black" ALIGN=CENTER><%=rsTeamStats("GA") %></TD>
				<TD CLASS="black" ALIGN=CENTER><%=FormatNumber(rsTeamStats("SPG"), 1) %></TD>
				<TD CLASS="black" ALIGN=CENTER><%=FormatNumber(rsTeamStats("SAPG"), 1) %></TD>
				<TD CLASS="black" ALIGN=CENTER><%=FormatNumber(rsTeamStats("GPG"), 1) %></TD>
				<TD CLASS="black" ALIGN=CENTER><%=FormatNumber(rsTeamStats("GAPG"), 1) %></TD>
				<TD CLASS="black" ALIGN=CENTER><%=rsTeamStats("PIM") %></TD>
			</TR>
			</tbody>
			</TABLE></CENTER><BR><BR>
</div>

<%
			rsTeamStats.close
			Set rsTeamStats = nothing

			Randomize
			strTempTable = Int(10000 * Rnd)

			strSQL = "SELECT tbl_Players.PlayerID, tbl_Players.FirstName, tbl_Players.LastName, " &_
				    "(SELECT COUNT (*) FROM tbl_Seasons INNER JOIN (tbl_Games INNER JOIN tbl_PlayerGames ON tbl_Games.GameID = tbl_PlayerGames.GameID) ON tbl_Seasons.SeasonID = tbl_Games.SeasonID WHERE (((tbl_PlayerGames.PlayerID)=tbl_Players.PlayerID) AND ((tbl_Seasons.SeasonName)='" & strSeason & "'))) AS GP, " &_
					"(SELECT COUNT (*) FROM (tbl_Seasons INNER JOIN tbl_Games ON tbl_Seasons.SeasonID = tbl_Games.SeasonID) INNER JOIN tbl_Goals ON tbl_Games.GameID = tbl_Goals.GameID WHERE tbl_Seasons.SeasonName='" & strSeason & "' AND tbl_Goals.PlayerID=tbl_Players.PlayerID) AS G, " &_
					"(SELECT COUNT (*) FROM (tbl_Seasons INNER JOIN tbl_Games ON tbl_Seasons.SeasonID = tbl_Games.SeasonID) INNER JOIN tbl_Goals ON tbl_Games.GameID = tbl_Goals.GameID WHERE (Assist1PlayerID=tbl_Players.PlayerID OR Assist2PlayerID=tbl_Players.PlayerID) AND tbl_Seasons.SeasonName='" & strSeason & "') AS A, " &_
					"[G]+[A] AS PTS, " &_
					"(SELECT Sum(Length) AS SumOfLength FROM tbl_Seasons INNER JOIN (tbl_Games INNER JOIN tbl_Penalties ON tbl_Games.GameID = tbl_Penalties.GameID) ON tbl_Seasons.SeasonID = tbl_Games.SeasonID GROUP BY tbl_Penalties.PlayerID, tbl_Seasons.SeasonName HAVING tbl_Penalties.PlayerID=tbl_Players.PlayerID AND tbl_Seasons.SeasonName='" & strSeason & "') AS PIM, " &_
					"(SELECT COUNT (*) FROM tbl_Seasons INNER JOIN (tbl_Games INNER JOIN tbl_Goals ON tbl_Games.GameID = tbl_Goals.GameID) ON tbl_Seasons.SeasonID = tbl_Games.SeasonID WHERE tbl_Seasons.SeasonName='" & strSeason & "' AND tbl_Goals.PPG=True AND tbl_Goals.PlayerID=tbl_Players.PlayerID) AS PPG, " &_
					"(SELECT COUNT (*) FROM tbl_Seasons INNER JOIN (tbl_Games INNER JOIN tbl_Goals ON tbl_Games.GameID = tbl_Goals.GameID) ON tbl_Seasons.SeasonID = tbl_Games.SeasonID WHERE (tbl_Seasons.SeasonName='" & strSeason & "' AND tbl_Goals.Assist1PlayerID=tbl_Players.PlayerID AND tbl_Goals.PPG=True) OR (tbl_Seasons.SeasonName='" & strSeason & "' AND tbl_Goals.PPG=True AND tbl_Goals.Assist2PlayerID=tbl_Players.PlayerID)) AS PPA, " &_
					"(SELECT COUNT (*) FROM tbl_Seasons INNER JOIN (tbl_Games INNER JOIN tbl_Goals ON tbl_Games.GameID = tbl_Goals.GameID) ON tbl_Seasons.SeasonID = tbl_Games.SeasonID WHERE tbl_Seasons.SeasonName='" & strSeason & "' AND tbl_Goals.SH=True AND tbl_Goals.PlayerID=tbl_Players.PlayerID) AS SHG, " &_
					"(SELECT COUNT (*) FROM tbl_Seasons INNER JOIN (tbl_Games INNER JOIN tbl_Goals ON tbl_Games.GameID = tbl_Goals.GameID) ON tbl_Seasons.SeasonID = tbl_Games.SeasonID WHERE (tbl_Goals.Assist1PlayerID=tbl_Players.PlayerID AND tbl_Goals.SH=True AND tbl_Seasons.SeasonName='" & strSeason & "') OR (tbl_Goals.SH=True AND tbl_Goals.Assist2PlayerID=tbl_Players.PlayerID AND tbl_Seasons.SeasonName='" & strSeason & "')) AS SHA, " &_
					"(SELECT COUNT (*) FROM tbl_Seasons INNER JOIN (tbl_Games INNER JOIN tbl_Goals ON tbl_Games.GameID = tbl_Goals.GameID) ON tbl_Seasons.SeasonID = tbl_Games.SeasonID WHERE tbl_Goals.GW=True AND tbl_Goals.PlayerID=tbl_Players.PlayerID AND tbl_Seasons.SeasonName='" & strSeason & "') AS GW, " &_
					"(SELECT COUNT (*) FROM tbl_Seasons INNER JOIN (tbl_Games INNER JOIN tbl_Goals ON tbl_Games.GameID = tbl_Goals.GameID) ON tbl_Seasons.SeasonID = tbl_Games.SeasonID WHERE tbl_Goals.GT=True AND tbl_Goals.PlayerID=tbl_Players.PlayerID AND tbl_Seasons.SeasonName='" & strSeason & "') AS GT, " &_
					"tbl_Players.Alternate INTO " & strTempTable & " FROM tbl_Players " &_
					"WHERE (((tbl_Players.PlayerID) In (SELECT tbl_PlayerGames.PlayerID FROM tbl_Games INNER JOIN tbl_PlayerGames ON tbl_Games.GameID = tbl_PlayerGames.GameID WHERE tbl_Games.SeasonID=" & strSeasonID & ")) AND ((tbl_Players.TeamID)=1) AND ((tbl_Players.Position)<>5))"
'		The following where clause seemed to produce weird stats for season 14
'					"WHERE (((tbl_Players.PlayerID) In (SELECT tbl_PlayerGames.PlayerID FROM tbl_Games INNER JOIN tbl_PlayerGames ON tbl_Games.GameID = tbl_PlayerGames.GameID WHERE tbl_Games.GameDate >= (SELECT StartDate FROM tbl_Seasons WHERE SeasonID=" & strSeasonID & ") AND tbl_Games.GameDate <= (SELECT EndDate FROM tbl_Seasons WHERE SeasonID=" & strSeasonID & "))) AND ((tbl_Players.TeamID)=1) AND ((tbl_Players.Position)<>5))"

			'response.write "<BR>SQL: " & strSQL & "<BR>"
			dbConn.Execute strSQL

			Set rsPlayers = Server.CreateObject("ADODB.Recordset")
			strSQL = "SELECT * FROM " & strTempTable & " WHERE Alternate=FALSE"

			If strSortBy <> ""  AND strSortBy <> "LastName" Then
				strSQL = strSQL & " ORDER BY " & strSortBy & " DESC, LastName"
			Else
				strSQL = strSQL & " ORDER BY LastName"
			End If

			rsPlayers.Open strSQL, dbConn, adOpenStatic

			rsPlayers.MoveFirst
%>
<div id="playerStats">
			<CENTER><H2>Player Stats</H2></CENTER>
			<HR>
			<H3>Skaters</H3>
			<EM>Click on any of the column headings to sort the players statistics, or click on the player's name for detailed player information.</EM>
			<CENTER>
			<TABLE WIDTH=600 CELLPADDING=2 CELLSPACING=0 BORDER=1 BORDERCOLOR="black" rules="all">
			<tbody id="playerStatsTable">
			<TR>
						<TD COLSPAN=12 ALIGN=RIGHT><A HREF="javascript: ShowLegend('skaterlegend')">Legend</A></TD>
			</TR>

			<TR BGCOLOR="#08479E">
				<TD WIDTH=160><A HREF="default.asp?sortby=LastName&SeasonID=<% =fmSeasonID%>"  ONMOUSEOVER="self.status='Player'; return true" ONMOUSEOUT="self.status=''; return true"><B>Player</B></A></TD>
				<TH WIDTH=40><A HREF="default.asp?sortby=GP&SeasonID=<% =fmSeasonID%>" ONMOUSEOVER="self.status='Games Played'; return true" ONMOUSEOUT="self.status=''; return true">GP</A></TH>
				<TH WIDTH=40><A HREF="default.asp?sortby=G&SeasonID=<% =fmSeasonID%>" ONMOUSEOVER="self.status='Goals'; return true" ONMOUSEOUT="self.status=''; return true">G</A></TH>
				<TH WIDTH=40><A HREF="default.asp?sortby=A&SeasonID=<% =fmSeasonID%>" ONMOUSEOVER="self.status='Assists'; return true" ONMOUSEOUT="self.status=''; return true">A</A></TH>
				<TH WIDTH=40><A HREF="default.asp?sortby=PTS&SeasonID=<% =fmSeasonID%>" ONMOUSEOVER="self.status='Points'; return true" ONMOUSEOUT="self.status=''; return true">PTS</A></TH>
				<TH WIDTH=40><A HREF="default.asp?sortby=PIM&SeasonID=<% =fmSeasonID%>" ONMOUSEOVER="self.status='Penalty Minutes'; return true" ONMOUSEOUT="self.status=''; return true">PIM</A></TH>
				<TH WIDTH=40><A HREF="default.asp?sortby=PPG&SeasonID=<% =fmSeasonID%>" ONMOUSEOVER="self.status='Power Play Goals'; return true" ONMOUSEOUT="self.status=''; return true">PPG</A></TH>
				<TH WIDTH=40><A HREF="default.asp?sortby=PPA&SeasonID=<% =fmSeasonID%>" ONMOUSEOVER="self.status='Power Play Assists'; return true" ONMOUSEOUT="self.status=''; return true">PPA</A></TH>
				<TH WIDTH=40><A HREF="default.asp?sortby=SHG&SeasonID=<% =fmSeasonID%>" ONMOUSEOVER="self.status='Short Handed Goals'; return true" ONMOUSEOUT="self.status=''; return true">SHG</A></TH>
				<TH WIDTH=40><A HREF="default.asp?sortby=SHA&SeasonID=<% =fmSeasonID%>" ONMOUSEOVER="self.status='Short Handed Assists'; return true" ONMOUSEOUT="self.status=''; return true">SHA</A></TH>
				<TH WIDTH=40><A HREF="default.asp?sortby=GW&SeasonID=<% =fmSeasonID%>" ONMOUSEOVER="self.status='Game Winning Goals'; return true" ONMOUSEOUT="self.status=''; return true">GW</A></TH>
				<TH WIDTH=40><A HREF="default.asp?sortby=GT&SeasonID=<% =fmSeasonID%>" ONMOUSEOVER="self.status='Game Tying Goals'; return true" ONMOUSEOUT="self.status=''; return true">GT</A></TH>
			</TR>
<%
			Do While Not rsPlayers.EOF

				strPlayerID = rsPlayers("PlayerID")
				strGP = rsPlayers("GP")
				strG = rsPlayers("G")
				strA = rsPlayers("A")
				strPTS = rsPlayers("PTS")
				If IsNull(rsPlayers("PIM") ) or rsPlayers("PIM") = "" Then
					strPIM = 0
				Else
					strPIM = rsPlayers("PIM")
				End If
				strPPG = rsPlayers("PPG")
				strPPA = rsPlayers("PPA")
				strSHG = rsPlayers("SHG")
				strSHA = rsPlayers("SHA")
				strGW = rsPlayers("GW")
				strGT = rsPlayers("GT")

				response.write "<TR BGCOLOR=""#DCDCDC"" class=""data"">"
				response.write "	<TD><A HREF=""../roster/default.asp?ID=" & strPlayerID & """>" & rsPlayers("FirstName") & " " & rsPlayers("LastName") & "</A></TD>"
%>
					<TD CLASS="black" ALIGN=CENTER><%=ChangeNullToZero(strGP) %></TD>
					<TD CLASS="black" ALIGN=CENTER><%=ChangeNullToZero(strG) %></TD>
					<TD CLASS="black" ALIGN=CENTER><%=ChangeNullToZero(strA) %></TD>
					<TD CLASS="black" ALIGN=CENTER><%=ChangeNullToZero(strPTS) %></TD>
					<TD CLASS="black" ALIGN=CENTER><%=ChangeNullToZero(strPIM) %></TD>
					<TD CLASS="black" ALIGN=CENTER><%=ChangeNullToZero(strPPG) %></TD>
					<TD CLASS="black" ALIGN=CENTER><%=ChangeNullToZero(strPPA) %></TD>
					<TD CLASS="black" ALIGN=CENTER><%=ChangeNullToZero(strSHG) %></TD>
					<TD CLASS="black" ALIGN=CENTER><%=ChangeNullToZero(strSHA) %></TD>
					<TD CLASS="black" ALIGN=CENTER><%=ChangeNullToZero(strGW) %></TD>
					<TD CLASS="black" ALIGN=CENTER><%=ChangeNullToZero(strGT) %></TD>
				</TR>
<%
				rsPlayers.MoveNext
			Loop

			rsPlayers.Close
			Set rsPlayers = nothing

			strSQL = "DROP TABLE " & strTempTable
			dbConn.Execute strSQL
%>
            </tbody>
			</TABLE>
			</CENTER>

<%

				Dim strW, strL, strT, strSA, strGA, strGAA, strSPCT, strSO
				Dim strGoalieTempTable

				Randomize
				strGoalieTempTable = Int(10000 * Rnd)

				strSQL = "SELECT tbl_Players.PlayerID, tbl_Players.FirstName, tbl_Players.LastName, " &_
						"(SELECT COUNT(*) FROM tbl_Seasons INNER JOIN (tbl_Games INNER JOIN tbl_PlayerGames ON tbl_Games.GameID = tbl_PlayerGames.GameID) ON tbl_Seasons.SeasonID = tbl_Games.SeasonID WHERE tbl_PlayerGames.PlayerID=tbl_Players.PlayerID AND tbl_Seasons.SeasonName='" & strSeason & "') AS GP, " &_
						"(SELECT COUNT(*) FROM tbl_Seasons INNER JOIN (tbl_Games INNER JOIN tbl_PlayerGames ON tbl_Games.GameID = tbl_PlayerGames.GameID) ON tbl_Seasons.SeasonID = tbl_Games.SeasonID WHERE (((IIf(([HomeTeamID]=tbl_Players.TeamID And ([FirstPeriodHomeGoals]+[SecondPeriodHomeGoals]+[ThirdPeriodHomeGoals]+[OvertimeHomeGoals])>([FirstPeriodAwayGoals]+[SecondPeriodAwayGoals]+[ThirdPeriodAwayGoals]+[OvertimeAwayGoals])) Or ([AwayTeamID]=tbl_Players.TeamID And ([FirstPeriodAwayGoals]+[SecondPeriodAwayGoals]+[ThirdPeriodAwayGoals]+[OvertimeAwayGoals])>([FirstPeriodHomeGoals]+[SecondPeriodHomeGoals]+[ThirdPeriodHomeGoals]+[OvertimeHomeGoals])),'True','False'))='True') AND ((tbl_PlayerGames.PlayerID)=tbl_Players.PlayerID) AND ((tbl_Seasons.SeasonName)='" & strSeason & "'))) AS W, " &_
						"(SELECT COUNT(*) FROM tbl_Seasons INNER JOIN (tbl_Games INNER JOIN tbl_PlayerGames ON tbl_Games.GameID = tbl_PlayerGames.GameID) ON tbl_Seasons.SeasonID = tbl_Games.SeasonID WHERE (((IIf(([HomeTeamID]=tbl_Players.TeamID And ([FirstPeriodHomeGoals]+[SecondPeriodHomeGoals]+[ThirdPeriodHomeGoals]+[OvertimeHomeGoals])<([FirstPeriodAwayGoals]+[SecondPeriodAwayGoals]+[ThirdPeriodAwayGoals]+[OvertimeAwayGoals])) Or ([AwayTeamID]=tbl_Players.TeamID And ([FirstPeriodAwayGoals]+[SecondPeriodAwayGoals]+[ThirdPeriodAwayGoals]+[OvertimeAwayGoals])<([FirstPeriodHomeGoals]+[SecondPeriodHomeGoals]+[ThirdPeriodHomeGoals]+[OvertimeHomeGoals])),'True','False'))='True') AND ((tbl_PlayerGames.PlayerID)=tbl_Players.PlayerID) AND ((tbl_Seasons.SeasonName)='" & strSeason & "'))) AS L, " &_
						"(SELECT COUNT(*) FROM tbl_Seasons INNER JOIN (tbl_Games INNER JOIN tbl_PlayerGames ON tbl_Games.GameID = tbl_PlayerGames.GameID) ON tbl_Seasons.SeasonID = tbl_Games.SeasonID WHERE ((([FirstPeriodHomeGoals]+[SecondPeriodHomeGoals]+[ThirdPeriodHomeGoals]+[OvertimeHomeGoals])=([FirstPeriodAwayGoals]+[SecondPeriodAwayGoals]+[ThirdPeriodAwayGoals]+[OvertimeAwayGoals])) AND ((tbl_PlayerGames.PlayerID)=tbl_Players.PlayerID) AND ((tbl_Seasons.SeasonName)='" & strSeason & "'))) AS T, " &_
						"(SELECT Sum(IIf([HomeTeamID]=tbl_Players.TeamID,[FirstPeriodAwayShots]+[SecondPeriodAwayShots]+[ThirdPeriodAwayShots]+[OvertimeAwayShots],[FirstPeriodHomeShots]+[SecondPeriodHomeShots]+[ThirdPeriodHomeShots]+[OvertimeHomeShots])) AS Expr1 FROM tbl_Seasons INNER JOIN (tbl_Games INNER JOIN tbl_PlayerGames ON tbl_Games.GameID = tbl_PlayerGames.GameID) ON tbl_Seasons.SeasonID = tbl_Games.SeasonID GROUP BY tbl_PlayerGames.PlayerID, tbl_Seasons.SeasonName HAVING (((tbl_PlayerGames.PlayerID)=tbl_Players.PlayerID) AND ((tbl_Seasons.SeasonName)='" & strSeason & "'))) AS SA, " &_
						"(SELECT Sum(IIf([HomeTeamID]=tbl_Players.TeamID,[FirstPeriodAwayGoals]+[SecondPeriodAwayGoals]+[ThirdPeriodAwayGoals]+[OvertimeAwayGoals],[FirstPeriodHomeGoals]+[SecondPeriodHomeGoals]+[ThirdPeriodHomeGoals]+[OvertimeHomeGoals])) AS Expr1 FROM tbl_Seasons INNER JOIN (tbl_Games INNER JOIN tbl_PlayerGames ON tbl_Games.GameID = tbl_PlayerGames.GameID) ON tbl_Seasons.SeasonID = tbl_Games.SeasonID GROUP BY tbl_PlayerGames.PlayerID, tbl_Seasons.SeasonName HAVING (((tbl_PlayerGames.PlayerID)=tbl_Players.PlayerID) AND ((tbl_Seasons.SeasonName)='" & strSeason & "'))) AS GA, " &_
						"Format([GA]/[GP],'0.00') AS GAA, " &_
						"(SELECT Count(*) FROM tbl_Seasons INNER JOIN (tbl_Games INNER JOIN tbl_PlayerGames ON tbl_Games.GameID = tbl_PlayerGames.GameID) ON tbl_Seasons.SeasonID = tbl_Games.SeasonID WHERE (((IIf([HomeTeamID]=tbl_Players.TeamID And [FirstPeriodAwayGoals]+[SecondPeriodAwayGoals]+[ThirdPeriodAwayGoals]+[OvertimeAwayGoals]=0,'True',IIf([AwayTeamID]=tbl_Players.TeamID And [FirstPeriodHomeGoals]+[SecondPeriodHomeGoals]+[ThirdPeriodHomeGoals]+[OvertimeHomeGoals]=0,'True')))='True') AND ((tbl_PlayerGames.PlayerID)=tbl_Players.PlayerID)) GROUP BY tbl_Seasons.SeasonName HAVING (((tbl_Seasons.SeasonName)='" & strSeason & "'))) AS SO, " &_
						"Format(([SA]-[GA])/[SA],'0.000') AS SPCT, " &_
						"(SELECT Sum(Length) FROM tbl_Seasons INNER JOIN (tbl_Games INNER JOIN tbl_Penalties ON tbl_Games.GameID = tbl_Penalties.GameID) ON tbl_Seasons.SeasonID = tbl_Games.SeasonID WHERE (((tbl_Penalties.PlayerID)=tbl_Players.PlayerID) AND ((tbl_Seasons.SeasonName)='" & strSeason & "'))) AS PIM, " &_
						"tbl_Players.Alternate INTO " & strGoalieTempTable & " FROM tbl_Players " &_
						"WHERE (((tbl_Players.PlayerID) IN (SELECT tbl_PlayerGames.PlayerID FROM tbl_Games INNER JOIN tbl_PlayerGames ON tbl_Games.GameID = tbl_PlayerGames.GameID WHERE tbl_Games.SeasonID=" & strSeasonID & ")) AND ((tbl_Players.TeamID)=1) AND ((tbl_Players.Position)=5))"
'		The following where clause seemed to produce weird stats for season 14
'						"WHERE (((tbl_Players.PlayerID) IN (SELECT tbl_PlayerGames.PlayerID FROM tbl_Games INNER JOIN tbl_PlayerGames ON tbl_Games.GameID = tbl_PlayerGames.GameID WHERE tbl_Games.GameDate >= (SELECT StartDate FROM tbl_Seasons WHERE SeasonID=" & strSeasonID & ") AND tbl_Games.GameDate <= (SELECT EndDate FROM tbl_Seasons WHERE SeasonID=" & strSeasonID & "))) AND ((tbl_Players.TeamID)=1) AND ((tbl_Players.Position)=5))"

				dbConn.Execute strSQL

				Set rsGoalies = Server.CreateObject("ADODB.RecordSet")
				rsGoalies.Open "SELECT * FROM " & strGoalieTempTable & " ORDER BY Alternate DESC, GP DESC", dbConn, adOpenStatic

				rsGoalies.MoveFirst

				If Not rsGoalies.BOF AND Not rsGoalies.EOF Then
%>
					<H3>Goalie</H3>
					<CENTER>
					<TABLE WIDTH=600 CELLPADDING=2 CELLSPACING=0 BORDER=1 BORDERCOLOR="black" rules="all">
					<tbody id="goalieStatsTable">
					<TR>
						<TD COLSPAN=11 ALIGN=RIGHT><A HREF="Javascript: ShowLegend('goalielegend')">Legend</A></TD>
					</TR>

					<TR BGCOLOR="#08479E">
						<TD WIDTH=200><B>Player</B></TD>
						<TH WIDTH=40>GP</TH>
						<TH WIDTH=40>W</TH>
						<TH WIDTH=40>L</TH>
						<TH WIDTH=40>T</TH>
						<TH WIDTH=40>SA</TH>
						<TH WIDTH=40>GA</TH>
						<TH WIDTH=40>GAA</TH>
						<TH WIDTH=40>SO</TH>
						<TH WIDTH=40>SV%</TH>
						<TH WIDTH=40>PIM</TH>
					</TR>
<%


					rsGoalies.MoveFirst

					Do While Not rsGoalies.EOF

						strPlayerID = rsGoalies("PlayerID")
						strGP = rsGoalies("GP")
						strW = rsGoalies("W")
						strL = rsGoalies("L")
						strT = rsGoalies("T")
						strSA = rsGoalies("SA")
						strGA = rsGoalies("GA")
						strGAA = rsGoalies("GAA")
						strSPCT = rsGoalies("SPCT")
						strSO = rsGoalies("SO")
						If IsNull(rsGoalies("PIM") ) or rsGoalies("PIM") = "" Then
							strPIM = 0
						Else
							strPIM = rsGoalies("PIM")
						End If
%>

						<TR BGCOLOR="#DCDCDC" class="data">

							<TD><A HREF="/roster/default.asp?ID=<% =strPlayerID %>&G=1"><% =rsGoalies("FirstName") %>&nbsp;<% =rsGoalies("LastName") %><% If rsGoalies("Alternate")=TRUE Then %><FONT COLOR="blue">*</FONT><% End If %></A></TD>
							<TD CLASS="black" ALIGN=CENTER><%=ChangeNullToZero(strGP) %></TD>
							<TD CLASS="black" ALIGN=CENTER><%=ChangeNullToZero(strW) %></TD>
							<TD CLASS="black" ALIGN=CENTER><%=ChangeNullToZero(strL) %></TD>
							<TD CLASS="black" ALIGN=CENTER><%=ChangeNullToZero(strT) %></TD>
							<TD CLASS="black" ALIGN=CENTER><%=ChangeNullToZero(strSA) %></TD>
							<TD CLASS="black" ALIGN=CENTER><%=ChangeNullToZero(strGA) %></TD>
							<TD CLASS="black" ALIGN=CENTER><%=ChangeNullToZero(strGAA) %></TD>
							<TD CLASS="black" ALIGN=CENTER><%=ChangeNullToZero(strSO) %></TD>
							<TD CLASS="black" ALIGN=CENTER><%=ChangeNullToZero(strSPCT) %></TD>
							<TD CLASS="black" ALIGN=CENTER><%=ChangeNullToZero(strPIM) %></TD>
						</TR>

<%

						rsGoalies.MoveNext
					Loop
%>
                    </tbody>
					</TABLE>
					</CENTER>
					&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<FONT COLOR="Blue">*</FONT>Alternate Goalie
					<BR>
					<h3>Player Search</h3>
			        <form action="playerSearchResults.asp" method="get">
			            Last name (type slowly for autocompletion): <input type="text" name="playerSearch" onkeyup="autoComplete(event.keyCode)" id="playerSearch"/><input type="submit" value="Search" />
			        </form>
</div>

<div id="pastSeasons">
					<HR>

<%
				End If

			rsGoalies.Close
			Set rsGoalies = nothing

			strSQL = "DROP TABLE " & strGoalieTempTable

			dbConn.Execute strSQL

		End If

		strSQL = "Select * FROM tbl_Seasons WHERE EndDate<Now()"

		If fmSeasonID <> "" Then
			strSQL = strSQL &  " AND SeasonID<>" & fmSeasonID
		End If

		Set rsSeason = dbConn.Execute(strSQL)

		response.write "<P id=""pastSeasonHandle""><B>"
		If fmSeasonID <> "" Then
			response.write "Other "
		End If

		response.write "Past Season Stats</B>:"
		response.write "<UL>"

		Do While Not rsSeason.EOF
			response.write "<LI><A HREF=""default.asp?SeasonID=" & rsSeason("SeasonID") & """>" & rsSeason("SeasonName") & "</A></LI>"
			rsSeason.MoveNext
		Loop

		response.write "</UL>"

		If fmSeasonID <> "" Then
			response.write "<TABLE WIDTH=100% CELLPADDING=0 CELLSPACING=0 BORDER=0><TR><TD ALIGN=RIGHT>"
			response.write "<A HREF=""default.asp"">Return to current season's stats.</A>"
			response.write "</TD></TR></TABLE>"
		End If

		rsSeason.Close
		Set rsSeason = nothing

		dbConn.Close
		Set dbConn= nothing
		
		Response.Write "<A HREF=""records.asp"">Records Book</A>"
%>
</div>
</div>
<br /><br /> <!-- br's add a bottom margin -->
	</TD>
</TR>
</TABLE>
</BODY>
</HTML>