<%
	Option Explicit
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<!--#include virtual="/include/adovbs.inc"-->
<HTML>
<HEAD>
<TITLE>DC Invaders Player Search</TITLE>
<LINK REL="stylesheet" HREF="../styles/hockey.css" TYPE="text/css">
</HEAD>
<BODY BACKGROUND="../images/stars.gif">
<TABLE WIDTH=660 CELLPADDING=0 CELLSPACING=0 BORDER=0>
<TR>
	<TD><CENTER><IMG SRC="../images/invader.gif" ORDER=0></CENTER>
<%
		Dim dbConn, strSQL, rsPlayers
		Dim playerSearch
		Dim playerId, firstName, lastName, position

        playerSearch = Request.QueryString("playerSearch")

		Set dbConn = Server.CreateObject("ADODB.Connection")
		'dbConn.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("/invaders/database/hockey.mdb")
		dbConn.Open "DSN=hockey"

		If playerSearch = "" Then
			strSQL = "SELECT PlayerID, FirstName, LastName FROM tbl_Players WHERE TeamID = 1 ORDER BY LastName, FirstName, PlayerId"
			Set rsPlayers = dbConn.Execute(strSQL)
		Else
			strSQL = "SELECT PlayerID, FirstName, LastName, pos.Position FROM tbl_Players p inner join tbl_Positions pos on p.Position = pos.PositionID WHERE TeamID = 1 AND LastName LIKE '" & playerSearch & "%' ORDER BY LastName, FirstName, PlayerId"
			Set rsPlayers = dbConn.Execute(strSQL)
		End If

		If rsPlayers.BOF AND rsPlayers.EOF Then
			Response.write "<H3>No players found.</H3>"
		Else
		    Response.write "<CENTER><H2>Invaders Players</H2></CENTER>"
		    rsPlayers.MoveFirst
		    
			Do While Not rsPlayers.EOF

				playerID = rsPlayers("PlayerID")
				firstName = rsPlayers("FirstName")
				lastName = rsPlayers("LastName")
				position = rsPlayers("Position")
				response.write "<A HREF=""../roster/default.asp?ID=" & playerID & """>" & firstName & " " & lastName & " (" & position & ")</A><BR>"
				rsPlayers.MoveNext
			Loop

			rsPlayers.Close
        End If
		dbConn.Close
		Set dbConn= nothing
%>
	</TD>
</TR>
</TABLE>
</BODY>
</HTML>