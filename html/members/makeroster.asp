<%
	Option Explicit
	'On Error Resume Next

	Dim strLogon
	Dim pnum
	Dim dbConn
	Dim rsRoster
	Dim strSQL
	Dim strPlayerID

	strLogon = Session("Logon")

	If isNull(strLogon) or strLogon <> 1 Then
		response.redirect "./"
	End If

	pnum = request.Form("pnum")
	strPlayerID = request.Form("ID")

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
	<TD><CENTER><IMG SRC="../images/invader.gif" WIDTH=216 HEIGHT=210 BORDER=0>

<%
	If pnum <> 1 Then
%>
	<H2><FONT COLOR="#FFFF66">Roster Maker</FONT></H2></CENTER>
	<P>This is a tool for the acting captain to use before games to generate a roster printout to give to the score keeper.</P>
	<P><B>1.</B> <EM>Choose the Season for which you wish to generate a roster.</EM></B>
	<FORM ACTION="makeroster.asp" METHOD=POST>
	<SELECT NAME="SeasonID">
		<OPTION VALUE="1" SELECTED>Winter Season</OPTION>
		<OPTION Value="2">Summer Season</OPTION>
	</SELECT></P>
	<P><B>2.</B> <EM>Check the players on the roster who you wish to keep on the line up.  Uncheck any players you wish not to appear on the linup.</EM></P>

<%

	Set rsRoster = dbConn.Execute("Select PlayerID, FirstName, LastName, Number, AlternateJerseyNumber, Captain, AssistantCaptain, tbl_Positions.Position FROM tbl_Players INNER JOIN tbl_Positions ON tbl_Players.Position = tbl_Positions.PositionID WHERE TeamID=1 AND Alternate=FALSE AND ACTIVE=TRUE ORDER BY tbl_Players.LastName, tbl_Players.FirstName")
%>
	<CENTER>
	<TABLE WIDTH=80% CELLPADDING=2 CELLSPACING=0 BORDER=1 BORDERCOLOR="black">
	<TR BGCOLOR="#08479E">
		<TD><B>Player Name</B></TD>
		<TD><B>Position</B></TD>
		<TD WIDTH=20><B>Number</B></TD>
		<TD WIDTH=20><B>Summer Number</B></TD>
		<TD WIDTH=20><B>Select</B></TD>
	</TR>

<%
	Do While Not rsRoster.EOF
%>
		<TR BGCOLOR="#DCDCDC">
			<TD CLASS="black"><%=rsRoster("FirstName") & " " & rsRoster("LastName") %>
			<%
				If rsRoster("Captain")=TRUE Then response.write " <FONT COLOR=""red"")>(C)</FONT>"
				If rsRoster("AssistantCaptain")=TRUE Then response.write " <FONT COLOR=""red"")>(A)</FONT>"
			%>
			</TD>
			<TD CLASS="black"><%=rsRoster("Position") %></TD>
			<TD CLASS="black"><%=rsRoster("Number") %></TD>
			<TD CLASS="black"><%=rsRoster("AlternateJerseyNumber") %></TD>
			<TD CLASS="black" ALIGN=CENTER><INPUT TYPE="CHECKBOX" NAME="ID" VALUE="<%=rsRoster("PlayerID") %>" CHECKED=TRUE></TD>
		</TR>

<%
		rsRoster.MoveNext
	Loop
%>

	</TABLE>
	</CENTER>
	<BR>
	<P><B>3.</B> <EM>Enter the names (and position and number if you have them) for any alternate players who need to be added to the roster.</EM></P>
	<CENTER>
	<TABLE WIDTH=80% CELLPADDING=2 CELLSPACING=0 BORDER=1 BORDERCOLOR="black">
	<TR BGCOLOR="#08479E">
		<TD><B>Player Name</B></TD>
		<TD><B>Position</B></TD>
		<TD WIDTH=20><B>Number</B></TD>
	</TR>

	<TR BGCOLOR="#DCDCDC">
		<TD CLASS="black"><INPUT TYPE="TEXT" NAME="Alt1" SIZE=30></TD>
		<TD CLASS="black"><INPUT TYPE="TEXT" NAME="Alt1Pos"></TD>
		<TD CLASS="black"><INPUT TYPE="TEXT" NAME="Alt1Num" SIZE=3></TD>
	</TR>

	<TR BGCOLOR="#DCDCDC">
		<TD CLASS="black"><INPUT TYPE="TEXT" NAME="Alt2" SIZE=30></TD>
		<TD CLASS="black"><INPUT TYPE="TEXT" NAME="Alt2Pos"></TD>
		<TD CLASS="black"><INPUT TYPE="TEXT" NAME="Alt2Num" SIZE=3></TD>
	</TR>

	<TR BGCOLOR="#DCDCDC">
		<TD CLASS="black"><INPUT TYPE="TEXT" NAME="Alt3" SIZE=30></TD>
		<TD CLASS="black"><INPUT TYPE="TEXT" NAME="Alt3Pos"></TD>
		<TD CLASS="black"><INPUT TYPE="TEXT" NAME="Alt3Num" SIZE=3></TD>
	</TR>
	</TABLE>
	<BR><BR>
	</CENTER>

	<P><B>4.</B> <EM>Click the "Generate Roster" button below to create a printer friendly roster to take to the game.</EM></P>
	<INPUT TYPE="SUBMIT" VALUE="Generate Roster">
	<INPUT TYPE="HIDDEN" NAME="pnum" VALUE=1>
	</FORM>
<%
	Else
		Dim arPlayers, strPlayers, player
		Dim i

		arPlayers = Split(Request.Form("ID"),", ")

		Do While i <= UBound(arPlayers)
			If i > 0 Then
				strPlayers = strPlayers & " OR "
			Else
				strPlayers = strPlayers & " ("
			End If
			strPlayers = strPlayers & "PlayerID=" & arPlayers(i)
			If i = UBound(arPlayers) Then
				strPlayers= strPlayers & ")"
			End If
			i = i + 1
		Loop


		strSQL = "SELECT FirstName, LastName, tbl_Positions.Position, Number, AlternateJerseyNumber, Captain, AssistantCaptain FROM tbl_Players INNER JOIN tbl_Positions ON tbl_Players.Position = tbl_Positions.PositionID WHERE " & strPlayers & " ORDER BY LastName, FirstName"
		'response.write strSQL

		Set rsRoster = dbConn.Execute(strSQL)
%>
	<H2><FONT COLOR="#FFFF66">DC Invaders Roster</FONT></H2></CENTER>
	<CENTER>
	<TABLE WIDTH=80% CELLPADDING=2 CELLSPACING=0 BORDER=1 BORDERCOLOR="black">
	<TR BGCOLOR="#08479E">
		<TD><B>Player Name</B></TD>
		<TD><B>Position</B></TD>
		<TD WIDTH=20><B>Number</B></TD>
	</TR>

<%
	Do While Not rsRoster.EOF
%>
		<TR BGCOLOR="#DCDCDC">
			<TD CLASS="black"><%=rsRoster("FirstName") & " " & rsRoster("LastName") %>
			<%
				If rsRoster("Captain")=TRUE Then response.write " <FONT COLOR=""red"">(C)"
				If rsRoster("AssistantCaptain")=TRUE Then response.write " <FONT COLOR=""red"">(A)"
			%>
			</TD>
			<TD CLASS="black"><%=rsRoster("Position") %></TD>
			<TD CLASS="black">
			<%
				If Request.Form("SeasonID")=1 Then response.write rsRoster("Number")
				If Request.Form("SeasonID")=2 Then response.write rsRoster("AlternateJerseyNumber")
			%></TD>
		</TR>

<%
		rsRoster.MoveNext
	Loop

	If Request.Form("Alt1")<>"" Then
%>
		<TR BGCOLOR="#DCDCDC">
			<TD CLASS="black"><% =Request.Form("Alt1") %></TD>
			<TD CLASS="black"><% If Request.Form("Alt1Pos")<>"" Then response.write Request.Form("Alt1Pos") Else response.write "&nbsp;" End If %></TD>
			<TD CLASS="black"><% If Request.Form("Alt1Num")<>"" Then response.write Request.Form("Alt1Num") Else response.write "&nbsp;" End If %></TD>
		</TR>
<%
	End If

	If Request.Form("Alt2")<>"" Then
%>
		<TR BGCOLOR="#DCDCDC">
			<TD CLASS="black"><% =Request.Form("Alt2") %></TD>
			<TD CLASS="black"><% If Request.Form("Alt2Pos")<>"" Then response.write Request.Form("Alt2Pos") Else response.write "&nbsp;" End If %></TD>
			<TD CLASS="black"><% If Request.Form("Alt2Num")<>"" Then response.write Request.Form("Alt2Num") Else response.write "&nbsp;" End If %></TD>
		</TR>
<%
	End If

	If Request.Form("Alt3")<>"" Then
%>
		<TR BGCOLOR="#DCDCDC">
			<TD CLASS="black"><% =Request.Form("Alt3") %></TD>
			<TD CLASS="black"><% If Request.Form("Alt3Pos")<>"" Then response.write Request.Form("Alt3Pos") Else response.write "&nbsp;" End If %></TD>
			<TD CLASS="black"><% If Request.Form("Alt3Num")<>"" Then response.write Request.Form("Alt3Num") Else response.write "&nbsp;" End If %></TD>
		</TR>
<%
	End If
%>
	</TABLE>
	</CENTER>
	<BR><BR><BR><BR>
<%
		rsRoster.Close
		Set rsRoster = nothing
	End If

	dbConn.Close
	Set dbConn = nothing
%>
</TD>
</TR>
</TABLE>

</BODY>
</HTML>