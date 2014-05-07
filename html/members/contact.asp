<%
	Option Explicit
	'On Error Resume Next
	
	Dim strLogon
	
	strLogon = Session("Logon")

	If isNull(strLogon) or strLogon <> 1 Then
		response.redirect "./"
	End If
	
	Dim strRecent
	Dim dbConn, rs, strSQL, i

	Function FormatPhone(strPhone)
		Dim strFormattedPhone
		
		strFormattedPhone = "(" & Left(strPhone,3) & ") " & Mid(strPhone,4,3) & "-" & Right(strPhone,4)
		
		FormatPhone=strFormattedPhone
	
	End Function

%>
<!-- #include virtual="/include/adovbs.inc"-->
<HTML>
<HEAD>
<TITLE>Hockey Tips</TITLE>
<LINK REL="stylesheet" HREF="/styles/hockey.css" TYPE="text/css">
</HEAD>
<BODY BACKGROUND="/images/stars.gif">
<TABLE WIDTH=660 CELLPADDING=0 CELLSPACING=0 BORDER=0>
<TR>
	<TD><CENTER><IMG SRC="/images/invader.gif" BORDER=0>
		<H2><FONT COLOR="#FFFF66">Team Contact Info</H2></CENTER>

		<TABLE WIDTH=100% CELLPADDING=0 CELLSPACING=0 BORDER=0>
		<TR>
			<TD ALIGN=RIGHT><A HREF="./">Back to members home.</A></TD>
		</TR>
		</TABLE>
		<BR><BR>
		
<%
	Set dbConn = Server.CreateObject("ADODB.Connection")
	'dbConn.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("/invaders/database/hockey.mdb")
	dbConn.Open "DSN=hockey"
	
	strSQL = "SELECT * FROM tbl_Players WHERE TeamID=1 AND Alternate=FALSE ORDER BY LastName"
	
	Set rs = dbConn.Execute(strSQL)
	
%>
		<TABLE WIDTH=100% CELLPADDING=2 CELLSPACING=0 BORDER=1 BORDERCOLOR="black" BGCOLOR="#EEEEEE">
		<TR BGCOLOR="#08479E">
			<TD><B>Name</B></TD>
			<TD><B>Address</B></TD>
			<TD><B>Home Phone</B></TD>
			<TD><B>Work Phone</B></TD>
			<TD><B>Cell Phone</B></TD>
			<TD><B>E-Mail</B></TD>
			<TD><B>Alternate E-Mail</B></TD>
		</TR>	

<%

	i = 0

	Do While Not rs.EOF

		If i Mod 2 = 1 Then
			response.write "<TR BGCOLOR=""#DCDCDC"">"
		Else
			response.write "<TR BGOLOR=""#C0C0C0"">"
		End If
		response.write "<TD CLASS=""black"" VALIGN=TOP><SMALL>" & rs("FirstName") & " " & rs("LastName") & "</SMALL></TD>"
		response.write "<TD CLASS=""black"" WIDTH=160><SMALL>"
			If Not IsNull(rs("Address1")) Then
				response.write rs("Address1") & "<BR>" & rs("City") & ", " & rs("State") & " " & rs("Zip")
			Else
				response.write "&nbsp;"
			End If
		response.write "</SMALL></TD>"
		response.write "<TD CLASS=""black""><SMALL>"
			If Not IsNull(rs("HomePhone")) Then
				response.write FormatPhone(rs("HomePhone"))
			Else
				response.write "&nbsp;"
			End If
		response.write "</SMALL></TD>"
		response.write "<TD CLASS=""black""><SMALL>"
			If Not IsNull(rs("WorkPhone")) Then
				response.write FormatPhone(rs("WorkPhone"))
			Else
				response.write "&nbsp;"
			End If
		response.write "</SMALL></TD>"
		response.write "<TD CLASS=""black""><SMALL>"
			If Not IsNull(rs("CellPhone")) Then
				response.write FormatPhone(rs("CellPhone"))
			Else
				response.write "&nbsp;"
			End If
		response.write "</SMALL></TD>"
		response.write "<TD CLASS=""black""><SMALL>"
			If Not IsNull(rs("email")) Then
				response.write "<A HREF=""mailto:" & rs("email") & """>" & rs("email") & "</A>"
			Else
				response.write "&nbsp;"
			End If
		response.write "</SMALL></TD>"
		response.write "<TD CLASS=""black""><SMALL>"
			If Not IsNull(rs("email2")) Then
				response.write "<A HREF=""mailto:" & rs("email2") & """>" & rs("email2") & "</A>"
			Else
				response.write "&nbsp;"
			End If
		response.write "</SMALL></TD>"
		response.write "</TR>"

		i = i + 1
		rs.MoveNext
	Loop

	response.write "</TABLE>"

	Set rs = nothing
		
	dbConn.Close
	Set dbConn = nothing
		
%>
	</TD>
</TR>
</TABLE>
<BR><BR><BR>
</BODY>
</HTML>