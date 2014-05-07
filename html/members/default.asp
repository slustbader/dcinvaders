<%
	Option Explicit
	'On Error Resume Next

	Dim pnum

	Dim strLogon
	
	strLogon = Session("Logon")

	pnum = Request.Form("pnum")	
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
		<H1><FONT COLOR="#FFFF66">Members Area</FONT></H1></CENTER>
<%	
		If pnum <> 1 AND strLogon <> 1 Then  %>
				<CENTER>
					<FORM METHOD=post ACTION="default.asp" NAME="Verify">
					<TABLE WIDTH=100% CELLPADDING=5 CELLSPACING=0 BORDER=0>
					<TR>
						<TD ALIGN=RIGHT>User Name:</TD>
						<TD><INPUT TYPE="TEXT" SIZE=20 NAME="UserID">
					</TR>
					
					<TR>
						<TD ALIGN=RIGHT>Password:</TD>
						<TD><INPUT TYPE="PASSWORD" SIZE=20 NAME="Password"></TD>
					</TR>
					</TABLE>
					<BR><BR>
					<INPUT TYPE="HIDDEN" NAME="pnum" VALUE=1>
					<INPUT TYPE=SUBMIT VALUE="Log On">
					</FORM>
				</CENTER>

<% 
		Else 
			If strLogon <> 1 Then
				Dim strUserID, strPassword
				Dim dbConn
				Dim rs
				Dim strSQL

				strUserID = Request.Form("UserID")
				strPassword = Request.Form("Password")
				'response.write "strUserID = " & strUserID & "<BR>Password = " & strPassword

				'Set dbConn = Server.CreateObject("ADODB.Connection")
				'dbConn.Open "DSN=bb"
				'dbConn.Open "DBQ=" & Server.Mappath("/database/bb.mdb") & ";DRIVER={Microsoft Access Driver (*.mdb)}"
				'Set rs = Server.CreateObject("ADODB.Recordset")					
				'strSQL = "SELECT UserName, Password FROM tbl_Admin WHERE UserName='" & strUserID & "'"

				'Set rs = dbConn.Execute(strSQL)

				'If Not rs.BOF AND Not rs.EOF Then
					'If rs("Password") <> strPassword then
					If (strUserID <> "Invader" AND strPassword <> "Spacepuck") Then
						response.write "You have entered an user name and/or invalid password.  <A HREF=""./"">Return</A> to the form and enter a valid password."
					Else
						Session("Logon") = 1
						response.write "<H3><A HREF=""contact.asp"">Team Contact Information</A></H3>"
						'response.write "<H3><A HREF=""poll.asp"">Current Polls</A></H3>"
						response.write "<H3><A HREF=""makeroster.asp"">Make Roster</A></H3>"
						response.write "<H3><A HREF=""/edit/links.html"">Update DB</A></H3>"
						response.write "<H3><FONT COLOR=""#FFFF66"">Message Board</FONT></H3>"

						response.write "<TABLE WIDTH=100% CELLPADDING=2 CELLSPACING=0 BORDER=0 BGCOLOR=""#08479E"">"
						response.write "<TR>"
						response.write "<TD><B>Recent Messages</B></TD>"
						response.write "<TD ALIGN=RIGHT><A HREF=""post.asp"">Post</A> | <A HREF=""messageboard.asp"">View All</A></TD>"
						response.write "</TR>"
						response.write "</TABLE>"

						'strSQL = "SELECT TOP 5 * FROM tbl_Posts ORDER BY PostDate DESC"

						'Set rs = dbConn.Execute(strSQL)

						'If NOT rs.BOF AND NOT rs.EOF Then

							'rs.MoveFirst

							'response.write "<UL>"

							'Do While Not rs.EOF
							'	response.write "<LI>"	
							'	response.write "<A HREF=""messageboard.asp?ID=" & rs("PostID") & """>" & rs("PostSubject") & "</A><BR>"
							'	response.write FormatDateTime(rs("PostDate"),1) & " - " & rs("PostAuthor") & "<BR>"
							'	response.write "</LI>"
							'	response.write "<BR>"

							'	rs.MoveNext
							'Loop

							'response.write "</UL>"


							'response.write "<BR><BR>"
						'Else
						'	response.write "<P>There have been no messages posted to the message board.</P><BR><BR><BR>"
						'End If
					End If
				'Else
				'	response.write "You have entered an invalid User Name.  Return to the form and enter a valid User Name"
				'End If

				'Set rs = nothing

				'dbConn.Close
				'Set dbConn = nothing
			Else

				'Set dbConn = Server.CreateObject("ADODB.Connection")
				'dbConn.Open "DBQ=" & Server.Mappath("/database/bb.mdb") & ";DRIVER={Microsoft Access Driver (*.mdb)}"
				'dbConn.Open "DSN=bb"
			
				response.write "<H3><A HREF=""contact.asp"">Team Contact Information</A></H3>"
				'response.write "<H3><A HREF=""poll.asp"">Current Polls</A></H3>"
				response.write "<H3><A HREF=""makeroster.asp"">Make Roster</A></H3>"
				'response.write "<H3>Message Board</H3>"

				'response.write "<TABLE WIDTH=100% CELLPADDING=2 CELLSPACING=0 BORDER=0 BGCOLOR=""#08479E"">"
				'response.write "<TR>"
				'response.write "<TD><B>Recent Messages</B></TD>"
				'response.write "<TD ALIGN=RIGHT><A HREF=""post.asp"">Post</A> | <A HREF=""messageboard.asp"">View All</A></TD>"
				'response.write "</TR>"
				'response.write "</TABLE>"
				
				'strSQL = "SELECT TOP 5 * FROM tbl_Posts ORDER BY PostDate DESC"

				'Set rs = dbConn.Execute(strSQL)

				'If NOT rs.BOF AND NOT rs.EOF Then

				'	rs.MoveFirst

				'	response.write "<UL>"

				'	Do While Not rs.EOF
				'		response.write "<LI>"	
				'		response.write "<A HREF=""messageboard.asp?ID=" & rs("PostID") & """>" & rs("PostSubject") & "</A><BR>"
				'		response.write FormatDateTime(rs("PostDate"),1) & " - " & rs("PostAuthor") & "<BR>"
				'		response.write "</LI>"
				'		response.write "<BR>"

				'		rs.MoveNext
				'	Loop

				'	response.write "</UL>"


				'	response.write "<BR><BR>"
				'Else
				'	response.write "<P>There have been no messages posted to the message board.</P><BR><BR><BR>"
				'End If				
			End If
		End If
	%>
	</TD>
</TR>
</TABLE>
</HTML>
</BODY>