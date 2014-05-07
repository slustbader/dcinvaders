<%
	Option Explicit
	'On Error Resume Next
	
	Dim strLogon
	
	strLogon = Session("Logon")

	If isNull(strLogon) or strLogon <> 1 Then
		response.redirect "./"
	End If
	
	Dim strPostID, strRecent
	Dim dbConn, rs, strSQL, i

	strPostID = Request.QueryString("ID")

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
		<H2><FONT COLOR="#FFFF66">Message Board</H2></CENTER>
		<TABLE WIDTH=100% CELLPADDING=0 CELLSPACING=0 BORDER=0>
		<TR>
			<TD ALIGN=RIGHT><A HREF="./">Back to members home.</A></TD>
		</TR>
		</TABLE>
		<BR><BR>
<%
	If strPostID = "" Then
		
		Set dbConn = Server.CreateObject("ADODB.Connection")
		'dbConn.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("/invaders/database/bb.mdb")
		dbConn.Open "DSN=bb"

		strSQL = "SELECT PostID, PostAuthor, PostSubject, PostDate FROM tbl_Posts ORDER BY PostDate Desc"
		
		Set rs = dbConn.Execute(strSQL)
		
		If NOT rs.BOF AND NOT rs.EOF Then
		
			rs.MoveFirst
	
			response.write "<TABLE WIDTH=100% CELLPADDING=2 CELLSPACING=0 BORDER=0 BGCOLOR=""#08479E"">"
			response.write "<TR>"
			response.write "<TD><B>All Posts</B></TD>"
			response.write "<TD ALIGN=RIGHT><A HREF=""post.asp"">Post Message</A></TD>"
			response.write "</TR>"
			response.write "</TABLE>"
			response.write "<UL>"
			
			Do While Not rs.EOF
				response.write "<LI>"	
				response.write "<A HREF=""messageboard.asp?ID=" & rs("PostID") & """>" & rs("PostSubject") & "</A><BR>"
				response.write FormatDateTime(rs("PostDate"),1) & " - " & rs("PostAuthor") & "<BR>"
				response.write "</LI>"
				response.write "<BR>"

				rs.MoveNext
			Loop
			
			response.write "</UL>"
		
		End If
		
		Set rs = nothing
		
		dbConn.Close
		Set dbConn = nothing

	Else
		Dim strPostDate

		Set dbConn = Server.CreateObject("ADODB.Connection")
		dbConn.Open "DSN=bb"
		'dbConn.Open "DBQ=" & Server.MapPath("/database/bb.mdb") & ";DRIVER={Microsoft Access Driver (*.mdb)}"
		
		strSQL = "SELECT PostID, PostAuthor, PostSubject, PostDate, PostText FROM tbl_Posts WHERE PostID=" & strPostID
		
		Set rs = dbConn.Execute(strSQL)		
		strPostDate = rs("PostDate")
		response.write "<A HREF=""post.asp"">Post</A> | <A HREF=""messageboard.asp"">List</A><BR><BR>"
		response.write "<TABLE WIDTH=100% CELLPADDING=0 CELLSAPCING=0 BORDER=0 BGCOLOR=""#999999""><TR><TD><BIG>" & rs("PostSubject") & "</BIG>" & "</TD><TD ALIGN=RIGHT VALIGN=TOP><SMALL>"
		response.write FormatDateTime(strPostDate,1) & ", "
		If DatePart("h", strPostDate)-5 > 0  AND DatePart("h", strPostDate)-5 < 12 Then
			response.write DatePart("h", strPostDate)-5 & ":"
			If DatePart("n", strPostDate) < 10 Then
				response.write "0"
			End If
			response.write DatePart("n", strPostDate)
			response.write " AM"
		Else
			response.write (DatePart("h", strPostDate)-17) & ":"
			If DatePart("n", strPostDate) < 10 Then
				response.write "0"
			End If
			response.write DatePart("n", strPostDate)
			response.write " PM"
		End If

		response.write "</SMALL></TD></TR>"
		response.write "<TR><TD COLSPAN=2 HEIGHT=10 CLASS=""black"">by: " & rs("PostAuthor") & "</TD></TR>"
		response.write "<TR><TD COLSPAN=2 HEIGHT=10 CLASS=""black"">" & rs("PostText") & "</TD></TR></TABLE>"
		
	End If

%>
	</TD>
</TR>
</TABLE>
</HTML>
</BODY>