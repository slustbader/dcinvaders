<%
	Option Explicit
	'On Error Resume Next

	Dim strLogon
	
	strLogon = Session("Logon")

	If isNull(strLogon) or strLogon <> 1 Then
		response.redirect "./"
	End If
	
	Dim dbConn
	Dim rs, rs2, rs3, rsPollResults
	Dim strSQL
	Dim strPollID	
	Dim strOld

	Set dbConn = Server.CreateObject("ADODB.Connection")
	'dbConn.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("/invaders/database/polls.mdb")
	dbConn.Open "DSN=polls"


	strPollID = Request.QueryString("PollID")
	strOld = Request.QueryString("Old")
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
	<TD><CENTER><IMG SRC="/images/invader.gif" BORDER=0></CENTER>
	
<%
	strSQL = "SELECT * FROM tbl_Polls WHERE PollOpenDate<Now and PollCloseDate>Now"
	
	If strPollID <> "" Then
		strSQL = strSQL & " AND PollID=" & strPollID
	End If
	
	If strOld <> "" Then
		strSQL = "SELECT * FROM tbl_Polls WHERE PollCloseDate<Now"
	End If
	
	Set rs = dbConn.Execute(strSQL)
	
	If rs.BOF AND rs.EOF Then
		If strOld <> "" Then
			response.write "<CENTER><H2><FONT COLOR=""#FFFF66"">Past Poll Results</FONT></H2></CENTER>"
			response.write "<P>There are currently no archived poll results.</P>"
		Else
			response.write "<CENTER><H2><FONT COLOR=""#FFFF66"">Current Poll Results</FONT></H2></CENTER>"
			response.write "<P>There are currently no active polls.</P>"
		End If
	Else
		response.write "<CENTER><H2><FONT COLOR=""#FFFF66"">Current Poll Results</FONT></H2></CENTER>"
		
		Do While Not rs. EOF
		
			response.write "<H4>" & rs("PollQuestion") & "</H4>"
			Set rs2 = dbConn.Execute("SELECT Choice, ChoiceID FROM tbl_Choices WHERE PollID=" & rs("PollID"))
			
			response.write "<TABLE CELLPADDING=2 CELLSPACING=0 BORDER=0><TR>"
			
			Do While Not rs2.EOF
				response.write "<TR><TD WIDTH=50>&nbsp;</TD>"
				response.write "<TD ALIGN=RIGHT><B>" & rs2("Choice") & ": </B>"
				Set rs3 = dbConn.Execute("SELECT Count(*) FROM tbl_Votes WHERE ChoiceID=" & rs2("ChoiceID"))
				response.write "</TD><TD>" & rs3.Fields(0) & "</TD>"
				response.write "</TR>"
				rs2.MoveNext
			Loop
			response.write "</TABLE>"
			rs.MoveNext
		Loop
	
		If strPollID <> "" Then
			response.write "<P><A HREF=""pollresults.asp"">Show all current poll results.</A></P>"
		End If
	
		If strOld = "" Then
			response.write "<P><A HREF=""pollresults.asp?Old=1"">View the results of past polls.</A></P>"
		End If
		
		rs3.Close
		Set rs3 = nothing
		
		rs2.Close
		Set rs2 = nothing
		
	End If
	
	rs.Close
	Set rs = nothing
%>	
	<BR><BR>
	</TD>
</TR>
</TABLE>

<%

	dbConn.Close
	Set dbConn = nothing
%>
</BODY>
</HTML>		