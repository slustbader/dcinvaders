<%
	Option Explicit
	'On Error Resume Next

	Dim strLogon
	
	strLogon = Session("Logon")

	If isNull(strLogon) or strLogon <> 1 Then
		response.redirect "./"
	End If

	Dim dbConn
	Dim rsPollQuestions
	Dim rsPollChoices
	Dim strSQL
	Dim strUserIP

	Set dbConn = Server.CreateObject("ADODB.Connection")
	'dbConn.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("/invaders/database/polls.mdb")
	dbConn.Open "DSN=polls"

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
	<TD><CENTER><IMG SRC="/images/invader.gif" BORDER=0>
	<H2><FONT COLOR="#FFFF66">Poll</FONT></H2></CENTER>
	<TABLE WIDTH=100% CELLPADDING=0 CELLSPACING=0 BORDER=0>
	<TR>
		<TD><A HREF="pollresults.asp">View Current Poll Results</A></TD>
		<TD ALIGN=RIGHT><A HREF="pollresults.asp?Old=1">View Old Poll Results</A></TD>
	</TR>
	</TABLE>
<%
	
		Set rsPollQuestions = dbConn.Execute("SELECT PollQuestion, PollID FROM tbl_Polls WHERE PollOpenDate<Now AND PollCloseDate>Now")
	
		Do While Not rsPollQuestions.EOF
		
			response.write "<P><B><FONT COLOR=yellow>" & rsPollQuestions("PollQuestion") & "</FONT></B></P>"
			
			response.write "<FORM ACTION=""vote.asp"" METHOD=GET>"
			
			strSQL = "SELECT tbl_Polls.PollID, PollQuestion, ChoiceID, Choice FROM tbl_Polls INNER JOIN tbl_Choices ON tbl_Polls.PollID=tbl_Choices.PollID WHERE tbl_Choices.PollID=" & rsPollQuestions("PollID")
			'response.write strSQL
			Set rsPollChoices = dbConn.Execute(strSQL)

			response.write "<TABLE WIDTH=400 CELLPADDING=0 CELLSPACING=0 BORDER=0>"
			response.write "<TR><TD WIDTH=50>&nbsp;</TD>"
			response.write "<TD>"
			
			Do While Not rsPollChoices.EOF
				response.write "<INPUT TYPE=""RADIO"" NAME=""" & rsPollChoices("PollID") & """ VALUE=""" & rsPollChoices("ChoiceID") & """> " & rsPollChoices("Choice") & "<BR>"
				rsPollChoices.MoveNext
			Loop

			response.write "</TD>"
			response.write "<TD ALIGN=LEFT><INPUT TYPE=""SUBMIT"" VALUE=""Vote""></TD>"
			response.write "</TR></TABLE>"
			response.write "</FORM>"
		
			response.write "<HR COLOR=blue>"
			rsPollQuestions.MoveNext
		Loop

		rsPollQuestions.Close
		Set rsPollQuestions = nothing

		rsPollChoices.Close
		Set rsPollChoices = nothing		
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