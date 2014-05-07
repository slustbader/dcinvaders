<%
	Option Explicit
	'On Error Resume Next

	Dim strLogon
	
	strLogon = Session("Logon")

	If isNull(strLogon) or strLogon <> 1 Then
		response.redirect "./"
	End If

	Dim dbConn
	Dim strSQL
	Dim strUserIP

	strUserIP = Request.ServerVariables("REMOTE_ADDR")	
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
	<TD><CENTER><H2><FONT COLOR="#FFFF66">Poll</FONT></H2></CENTER>
		<FORM ACTION="poll.asp" METHOD=GET>
<%

		Dim i
		Dim rsCheckVotes
		
		If Request.QueryString.Count < 1 Then
			response.write "You did not make a selection.  Please <A HREF=""poll.asp"">return</A> to the voting both and make a selection before clicking on the &quot;vote&quot; button."
		Else
		
			For i = 1 To Request.QueryString.Count
				strSQL = "SELECT tbl_Votes.VoterIP, tbl_Choices.PollID, tbl_Polls.PollQuestion FROM tbl_Polls INNER JOIN (tbl_Choices INNER JOIN tbl_Votes ON tbl_Choices.ChoiceID = tbl_Votes.ChoiceID) ON tbl_Polls.PollID = tbl_Choices.PollID WHERE tbl_Votes.VoterIP='" & strUserIP & "' AND tbl_Choices.PollID=" & Request.QueryString.Key(i)
	
				Set rsCheckVotes = dbConn.Execute(strSQL)
				If rsCheckVotes.BOF AND rsCheckVotes.EOF Then
					strSQL = "INSERT INTO tbl_Votes (VoterIP, ChoiceID) VALUES ('" & strUserIP & "', '" & Request.QueryString.Item(i) & "')"
					dbConn.Execute(strSQL)
					response.write "<P>Your vote has successfully been registered.</P>"
				Else
					response.write "<P><B>" & rsCheckVotes("PollQuestion") & "</B><BR><FONT COLOR=red>Error:</FONT> Our records indicate that you have already voted on this particular poll.  You are only allowed one vote per poll.  Thank you, come again.</P>"
				End If
			Next
			
			response.write "<P><A HREF=""pollresults.asp?PollID=" & Request.QueryString.Key(1) & """>Click here</A> to see the current poll results.</P>"
	
		End If
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