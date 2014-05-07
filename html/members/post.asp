<%
	Option Explicit
	'On Error Resume Next

	Dim strLogon, pnum
	
	strLogon = Session("Logon")

	If isNull(strLogon) or strLogon <> 1 Then
		response.redirect "./"
	End If	
	
	pnum = Request.Form("pnum")
	
	If pnum <> 1 Then
%>
<HTML>
<HEAD>
<TITLE>Hockey Tips</TITLE>
<SCRIPT LANGUAGE="JavaScript">
<!--
	function submitform(){
		if (document.forms[0].name.value=="" ||document.forms[0].subject.value=="" || document.forms[0].body.value==""){
			alert("You must complete all fields to post a message");		
		}
		else{	
			document.forms[0].submit();
		}
	}

	function cancelform(){
		document.location.href="./";
	}

//-->
</SCRIPT>
<LINK REL="stylesheet" HREF="/styles/hockey.css" TYPE="text/css">
</HEAD>
<BODY BACKGROUND="/images/stars.gif">
<TABLE WIDTH=660 CELLPADDING=0 CELLSPACING=0 BORDER=0>
<TR>
	<TD><CENTER><IMG SRC="/images/invader.gif" BORDER=0></CENTER>
		<H2><B>Post Message</B></H2>
		<FORM ACTION="post.asp" METHOD=POST>
		<TABLE WIDTH=100% CELLPADDING=0 CELLSPACING=0 BORDER=0>
		<TR>
			<TD WIDTH=100 HEIGHT=40><B>Name:</B>:</TD>
			<TD><INPUT TYPE="TEXTBOX" SIZE=20 NAME="name"></TD>
		</TR>
		
		<TR>
			<TD WIDTH=100 HEIGHT=40><B>Subject</B>:</TD>
			<TD><INPUT TYPE="TEXTBOX" SIZE=50 NAME="subject"></TD>
		</TR>

		<TR>
			<TD VALIGN=TOP><B>Message:</B></TD>
			<TD><TEXTAREA COLS=50 ROWS=10 NAME="body"></TEXTAREA></TD>
		</TR>
		</TABLE>
		<BR><BR>
		<INPUT TYPE="HIDDEN" NAME="pnum" VALUE=1>
		<INPUT TYPE="BUTTON" VALUE="Submit" ONCLICK="submitform()">&nbsp;&nbsp;<INPUT TYPE="BUTTON" VALUE="Cancel" ONCLICK="cancelform()">
		</FORM>
	</TD>
</TR>
</TABLE>
</BODY>
</HTML>		
<%
	Else
%>
<!--#include virtual="/include/adovbs.inc"-->
<%

		Dim dbConn, rs, strSQL
		Dim strName, strSubject, strBody
		
		strName = Request.Form("name")
		strSubject = Request.Form("subject")
		strBody = Request.Form("body")
		
		strSubject = Replace(strSubject, "'", "''")
		strBody = Replace(strBody, "'", "''")
		
		Set dbConn = Server.CreateObject("ADODB.Connection")
		'dbConn.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("/invaders/database/bb.mdb")
		dbConn.Open "DSN=bb"
	
		Set rs = Server.CreateObject("ADODB.Recordset")
	
		'strSQL = "INSERT INTO tbl_Posts (PostAuthor, PostSubject, PostText, PostDate) VALUES ('" & strName & "', '" & strSubject & "',  '" & strBody & "', #" & Now() & "#)"
		strSQL = "SELECT * FROM tbl_Posts"
		
		rs.Open strSQL, dbConn, adOpenKeyset, adLockOptimistic
		
		rs.AddNew
		
		rs("PostAuthor") = strName
		rs("PostSubject") = strSubject
		rs("PostText") = strBody
		rs("PostDate") = Now()
		
		rs.Update
		
		rs.Close
		Set rs = nothing
		
		dbConn.Execute(strSQL)
		
		dbConn.Close
		Set dbConn = nothing
		
		response.redirect "./"
		
		If Err.Number > 0 then
			dbConn.Close
			Set dbConn = nothing		
			response.write "<HTML><HEAD></HEAD><BODY><FONT COLOR=red><B>Error:</B> There was a problem posting your message to the message board.  If the problem persists, contact the <A HREF=""mailto:swimn10s@yahoo.com"">website administrator</A></BODY></HTML>"
		End If
	End If		
%>