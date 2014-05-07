<%
	Option Explicit
	'On Error Resume Next
	
	Dim dbConn, rs, strSQL
	Dim strNow
	Dim strPracticeDate, strStartTime, strEndTime, strLocation, strLocationID, strComments
	
	Function FixDate(strDate)
		
		Dim strFixedDate
		
		If DatePart("h", strDate) > 12 Then
			strFixedDate = DatePart("h",strDate)-12 & ":"
			If DatePart("n",strDate) < 10 Then
				strFixedDate = strFixedDate & "0"
			End If
			strFixedDate = strFixedDate & DatePart("n", strDate)
			strFixedDate = strFixedDate & " PM"
		ElseIf DatePart("h", strDate)=0 Then
			strFixedDate = DatePart("h", strDate) & ":"
			If DatePart("n", strDate) < 10 Then
				strFixedDate = strFixedDate & "0"
			End If
			strFixedDate = strFixedDate & DatePart("n", strDate) & " AM"
		End If
	
		FixDate = strFixedDate
	End Function


	strNow = Now()
	
	Set dbConn = Server.CreateObject("ADODB.Connection")
	'dbConn.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("/invaders/database/hockey.mdb")
	dbConn.Open "DSN=hockey"
	
	Set rs = Server.CreateObject("ADODB.Recordset")
	
	strSQL = "SELECT tbl_Practices.PracticeLocationID, tbl_Locations.LocationName, tbl_Practices.PracticeDate, tbl_Practices.StartTime, tbl_Practices.EndTime, tbl_Practices.Comments FROM tbl_Practices INNER JOIN tbl_Locations ON tbl_Practices.PracticeLocationID = tbl_Locations.LocationID"
	
	If Not IsEmpty(Request.QueryString("ID")) Then
		strSQL = strSQL & " WHERE PracticeID=" & Request.QueryString("ID")
	End If
	
	strSQL = strSQL & " ORDER BY PracticeDate"
	rs.Open strSQL, dbConn, adOpenStatic
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
	<H2><FONT COLOR="#FFFF66">Scheduled Practices</FONT></H2></CENTER>


<%
		If IsEmpty(Request.QueryString("ID")) Then
			response.write "<H3>Upcoming Practices</H3>"

			rs.Filter = "PracticeDate > #" & Now()-1 & "#"
			
			If Not rs.BOF AND Not rs.EOF Then
			
				rs.MoveFirst

				Do While Not rs.EOF

%>

				<B STYLE="color: #FF0000">Date:</B> <%=FormatDateTime(rs("PracticeDate"), 1) %><BR>
				<B STYLE="color: #FF0000">Time:</B><%=FixDate(rs("StartTime")) %> to <%=FixDate(rs("EndTime")) %><BR>
				<B STYLE="color: #FF0000">Location:</B> <A HREF="../rinks/default.asp?ID=<%=rs("PracticeLocationID") %>"><%=rs("LocationName") %></A><BR>
				<HR>
<%
					rs.MoveNext
				Loop
			Else
				response.write "None"
			End If

			response.write "<H3>Past Practices</H3>"
			
			rs.Filter = "PracticeDate < #" & Now()-1 & "#"
			
			If Not rs.BOF AND Not rs.EOF Then
				rs.MoveFirst

				Do While Not rs.EOF

%>

				<B STYLE="color: #FF0000">Date:</B> <%=FormatDateTime(rs("PracticeDate"), 1) %><BR>
				<B STYLE="color: #FF0000">Time:</B> <%=FixDate(rs("StartTime")) %> to <%=FixDate(rs("EndTime")) %><BR>
				<B STYLE="color: #FF0000">Location:</B> <A HREF="../rinks/default.asp?ID=<%=rs("PracticeLocationID") %>"><%=rs("LocationName") %></A><BR>
				<HR>
<%
					rs.MoveNext
				Loop
			End If
				
		End If
		
		rs.Close
		Set rs = nothing
		
		dbConn.Close
		Set dbConn = nothing
%>
			
	</TD>
</TR>
</TABLE>
</BODY>
</HTML>