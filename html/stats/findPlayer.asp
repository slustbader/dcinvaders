<%
	Option Explicit
	Dim search

	search = Request.QueryString("lastName")

	Dim dbConn, strSQL, rsPlayers
	Dim lastName

	Set dbConn = Server.CreateObject("ADODB.Connection")
	'dbConn.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("/invaders/database/hockey.mdb")
	dbConn.Open "DSN=hockey"

	If search <> "" Then
		strSQL = "SELECT LastName FROM tbl_Players WHERE LastName like '" & search & "%' and teamid = 1 ORDER BY lastname "
		Set rsPlayers = dbConn.Execute(strSQL)
		If rsPlayers.BOF AND rsPlayers.EOF Then
		    Response.Status = "304"
	    Else
	        lastName = rsPlayers("LastName")
	        Response.ContentType = "text/xml"
		    Response.AddHeader "Cache-Control", "no-cache"
		    Response.write "<lastName>" & lastName & "</lastName>"
	    End If
	    rsPlayers.close
	Else
	    Response.Status = "204"
	End If
%>