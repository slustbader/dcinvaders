<%
    Option Explicit
    On Error Resume Next 

    Dim sMsg, sErr, sMDB, sConn, vID, vGameID,vSeasonID,vHomeTeamID,vAwayTeamID,vLocationID,vGameDate,vGameTime,vFirstPeriodHomeShots,vFirstPeriodAwayShots,vFirstPeriodHomeGoals,vFirstPeriodAwayGoals,vSecondPeriodHomeShots,vSecondPeriodAwayShots,vSecondPeriodHomeGoals,vSecondPeriodAwayGoals,vThirdPeriodHomeShots,vThirdPeriodAwayShots,vThirdPeriodHomeGoals,vThirdPeriodAwayGoals,vOvertimeHomeShots,vOvertimeAwayShots,vOvertimeHomeGoals,vOvertimeAwayGoals,vAttendance,vReferee1,vReferee2,vScoreKeeper,vNotes,vPreviewTitle,vPreview,vRecapTitle,vRecap,vScored
    Dim teamIds(100),teamNames(100),locationIds(100),locationNames(100),seasonIds(100),seasonNames(100)
    Dim lastGameID

    Call RetrieveTeams()
    Call RetrieveLocations()
    Call RetrieveSeasons()
    
    lastGameID = getMaxGameId

    vSeasonID = Request.Form("txtSeasonID")
    vHomeTeamID = Request.Form("txtHomeTeamID")
    vAwayTeamID = Request.Form("txtAwayTeamID")
    vLocationID = Request.Form("txtLocationID")
    vGameDate = Request.Form("txtGameDate")
    vGameTime = Request.Form("txtGameTime")
    vFirstPeriodHomeShots = Request.Form("txtFirstPeriodHomeShots")
    vFirstPeriodAwayShots = Request.Form("txtFirstPeriodAwayShots")
    vFirstPeriodHomeGoals = Request.Form("txtFirstPeriodHomeGoals")
    vFirstPeriodAwayGoals = Request.Form("txtFirstPeriodAwayGoals")
    vSecondPeriodHomeShots = Request.Form("txtSecondPeriodHomeShots")
    vSecondPeriodAwayShots = Request.Form("txtSecondPeriodAwayShots")
    vSecondPeriodHomeGoals = Request.Form("txtSecondPeriodHomeGoals")
    vSecondPeriodAwayGoals = Request.Form("txtSecondPeriodAwayGoals")
    vThirdPeriodHomeShots = Request.Form("txtThirdPeriodHomeShots")
    vThirdPeriodAwayShots = Request.Form("txtThirdPeriodAwayShots")
    vThirdPeriodHomeGoals = Request.Form("txtThirdPeriodHomeGoals")
    vThirdPeriodAwayGoals = Request.Form("txtThirdPeriodAwayGoals")
    vOvertimeHomeShots = Request.Form("txtOvertimeHomeShots")
    vOvertimeAwayShots = Request.Form("txtOvertimeAwayShots")
    vOvertimeHomeGoals = Request.Form("txtOvertimeHomeGoals")
    vOvertimeAwayGoals = Request.Form("txtOvertimeAwayGoals")
    vAttendance = Request.Form("txtAttendance")
    vReferee1 = Request.Form("txtReferee1")
    vReferee2 = Request.Form("txtReferee2")
    vScoreKeeper = Request.Form("txtScoreKeeper")
    vNotes = Request.Form("txtNotes")
    vPreviewTitle = Request.Form("txtPreviewTitle")
    vPreview = Request.Form("txtPreview")
    vRecapTitle = Request.Form("txtRecapTitle")
    vRecap = Request.Form("txtRecap")
    vScored = Request.Form("txtScored")


    vID = Request.QueryString("ID")
    If vID = "" Then vID = 0
    If Request("btnNew") <> "" Then 
        vID = -1
        vSeasonID = ""
        vHomeTeamID = ""
        vAwayTeamID = ""
        vLocationID = ""
        vGameDate = ""
        vGameTime = ""
        vFirstPeriodHomeShots = "0"
        vFirstPeriodAwayShots = "0"
        vFirstPeriodHomeGoals = "0"
        vFirstPeriodAwayGoals = "0"
        vSecondPeriodHomeShots = "0"
        vSecondPeriodAwayShots = "0"
        vSecondPeriodHomeGoals = "0"
        vSecondPeriodAwayGoals = "0"
        vThirdPeriodHomeShots = "0"
        vThirdPeriodAwayShots = "0"
        vThirdPeriodHomeGoals = "0"
        vThirdPeriodAwayGoals = "0"
        vOvertimeHomeShots = "0"
        vOvertimeAwayShots = "0"
        vOvertimeHomeGoals = "0"
        vOvertimeAwayGoals = "0"
        vAttendance = ""
        vReferee1 = ""
        vReferee2 = ""
        vScoreKeeper = ""
        vNotes = ""
        vPreviewTitle = ""
        vPreview = ""
        vRecapTitle = ""
        vRecap = ""
        vScored = "false"
        sMsg = "<font color=darkgreen><b>New Record</b></font>"
    End If

    If Request("btnGet") <> "" Then vID=-99

    If Request("btnSave") <> "" Then
        If vID = "-1" Then
            vID = InsertNewRecord()
        Else
            Call UpdateRecord(vID)
            sMsg = "<font color=darkred><b>Updated Record #" & vID & "</b></font>"
        End If
    Else
        Call RetrieveRecord(vID)
        If cLng(vID) > 0 Then sMsg = "<font color=darkblue><b>Retrieved Record #" & vID & "</b></font>"
        If cLng(vID) = -99 Then sMsg = "<font color=darkpurple><b>Lookup Max (last) Record</b></font>"
    End If

    If sErr = "" Then sErr = Err.Description
    If sErr <> "" Then sMsg = sErr
    

%>
<FORM method="post" action="_Games.asp?ID=<%=vID%>" id="frm_Games" name="frm_Games">

    <table border="1" width="440" id="_Games">
        <tr valign="top"><td width="180"> <INPUT type="submit" value="Save" id="btnSave" name="btnSave"> &nbsp; 
            <INPUT type="submit" value="New" id="btnNew" name="btnNew"> &nbsp; 
            <INPUT type="submit" value="Last" id="btnGet" name="btnGet">
        </td><td width="260"> 
            <%=sMsg%> <INPUT type="hidden" id="txtGameID" name="txtGameID" value="<%=Trim(vGameID)%>">
        </td></tr>

<tr><td colspan=2 align=center>
<%
If vID > 2 Then
%>
<a href="_Games.asp?ID=2">&lt;&lt; First</a>
&nbsp;&nbsp;&nbsp;<a href="_Games.asp?ID=<%= vID - 1 %>">&lt; Prev</a>
<%
Else
%>
&lt;&lt; First&nbsp;&nbsp;&nbsp;&lt; Prev
<%
End If
 %>
<%
If vID < lastGameId Then
%>
&nbsp;&nbsp;&nbsp;<a href="_Games.asp?ID=<%= vID + 1 %>">Next &gt;</a>
&nbsp;&nbsp;&nbsp;<a href="_Games.asp?ID=<%= lastGameId %>">Last &gt;&gt;</a>
<%
Else
%>
&nbsp;&nbsp;&nbsp;Next &gt;&nbsp;&nbsp;&nbsp;Last &gt;&gt;
<%
End If
 %>
</td></tr>        <tr valign="top"><td width="180">  Season
        </td><td width="260"> 
			<select name="txtSeasonID">
        		<% 
        		dim i, selected
        		For i = 0 to UBound(seasonIds) 
        			selected = ""
        			If vSeasonID = seasonIds(i) Then
        				selected = "SELECTED"
        			End If
        			If seasonIds(i) <> "" Then 
        				Response.write "<option " & selected & " value=""" & seasonIds(i)& """>" & seasonNames(i) & "</option>"
        			End If
        		 Next %>
        	</select>
        </td></tr>

        <tr valign="top"><td width="180">  Home Team
        </td><td width="260"> 
			<select name="txtHomeTeamId">
        		<% 
        		For i = 0 to UBound(teamIds) 
        			selected = ""
        			If vHomeTeamID = teamIds(i) Then
        				selected = "SELECTED"
        			End If
        			If teamIds(i) <> "" Then 
        				Response.write "<option " & selected & " value=""" & teamIds(i)& """>" & teamNames(i) & "</option>"
        			End If
        		 Next %>
        	</select>
        </td></tr>

        <tr valign="top"><td width="180">  Away Team
        </td><td width="260"> 
			<select name="txtAwayTeamID">
        		<% 
        		For i = 0 to UBound(teamIds) 
        			selected = ""
        			If vAwayTeamID = teamIds(i) Then
        				selected = "SELECTED"
        			End If
        			If teamIds(i) <> "" Then 
        				Response.write "<option " & selected & " value=""" & teamIds(i)& """>" & teamNames(i) & "</option>"
        			End If
        		 Next %>
        	</select>
        </td></tr>

        <tr valign="top"><td width="180">  Location
        </td><td width="260"> 
			<select name="txtLocationID">
        		<% 
        		For i = 0 to UBound(locationIds) 
        			selected = ""
        			If vLocationID = locationIds(i) Then
        				selected = "SELECTED"
        			End If
        			If locationIds(i) <> "" Then 
        				Response.write "<option " & selected & " value=""" & locationIds(i)& """>" & locationNames(i) & "</option>"
        			End If
        		 Next %>
        	</select>
        </td></tr>

        <tr valign="top"><td width="180">  Game Date
        </td><td width="260"> 
            <INPUT style="FONT-SIZE: xx-small; WIDTH: 75px" id="txtGameDate" name="txtGameDate" value="<%=Trim(vGameDate)%>">
        </td></tr>

        <tr valign="top"><td width="180">  Game Time
        </td><td width="260"> 
            <INPUT style="FONT-SIZE: xx-small; WIDTH: 75px" id="txtGameTime" name="txtGameTime" value="<%=Trim(vGameTime)%>">
        </td></tr>
        <tr><td colspan="2" style="color: Red">You shouldn't have to edit anything below this point</td></tr>
        <tr valign="top"><td width="180">  First Period Home Shots
        </td><td width="260"> 
            <INPUT style="FONT-SIZE: xx-small; WIDTH: 75px" id="txtFirstPeriodHomeShots" name="txtFirstPeriodHomeShots" value="<%=Trim(vFirstPeriodHomeShots)%>">
        </td></tr>

        <tr valign="top"><td width="180">  First Period Away Shots
        </td><td width="260"> 
            <INPUT style="FONT-SIZE: xx-small; WIDTH: 75px" id="txtFirstPeriodAwayShots" name="txtFirstPeriodAwayShots" value="<%=Trim(vFirstPeriodAwayShots)%>">
        </td></tr>

        <tr valign="top"><td width="180">  First Period Home Goals
        </td><td width="260"> 
            <INPUT style="FONT-SIZE: xx-small; WIDTH: 75px" id="txtFirstPeriodHomeGoals" name="txtFirstPeriodHomeGoals" value="<%=Trim(vFirstPeriodHomeGoals)%>">
        </td></tr>

        <tr valign="top"><td width="180">  First Period Away Goals
        </td><td width="260"> 
            <INPUT style="FONT-SIZE: xx-small; WIDTH: 75px" id="txtFirstPeriodAwayGoals" name="txtFirstPeriodAwayGoals" value="<%=Trim(vFirstPeriodAwayGoals)%>">
        </td></tr>

        <tr valign="top"><td width="180">  Second Period Home Shots
        </td><td width="260"> 
            <INPUT style="FONT-SIZE: xx-small; WIDTH: 75px" id="txtSecondPeriodHomeShots" name="txtSecondPeriodHomeShots" value="<%=Trim(vSecondPeriodHomeShots)%>">
        </td></tr>

        <tr valign="top"><td width="180">  Second Period Away Shots
        </td><td width="260"> 
            <INPUT style="FONT-SIZE: xx-small; WIDTH: 75px" id="txtSecondPeriodAwayShots" name="txtSecondPeriodAwayShots" value="<%=Trim(vSecondPeriodAwayShots)%>">
        </td></tr>

        <tr valign="top"><td width="180">  Second Period Home Goals
        </td><td width="260"> 
            <INPUT style="FONT-SIZE: xx-small; WIDTH: 75px" id="txtSecondPeriodHomeGoals" name="txtSecondPeriodHomeGoals" value="<%=Trim(vSecondPeriodHomeGoals)%>">
        </td></tr>

        <tr valign="top"><td width="180">  Second Period Away Goals
        </td><td width="260"> 
            <INPUT style="FONT-SIZE: xx-small; WIDTH: 75px" id="txtSecondPeriodAwayGoals" name="txtSecondPeriodAwayGoals" value="<%=Trim(vSecondPeriodAwayGoals)%>">
        </td></tr>

        <tr valign="top"><td width="180">  Third Period Home Shots
        </td><td width="260"> 
            <INPUT style="FONT-SIZE: xx-small; WIDTH: 75px" id="txtThirdPeriodHomeShots" name="txtThirdPeriodHomeShots" value="<%=Trim(vThirdPeriodHomeShots)%>">
        </td></tr>

        <tr valign="top"><td width="180">  Third Period Away Shots
        </td><td width="260"> 
            <INPUT style="FONT-SIZE: xx-small; WIDTH: 75px" id="txtThirdPeriodAwayShots" name="txtThirdPeriodAwayShots" value="<%=Trim(vThirdPeriodAwayShots)%>">
        </td></tr>

        <tr valign="top"><td width="180">  Third Period Home Goals
        </td><td width="260"> 
            <INPUT style="FONT-SIZE: xx-small; WIDTH: 75px" id="txtThirdPeriodHomeGoals" name="txtThirdPeriodHomeGoals" value="<%=Trim(vThirdPeriodHomeGoals)%>">
        </td></tr>

        <tr valign="top"><td width="180">  Third Period Away Goals
        </td><td width="260"> 
            <INPUT style="FONT-SIZE: xx-small; WIDTH: 75px" id="txtThirdPeriodAwayGoals" name="txtThirdPeriodAwayGoals" value="<%=Trim(vThirdPeriodAwayGoals)%>">
        </td></tr>

        <tr valign="top"><td width="180">  Overtime Home Shots
        </td><td width="260"> 
            <INPUT style="FONT-SIZE: xx-small; WIDTH: 75px" id="txtOvertimeHomeShots" name="txtOvertimeHomeShots" value="<%=Trim(vOvertimeHomeShots)%>">
        </td></tr>

        <tr valign="top"><td width="180">  Overtime Away Shots
        </td><td width="260"> 
            <INPUT style="FONT-SIZE: xx-small; WIDTH: 75px" id="txtOvertimeAwayShots" name="txtOvertimeAwayShots" value="<%=Trim(vOvertimeAwayShots)%>">
        </td></tr>

        <tr valign="top"><td width="180">  Overtime Home Goals
        </td><td width="260"> 
            <INPUT style="FONT-SIZE: xx-small; WIDTH: 75px" id="txtOvertimeHomeGoals" name="txtOvertimeHomeGoals" value="<%=Trim(vOvertimeHomeGoals)%>">
        </td></tr>

        <tr valign="top"><td width="180">  Overtime Away Goals
        </td><td width="260"> 
            <INPUT style="FONT-SIZE: xx-small; WIDTH: 75px" id="txtOvertimeAwayGoals" name="txtOvertimeAwayGoals" value="<%=Trim(vOvertimeAwayGoals)%>">
        </td></tr>

        <tr valign="top"><td width="180">  Attendance
        </td><td width="260"> 
            <INPUT style="FONT-SIZE: xx-small; WIDTH: 75px" id="txtAttendance" name="txtAttendance" value="<%=Trim(vAttendance)%>">
        </td></tr>

        <tr valign="top"><td width="180">  Referee1
        </td><td width="260"> 
            <INPUT style="FONT-SIZE: xx-small; WIDTH: 100px" id="txtReferee1" name="txtReferee1" value="<%=Trim(vReferee1)%>">
        </td></tr>

        <tr valign="top"><td width="180">  Referee2
        </td><td width="260"> 
            <INPUT style="FONT-SIZE: xx-small; WIDTH: 100px" id="txtReferee2" name="txtReferee2" value="<%=Trim(vReferee2)%>">
        </td></tr>

        <tr valign="top"><td width="180">  Score Keeper
        </td><td width="260"> 
            <INPUT style="FONT-SIZE: xx-small; WIDTH: 100px" id="txtScoreKeeper" name="txtScoreKeeper" value="<%=Trim(vScoreKeeper)%>">
        </td></tr>

        <tr valign="top"><td width="180">  Notes
        </td><td width="260"> 
            <INPUT style="FONT-SIZE: xx-small; WIDTH: 75px" id="txtNotes" name="txtNotes" value="<%=Trim(vNotes)%>">
        </td></tr>

        <tr valign="top"><td width="180">  Preview Title
        </td><td width="260"> 
            <INPUT style="FONT-SIZE: xx-small; WIDTH: 200px" id="txtPreviewTitle" name="txtPreviewTitle" value="<%=Trim(vPreviewTitle)%>">
        </td></tr>

        <tr valign="top"><td width="180">  Preview
        </td><td width="260"> 
            <INPUT style="FONT-SIZE: xx-small; WIDTH: 75px" id="txtPreview" name="txtPreview" value="<%=Trim(vPreview)%>">
        </td></tr>

        <tr valign="top"><td width="180">  Recap Title
        </td><td width="260"> 
            <INPUT style="FONT-SIZE: xx-small; WIDTH: 260px" id="txtRecapTitle" name="txtRecapTitle" value="<%=Trim(vRecapTitle)%>">
        </td></tr>

        <tr valign="top"><td width="180">  Recap
        </td><td width="260"> 
            <INPUT style="FONT-SIZE: xx-small; WIDTH: 75px" id="txtRecap" name="txtRecap" value="<%=Trim(vRecap)%>">
        </td></tr>

        <tr valign="top"><td width="180">  Scored
        </td><td width="260"> 
            <INPUT style="FONT-SIZE: xx-small; WIDTH: 75px" id="txtScored" name="txtScored" value="<%=Trim(vScored)%>">
            (true/false)
        </td></tr>
    </table>

<%

    Function InsertNewRecord()
        Dim sSQL, vID, sConnect, cnnDBS, rsDataTable

        Set cnnDBS = server.CreateObject("adodb.connection")
        
		'cnnDBS.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("/invaders/database/hockey.mdb")
		cnnDBS.Open "DSN=hockey"

        sSQL = "INSERT INTO tbl_Games (SeasonID,HomeTeamID,AwayTeamID,LocationID,GameDate,GameTime,FirstPeriodHomeShots,FirstPeriodAwayShots,FirstPeriodHomeGoals,FirstPeriodAwayGoals,SecondPeriodHomeShots,SecondPeriodAwayShots,SecondPeriodHomeGoals,SecondPeriodAwayGoals,ThirdPeriodHomeShots,ThirdPeriodAwayShots,ThirdPeriodHomeGoals,ThirdPeriodAwayGoals,OvertimeHomeShots,OvertimeAwayShots,OvertimeHomeGoals,OvertimeAwayGoals,Attendance,Referee1,Referee2,ScoreKeeper,Notes,PreviewTitle,Preview,RecapTitle,Recap,Scored) "
        sSQL = sSQL & " VALUES(" & vSeasonID & "," & vHomeTeamID & "," & vAwayTeamID & "," & vLocationID & ",#" & vGameDate & "#,#" & vGameTime & "#," & vFirstPeriodHomeShots & "," & vFirstPeriodAwayShots & "," & vFirstPeriodHomeGoals & "," & vFirstPeriodAwayGoals & "," & vSecondPeriodHomeShots & "," & vSecondPeriodAwayShots & "," & vSecondPeriodHomeGoals & "," & vSecondPeriodAwayGoals & "," & vThirdPeriodHomeShots & "," & vThirdPeriodAwayShots & "," & vThirdPeriodHomeGoals & "," & vThirdPeriodAwayGoals & "," & vOvertimeHomeShots & "," & vOvertimeAwayShots & "," & vOvertimeHomeGoals & "," & vOvertimeAwayGoals & "," & vAttendance & ",'" & Left(Replace(vReferee1,"'","''"),50) & "','" & Left(Replace(vReferee2,"'","''"),50) & "','" & Left(Replace(vScoreKeeper,"'","''"),50) & "','" & Left(Replace(vNotes,"'","''"),0) & "','" & Left(Replace(vPreviewTitle,"'","''"),100) & "','" & Left(Replace(vPreview,"'","''"),0) & "','" & Left(Replace(vRecapTitle,"'","''"),200) & "','" & Left(Replace(vRecap,"'","''"),0) & "'," & vScored & ") "

        sSQL = Replace(sSQL,"'Null'","Null")
        sSQL = Replace(sSQL,",,",",Null,")
        sSQL = Replace(sSQL,",,",",Null,")
        cnnDBS.Execute (sSQL)

        sSQL = "SELECT Max([GameID]) As ID FROM tbl_Games"
        Set rsDataTable = server.CreateObject("adodb.recordset") 
        Set rsDataTable = cnnDBS.Execute(sSQL)
        InsertNewRecord = rsDataTable("ID")

        If rsDataTable.BOF And rsDataTable.EOF Then
            InsertNewRecord = 0
        Else
            InsertNewRecord = rsDataTable("ID")
        End If

        Set rsDataTable = Nothing
        Set cnnDBS = Nothing

    End Function

    Function UpdateRecord(vID)
        Dim sSQL, sConnect, cnnDBS

        Set cnnDBS = server.CreateObject("adodb.connection")
        
        sConn = "DSN=hockey"

        cnnDBS.Open sConn 

        sSQL = "UPDATE tbl_Games SET "
        sSQL = sSQL & "SeasonID=" & vSeasonID & ","
        sSQL = sSQL & "HomeTeamID=" & vHomeTeamID & ","
        sSQL = sSQL & "AwayTeamID=" & vAwayTeamID & ","
        sSQL = sSQL & "LocationID=" & vLocationID & ","
        sSQL = sSQL & "GameDate=#" & vGameDate & "#,"
        sSQL = sSQL & "GameTime=#" & vGameTime & "#,"
        sSQL = sSQL & "FirstPeriodHomeShots=" & vFirstPeriodHomeShots & ","
        sSQL = sSQL & "FirstPeriodAwayShots=" & vFirstPeriodAwayShots & ","
        sSQL = sSQL & "FirstPeriodHomeGoals=" & vFirstPeriodHomeGoals & ","
        sSQL = sSQL & "FirstPeriodAwayGoals=" & vFirstPeriodAwayGoals & ","
        sSQL = sSQL & "SecondPeriodHomeShots=" & vSecondPeriodHomeShots & ","
        sSQL = sSQL & "SecondPeriodAwayShots=" & vSecondPeriodAwayShots & ","
        sSQL = sSQL & "SecondPeriodHomeGoals=" & vSecondPeriodHomeGoals & ","
        sSQL = sSQL & "SecondPeriodAwayGoals=" & vSecondPeriodAwayGoals & ","
        sSQL = sSQL & "ThirdPeriodHomeShots=" & vThirdPeriodHomeShots & ","
        sSQL = sSQL & "ThirdPeriodAwayShots=" & vThirdPeriodAwayShots & ","
        sSQL = sSQL & "ThirdPeriodHomeGoals=" & vThirdPeriodHomeGoals & ","
        sSQL = sSQL & "ThirdPeriodAwayGoals=" & vThirdPeriodAwayGoals & ","
        sSQL = sSQL & "OvertimeHomeShots=" & vOvertimeHomeShots & ","
        sSQL = sSQL & "OvertimeAwayShots=" & vOvertimeAwayShots & ","
        sSQL = sSQL & "OvertimeHomeGoals=" & vOvertimeHomeGoals & ","
        sSQL = sSQL & "OvertimeAwayGoals=" & vOvertimeAwayGoals & ","
        sSQL = sSQL & "Attendance=" & vAttendance & ","
        sSQL = sSQL & "Referee1='" & Left(Replace(vReferee1,"'","''"),50) & "',"
        sSQL = sSQL & "Referee2='" & Left(Replace(vReferee2,"'","''"),50) & "',"
        sSQL = sSQL & "ScoreKeeper='" & Left(Replace(vScoreKeeper,"'","''"),50) & "',"
        sSQL = sSQL & "Notes='" & Left(Replace(vNotes,"'","''"),0) & "',"
        sSQL = sSQL & "PreviewTitle='" & Left(Replace(vPreviewTitle,"'","''"),100) & "',"
        sSQL = sSQL & "Preview='" & Left(Replace(vPreview,"'","''"),0) & "',"
        sSQL = sSQL & "RecapTitle='" & Left(Replace(vRecapTitle,"'","''"),200) & "',"
        sSQL = sSQL & "Recap='" & Left(Replace(vRecap,"'","''"),0) & "',"
        sSQL = sSQL & "Scored=" & vScored & ""
        sSQL = sSQL & " WHERE GameID = " & vID


        sSQL = Replace(sSQL,"=''","=Null")
        sSQL = Replace(sSQL,"=,","=Null,")
        cnnDBS.Execute (sSQL)

        Set cnnDBS = Nothing

    End Function


    Function RetrieveRecord(vID)
        Dim sSQL, sConnect, cnnDBS, rsDataTable, sFilter
        On Error Resume Next 

        Set cnnDBS = server.CreateObject("adodb.connection")
        
        sConn = "DSN=hockey"

        cnnDBS.Open sConn 

        If Err.number <> 0 Then sErr = "<font color=red><b>ERROR:  Assign WRITE permissions to Web Folder containing mdb file<br>" & Err.Description  & "</b></font>"

        Set rsDataTable = server.CreateObject("adodb.recordset")
        If vID=-99 Then sFilter = "(DMax('GameID','tbl_Games'))" Else sFilter = vID
        sSQL = "SELECT * FROM tbl_Games WHERE GameID = " & sFilter
        Set rsDataTable = cnnDBS.Execute(sSQL)

        If Err.number <> 0 Then sErr = "<font color=red><b>ERROR: " & Err.Description  & "</b></font>"

        If Not rsDataTable.BOF And Not rsDataTable.EOF Then
            vID=rsDataTable("GameID")
            vSeasonID=rsDataTable("SeasonID")
            vHomeTeamID=rsDataTable("HomeTeamID")
            vAwayTeamID=rsDataTable("AwayTeamID")
            vLocationID=rsDataTable("LocationID")
            vGameDate=rsDataTable("GameDate")
            vGameTime=rsDataTable("GameTime")
            vFirstPeriodHomeShots=rsDataTable("FirstPeriodHomeShots")
            vFirstPeriodAwayShots=rsDataTable("FirstPeriodAwayShots")
            vFirstPeriodHomeGoals=rsDataTable("FirstPeriodHomeGoals")
            vFirstPeriodAwayGoals=rsDataTable("FirstPeriodAwayGoals")
            vSecondPeriodHomeShots=rsDataTable("SecondPeriodHomeShots")
            vSecondPeriodAwayShots=rsDataTable("SecondPeriodAwayShots")
            vSecondPeriodHomeGoals=rsDataTable("SecondPeriodHomeGoals")
            vSecondPeriodAwayGoals=rsDataTable("SecondPeriodAwayGoals")
            vThirdPeriodHomeShots=rsDataTable("ThirdPeriodHomeShots")
            vThirdPeriodAwayShots=rsDataTable("ThirdPeriodAwayShots")
            vThirdPeriodHomeGoals=rsDataTable("ThirdPeriodHomeGoals")
            vThirdPeriodAwayGoals=rsDataTable("ThirdPeriodAwayGoals")
            vOvertimeHomeShots=rsDataTable("OvertimeHomeShots")
            vOvertimeAwayShots=rsDataTable("OvertimeAwayShots")
            vOvertimeHomeGoals=rsDataTable("OvertimeHomeGoals")
            vOvertimeAwayGoals=rsDataTable("OvertimeAwayGoals")
            vAttendance=rsDataTable("Attendance")
            vReferee1=rsDataTable("Referee1")
            vReferee2=rsDataTable("Referee2")
            vScoreKeeper=rsDataTable("ScoreKeeper")
            vNotes=rsDataTable("Notes")
            vPreviewTitle=rsDataTable("PreviewTitle")
            vPreview=rsDataTable("Preview")
            vRecapTitle=rsDataTable("RecapTitle")
            vRecap=rsDataTable("Recap")
            vScored=rsDataTable("Scored")

        End If
        Set rsDataTable = Nothing
        Set cnnDBS = Nothing

    End Function


    Function RetrieveTeams()
        Dim sSQL, sConnect, cnnDBS, rsDataTable, sFilter
        On Error Resume Next 
        Set cnnDBS = server.CreateObject("adodb.connection")
        cnnDBS.Open "DSN=hockey" 
        If Err.number <> 0 Then sErr = "<font color=red><b>ERROR:  Assign WRITE permissions to Web Folder containing mdb file<br>" & Err.Description  & "</b></font>"
        Set rsDataTable = server.CreateObject("adodb.recordset")
        sSQL = "SELECT TeamID, TeamName FROM tbl_Teams order by TeamName"
        Set rsDataTable = cnnDBS.Execute(sSQL)
        If Err.number <> 0 Then sErr = "<font color=red><b>ERROR: " & Err.Description  & "</b></font>"
        Dim i
        i = 0
        While Not rsDataTable.EOF
        	teamIds(i) = rsDataTable("TeamID")
        	teamNames(i) = rsDataTable("TeamName")
        	i = i + 1
        	rsDataTable.MoveNext
        WEND
        Set rsDataTable = Nothing
        Set cnnDBS = Nothing
    End Function

    Function RetrieveLocations()
        Dim sSQL, sConnect, cnnDBS, rsDataTable, sFilter
        On Error Resume Next 
        Set cnnDBS = server.CreateObject("adodb.connection")
        cnnDBS.Open "DSN=hockey" 
        If Err.number <> 0 Then sErr = "<font color=red><b>ERROR:  Assign WRITE permissions to Web Folder containing mdb file<br>" & Err.Description  & "</b></font>"
        Set rsDataTable = server.CreateObject("adodb.recordset")
        sSQL = "SELECT LocationID, LocationName FROM tbl_Locations order by LocationName"
        Set rsDataTable = cnnDBS.Execute(sSQL)
        If Err.number <> 0 Then sErr = "<font color=red><b>ERROR: " & Err.Description  & "</b></font>"
        Dim i
        i = 0
        While Not rsDataTable.EOF
        	locationIds(i) = rsDataTable("LocationID")
        	locationNames(i) = rsDataTable("LocationName")
        	i = i + 1
        	rsDataTable.MoveNext
        WEND
        Set rsDataTable = Nothing
        Set cnnDBS = Nothing
    End Function
    
    Function RetrieveSeasons()
        Dim sSQL, sConnect, cnnDBS, rsDataTable, sFilter
        On Error Resume Next 
        Set cnnDBS = server.CreateObject("adodb.connection")
        cnnDBS.Open "DSN=hockey" 
        If Err.number <> 0 Then sErr = "<font color=red><b>ERROR:  Assign WRITE permissions to Web Folder containing mdb file<br>" & Err.Description  & "</b></font>"
        Set rsDataTable = server.CreateObject("adodb.recordset")
        sSQL = "SELECT SeasonID, SeasonName FROM tbl_Seasons order by SeasonID desc"
        Set rsDataTable = cnnDBS.Execute(sSQL)
        If Err.number <> 0 Then sErr = "<font color=red><b>ERROR: " & Err.Description  & "</b></font>"
        Dim i
        i = 0
        While Not rsDataTable.EOF
        	seasonIds(i) = rsDataTable("SeasonID")
        	seasonNames(i) = rsDataTable("SeasonName")
        	i = i + 1
        	rsDataTable.MoveNext
        WEND
        Set rsDataTable = Nothing
        Set cnnDBS = Nothing
    End Function

    Function getMaxGameId()
        Dim sSQL, sConnect, cnnDBS, rsDataTable, lastId
        On Error Resume Next 
        Set cnnDBS = server.CreateObject("adodb.connection")
        cnnDBS.Open "DSN=hockey" 
        If Err.number <> 0 Then sErr = "<font color=red><b>ERROR:  Assign WRITE permissions to Web Folder containing mdb file<br>" & Err.Description  & "</b></font>"
        Set rsDataTable = server.CreateObject("adodb.recordset")
        sSQL = "SELECT top 1 GameID FROM tbl_Games order by GameID desc"
        Set rsDataTable = cnnDBS.Execute(sSQL)
        If Err.number <> 0 Then sErr = "<font color=red><b>ERROR: " & Err.Description  & "</b></font>"
        lastId = rsDataTable("GameID")
        Set rsDataTable = Nothing
        Set cnnDBS = Nothing
        getMaxGameId = lastId
    End Function
    
    %>

</FORM>

