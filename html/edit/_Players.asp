<%
    Option Explicit
    On Error Resume Next 

    Dim sMsg, sErr, sMDB, sConn, vID, vPlayerID,vTeamID,vFirstName,vLastName,vNickName,vNumber,vAlternateJerseyNumber,vActive,vDateActivated,vDateDeactivated,vAlternate,vPosition,vLeftRight,vCaptain,vAssistantCaptain,vHeight,vWeight,vAddress1,vAddress2,vCity,vState,vZip,vFax,vHomePhone,vWorkPhone,vCellPhone,vEmail,vEmail2,vEmail3,vBirthdate,vBirthplace,vPhoto,vHobbies,vFavoriteTeam,vFavoriteHockeyPlayer,vFavoriteFood
    Dim teamIds(100),teamNames(100)

    Call RetrieveTeams()

    vTeamID = Request.Form("txtTeamID")
    vFirstName = Request.Form("txtFirstName")
    vLastName = Request.Form("txtLastName")
    vNickName = Request.Form("txtNickName")
    vNumber = Request.Form("txtNumber")
    vAlternateJerseyNumber = Request.Form("txtAlternateJerseyNumber")
    vActive = Request.Form("txtActive")
    vDateActivated = Request.Form("txtDateActivated")
    vDateDeactivated = Request.Form("txtDateDeactivated")
    vAlternate = Request.Form("txtAlternate")
    vPosition = Request.Form("txtPosition")
    vLeftRight = Request.Form("txtLeftRight")
    vCaptain = Request.Form("txtCaptain")
    vAssistantCaptain = Request.Form("txtAssistantCaptain")
    vHeight = Request.Form("txtHeight")
    vWeight = Request.Form("txtWeight")
    vAddress1 = Request.Form("txtAddress1")
    vAddress2 = Request.Form("txtAddress2")
    vCity = Request.Form("txtCity")
    vState = Request.Form("txtState")
    vZip = Request.Form("txtZip")
    vFax = Request.Form("txtFax")
    vHomePhone = Request.Form("txtHomePhone")
    vWorkPhone = Request.Form("txtWorkPhone")
    vCellPhone = Request.Form("txtCellPhone")
    vEmail = Request.Form("txtEmail")
    vEmail2 = Request.Form("txtEmail2")
    vEmail3 = Request.Form("txtEmail3")
    vBirthdate = Request.Form("txtBirthdate")
    vBirthplace = Request.Form("txtBirthplace")
    vPhoto = Request.Form("txtPhoto")
    vHobbies = Request.Form("txtHobbies")
    vFavoriteTeam = Request.Form("txtFavoriteTeam")
    vFavoriteHockeyPlayer = Request.Form("txtFavoriteHockeyPlayer")
    vFavoriteFood = Request.Form("txtFavoriteFood")


    vID = Request.QueryString("ID")
    If vID = "" Then vID = 0
    If Request("btnNew") <> "" Then 
        vID = -1
        vTeamID = ""
        vFirstName = ""
        vLastName = ""
        vNickName = ""
        vNumber = ""
        vAlternateJerseyNumber = ""
        vActive = ""
        vDateActivated = ""
        vDateDeactivated = ""
        vAlternate = ""
        vPosition = ""
        vLeftRight = ""
        vCaptain = ""
        vAssistantCaptain = ""
        vHeight = ""
        vWeight = ""
        vAddress1 = ""
        vAddress2 = ""
        vCity = ""
        vState = ""
        vZip = ""
        vFax = ""
        vHomePhone = ""
        vWorkPhone = ""
        vCellPhone = ""
        vEmail = ""
        vEmail2 = ""
        vEmail3 = ""
        vBirthdate = ""
        vBirthplace = ""
        vPhoto = ""
        vHobbies = ""
        vFavoriteTeam = ""
        vFavoriteHockeyPlayer = ""
        vFavoriteFood = ""
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
<FORM method=post action=_Players.asp?ID=<%=vID%> id=frm_Players & name=frm_Players>

    <table border="1" width="440" id="_Players">
        <tr valign="top"><td width="180"> <INPUT type="submit" value="Save" id="btnSave" name="btnSave"> &nbsp; 
            <INPUT type="submit" value="New" id="btnNew" name="btnNew"> &nbsp; 
            <INPUT type="submit" value="Last" id="btnGet" name="btnGet">
        </td><td width="260"> 
            <%=sMsg%> <INPUT type="hidden" id="txtPlayerID" name="txtPlayerID" value="<%=Trim(vPlayerID)%>">
        </td></tr>

<tr><td colspan=2 align=center>&nbsp;<a href="_Players.asp?ID=1">Jump To Record #1</a><br>
&nbsp;<a href="_Players.asp?ID=2">Jump To Record #2</a><br>
&nbsp;<a href="_Players.asp?ID=3">Jump To Record #3</a><br>
&nbsp;</td></tr>        <tr valign="top"><td width="180">  Team I D
        </td><td width="260"> 
			<select name="txtTeamID">
        		<% 
        		dim i, selected
        		For i = 0 to UBound(teamIds) 
        			selected = ""
        			If vTeamID = teamIds(i) Then
        				selected = "SELECTED"
        			End If
        			If teamIds(i) <> "" Then 
        				Response.write "<option " & selected & " value=""" & teamIds(i)& """>" & teamNames(i) & "</option>"
        			End If
        		 Next %>
        	</select>
        </td></tr>

        <tr valign="top"><td width="180">  First Name
        </td><td width="260"> 
            <INPUT style="FONT-SIZE: xx-small; WIDTH: 100px" id="txtFirstName" name="txtFirstName" value="<%=Trim(vFirstName)%>">
        </td></tr>

        <tr valign="top"><td width="180">  Last Name
        </td><td width="260"> 
            <INPUT style="FONT-SIZE: xx-small; WIDTH: 100px" id="txtLastName" name="txtLastName" value="<%=Trim(vLastName)%>">
        </td></tr>

        <tr valign="top"><td width="180">  Nick Name
        </td><td width="260"> 
            <INPUT style="FONT-SIZE: xx-small; WIDTH: 100px" id="txtNickName" name="txtNickName" value="<%=Trim(vNickName)%>">
        </td></tr>

        <tr valign="top"><td width="180">  Number
        </td><td width="260"> 
            <INPUT style="FONT-SIZE: xx-small; WIDTH: 75px" id="txtNumber" name="txtNumber" value="<%=Trim(vNumber)%>">
        </td></tr>

        <tr valign="top"><td width="180">  Alternate Jersey Number
        </td><td width="260"> 
            <INPUT style="FONT-SIZE: xx-small; WIDTH: 75px" id="txtAlternateJerseyNumber" name="txtAlternateJerseyNumber" value="<%=Trim(vAlternateJerseyNumber)%>">
        </td></tr>

        <tr valign="top"><td width="180">  Active
        </td><td width="260"> 
            <INPUT style="FONT-SIZE: xx-small; WIDTH: 75px" id="txtActive" name="txtActive" value="<%=Trim(vActive)%>">
            (true/false)
        </td></tr>

        <tr valign="top"><td width="180">  Date Activated
        </td><td width="260"> 
            <INPUT style="FONT-SIZE: xx-small; WIDTH: 75px" id="txtDateActivated" name="txtDateActivated" value="<%=Trim(vDateActivated)%>">
        </td></tr>

        <tr valign="top"><td width="180">  Date Deactivated
        </td><td width="260"> 
            <INPUT style="FONT-SIZE: xx-small; WIDTH: 75px" id="txtDateDeactivated" name="txtDateDeactivated" value="<%=Trim(vDateDeactivated)%>">
        </td></tr>

        <tr valign="top"><td width="180">  Alternate
        </td><td width="260"> 
            <INPUT style="FONT-SIZE: xx-small; WIDTH: 75px" id="txtAlternate" name="txtAlternate" value="<%=Trim(vAlternate)%>">
        </td></tr>

        <tr valign="top"><td width="180">  Position
        </td><td width="260"> 
            <INPUT style="FONT-SIZE: xx-small; WIDTH: 75px" id="txtPosition" name="txtPosition" value="<%=Trim(vPosition)%>"><br>
            (Left Wing, Right Wing, Center, Defenseman, Goalie)
        </td></tr>

        <tr valign="top"><td width="180">  Left/Right
        </td><td width="260"> 
            <INPUT style="FONT-SIZE: xx-small; WIDTH: 100px" id="txtLeftRight" name="txtLeftRight" value="<%=Trim(vLeftRight)%>">
        </td></tr>

        <tr valign="top"><td width="180">  Captain
        </td><td width="260"> 
            <INPUT style="FONT-SIZE: xx-small; WIDTH: 75px" id="txtCaptain" name="txtCaptain" value="<%=Trim(vCaptain)%>">
            (true/false)
        </td></tr>

        <tr valign="top"><td width="180">  Assistant Captain
        </td><td width="260"> 
            <INPUT style="FONT-SIZE: xx-small; WIDTH: 75px" id="txtAssistantCaptain" name="txtAssistantCaptain" value="<%=Trim(vAssistantCaptain)%>">
            (true/false)
        </td></tr>

        <tr valign="top"><td width="180">  Height
        </td><td width="260"> 
            <INPUT style="FONT-SIZE: xx-small; WIDTH: 75px" id="txtHeight" name="txtHeight" value="<%=Trim(vHeight)%>">
            (inches)
        </td></tr>

        <tr valign="top"><td width="180">  Weight
        </td><td width="260"> 
            <INPUT style="FONT-SIZE: xx-small; WIDTH: 75px" id="txtWeight" name="txtWeight" value="<%=Trim(vWeight)%>">
        </td></tr>

        <tr valign="top"><td width="180">  Address1
        </td><td width="260"> 
            <INPUT style="FONT-SIZE: xx-small; WIDTH: 100px" id="txtAddress1" name="txtAddress1" value="<%=Trim(vAddress1)%>">
        </td></tr>

        <tr valign="top"><td width="180">  Address2
        </td><td width="260"> 
            <INPUT style="FONT-SIZE: xx-small; WIDTH: 100px" id="txtAddress2" name="txtAddress2" value="<%=Trim(vAddress2)%>">
        </td></tr>

        <tr valign="top"><td width="180">  City
        </td><td width="260"> 
            <INPUT style="FONT-SIZE: xx-small; WIDTH: 100px" id="txtCity" name="txtCity" value="<%=Trim(vCity)%>">
        </td></tr>

        <tr valign="top"><td width="180">  State
        </td><td width="260"> 
            <INPUT style="FONT-SIZE: xx-small; WIDTH: 75px" id="txtState" name="txtState" value="<%=Trim(vState)%>">
        </td></tr>

        <tr valign="top"><td width="180">  Zip
        </td><td width="260"> 
            <INPUT style="FONT-SIZE: xx-small; WIDTH: 75px" id="txtZip" name="txtZip" value="<%=Trim(vZip)%>">
        </td></tr>

        <tr valign="top"><td width="180">  Fax
        </td><td width="260"> 
            <INPUT style="FONT-SIZE: xx-small; WIDTH: 100px" id="txtFax" name="txtFax" value="<%=Trim(vFax)%>">
        </td></tr>

        <tr valign="top"><td width="180">  Home Phone
        </td><td width="260"> 
            <INPUT style="FONT-SIZE: xx-small; WIDTH: 100px" id="txtHomePhone" name="txtHomePhone" value="<%=Trim(vHomePhone)%>">
        </td></tr>

        <tr valign="top"><td width="180">  Work Phone
        </td><td width="260"> 
            <INPUT style="FONT-SIZE: xx-small; WIDTH: 100px" id="txtWorkPhone" name="txtWorkPhone" value="<%=Trim(vWorkPhone)%>">
        </td></tr>

        <tr valign="top"><td width="180">  Cell Phone
        </td><td width="260"> 
            <INPUT style="FONT-SIZE: xx-small; WIDTH: 100px" id="txtCellPhone" name="txtCellPhone" value="<%=Trim(vCellPhone)%>">
        </td></tr>

        <tr valign="top"><td width="180">  Email
        </td><td width="260"> 
            <INPUT style="FONT-SIZE: xx-small; WIDTH: 100px" id="txtEmail" name="txtEmail" value="<%=Trim(vEmail)%>">
        </td></tr>

        <tr valign="top"><td width="180">  Email2
        </td><td width="260"> 
            <INPUT style="FONT-SIZE: xx-small; WIDTH: 100px" id="txtEmail2" name="txtEmail2" value="<%=Trim(vEmail2)%>">
        </td></tr>

        <tr valign="top"><td width="180">  Email3
        </td><td width="260"> 
            <INPUT style="FONT-SIZE: xx-small; WIDTH: 100px" id="txtEmail3" name="txtEmail3" value="<%=Trim(vEmail3)%>">
        </td></tr>

        <tr valign="top"><td width="180">  Birthdate
        </td><td width="260"> 
            <INPUT style="FONT-SIZE: xx-small; WIDTH: 75px" id="txtBirthdate" name="txtBirthdate" value="<%=Trim(vBirthdate)%>">
        </td></tr>

        <tr valign="top"><td width="180">  Birthplace
        </td><td width="260"> 
            <INPUT style="FONT-SIZE: xx-small; WIDTH: 100px" id="txtBirthplace" name="txtBirthplace" value="<%=Trim(vBirthplace)%>">
        </td></tr>

        <tr valign="top"><td width="180">  Photo
        </td><td width="260"> 
            <INPUT style="FONT-SIZE: xx-small; WIDTH: 140px" id="txtPhoto" name="txtPhoto" value="<%=Trim(vPhoto)%>">
        </td></tr>

        <tr valign="top"><td width="180">  Hobbies
        </td><td width="260"> 
            <INPUT style="FONT-SIZE: xx-small; WIDTH: 75px" id="txtHobbies" name="txtHobbies" value="<%=Trim(vHobbies)%>">
        </td></tr>

        <tr valign="top"><td width="180">  Favorite Team
        </td><td width="260"> 
            <INPUT style="FONT-SIZE: xx-small; WIDTH: 100px" id="txtFavoriteTeam" name="txtFavoriteTeam" value="<%=Trim(vFavoriteTeam)%>">
        </td></tr>

        <tr valign="top"><td width="180">  Favorite Hockey Player
        </td><td width="260"> 
            <INPUT style="FONT-SIZE: xx-small; WIDTH: 100px" id="txtFavoriteHockeyPlayer" name="txtFavoriteHockeyPlayer" value="<%=Trim(vFavoriteHockeyPlayer)%>">
        </td></tr>

        <tr valign="top"><td width="180">  Favorite Food
        </td><td width="260"> 
            <INPUT style="FONT-SIZE: xx-small; WIDTH: 100px" id="txtFavoriteFood" name="txtFavoriteFood" value="<%=Trim(vFavoriteFood)%>">
        </td></tr>

    </table>

<%

    Function InsertNewRecord()
        Dim sSQL, vID, sConnect, cnnDBS, rsDataTable

        Set cnnDBS = server.CreateObject("adodb.connection")
        
		'cnnDBS.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("/invaders/database/hockey.mdb")
		cnnDBS.Open "DSN=hockey"

        sSQL = "INSERT INTO tbl_Players (TeamID,FirstName,LastName,NickName,Number,AlternateJerseyNumber,Active,DateActivated,DateDeactivated,Alternate,Position,LeftRight,Captain,AssistantCaptain,Height,Weight,Address1,Address2,City,State,Zip,Fax,HomePhone,WorkPhone,CellPhone,Email,Email2,Email3,Birthdate,Birthplace,Photo,Hobbies,FavoriteTeam,FavoriteHockeyPlayer,FavoriteFood) "
        sSQL = sSQL & " VALUES(" & vTeamID & ",'" & Left(Replace(vFirstName,"'","''"),50) & "','" & Left(Replace(vLastName,"'","''"),50) & "','" & Left(Replace(vNickName,"'","''"),50) & "'," & vNumber & "," & vAlternateJerseyNumber & "," & vActive & ",#" & vDateActivated & "#,#" & vDateDeactivated & "#," & vAlternate & "," & vPosition & ",'" & Left(Replace(vLeftRight,"'","''"),50) & "'," & vCaptain & "," & vAssistantCaptain & "," & vHeight & "," & vWeight & ",'" & Left(Replace(vAddress1,"'","''"),50) & "','" & Left(Replace(vAddress2,"'","''"),50) & "','" & Left(Replace(vCity,"'","''"),50) & "','" & Left(Replace(vState,"'","''"),2) & "','" & Left(Replace(vZip,"'","''"),5) & "','" & Left(Replace(vFax,"'","''"),50) & "','" & Left(Replace(vHomePhone,"'","''"),50) & "','" & Left(Replace(vWorkPhone,"'","''"),50) & "','" & Left(Replace(vCellPhone,"'","''"),50) & "','" & Left(Replace(vEmail,"'","''"),50) & "','" & Left(Replace(vEmail2,"'","''"),50) & "','" & Left(Replace(vEmail3,"'","''"),50) & "',#" & vBirthdate & "#,'" & Left(Replace(vBirthplace,"'","''"),50) & "','" & Left(Replace(vPhoto,"'","''"),70) & "','" & Left(Replace(vHobbies,"'","''"),0) & "','" & Left(Replace(vFavoriteTeam,"'","''"),50) & "','" & Left(Replace(vFavoriteHockeyPlayer,"'","''"),50) & "','" & Left(Replace(vFavoriteFood,"'","''"),50) & "') "

        sSQL = Replace(sSQL,"'Null'","Null")
        sSQL = Replace(sSQL,",,",",Null,")
        sSQL = Replace(sSQL,",,",",Null,")
        cnnDBS.Execute (sSQL)

        sSQL = "SELECT Max([PlayerID]) As ID FROM tbl_Players"
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

        sSQL = "UPDATE tbl_Players SET "
        sSQL = sSQL & "TeamID=" & vTeamID & ","
        sSQL = sSQL & "FirstName='" & Left(Replace(vFirstName,"'","''"),50) & "',"
        sSQL = sSQL & "LastName='" & Left(Replace(vLastName,"'","''"),50) & "',"
        sSQL = sSQL & "NickName='" & Left(Replace(vNickName,"'","''"),50) & "',"
        sSQL = sSQL & "Number=" & vNumber & ","
        sSQL = sSQL & "AlternateJerseyNumber=" & vAlternateJerseyNumber & ","
        sSQL = sSQL & "Active=" & vActive & ","
        sSQL = sSQL & "DateActivated=#" & vDateActivated & "#,"
        sSQL = sSQL & "DateDeactivated=#" & vDateDeactivated & "#,"
        sSQL = sSQL & "Alternate=" & vAlternate & ","
        sSQL = sSQL & "Position=" & vPosition & ","
        sSQL = sSQL & "LeftRight='" & Left(Replace(vLeftRight,"'","''"),50) & "',"
        sSQL = sSQL & "Captain=" & vCaptain & ","
        sSQL = sSQL & "AssistantCaptain=" & vAssistantCaptain & ","
        sSQL = sSQL & "Height=" & vHeight & ","
        sSQL = sSQL & "Weight=" & vWeight & ","
        sSQL = sSQL & "Address1='" & Left(Replace(vAddress1,"'","''"),50) & "',"
        sSQL = sSQL & "Address2='" & Left(Replace(vAddress2,"'","''"),50) & "',"
        sSQL = sSQL & "City='" & Left(Replace(vCity,"'","''"),50) & "',"
        sSQL = sSQL & "State='" & Left(Replace(vState,"'","''"),2) & "',"
        sSQL = sSQL & "Zip='" & Left(Replace(vZip,"'","''"),5) & "',"
        sSQL = sSQL & "Fax='" & Left(Replace(vFax,"'","''"),50) & "',"
        sSQL = sSQL & "HomePhone='" & Left(Replace(vHomePhone,"'","''"),50) & "',"
        sSQL = sSQL & "WorkPhone='" & Left(Replace(vWorkPhone,"'","''"),50) & "',"
        sSQL = sSQL & "CellPhone='" & Left(Replace(vCellPhone,"'","''"),50) & "',"
        sSQL = sSQL & "Email='" & Left(Replace(vEmail,"'","''"),50) & "',"
        sSQL = sSQL & "Email2='" & Left(Replace(vEmail2,"'","''"),50) & "',"
        sSQL = sSQL & "Email3='" & Left(Replace(vEmail3,"'","''"),50) & "',"
        sSQL = sSQL & "Birthdate=#" & vBirthdate & "#,"
        sSQL = sSQL & "Birthplace='" & Left(Replace(vBirthplace,"'","''"),50) & "',"
        sSQL = sSQL & "Photo='" & Left(Replace(vPhoto,"'","''"),70) & "',"
        sSQL = sSQL & "Hobbies='" & Left(Replace(vHobbies,"'","''"),0) & "',"
        sSQL = sSQL & "FavoriteTeam='" & Left(Replace(vFavoriteTeam,"'","''"),50) & "',"
        sSQL = sSQL & "FavoriteHockeyPlayer='" & Left(Replace(vFavoriteHockeyPlayer,"'","''"),50) & "',"
        sSQL = sSQL & "FavoriteFood='" & Left(Replace(vFavoriteFood,"'","''"),50) & "'"
        sSQL = sSQL & " WHERE PlayerID = " & vID


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
        If vID=-99 Then sFilter = "(DMax('PlayerID','tbl_Players'))" Else sFilter = vID
        sSQL = "SELECT * FROM tbl_Players WHERE PlayerID = " & sFilter
        Set rsDataTable = cnnDBS.Execute(sSQL)

        If Err.number <> 0 Then sErr = "<font color=red><b>ERROR: " & Err.Description  & "</b></font>"

        If Not rsDataTable.BOF And Not rsDataTable.EOF Then
            vID=rsDataTable("PlayerID")
            vTeamID=rsDataTable("TeamID")
            vFirstName=rsDataTable("FirstName")
            vLastName=rsDataTable("LastName")
            vNickName=rsDataTable("NickName")
            vNumber=rsDataTable("Number")
            vAlternateJerseyNumber=rsDataTable("AlternateJerseyNumber")
            vActive=rsDataTable("Active")
            vDateActivated=rsDataTable("DateActivated")
            vDateDeactivated=rsDataTable("DateDeactivated")
            vAlternate=rsDataTable("Alternate")
            vPosition=rsDataTable("Position")
            vLeftRight=rsDataTable("LeftRight")
            vCaptain=rsDataTable("Captain")
            vAssistantCaptain=rsDataTable("AssistantCaptain")
            vHeight=rsDataTable("Height")
            vWeight=rsDataTable("Weight")
            vAddress1=rsDataTable("Address1")
            vAddress2=rsDataTable("Address2")
            vCity=rsDataTable("City")
            vState=rsDataTable("State")
            vZip=rsDataTable("Zip")
            vFax=rsDataTable("Fax")
            vHomePhone=rsDataTable("HomePhone")
            vWorkPhone=rsDataTable("WorkPhone")
            vCellPhone=rsDataTable("CellPhone")
            vEmail=rsDataTable("Email")
            vEmail2=rsDataTable("Email2")
            vEmail3=rsDataTable("Email3")
            vBirthdate=rsDataTable("Birthdate")
            vBirthplace=rsDataTable("Birthplace")
            vPhoto=rsDataTable("Photo")
            vHobbies=rsDataTable("Hobbies")
            vFavoriteTeam=rsDataTable("FavoriteTeam")
            vFavoriteHockeyPlayer=rsDataTable("FavoriteHockeyPlayer")
            vFavoriteFood=rsDataTable("FavoriteFood")

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
%>

</FORM>

