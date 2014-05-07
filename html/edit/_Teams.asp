<%
    Option Explicit
    On Error Resume Next 

    Dim sMsg, sErr, sMDB, sConn, vID, vTeamID,vTeamName,vLocationID,vSummer2001TeamNumber

    vTeamName = Request.Form("txtTeamName")
    vLocationID = Request.Form("txtLocationID")
    vSummer2001TeamNumber = Request.Form("txtSummer2001TeamNumber")


    vID = Request.QueryString("ID")
    If vID = "" Then vID = 0
    If Request("btnNew") <> "" Then 
        vID = -1
        vTeamName = ""
        vLocationID = ""
        vSummer2001TeamNumber = ""
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
<FORM method=post action=_Teams.asp?ID=<%=vID%> id=frm_Teams & name=frm_Teams>

    <table border="1" width="440" id="_Teams">
        <tr valign="top"><td width="180"> <INPUT type="submit" value="Save" id="btnSave" name="btnSave"> &nbsp; 
            <INPUT type="submit" value="New" id="btnNew" name="btnNew"> &nbsp; 
            <INPUT type="submit" value="Last" id="btnGet" name="btnGet">
        </td><td width="260"> 
            <%=sMsg%> <INPUT type="hidden" id="txtTeamID" name="txtTeamID" value="<%=Trim(vTeamID)%>">
        </td></tr>

<tr><td colspan=2 align=center>&nbsp;<a href="_Teams.asp?ID=1">Jump To Record #1</a><br>
&nbsp;<a href="_Teams.asp?ID=2">Jump To Record #2</a><br>
&nbsp;<a href="_Teams.asp?ID=3">Jump To Record #3</a><br>
&nbsp;</td></tr>        <tr valign="top"><td width="180">  Team Name
        </td><td width="260"> 
            <INPUT style="FONT-SIZE: xx-small; WIDTH: 100px" id="txtTeamName" name="txtTeamName" value="<%=Trim(vTeamName)%>">
        </td></tr>

        <tr valign="top"><td width="180">  Location I D
        </td><td width="260"> 
            <INPUT style="FONT-SIZE: xx-small; WIDTH: 75px" id="txtLocationID" name="txtLocationID" value="<%=Trim(vLocationID)%>">
        </td></tr>

        <tr valign="top"><td width="180">  Summer2001 Team Number
        </td><td width="260"> 
            <INPUT style="FONT-SIZE: xx-small; WIDTH: 75px" id="txtSummer2001TeamNumber" name="txtSummer2001TeamNumber" value="<%=Trim(vSummer2001TeamNumber)%>">
        </td></tr>

    </table>

<%

    Function InsertNewRecord()
        Dim sSQL, vID, sConnect, cnnDBS, rsDataTable

        Set cnnDBS = server.CreateObject("adodb.connection")
        
		'cnnDBS.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("/invaders/database/hockey.mdb")
		cnnDBS.Open "DSN=hockey"

        sSQL = "INSERT INTO tbl_Teams (TeamName,LocationID,Summer2001TeamNumber) "
        sSQL = sSQL & " VALUES('" & Left(Replace(vTeamName,"'","''"),50) & "'," & vLocationID & "," & vSummer2001TeamNumber & ") "

        sSQL = Replace(sSQL,"'Null'","Null")
        sSQL = Replace(sSQL,",,",",Null,")
        sSQL = Replace(sSQL,",,",",Null,")
        cnnDBS.Execute (sSQL)

        sSQL = "SELECT Max([TeamID]) As ID FROM tbl_Teams"
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

        sSQL = "UPDATE tbl_Teams SET "
        sSQL = sSQL & "TeamName='" & Left(Replace(vTeamName,"'","''"),50) & "',"
        sSQL = sSQL & "LocationID=" & vLocationID & ","
        sSQL = sSQL & "Summer2001TeamNumber=" & vSummer2001TeamNumber & ""
        sSQL = sSQL & " WHERE TeamID = " & vID


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
        If vID=-99 Then sFilter = "(DMax('TeamID','tbl_Teams'))" Else sFilter = vID
        sSQL = "SELECT * FROM tbl_Teams WHERE TeamID = " & sFilter
        Set rsDataTable = cnnDBS.Execute(sSQL)

        If Err.number <> 0 Then sErr = "<font color=red><b>ERROR: " & Err.Description  & "</b></font>"

        If Not rsDataTable.BOF And Not rsDataTable.EOF Then
            vID=rsDataTable("TeamID")
            vTeamName=rsDataTable("TeamName")
            vLocationID=rsDataTable("LocationID")
            vSummer2001TeamNumber=rsDataTable("Summer2001TeamNumber")

        End If
        Set rsDataTable = Nothing
        Set cnnDBS = Nothing

    End Function

%>

</FORM>

