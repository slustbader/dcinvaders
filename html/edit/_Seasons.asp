<%
    Option Explicit
    On Error Resume Next 

    Dim sMsg, sErr, sMDB, sConn, vID, vSeasonID,vSeasonName,vStartDate,vEndDate,vAltJerseySeason

    vSeasonName = Request.Form("txtSeasonName")
    vStartDate = Request.Form("txtStartDate")
    vEndDate = Request.Form("txtEndDate")
    vAltJerseySeason = Request.Form("txtAltJerseySeason")


    vID = Request.QueryString("ID")
    If vID = "" Then vID = 0
    If Request("btnNew") <> "" Then 
        vID = -1
        vSeasonName = ""
        vStartDate = ""
        vEndDate = ""
        vAltJerseySeason = ""
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
<h1>Edit Seasons</h1>
<FORM method=post action=_Seasons.asp?ID=<%=vID%> id=frm_Seasons & name=frm_Seasons>

    <table border="1" width="440" id="_Seasons">
        <tr valign="top"><td width="180"> <INPUT type="submit" value="Save" id="btnSave" name="btnSave"> &nbsp; 
            <INPUT type="submit" value="New" id="btnNew" name="btnNew"> &nbsp; 
            <INPUT type="submit" value="Last" id="btnGet" name="btnGet">
        </td><td width="260"> 
            <%=sMsg%> <INPUT type="hidden" id="txtSeasonID" name="txtSeasonID" value="<%=Trim(vSeasonID)%>">
        </td></tr>

<tr><td colspan=2 align=center>&nbsp;<a href="_Seasons.asp?ID=1">Jump To Record #1</a><br>
&nbsp;<a href="_Seasons.asp?ID=2">Jump To Record #2</a><br>
&nbsp;<a href="_Seasons.asp?ID=3">Jump To Record #3</a><br>
&nbsp;</td></tr>        <tr valign="top"><td width="180">  Season Name
        </td><td width="260"> 
            <INPUT style="FONT-SIZE: xx-small; WIDTH: 100px" id="txtSeasonName" name="txtSeasonName" value="<%=Trim(vSeasonName)%>">
        </td></tr>

        <tr valign="top"><td width="180">  Start Date
        </td><td width="260"> 
            <INPUT style="FONT-SIZE: xx-small; WIDTH: 75px" id="txtStartDate" name="txtStartDate" value="<%=Trim(vStartDate)%>">
        </td></tr>

        <tr valign="top"><td width="180">  End Date
        </td><td width="260"> 
            <INPUT style="FONT-SIZE: xx-small; WIDTH: 75px" id="txtEndDate" name="txtEndDate" value="<%=Trim(vEndDate)%>">
        </td></tr>

        <tr valign="top"><td width="180">  Alt Jersey Season
        </td><td width="260"> 
            <INPUT style="FONT-SIZE: xx-small; WIDTH: 75px" id="txtAltJerseySeason" name="txtAltJerseySeason" value="<%=Trim(vAltJerseySeason)%>">
            (true/false)
        </td></tr>

    </table>

<%

    Function InsertNewRecord()
        Dim sSQL, vID, sConnect, dbConn, rsDataTable

        Set dbConn = server.CreateObject("ADODB.Connection")
        
		'cnnDBS.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("/invaders/database/hockey.mdb")
		cnnDBS.Open "DSN=hockey"

        sSQL = "INSERT INTO tbl_Seasons (SeasonName,StartDate,EndDate,AltJerseySeason) "
        sSQL = sSQL & " VALUES('" & Left(Replace(vSeasonName,"'","''"),50) & "',#" & vStartDate & "#,#" & vEndDate & "#," & vAltJerseySeason & ") "

        sSQL = Replace(sSQL,"'Null'","Null")
        sSQL = Replace(sSQL,",,",",Null,")
        sSQL = Replace(sSQL,",,",",Null,")
        dbConn.Execute (sSQL)

        sSQL = "SELECT Max([SeasonID]) As ID FROM tbl_Seasons"
        Set rsDataTable = server.CreateObject("adodb.recordset") 
        Set rsDataTable = dbConn.Execute(sSQL)
        InsertNewRecord = rsDataTable("ID")

        If rsDataTable.BOF And rsDataTable.EOF Then
            InsertNewRecord = 0
        Else
            InsertNewRecord = rsDataTable("ID")
        End If

        Set rsDataTable = Nothing
        Set dbConn = Nothing

    End Function

    Function UpdateRecord(vID)
        Dim sSQL, sConnect, dbConn

        Set dbConn = server.CreateObject("adodb.connection")
        
 		sConn = "DSN=hockey"

        dbConn.Open sConn 

        sSQL = "UPDATE tbl_Seasons SET "
        sSQL = sSQL & "SeasonName='" & Left(Replace(vSeasonName,"'","''"),50) & "',"
        sSQL = sSQL & "StartDate=#" & vStartDate & "#,"
        sSQL = sSQL & "EndDate=#" & vEndDate & "#,"
        sSQL = sSQL & "AltJerseySeason=" & vAltJerseySeason & ""
        sSQL = sSQL & " WHERE SeasonID = " & vID


        sSQL = Replace(sSQL,"=''","=Null")
        sSQL = Replace(sSQL,"=,","=Null,")
        dbConn.Execute (sSQL)

        Set dbConn = Nothing

    End Function


    Function RetrieveRecord(vID)
        Dim sSQL, sConnect, dbConn, rsDataTable, sFilter
        On Error Resume Next 

        Set dbConn = server.CreateObject("adodb.connection")
        
 		sConn = "DSN=hockey"

        dbConn.Open sConn 

        If Err.number <> 0 Then sErr = "<font color=red><b>ERROR:  Assign WRITE permissions to Web Folder containing mdb file<br>" & Err.Description  & "</b></font>"

        Set rsDataTable = server.CreateObject("adodb.recordset")
        If vID=-99 Then sFilter = "(DMax('SeasonID','tbl_Seasons'))" Else sFilter = vID
        sSQL = "SELECT * FROM tbl_Seasons WHERE SeasonID = " & sFilter
        Set rsDataTable = dbConn.Execute(sSQL)

        If Err.number <> 0 Then sErr = "<font color=red><b>ERROR: " & Err.Description  & "</b></font>"

        If Not rsDataTable.BOF And Not rsDataTable.EOF Then
            vID=rsDataTable("SeasonID")
            vSeasonName=rsDataTable("SeasonName")
            vStartDate=rsDataTable("StartDate")
            vEndDate=rsDataTable("EndDate")
            vAltJerseySeason=rsDataTable("AltJerseySeason")

        End If
        Set rsDataTable = Nothing
        Set dbConn = Nothing

    End Function

%>

</FORM>

