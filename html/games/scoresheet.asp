<HTML>
<HEAD>
<TITLE>frm_Scoresheet</TITLE>
<LINK REL="stylesheet" HREF="styles/hockey.css" TYPE="text/css">
</HEAD>
<BODY BACKGROUND="images/stars.gif"><%
If IsObject(Session("hockey_conn")) Then
    Set conn = Session("hockey_conn")
Else
    Set conn = Server.CreateObject("ADODB.Connection")
    conn.open "hockey","",""
    Set Session("??hockey_conn") = conn
End If
%>
<%
If IsObject(Session("frm_Scoresheet_rs")) Then
    Set rs = Session("frm_Scoresheet_rs")
Else
    sql = "SELECT tbl_Games.GameID, tbl_Games.GameDate, tbl_Games.GameTime, tbl_Games.FirstPeriodHomeShots, tbl_Games.FirstPeriodAwayShots, tbl_Games.FirstPeriodHomeGoals, tbl_Games.FirstPeriodAwayGoals, tbl_Games.SecondPeriodHomeShots, tbl_Games.SecondPeriodAwayShots, tbl_Games.SecondPeriodHomeGoals, tbl_Games.SecondPeriodAwayGoals, tbl_Games.ThirdPeriodHomeShots, tbl_Games.ThirdPeriodAwayShots, tbl_Games.ThirdPeriodHomeGoals, tbl_Games.ThirdPeriodAwayGoals, tbl_Games.OvertimeHomeShots, tbl_Games.OvertimeAwayShots, tbl_Games.OvertimeHomeGoals, tbl_Games.OvertimeAwayGoals, tbl_Games.Referee1, tbl_Games.Referee2, tbl_Games.ScoreKeeper, tbl_Games.Preview, tbl_Games.Recap, tbl_Games.HomeTeamID, tbl_Games.AwayTeamID, tbl_Games.LocationID, tbl_Games.Notes, tbl_Games.Attendance, tbl_Games.Scored, tbl_Seasons.AltJerseySeason FROM tbl_Seasons INNER JOIN tbl_Games ON tbl_Seasons.SeasonID = tbl_Games.SeasonID ORDER BY tbl_Games.GameDate "
    Set rs = Server.CreateObject("ADODB.Recordset")
    rs.Open sql, conn, 3, 3
    If rs.eof Then
        rs.AddNew
    End If
    Set Session("frm_Scoresheet_rs") = rs
End If
%>
<TABLE BORDER=1 BGCOLOR=#ffffff CELLSPACING=0><FONT FACE="Arial" COLOR=#000000><CAPTION><B>frm_Scoresheet</B></CAPTION></FONT>

<THEAD>
<TR>
<TH BGCOLOR=#c0c0c0 BORDERCOLOR=#000000 ><FONT style=FONT-SIZE:10pt FACE="Arial" COLOR=#000000>txt_GameDate</FONT></TH>
<TH BGCOLOR=#c0c0c0 BORDERCOLOR=#000000 ><FONT style=FONT-SIZE:10pt FACE="Arial" COLOR=#000000>Text11</FONT></TH>
<TH BGCOLOR=#c0c0c0 BORDERCOLOR=#000000 ><FONT style=FONT-SIZE:10pt FACE="Arial" COLOR=#000000>txt_FirstPeriodAwayGoals</FONT></TH>
<TH BGCOLOR=#c0c0c0 BORDERCOLOR=#000000 ><FONT style=FONT-SIZE:10pt FACE="Arial" COLOR=#000000>txt_FirstPeriodHomeGoals</FONT></TH>
<TH BGCOLOR=#c0c0c0 BORDERCOLOR=#000000 ><FONT style=FONT-SIZE:10pt FACE="Arial" COLOR=#000000>txt_SecondPeriodAwayGoals</FONT></TH>
<TH BGCOLOR=#c0c0c0 BORDERCOLOR=#000000 ><FONT style=FONT-SIZE:10pt FACE="Arial" COLOR=#000000>txt_SecondPeriodHomeGoals</FONT></TH>
<TH BGCOLOR=#c0c0c0 BORDERCOLOR=#000000 ><FONT style=FONT-SIZE:10pt FACE="Arial" COLOR=#000000>txt_ThirdPeriodAwayGoals</FONT></TH>
<TH BGCOLOR=#c0c0c0 BORDERCOLOR=#000000 ><FONT style=FONT-SIZE:10pt FACE="Arial" COLOR=#000000>txt_ThirdPeriodHomeGoals</FONT></TH>
<TH BGCOLOR=#c0c0c0 BORDERCOLOR=#000000 ><FONT style=FONT-SIZE:10pt FACE="Arial" COLOR=#000000>txt_OvertimeAwayGoals</FONT></TH>
<TH BGCOLOR=#c0c0c0 BORDERCOLOR=#000000 ><FONT style=FONT-SIZE:10pt FACE="Arial" COLOR=#000000>txt_OvertimeHomeGoals</FONT></TH>
<TH BGCOLOR=#c0c0c0 BORDERCOLOR=#000000 ><FONT style=FONT-SIZE:10pt FACE="Arial" COLOR=#000000>txt_FinalAwayGoals</FONT></TH>
<TH BGCOLOR=#c0c0c0 BORDERCOLOR=#000000 ><FONT style=FONT-SIZE:10pt FACE="Arial" COLOR=#000000>txt_FinalHomeGoals</FONT></TH>
<TH BGCOLOR=#c0c0c0 BORDERCOLOR=#000000 ><FONT style=FONT-SIZE:10pt FACE="Arial" COLOR=#000000>Text42</FONT></TH>
<TH BGCOLOR=#c0c0c0 BORDERCOLOR=#000000 ><FONT style=FONT-SIZE:10pt FACE="Arial" COLOR=#000000>cbo_AwayTeamID</FONT></TH>
<TH BGCOLOR=#c0c0c0 BORDERCOLOR=#000000 ><FONT style=FONT-SIZE:10pt FACE="Arial" COLOR=#000000>cbo_HomeTeamID</FONT></TH>
<TH BGCOLOR=#c0c0c0 BORDERCOLOR=#000000 ><FONT style=FONT-SIZE:10pt FACE="Arial" COLOR=#000000>Text68</FONT></TH>
<TH BGCOLOR=#c0c0c0 BORDERCOLOR=#000000 ><FONT style=FONT-SIZE:10pt FACE="Arial" COLOR=#000000>Text70</FONT></TH>
<TH BGCOLOR=#c0c0c0 BORDERCOLOR=#000000 ><FONT style=FONT-SIZE:10pt FACE="Arial" COLOR=#000000>Text71</FONT></TH>
<TH BGCOLOR=#c0c0c0 BORDERCOLOR=#000000 ><FONT style=FONT-SIZE:10pt FACE="Arial" COLOR=#000000>Text72</FONT></TH>
<TH BGCOLOR=#c0c0c0 BORDERCOLOR=#000000 ><FONT style=FONT-SIZE:10pt FACE="Arial" COLOR=#000000>Text73</FONT></TH>
<TH BGCOLOR=#c0c0c0 BORDERCOLOR=#000000 ><FONT style=FONT-SIZE:10pt FACE="Arial" COLOR=#000000>Text74</FONT></TH>
<TH BGCOLOR=#c0c0c0 BORDERCOLOR=#000000 ><FONT style=FONT-SIZE:10pt FACE="Arial" COLOR=#000000>Text75</FONT></TH>
<TH BGCOLOR=#c0c0c0 BORDERCOLOR=#000000 ><FONT style=FONT-SIZE:10pt FACE="Arial" COLOR=#000000>Text76</FONT></TH>
<TH BGCOLOR=#c0c0c0 BORDERCOLOR=#000000 ><FONT style=FONT-SIZE:10pt FACE="Arial" COLOR=#000000>Combo81</FONT></TH>
<TH BGCOLOR=#c0c0c0 BORDERCOLOR=#000000 ><FONT style=FONT-SIZE:10pt FACE="Arial" COLOR=#000000>Text103</FONT></TH>
<TH BGCOLOR=#c0c0c0 BORDERCOLOR=#000000 ><FONT style=FONT-SIZE:10pt FACE="Arial" COLOR=#000000>Text114</FONT></TH>
<TH BGCOLOR=#c0c0c0 BORDERCOLOR=#000000 ><FONT style=FONT-SIZE:10pt FACE="Arial" COLOR=#000000>Text118</FONT></TH>
<TH BGCOLOR=#c0c0c0 BORDERCOLOR=#000000 ><FONT style=FONT-SIZE:10pt FACE="Arial" COLOR=#000000>Text128</FONT></TH>
<TH BGCOLOR=#c0c0c0 BORDERCOLOR=#000000 ><FONT style=FONT-SIZE:10pt FACE="Arial" COLOR=#000000>Check131</FONT></TH>
<TH BGCOLOR=#c0c0c0 BORDERCOLOR=#000000 ><FONT style=FONT-SIZE:10pt FACE="Arial" COLOR=#000000>chk_AltSeason</FONT></TH>

</TR>
</THEAD>
<TBODY>
<%
On Error Resume Next
rs.MoveFirst
do while Not rs.eof
 %>
<TR VALIGN=TOP>
<TD BORDERCOLOR=#000000 ><U><FONT style=FONT-SIZE:8pt FACE="Lucida Handwriting" COLOR=#000000><%=Server.HTMLEncode(rs.Fields("GameDate").Value)%><BR></FONT></U></TD>
<TD BORDERCOLOR=#000000 ><U><FONT style=FONT-SIZE:8pt FACE="Lucida Handwriting" COLOR=#000000><%=Server.HTMLEncode(rs.Fields("GameTime").Value)%><BR></FONT></U></TD>
<TD BORDERCOLOR=#000000 ><B><FONT style=FONT-SIZE:12pt FACE="Lucida Handwriting" COLOR=#000000><%=Server.HTMLEncode(rs.Fields("FirstPeriodAwayGoals").Value)%><BR></FONT></B></TD>
<TD BORDERCOLOR=#000000 ><B><FONT style=FONT-SIZE:12pt FACE="Lucida Handwriting" COLOR=#000000><%=Server.HTMLEncode(rs.Fields("FirstPeriodHomeGoals").Value)%><BR></FONT></B></TD>
<TD BORDERCOLOR=#000000 ><B><FONT style=FONT-SIZE:12pt FACE="Lucida Handwriting" COLOR=#000000><%=Server.HTMLEncode(rs.Fields("SecondPeriodAwayGoals").Value)%><BR></FONT></B></TD>
<TD BORDERCOLOR=#000000 ><B><FONT style=FONT-SIZE:12pt FACE="Lucida Handwriting" COLOR=#000000><%=Server.HTMLEncode(rs.Fields("SecondPeriodHomeGoals").Value)%><BR></FONT></B></TD>
<TD BORDERCOLOR=#000000 ><B><FONT style=FONT-SIZE:12pt FACE="Lucida Handwriting" COLOR=#000000><%=Server.HTMLEncode(rs.Fields("ThirdPeriodAwayGoals").Value)%><BR></FONT></B></TD>
<TD BORDERCOLOR=#000000 ><B><FONT style=FONT-SIZE:12pt FACE="Lucida Handwriting" COLOR=#000000><%=Server.HTMLEncode(rs.Fields("ThirdPeriodHomeGoals").Value)%><BR></FONT></B></TD>
<TD BORDERCOLOR=#000000 ><B><FONT style=FONT-SIZE:12pt FACE="Lucida Handwriting" COLOR=#000000><%=Server.HTMLEncode(rs.Fields("OvertimeAwayGoals").Value)%><BR></FONT></B></TD>
<TD BORDERCOLOR=#000000 ><B><FONT style=FONT-SIZE:12pt FACE="Lucida Handwriting" COLOR=#000000><%=Server.HTMLEncode(rs.Fields("OvertimeHomeGoals").Value)%><BR></FONT></B></TD>
<TD BORDERCOLOR=#000000 ><B><FONT style=FONT-SIZE:12pt FACE="Lucida Handwriting" COLOR=#000000><%=Server.HTMLEncode(rs.Fields("0").Value)%><BR></FONT></B></TD>
<TD BORDERCOLOR=#000000 ><B><FONT style=FONT-SIZE:12pt FACE="Lucida Handwriting" COLOR=#000000><%=Server.HTMLEncode(rs.Fields("2").Value)%><BR></FONT></B></TD>
<TD BORDERCOLOR=#000000  ALIGN=CENTER><FONT style=FONT-SIZE:8pt FACE="Lucida Handwriting" COLOR=#000000><%=Server.HTMLEncode(rs.Fields("ScoreKeeper").Value)%><BR></FONT></TD>
<TD BORDERCOLOR=#000000 ><FONT style=FONT-SIZE:10pt FACE="Lucida Handwriting" COLOR=#000000><%=Server.HTMLEncode(rs.Fields("AwayTeamID").Value)%><BR></FONT></TD>
<TD BORDERCOLOR=#000000 ><FONT style=FONT-SIZE:10pt FACE="Lucida Handwriting" COLOR=#000000><%=Server.HTMLEncode(rs.Fields("HomeTeamID").Value)%><BR></FONT></TD>
<TD BORDERCOLOR=#000000 ><FONT style=FONT-SIZE:8pt FACE="Lucida Handwriting" COLOR=#000000><%=Server.HTMLEncode(rs.Fields("FirstPeriodHomeShots").Value)%><BR></FONT></TD>
<TD BORDERCOLOR=#000000 ><FONT style=FONT-SIZE:8pt FACE="Lucida Handwriting" COLOR=#000000><%=Server.HTMLEncode(rs.Fields("SecondPeriodHomeShots").Value)%><BR></FONT></TD>
<TD BORDERCOLOR=#000000 ><FONT style=FONT-SIZE:8pt FACE="Lucida Handwriting" COLOR=#000000><%=Server.HTMLEncode(rs.Fields("ThirdPeriodHomeShots").Value)%><BR></FONT></TD>
<TD BORDERCOLOR=#000000 ><FONT style=FONT-SIZE:8pt FACE="Lucida Handwriting" COLOR=#000000><%=Server.HTMLEncode(rs.Fields("15").Value)%><BR></FONT></TD>
<TD BORDERCOLOR=#000000 ><FONT style=FONT-SIZE:8pt FACE="Lucida Handwriting" COLOR=#000000><%=Server.HTMLEncode(rs.Fields("FirstPeriodAwayShots").Value)%><BR></FONT></TD>
<TD BORDERCOLOR=#000000 ><FONT style=FONT-SIZE:8pt FACE="Lucida Handwriting" COLOR=#000000><%=Server.HTMLEncode(rs.Fields("SecondPeriodAwayShots").Value)%><BR></FONT></TD>
<TD BORDERCOLOR=#000000 ><FONT style=FONT-SIZE:8pt FACE="Lucida Handwriting" COLOR=#000000><%=Server.HTMLEncode(rs.Fields("ThirdPeriodAwayShots").Value)%><BR></FONT></TD>
<TD BORDERCOLOR=#000000 ><FONT style=FONT-SIZE:8pt FACE="Lucida Handwriting" COLOR=#000000><%=Server.HTMLEncode(rs.Fields("12").Value)%><BR></FONT></TD>
<TD BORDERCOLOR=#000000 ><FONT style=FONT-SIZE:8pt FACE="MS Sans Serif" COLOR=#000000><%=Server.HTMLEncode(rs.Fields("LocationID").Value)%><BR></FONT></TD>
<TD BORDERCOLOR=#000000 ><FONT style=FONT-SIZE:8pt FACE="MS Sans Serif" COLOR=#000000><%=Server.HTMLEncode(rs.Fields("Notes").Value)%><BR></FONT></TD>
<TD BORDERCOLOR=#000000 ><FONT style=FONT-SIZE:8pt FACE="Lucida Handwriting" COLOR=#000000><%=Server.HTMLEncode(rs.Fields("Referee1").Value)%><BR></FONT></TD>
<TD BORDERCOLOR=#000000 ><FONT style=FONT-SIZE:8pt FACE="Lucida Handwriting" COLOR=#000000><%=Server.HTMLEncode(rs.Fields("Referee2").Value)%><BR></FONT></TD>
<TD BORDERCOLOR=#000000  ALIGN=RIGHT><FONT style=FONT-SIZE:8pt FACE="MS Sans Serif" COLOR=#000000><%=Server.HTMLEncode(rs.Fields("Attendance").Value)%><BR></FONT></TD>
<TD BORDERCOLOR=#000000  ALIGN=RIGHT><B><FONT style=FONT-SIZE:10pt FACE="System" COLOR=#000000><%=Server.HTMLEncode(rs.Fields("Scored").Value)%><BR></FONT></B></TD>
<TD BORDERCOLOR=#000000  ALIGN=RIGHT><B><FONT style=FONT-SIZE:10pt FACE="System" COLOR=#000000><%=Server.HTMLEncode(rs.Fields("AltJerseySeason").Value)%><BR></FONT></B></TD>

</TR>
<%
rs.MoveNext
loop%>
</TBODY>
<TFOOT></TFOOT>
</TABLE></BODY>
</HTML