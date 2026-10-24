<%@ Language=VBScript CodePage=65001 %>
<!--#include file="includes/db.asp"-->
<%
Response.ContentType = "application/json"

If Session("user_id") & "" = "" Then
    Response.Status = "401 Unauthorized"
    Response.Write "{""ok"":false,""error"":""unauthorized""}"
    CloseDbConnection
    Response.End
End If

Dim songId, cmd, plays
songId = Request.Form("id") & ""

If songId = "" Or Not IsNumeric(songId) Then
    Response.Status = "400 Bad Request"
    Response.Write "{""ok"":false,""error"":""invalid_id""}"
    CloseDbConnection
    Response.End
End If

On Error Resume Next
Set cmd = NewCmd("UPDATE TBL_SONG SET PLAYS = ISNULL(PLAYS, 0) + 1 WHERE ID_SONG = ?")
AddInt cmd, "@id", CLng(songId)
cmd.Execute

If Err.Number <> 0 Then
    Response.Status = "500 Internal Server Error"
    Response.Write "{""ok"":false,""error"":""db_error""}"
    Err.Clear
    CloseDbConnection
    Response.End
End If
On Error GoTo 0

plays = ScalarVal("SELECT PLAYS FROM TBL_SONG WHERE ID_SONG = " & CLng(songId))
If IsNull(plays) Then plays = 0

Response.Write "{""ok"":true,""id"":" & CLng(songId) & ",""plays"":" & CLng(plays) & "}"
CloseDbConnection
%>
