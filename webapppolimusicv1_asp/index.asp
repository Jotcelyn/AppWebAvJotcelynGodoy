<%@ Language=VBScript CodePage=65001 %>
<%
Response.Buffer = True
If Session("user_id") & "" <> "" Then
    Response.Redirect "home.asp"
Else
    Response.Redirect "login.asp"
End If
%>
