<%@ Language=VBScript CodePage=65001 %>
<%
Response.Buffer = True
Session.Contents.RemoveAll()
Session.Abandon
Response.Redirect "login.asp?bye=1"
%>
