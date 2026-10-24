<%
' ============================================================
'  PoliMusic - Sesion, roles y mensajes flash
' ============================================================
Sub SetFlash(msg, kind)
    Session("flash_msg") = msg
    Session("flash_kind") = kind
End Sub

Function IsLoggedIn()
    IsLoggedIn = (Session("user_id") & "" <> "")
End Function

Function IsAdmin()
    IsAdmin = (Session("user_type") & "" = "1")
End Function

Sub RequireLogin()
    If Not IsLoggedIn() Then
        Response.Redirect PageDepth & "login.asp"
    End If
End Sub

Sub RequireAdmin()
    RequireLogin
    If Not IsAdmin() Then
        SetFlash "Acceso restringido: solo el usuario administrador puede gestionar el catalogo.", "error"
        Response.Redirect PageDepth & "home.asp"
    End If
End Sub

Function RoleLabel()
    If IsAdmin() Then
        RoleLabel = "Administrador"
    Else
        RoleLabel = "Usuario"
    End If
End Function
%>
