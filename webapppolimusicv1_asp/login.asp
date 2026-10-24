<%@ Language=VBScript CodePage=65001 %>
<% Dim PageDepth : PageDepth = "" %>
<!--#include file="includes/db.asp"-->
<!--#include file="includes/icons.asp"-->
<%
Response.Buffer = True

If Session("user_id") & "" <> "" Then
    Response.Redirect "home.asp"
End If

Dim errMsg, infoMsg, loginUser, loginPass, rs, cmd, postUser
errMsg = ""
infoMsg = ""
loginUser = ""

If Request.QueryString("bye") = "1" Then
    infoMsg = "Tu sesión se cerró correctamente. ¡Hasta pronto!"
End If

If UCase(Request.ServerVariables("REQUEST_METHOD")) = "POST" Then
    postUser = Trim(Request.Form("username"))
    loginUser = postUser
    loginPass = Trim(Request.Form("password"))

    If postUser = "" Or loginPass = "" Then
        errMsg = "Ingresa tu usuario y contraseña para continuar."
    Else
        Set cmd = NewCmd("SELECT TOP 1 ID_USER, USERNAME, EMAIL, USER_PHOTO, USER_TYPE " & _
                         "FROM TBL_USER WHERE USERNAME = ? AND (PASSWORD = ? OR PASSWORD = " & _
                         "UPPER(CONVERT(varchar(40), HASHBYTES('SHA1', CAST(? AS varchar(50))), 2)))")
        AddStr cmd, "@username", postUser
        AddStr cmd, "@password", loginPass
        AddStr cmd, "@password2", loginPass
        Set rs = cmd.Execute

        If Not rs.EOF Then
            Session("user_id") = rs("ID_USER")
            Session("username") = rs("USERNAME") & ""
            Session("user_email") = rs("EMAIL") & ""
            Session("user_photo") = rs("USER_PHOTO") & ""
            Session("user_type") = rs("USER_TYPE")
            Session.Timeout = 90
            Session("flash_msg") = "Bienvenido a PoliMusic, " & (rs("USERNAME") & "") & ". Tu sesión se inició correctamente."
            Session("flash_kind") = "success"
            rs.Close
            Set rs = Nothing
            Response.Redirect "home.asp?welcome=1"
        Else
            errMsg = "Usuario o contraseña incorrectos. Verifica tus datos e inténtalo otra vez."
        End If
    End If
End If

CloseDbConnection
%>
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <meta name="theme-color" content="#0E0B16">
    <title>Iniciar sesión · <%=APP_NAME%></title>
    <link rel="icon" href="data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 32 32'%3E%3Crect width='32' height='32' rx='8' fill='%230E0B16'/%3E%3Cpath d='M5 20.5c3-7.5 5.5-7.5 8.5 0s5.5 7.5 8.5 0 4-6 5-4.5' stroke='%23B9A3F5' stroke-width='2.2' fill='none' stroke-linecap='round'/%3E%3C/svg%3E">
    <link rel="stylesheet" href="assets/css/app.css">
</head>
<body class="theme-user">
<div class="ambient" aria-hidden="true">
    <span class="ambient__orb ambient__orb--one"></span>
    <span class="ambient__orb ambient__orb--two"></span>
    <span class="ambient__grid"></span>
</div>

<div class="auth">
    <section class="auth__brand">
        <div class="auth__logo">
            <span class="brand__mark">
                <svg viewBox="0 0 32 32" fill="none" aria-hidden="true">
                    <path d="M5 20.5c3-7.5 5.5-7.5 8.5 0s5.5 7.5 8.5 0 4-6 5-4.5" stroke="currentColor" stroke-width="2.1" stroke-linecap="round"/>
                    <circle cx="5" cy="20.5" r="2" fill="currentColor"/>
                    <circle cx="27" cy="16" r="2" fill="currentColor"/>
                </svg>
            </span>
            <div>
                <strong><%=APP_NAME%></strong>
                <small>music streaming</small>
            </div>
        </div>

        <div>
            <h1 class="auth__headline">Tu biblioteca, <em>elevada</em> a otra atmósfera.</h1>
            <p class="auth__sub">
                Un catálogo musical con identidad propia: escucha tus canciones, recorre artistas,
                álbumes y playlists, y administra todo el contenido desde un panel diseñado para cada rol.
            </p>

            <div class="auth__features">
                <div class="auth__feature">
                    <%=Icon("headphones")%>
                    <span><strong>Reproductor integrado</strong> con cola, aleatorio y repetición.</span>
                </div>
                <div class="auth__feature">
                    <%=Icon("sparkles")%>
                    <span><strong>Dos experiencias visuales</strong> según tu rol: lavanda o dorado.</span>
                </div>
                <div class="auth__feature">
                    <%=Icon("shield")%>
                    <span><strong>Panel administrativo</strong> con CRUD completo del catálogo.</span>
                </div>
            </div>
        </div>

        <p class="auth__foot">Proyecto académico · ASP clásico · SQL Server · <span data-year>2026</span></p>
    </section>

    <section class="auth__panel">
        <div class="auth__card">
            <h2>Bienvenido de nuevo</h2>
            <p>Inicia sesión para entrar a tu biblioteca de música.</p>

            <% If errMsg <> "" Then %>
            <div class="auth__error" style="margin-top:20px">
                <%=Icon("alert")%>
                <span><%=H(errMsg)%></span>
            </div>
            <% End If %>

            <% If infoMsg <> "" Then %>
            <div class="auth__error" style="margin-top:20px;background:rgba(110,231,183,.1);border-color:rgba(110,231,183,.28);color:#6EE7B7">
                <%=Icon("check")%>
                <span><%=H(infoMsg)%></span>
            </div>
            <% End If %>

            <form class="auth__form" method="post" action="login.asp" autocomplete="on">
                <div class="field">
                    <label for="username">Usuario</label>
                    <input type="text" id="username" name="username" value="<%=H(loginUser)%>"
                           placeholder="Ej. userNormal" autocomplete="username" required>
                </div>
                <div class="field">
                    <label for="password">Contraseña</label>
                    <input type="password" id="password" name="password"
                           placeholder="Tu contraseña" autocomplete="current-password" required>
                </div>
                <button type="submit" class="btn btn--primary btn--block">
                    Entrar a PoliMusic
                    <%=Icon("arrow-right")%>
                </button>
            </form>

            <div class="auth__divider">Cuentas de prueba</div>

            <div class="demo-accounts">
                <button type="button" class="demo-account" data-demo-user="userNormal" data-demo-pass="12345678">
                    <span class="demo-account__dot"><%=Icon("headphones")%></span>
                    <span class="demo-account__meta">
                        <strong>userNormal</strong>
                        <span>Contraseña: 12345678</span>
                    </span>
                    <span class="demo-account__tag">Usuario</span>
                </button>
                <button type="button" class="demo-account" data-demo-user="userAdmin" data-demo-pass="12345678">
                    <span class="demo-account__dot"><%=Icon("shield")%></span>
                    <span class="demo-account__meta">
                        <strong>userAdmin</strong>
                        <span>Contraseña: 12345678</span>
                    </span>
                    <span class="demo-account__tag">Admin</span>
                </button>
            </div>

            <p class="auth__note">
                El rol se obtiene de <strong>TBL_USER.USER_TYPE</strong>: 0 = usuario, 1 = administrador.
                Solo el administrador puede crear, editar y eliminar registros.
            </p>
        </div>
    </section>
</div>

<script src="assets/js/app.js"></script>
</body>
</html>
