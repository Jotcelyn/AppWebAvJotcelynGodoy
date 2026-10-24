<%@ Language=VBScript CodePage=65001 %>
<% Dim PageDepth : PageDepth = "" %>
<!--#include file="includes/db.asp"-->
<!--#include file="includes/auth.asp"-->
<!--#include file="includes/admin_helpers.asp"-->
<%
RequireAdmin

Dim PageTitle, ActiveNav
PageTitle = "Administrar usuarios"
ActiveNav = "admin-users"

Dim action, postId, errMsg, formUser, formPass, formEmail, formBirthday, formPhoto, formType
Dim isEdit, openModal, editId, c, rsEdit, newId, rsNew, bday

isEdit = False
openModal = False
editId = ""
formUser = ""
formPass = ""
formEmail = ""
formBirthday = ""
formPhoto = "~/userImages/default.jpg"
formType = "0"

If UCase(Request.ServerVariables("REQUEST_METHOD")) = "POST" Then
    action = Request.Form("action") & ""
    postId = Trim(Request.Form("id") & "")

    If action = "delete" Then
        If postId <> "" And IsNumeric(postId) Then
            If CLng(postId) = CLng(Session("user_id")) Then
                SetFlash "No puedes eliminar la cuenta con la que estás conectado.", "error"
            Else
                Dim adminsLeft
                adminsLeft = ScalarVal("SELECT COUNT(*) FROM TBL_USER WHERE USER_TYPE = 1")
                If ScalarVal("SELECT USER_TYPE FROM TBL_USER WHERE ID_USER = " & CLng(postId)) = 1 And CLng(adminsLeft) <= 1 Then
                    SetFlash "Debe existir al menos un usuario administrador en el sistema.", "error"
                Else
                    On Error Resume Next
                    gConn.Execute "DELETE FROM TBL_PLAYLIST_SONG WHERE ID_PLAYLIST IN (SELECT ID_PLAYLIST FROM TBL_PLAYLIST WHERE ID_USER = " & CLng(postId) & ")"
                    gConn.Execute "DELETE FROM TBL_PLAYLIST WHERE ID_USER = " & CLng(postId)
                    gConn.Execute "DELETE FROM TBL_USER WHERE ID_USER = " & CLng(postId)
                    If Err.Number <> 0 Then
                        SetFlash "No se pudo eliminar el usuario: " & Err.Description, "error"
                        Err.Clear
                    Else
                        SetFlash "Usuario eliminado correctamente.", "success"
                    End If
                    On Error GoTo 0
                End If
            End If
        End If
        Response.Redirect "admin-users.asp"
    End If

    If action = "save" Then
        formUser = TrimText(Request.Form("username"), 50)
        formPass = Trim(Request.Form("password") & "")
        formEmail = TrimText(Request.Form("email"), 50)
        formBirthday = Trim(Request.Form("birthday") & "")
        formPhoto = TrimText(Request.Form("user_photo"), 255)
        formType = Request.Form("user_type") & ""

        If formUser = "" Then
            errMsg = "El nombre de usuario es obligatorio."
        ElseIf formEmail = "" Then
            errMsg = "El correo electrónico es obligatorio."
        ElseIf postId = "" And formPass = "" Then
            errMsg = "La contraseña es obligatoria para un usuario nuevo."
        ElseIf InStr(formEmail, "@") = 0 Then
            errMsg = "Ingresa un correo electrónico válido."
        End If

        If formPhoto = "" Then formPhoto = "~/userImages/default.jpg"
        If formType <> "1" Then formType = "0"
        bday = ParseIsoDate(formBirthday)

        If errMsg <> "" Then
            SetFlash errMsg, "error"
            If postId <> "" Then
                Response.Redirect "admin-users.asp?edit=" & postId
            Else
                Response.Redirect "admin-users.asp?new=1"
            End If
        End If

        On Error Resume Next
        If postId = "" Then
            Set c = NewCmd("INSERT INTO TBL_USER (USERNAME, PASSWORD, EMAIL, BIRTHDAY, USER_PHOTO, USER_TYPE) " & _
                           "OUTPUT INSERTED.ID_USER AS NEW_ID VALUES (?, ?, ?, ?, ?, ?)")
            AddStr c, "@username", formUser
            AddStr c, "@password", formPass
            AddStr c, "@email", formEmail
            AddDateOrNull c, "@birthday", bday
            AddStr c, "@photo", formPhoto
            AddInt c, "@type", CLng(formType)
            Set rsNew = c.Execute
            If Err.Number = 0 Then
                newId = rsNew("NEW_ID")
                rsNew.Close
            End If
        Else
            If formPass <> "" Then
                Set c = NewCmd("UPDATE TBL_USER SET USERNAME = ?, PASSWORD = ?, EMAIL = ?, BIRTHDAY = ?, USER_PHOTO = ?, USER_TYPE = ? WHERE ID_USER = ?")
                AddStr c, "@username", formUser
                AddStr c, "@password", formPass
                AddStr c, "@email", formEmail
                AddDateOrNull c, "@birthday", bday
                AddStr c, "@photo", formPhoto
                AddInt c, "@type", CLng(formType)
                AddInt c, "@id", CLng(postId)
                c.Execute
            Else
                Set c = NewCmd("UPDATE TBL_USER SET USERNAME = ?, EMAIL = ?, BIRTHDAY = ?, USER_PHOTO = ?, USER_TYPE = ? WHERE ID_USER = ?")
                AddStr c, "@username", formUser
                AddStr c, "@email", formEmail
                AddDateOrNull c, "@birthday", bday
                AddStr c, "@photo", formPhoto
                AddInt c, "@type", CLng(formType)
                AddInt c, "@id", CLng(postId)
                c.Execute
            End If
        End If

        If Err.Number <> 0 Then
            SetFlash "No se pudo guardar el usuario: " & Err.Description, "error"
            Err.Clear
            On Error GoTo 0
            If postId <> "" Then
                Response.Redirect "admin-users.asp?edit=" & postId
            Else
                Response.Redirect "admin-users.asp?new=1"
            End If
        End If
        On Error GoTo 0

        If postId = "" Then
            SetFlash "Usuario «" & formUser & "» creado correctamente.", "success"
        Else
            SetFlash "Usuario «" & formUser & "» actualizado correctamente.", "success"
        End If
        Response.Redirect "admin-users.asp"
    End If
End If

' ---------- Edición ----------
editId = Request.QueryString("edit") & ""
If editId <> "" And IsNumeric(editId) Then
    Set c = NewCmd("SELECT ID_USER, USERNAME, PASSWORD, EMAIL, BIRTHDAY, USER_PHOTO, USER_TYPE FROM TBL_USER WHERE ID_USER = ?")
    AddInt c, "@id", CLng(editId)
    Set rsEdit = c.Execute
    If Not rsEdit.EOF Then
        isEdit = True
        openModal = True
        editId = rsEdit("ID_USER")
        formUser = rsEdit("USERNAME") & ""
        formEmail = rsEdit("EMAIL") & ""
        formBirthday = IsoDate(rsEdit("BIRTHDAY"))
        formPhoto = rsEdit("USER_PHOTO") & ""
        formType = rsEdit("USER_TYPE") & ""
        rsEdit.Close
    Else
        rsEdit.Close
    End If
End If

If Request.QueryString("new") = "1" Then openModal = True

Dim rsUsers, totalUsers
Set rsUsers = SqlRs("SELECT ID_USER, USERNAME, EMAIL, BIRTHDAY, USER_PHOTO, USER_TYPE FROM TBL_USER ORDER BY ID_USER")
totalUsers = ScalarVal("SELECT COUNT(*) FROM TBL_USER")
%>
<!--#include file="includes/layout_top.asp"-->

<div class="page-head">
    <div>
        <span class="page-head__eyebrow"><%=Icon("shield")%> Panel administrativo</span>
        <h1>Usuarios</h1>
        <p>Gestiona las cuentas del sistema y su rol. El tipo 0 corresponde a usuario normal y el tipo 1 a administrador.</p>
    </div>
    <div class="page-head__actions">
        <a class="btn btn--ghost" href="home.asp"><%=Icon("arrow-left")%> Ir al inicio</a>
        <a class="btn btn--primary" href="admin-users.asp?new=1"><%=Icon("plus")%> Nuevo usuario</a>
    </div>
</div>

<div class="admin-bar">
    <span class="admin-bar__count"><strong><%=totalUsers%></strong> usuarios registrados</span>
    <span class="admin-bar__spacer"></span>
    <label class="mini-search">
        <%=Icon("search")%>
        <input type="search" placeholder="Filtrar usuarios..." data-mini-search="#usersTable" aria-label="Filtrar usuarios">
    </label>
</div>

<div class="table-wrap">
    <table class="song-table" id="usersTable">
        <thead>
            <tr>
                <th style="width:54px">ID</th>
                <th>Usuario</th>
                <th class="hide-sm">Correo</th>
                <th class="hide-sm">Cumpleaños</th>
                <th>Rol</th>
                <th class="cell-actions">Acciones</th>
            </tr>
        </thead>
        <tbody>
        <%
        Dim uId, uName, uEmail, uBirth, uPhoto, uType, uAvatar, uSearch
        Do While Not rsUsers.EOF
            uId = rsUsers("ID_USER")
            uName = rsUsers("USERNAME") & ""
            uEmail = rsUsers("EMAIL") & ""
            uBirth = rsUsers("BIRTHDAY")
            uType = rsUsers("USER_TYPE") & ""
            uPhoto = ImgUrl(rsUsers("USER_PHOTO") & "", PageDepth & "assets/img/default.jpg")
            uSearch = uId & " " & uName & " " & uEmail
        %>
            <tr data-search-item="<%=H(uSearch)%>">
                <td class="muted" style="font-variant-numeric:tabular-nums">#<%=uId%></td>
                <td>
                    <div class="song-cell">
                        <img class="avatar" src="<%=uPhoto%>" alt="" style="width:38px;height:38px">
                        <div class="song-name">
                            <strong><%=H(uName)%></strong>
                            <span><% If CStr(uId) = CStr(Session("user_id")) Then %>Tu cuenta<% Else %>Cuenta activa<% End If %></span>
                        </div>
                    </div>
                </td>
                <td class="hide-sm"><%=H(uEmail)%></td>
                <td class="hide-sm"><%=PrettyDate(uBirth)%></td>
                <td>
                    <% If uType = "1" Then %>
                    <span class="tag tag--accent"><%=Icon("shield")%> Administrador</span>
                    <% Else %>
                    <span class="tag"><%=Icon("headphones")%> Usuario</span>
                    <% End If %>
                </td>
                <td class="cell-actions">
                    <div class="cell-actions__inner">
                        <a class="icon-btn icon-btn--sm" href="admin-users.asp?edit=<%=uId%>" title="Editar" aria-label="Editar <%=H(uName)%>"><%=Icon("edit")%></a>
                        <% If CStr(uId) <> CStr(Session("user_id")) Then %>
                        <form method="post" action="admin-users.asp" style="display:inline"
                              data-confirm="¿Eliminar al usuario «<%=H(uName)%>»? También se eliminarán sus playlists.">
                            <input type="hidden" name="action" value="delete">
                            <input type="hidden" name="id" value="<%=uId%>">
                            <button type="submit" class="icon-btn icon-btn--sm icon-btn--danger" title="Eliminar" aria-label="Eliminar <%=H(uName)%>"><%=Icon("trash")%></button>
                        </form>
                        <% Else %>
                        <span class="icon-btn icon-btn--sm" title="No puedes eliminar tu propia cuenta" style="opacity:.3;cursor:not-allowed"><%=Icon("trash")%></span>
                        <% End If %>
                    </div>
                </td>
            </tr>
        <%
            rsUsers.MoveNext
        Loop
        rsUsers.Close
        Set rsUsers = Nothing
        %>
        </tbody>
    </table>
</div>

<div class="modal" id="userForm" <% If openModal Then %>data-auto-open="true"<% End If %> role="dialog" aria-modal="true" aria-label="Formulario de usuario">
    <div class="modal__backdrop" data-close-modal></div>
    <div class="modal__card">
        <form method="post" action="admin-users.asp">
            <div class="modal__head">
                <div>
                    <h3><% If isEdit Then %>Editar usuario<% Else %>Nuevo usuario<% End If %></h3>
                    <p>Define las credenciales de acceso y el rol dentro del sistema.</p>
                </div>
                <button type="button" class="icon-btn" data-close-modal aria-label="Cerrar"><%=Icon("close")%></button>
            </div>
            <div class="modal__body">
                <input type="hidden" name="action" value="save">
                <input type="hidden" name="id" value="<% If isEdit Then %><%=editId%><% End If %>">
                <div class="form-grid">
                    <div class="field">
                        <label for="username">Usuario *</label>
                        <input type="text" id="username" name="username" maxlength="50" required
                               value="<%=H(formUser)%>" placeholder="Ej. usuario.nuevo">
                    </div>
                    <div class="field">
                        <label for="password">Contraseña <% If Not isEdit Then %>*<% End If %></label>
                        <input type="password" id="password" name="password" maxlength="50"
                               placeholder="<% If isEdit Then %>Dejar en blanco para conservar<% Else %>Mínimo recomendado: 8 caracteres<% End If %>">
                    </div>
                    <div class="field field--full">
                        <label for="email">Correo electrónico *</label>
                        <input type="email" id="email" name="email" maxlength="50" required
                               value="<%=H(formEmail)%>" placeholder="usuario@epn.edu.ec">
                    </div>
                    <div class="field">
                        <label for="birthday">Fecha de nacimiento</label>
                        <input type="date" id="birthday" name="birthday" value="<%=H(formBirthday)%>">
                    </div>
                    <div class="field">
                        <label for="user_type">Rol del usuario</label>
                        <select id="user_type" name="user_type">
                            <option value="0"<% If formType <> "1" Then %> selected<% End If %>>Usuario normal</option>
                            <option value="1"<% If formType = "1" Then %> selected<% End If %>>Administrador</option>
                        </select>
                    </div>
                    <div class="field field--full">
                        <label for="user_photo">Ruta de la foto de perfil</label>
                        <input type="text" id="user_photo" name="user_photo" maxlength="255"
                               value="<%=H(formPhoto)%>" placeholder="~/userImages/default.jpg">
                        <span class="field__hint">Archivos disponibles: <code>default.jpg</code> y <code>root.png</code> dentro de <code>/assets/img</code>.</span>
                    </div>
                </div>
            </div>
            <div class="modal__foot">
                <span class="modal__foot-note">Tabla: TBL_USER</span>
                <button type="button" class="btn btn--ghost" data-close-modal>Cancelar</button>
                <button type="submit" class="btn btn--primary"><%=Icon("check")%> Guardar usuario</button>
            </div>
        </form>
    </div>
</div>

<!--#include file="includes/layout_bottom.asp"-->
