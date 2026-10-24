<%@ Language=VBScript CodePage=65001 %>
<% Dim PageDepth : PageDepth = "" %>
<!--#include file="includes/db.asp"-->
<!--#include file="includes/auth.asp"-->
<!--#include file="includes/admin_helpers.asp"-->
<%
RequireAdmin

Dim PageTitle, ActiveNav
PageTitle = "Administrar playlists"
ActiveNav = "admin-playlists"

Dim action, postId, errMsg, formTitle, formDate, formType, formPhoto, formUser, selSongs
Dim isEdit, openModal, editId, c, rsEdit, newId, rsNew, pdate

isEdit = False
openModal = False
editId = ""
formTitle = ""
formDate = IsoDate(Now())
formType = "0"
formPhoto = "~/playlistImages/default.jpg"
formUser = ""
selSongs = ""

If UCase(Request.ServerVariables("REQUEST_METHOD")) = "POST" Then
    action = Request.Form("action") & ""
    postId = Trim(Request.Form("id") & "")

    If action = "delete" Then
        If postId <> "" And IsNumeric(postId) Then
            On Error Resume Next
            gConn.Execute "DELETE FROM TBL_PLAYLIST_SONG WHERE ID_PLAYLIST = " & CLng(postId)
            gConn.Execute "DELETE FROM TBL_PLAYLIST WHERE ID_PLAYLIST = " & CLng(postId)
            If Err.Number <> 0 Then
                SetFlash "No se pudo eliminar la playlist: " & Err.Description, "error"
                Err.Clear
            Else
                SetFlash "Playlist eliminada correctamente.", "success"
            End If
            On Error GoTo 0
        End If
        Response.Redirect "admin-playlists.asp"
    End If

    If action = "save" Then
        formTitle = TrimText(Request.Form("title"), 50)
        formDate = Trim(Request.Form("creation_date") & "")
        formType = Request.Form("type") & ""
        formPhoto = TrimText(Request.Form("photo"), 255)
        formUser = Trim(Request.Form("id_user") & "")
        selSongs = CleanIntList(Request.Form("songs") & "")

        If formTitle = "" Then errMsg = "El título de la playlist es obligatorio."
        If formUser = "" Or Not IsNumeric(formUser) Then errMsg = "Selecciona un usuario propietario."
        If formPhoto = "" Then formPhoto = "~/playlistImages/default.jpg"
        If formType <> "1" Then formType = "0"

        pdate = ParseIsoDate(formDate)
        If IsNull(pdate) Then pdate = Now()

        If errMsg <> "" Then
            SetFlash errMsg, "error"
            If postId <> "" Then
                Response.Redirect "admin-playlists.asp?edit=" & postId
            Else
                Response.Redirect "admin-playlists.asp?new=1"
            End If
        End If

        On Error Resume Next
        If postId = "" Then
            Set c = NewCmd("INSERT INTO TBL_PLAYLIST (ID_USER, TITLE, CREATION_DATE, TYPE, PHOTO) OUTPUT INSERTED.ID_PLAYLIST AS NEW_ID VALUES (?, ?, ?, ?, ?)")
            AddInt c, "@user", CLng(formUser)
            AddStr c, "@title", formTitle
            AddDate c, "@date", pdate
            AddInt c, "@type", CLng(formType)
            AddStr c, "@photo", formPhoto
            Set rsNew = c.Execute
            If Err.Number = 0 Then
                newId = rsNew("NEW_ID")
                rsNew.Close
            End If
        Else
            newId = CLng(postId)
            Set c = NewCmd("UPDATE TBL_PLAYLIST SET ID_USER = ?, TITLE = ?, CREATION_DATE = ?, TYPE = ?, PHOTO = ? WHERE ID_PLAYLIST = ?")
            AddInt c, "@user", CLng(formUser)
            AddStr c, "@title", formTitle
            AddDate c, "@date", pdate
            AddInt c, "@type", CLng(formType)
            AddStr c, "@photo", formPhoto
            AddInt c, "@id", CLng(postId)
            c.Execute
        End If

        If Err.Number <> 0 Then
            SetFlash "No se pudo guardar la playlist: " & Err.Description, "error"
            Err.Clear
            On Error GoTo 0
            If postId <> "" Then
                Response.Redirect "admin-playlists.asp?edit=" & postId
            Else
                Response.Redirect "admin-playlists.asp?new=1"
            End If
        End If

        gConn.Execute "DELETE FROM TBL_PLAYLIST_SONG WHERE ID_PLAYLIST = " & newId
        Dim sParts, si
        If selSongs <> "" Then
            sParts = Split(selSongs, ",")
            For si = 0 To UBound(sParts)
                gConn.Execute "INSERT INTO TBL_PLAYLIST_SONG (ID_PLAYLIST, ID_SONG) VALUES (" & newId & ", " & sParts(si) & ")"
            Next
        End If
        On Error GoTo 0

        If postId = "" Then
            SetFlash "Playlist «" & formTitle & "» creada correctamente.", "success"
        Else
            SetFlash "Playlist «" & formTitle & "» actualizada correctamente.", "success"
        End If
        Response.Redirect "admin-playlists.asp"
    End If
End If

editId = Request.QueryString("edit") & ""
If editId <> "" And IsNumeric(editId) Then
    Set c = NewCmd("SELECT ID_PLAYLIST, ID_USER, TITLE, CREATION_DATE, TYPE, ISNULL(PHOTO,'') AS PHOTO FROM TBL_PLAYLIST WHERE ID_PLAYLIST = ?")
    AddInt c, "@id", CLng(editId)
    Set rsEdit = c.Execute
    If Not rsEdit.EOF Then
        isEdit = True
        openModal = True
        editId = rsEdit("ID_PLAYLIST")
        formTitle = rsEdit("TITLE") & ""
        formDate = IsoDate(rsEdit("CREATION_DATE"))
        formType = rsEdit("TYPE") & ""
        formPhoto = rsEdit("PHOTO") & ""
        formUser = rsEdit("ID_USER") & ""
        rsEdit.Close

        Dim relRs
        Set c = NewCmd("SELECT ID_SONG FROM TBL_PLAYLIST_SONG WHERE ID_PLAYLIST = ?")
        AddInt c, "@id", CLng(editId)
        Set relRs = c.Execute
        Do While Not relRs.EOF
            selSongs = selSongs & relRs(0) & ","
            relRs.MoveNext
        Loop
        relRs.Close
        Set relRs = Nothing
    Else
        rsEdit.Close
    End If
End If

If Request.QueryString("new") = "1" Then openModal = True

Dim rsPlaylists, totalPlaylists, rsUsers, rsSongs
Set rsPlaylists = SqlRs("SELECT p.ID_PLAYLIST, p.TITLE, p.TYPE, p.PHOTO, p.CREATION_DATE, p.ID_USER, " & _
    "ISNULL((SELECT u.USERNAME FROM TBL_USER u WHERE u.ID_USER = p.ID_USER), 'Sin usuario') AS OWNER, " & _
    "(SELECT COUNT(*) FROM TBL_PLAYLIST_SONG ps WHERE ps.ID_PLAYLIST = p.ID_PLAYLIST) AS SONG_COUNT " & _
    "FROM TBL_PLAYLIST p ORDER BY p.ID_PLAYLIST")
totalPlaylists = ScalarVal("SELECT COUNT(*) FROM TBL_PLAYLIST")
Set rsUsers = SqlRs("SELECT ID_USER, USERNAME FROM TBL_USER ORDER BY USERNAME")
Set rsSongs = SqlRs("SELECT ID_SONG, SONG_NAME FROM TBL_SONG ORDER BY SONG_NAME")
%>
<!--#include file="includes/layout_top.asp"-->

<div class="page-head">
    <div>
        <span class="page-head__eyebrow"><%=Icon("shield")%> Panel administrativo</span>
        <h1>Playlists</h1>
        <p>Administra las colecciones de todos los usuarios, su visibilidad y las canciones que las componen.</p>
    </div>
    <div class="page-head__actions">
        <a class="btn btn--ghost" href="home.asp"><%=Icon("arrow-left")%> Ir al inicio</a>
        <a class="btn btn--primary" href="admin-playlists.asp?new=1"><%=Icon("plus")%> Nueva playlist</a>
    </div>
</div>

<div class="admin-bar">
    <span class="admin-bar__count"><strong><%=totalPlaylists%></strong> playlists registradas</span>
    <span class="admin-bar__spacer"></span>
    <label class="mini-search">
        <%=Icon("search")%>
        <input type="search" placeholder="Filtrar playlists..." data-mini-search="#playlistsTable" aria-label="Filtrar playlists">
    </label>
</div>

<div class="table-wrap">
    <table class="song-table" id="playlistsTable">
        <thead>
            <tr>
                <th style="width:54px">ID</th>
                <th>Playlist</th>
                <th class="hide-sm">Propietario</th>
                <th>Visibilidad</th>
                <th class="hide-sm">Canciones</th>
                <th class="cell-actions">Acciones</th>
            </tr>
        </thead>
        <tbody>
        <%
        Dim pId, pTitle, pType, pPhoto, pOwner, pCount, plDate, pSearch
        Do While Not rsPlaylists.EOF
            pId = rsPlaylists("ID_PLAYLIST")
            pTitle = rsPlaylists("TITLE") & ""
            pType = rsPlaylists("TYPE") & ""
            pPhoto = ImgUrl(rsPlaylists("PHOTO") & "", PageDepth & "assets/img/default.jpg")
            pOwner = rsPlaylists("OWNER") & ""
            pCount = rsPlaylists("SONG_COUNT")
            plDate = rsPlaylists("CREATION_DATE")
            pSearch = pId & " " & pTitle & " " & pOwner
        %>
            <tr data-search-item="<%=H(pSearch)%>">
                <td class="muted" style="font-variant-numeric:tabular-nums">#<%=pId%></td>
                <td>
                    <div class="song-cell">
                        <div class="song-cover"><img src="<%=pPhoto%>" alt=""></div>
                        <div class="song-name">
                            <strong><%=H(pTitle)%></strong>
                            <span><%=PrettyDate(plDate)%></span>
                        </div>
                    </div>
                </td>
                <td class="hide-sm"><span class="tag"><%=Icon("user")%> <%=H(pOwner)%></span></td>
                <td>
                    <% If pType = "1" Then %>
                    <span class="tag tag--accent"><%=Icon("lock")%> Privada</span>
                    <% Else %>
                    <span class="tag"><%=Icon("users")%> Pública</span>
                    <% End If %>
                </td>
                <td class="hide-sm"><span class="plays"><%=Icon("music")%> <%=pCount%></span></td>
                <td class="cell-actions">
                    <div class="cell-actions__inner">
                        <a class="icon-btn icon-btn--sm" href="admin-playlists.asp?edit=<%=pId%>" title="Editar" aria-label="Editar <%=H(pTitle)%>"><%=Icon("edit")%></a>
                        <form method="post" action="admin-playlists.asp" style="display:inline"
                              data-confirm="¿Eliminar la playlist «<%=H(pTitle)%>»? Se eliminarán también sus canciones asociadas.">
                            <input type="hidden" name="action" value="delete">
                            <input type="hidden" name="id" value="<%=pId%>">
                            <button type="submit" class="icon-btn icon-btn--sm icon-btn--danger" title="Eliminar" aria-label="Eliminar <%=H(pTitle)%>"><%=Icon("trash")%></button>
                        </form>
                    </div>
                </td>
            </tr>
        <%
            rsPlaylists.MoveNext
        Loop
        rsPlaylists.Close
        Set rsPlaylists = Nothing
        %>
        </tbody>
    </table>
</div>

<div class="modal" id="playlistForm" <% If openModal Then %>data-auto-open="true"<% End If %> role="dialog" aria-modal="true" aria-label="Formulario de playlist">
    <div class="modal__backdrop" data-close-modal></div>
    <div class="modal__card">
        <form method="post" action="admin-playlists.asp">
            <div class="modal__head">
                <div>
                    <h3><% If isEdit Then %>Editar playlist<% Else %>Nueva playlist<% End If %></h3>
                    <p>Asigna un propietario, define la visibilidad y selecciona las canciones.</p>
                </div>
                <button type="button" class="icon-btn" data-close-modal aria-label="Cerrar"><%=Icon("close")%></button>
            </div>
            <div class="modal__body">
                <input type="hidden" name="action" value="save">
                <input type="hidden" name="id" value="<% If isEdit Then %><%=editId%><% End If %>">
                <div class="form-grid">
                    <div class="field field--full">
                        <label for="title">Título de la playlist *</label>
                        <input type="text" id="title" name="title" maxlength="50" required
                               value="<%=H(formTitle)%>" placeholder="Ej. Playlist 1">
                    </div>
                    <div class="field">
                        <label for="id_user">Propietario *</label>
                        <select id="id_user" name="id_user" required>
                            <option value="">Selecciona un usuario</option>
                            <%
                            Dim usId, usName
                            Do While Not rsUsers.EOF
                                usId = rsUsers("ID_USER")
                                usName = rsUsers("USERNAME") & ""
                            %>
                            <option value="<%=usId%>"<% If CStr(formUser) = CStr(usId) Then %> selected<% End If %>><%=H(usName)%></option>
                            <%
                                rsUsers.MoveNext
                            Loop
                            %>
                        </select>
                    </div>
                    <div class="field">
                        <label for="creation_date">Fecha de creación</label>
                        <input type="date" id="creation_date" name="creation_date" value="<%=H(formDate)%>">
                    </div>
                    <div class="field">
                        <label for="type">Visibilidad</label>
                        <select id="type" name="type">
                            <option value="0"<% If formType <> "1" Then %> selected<% End If %>>Pública</option>
                            <option value="1"<% If formType = "1" Then %> selected<% End If %>>Privada</option>
                        </select>
                    </div>
                    <div class="field">
                        <label for="photo">Ruta de la portada</label>
                        <input type="text" id="photo" name="photo" maxlength="255"
                               value="<%=H(formPhoto)%>" placeholder="~/playlistImages/playlist1.png">
                    </div>
                    <div class="field field--full">
                        <label>Canciones de la playlist</label>
                        <% If rsSongs.EOF Then %>
                        <span class="field__hint">No hay canciones registradas para asociar.</span>
                        <% Else %>
                        <div class="check-panel">
                            <%
                            Dim soId, soName
                            Do While Not rsSongs.EOF
                                soId = rsSongs("ID_SONG")
                                soName = rsSongs("SONG_NAME") & ""
                            %>
                            <label class="check-pill">
                                <input type="checkbox" name="songs" value="<%=soId%>"<%=SelectedAttr(selSongs, soId)%>>
                                <span><%=H(soName)%></span>
                            </label>
                            <%
                                rsSongs.MoveNext
                            Loop
                            %>
                        </div>
                        <% End If %>
                    </div>
                </div>
            </div>
            <div class="modal__foot">
                <span class="modal__foot-note">Tabla: TBL_PLAYLIST · TBL_PLAYLIST_SONG</span>
                <button type="button" class="btn btn--ghost" data-close-modal>Cancelar</button>
                <button type="submit" class="btn btn--primary"><%=Icon("check")%> Guardar playlist</button>
            </div>
        </form>
    </div>
</div>

<!--#include file="includes/layout_bottom.asp"-->
