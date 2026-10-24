<%@ Language=VBScript CodePage=65001 %>
<% Dim PageDepth : PageDepth = "" %>
<!--#include file="includes/db.asp"-->
<!--#include file="includes/auth.asp"-->
<!--#include file="includes/admin_helpers.asp"-->
<%
RequireAdmin

Dim PageTitle, ActiveNav
PageTitle = "Administrar álbumes"
ActiveNav = "admin-albums"

Dim action, postId, errMsg, formTitle, formYear, formPhoto, selSongs
Dim isEdit, openModal, editId, c, rsEdit, newId, rsNew

isEdit = False
openModal = False
editId = ""
formTitle = ""
formYear = ""
formPhoto = ""
selSongs = ""

If UCase(Request.ServerVariables("REQUEST_METHOD")) = "POST" Then
    action = Request.Form("action") & ""
    postId = Trim(Request.Form("id") & "")

    If action = "delete" Then
        If postId <> "" And IsNumeric(postId) Then
            On Error Resume Next
            gConn.Execute "DELETE FROM TBL_TRACKLIST WHERE ID_ALBUM = " & CLng(postId)
            gConn.Execute "DELETE FROM TBL_ALBUM WHERE ID_ALBUM = " & CLng(postId)
            If Err.Number <> 0 Then
                SetFlash "No se pudo eliminar el álbum: " & Err.Description, "error"
                Err.Clear
            Else
                SetFlash "Álbum eliminado correctamente.", "success"
            End If
            On Error GoTo 0
        End If
        Response.Redirect "admin-albums.asp"
    End If

    If action = "save" Then
        formTitle = TrimText(Request.Form("title_album"), 50)
        formYear = Trim(Request.Form("release_year") & "")
        formPhoto = TrimText(Request.Form("album_image"), 255)
        selSongs = CleanIntList(Request.Form("songs") & "")

        If formTitle = "" Then errMsg = "El título del álbum es obligatorio."
        If formYear <> "" And Not IsNumeric(formYear) Then errMsg = "El año de lanzamiento debe ser numérico."

        If errMsg <> "" Then
            SetFlash errMsg, "error"
            If postId <> "" Then
                Response.Redirect "admin-albums.asp?edit=" & postId
            Else
                Response.Redirect "admin-albums.asp?new=1"
            End If
        End If

        On Error Resume Next
        If postId = "" Then
            Set c = NewCmd("INSERT INTO TBL_ALBUM (TITLE_ALBUM, RELEASE_YEAR, ALBUM_IMAGE_PATH) OUTPUT INSERTED.ID_ALBUM AS NEW_ID VALUES (?, ?, ?)")
            AddStr c, "@title", formTitle
            AddIntOrNull c, "@year", formYear
            AddStr c, "@image", formPhoto
            Set rsNew = c.Execute
            If Err.Number = 0 Then
                newId = rsNew("NEW_ID")
                rsNew.Close
            End If
        Else
            newId = CLng(postId)
            Set c = NewCmd("UPDATE TBL_ALBUM SET TITLE_ALBUM = ?, RELEASE_YEAR = ?, ALBUM_IMAGE_PATH = ? WHERE ID_ALBUM = ?")
            AddStr c, "@title", formTitle
            AddIntOrNull c, "@year", formYear
            AddStr c, "@image", formPhoto
            AddInt c, "@id", CLng(postId)
            c.Execute
        End If

        If Err.Number <> 0 Then
            SetFlash "No se pudo guardar el álbum: " & Err.Description, "error"
            Err.Clear
            On Error GoTo 0
            If postId <> "" Then
                Response.Redirect "admin-albums.asp?edit=" & postId
            Else
                Response.Redirect "admin-albums.asp?new=1"
            End If
        End If

        gConn.Execute "DELETE FROM TBL_TRACKLIST WHERE ID_ALBUM = " & newId
        Dim sParts, si
        If selSongs <> "" Then
            sParts = Split(selSongs, ",")
            For si = 0 To UBound(sParts)
                gConn.Execute "INSERT INTO TBL_TRACKLIST (ID_ALBUM, ID_SONG) VALUES (" & newId & ", " & sParts(si) & ")"
            Next
        End If
        On Error GoTo 0

        If postId = "" Then
            SetFlash "Álbum «" & formTitle & "» creado correctamente.", "success"
        Else
            SetFlash "Álbum «" & formTitle & "» actualizado correctamente.", "success"
        End If
        Response.Redirect "admin-albums.asp"
    End If
End If

editId = Request.QueryString("edit") & ""
If editId <> "" And IsNumeric(editId) Then
    Set c = NewCmd("SELECT ID_ALBUM, TITLE_ALBUM, RELEASE_YEAR, ISNULL(ALBUM_IMAGE_PATH,'') AS ALBUM_IMAGE_PATH FROM TBL_ALBUM WHERE ID_ALBUM = ?")
    AddInt c, "@id", CLng(editId)
    Set rsEdit = c.Execute
    If Not rsEdit.EOF Then
        isEdit = True
        openModal = True
        editId = rsEdit("ID_ALBUM")
        formTitle = rsEdit("TITLE_ALBUM") & ""
        formYear = rsEdit("RELEASE_YEAR") & ""
        formPhoto = rsEdit("ALBUM_IMAGE_PATH") & ""
        rsEdit.Close

        Dim relRs
        Set c = NewCmd("SELECT ID_SONG FROM TBL_TRACKLIST WHERE ID_ALBUM = ?")
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

Dim rsAlbums, totalAlbums, rsSongs
Set rsAlbums = SqlRs("SELECT al.ID_ALBUM, al.TITLE_ALBUM, al.RELEASE_YEAR, al.ALBUM_IMAGE_PATH, " & _
    "(SELECT COUNT(*) FROM TBL_TRACKLIST t WHERE t.ID_ALBUM = al.ID_ALBUM) AS SONG_COUNT " & _
    "FROM TBL_ALBUM al ORDER BY al.RELEASE_YEAR DESC, al.TITLE_ALBUM")
totalAlbums = ScalarVal("SELECT COUNT(*) FROM TBL_ALBUM")
Set rsSongs = SqlRs("SELECT ID_SONG, SONG_NAME FROM TBL_SONG ORDER BY SONG_NAME")
%>
<!--#include file="includes/layout_top.asp"-->

<div class="page-head">
    <div>
        <span class="page-head__eyebrow"><%=Icon("shield")%> Panel administrativo</span>
        <h1>Álbumes</h1>
        <p>Crea colecciones musicales y define su tracklist. La relación entre álbum y canción vive en TBL_TRACKLIST.</p>
    </div>
    <div class="page-head__actions">
        <a class="btn btn--ghost" href="home.asp"><%=Icon("arrow-left")%> Ir al inicio</a>
        <a class="btn btn--primary" href="admin-albums.asp?new=1"><%=Icon("plus")%> Nuevo álbum</a>
    </div>
</div>

<div class="admin-bar">
    <span class="admin-bar__count"><strong><%=totalAlbums%></strong> álbumes registrados</span>
    <span class="admin-bar__spacer"></span>
    <label class="mini-search">
        <%=Icon("search")%>
        <input type="search" placeholder="Filtrar álbumes..." data-mini-search="#albumsTable" aria-label="Filtrar álbumes">
    </label>
</div>

<div class="table-wrap">
    <table class="song-table" id="albumsTable">
        <thead>
            <tr>
                <th style="width:54px">ID</th>
                <th>Álbum</th>
                <th class="hide-sm">Año</th>
                <th>Canciones</th>
                <th class="cell-actions">Acciones</th>
            </tr>
        </thead>
        <tbody>
        <%
        Dim bId, bTitle, bYear, bPhoto, bCount, bSearch
        Do While Not rsAlbums.EOF
            bId = rsAlbums("ID_ALBUM")
            bTitle = rsAlbums("TITLE_ALBUM") & ""
            bYear = rsAlbums("RELEASE_YEAR") & ""
            bPhoto = ImgUrl(rsAlbums("ALBUM_IMAGE_PATH") & "", PageDepth & "assets/img/default.jpg")
            bCount = rsAlbums("SONG_COUNT")
            If bYear = "" Then bYear = "s/f"
            bSearch = bId & " " & bTitle & " " & bYear
        %>
            <tr data-search-item="<%=H(bSearch)%>">
                <td class="muted" style="font-variant-numeric:tabular-nums">#<%=bId%></td>
                <td>
                    <div class="song-cell">
                        <div class="song-cover"><img src="<%=bPhoto%>" alt=""></div>
                        <div class="song-name">
                            <strong><%=H(bTitle)%></strong>
                            <span><%=H(rsAlbums("ALBUM_IMAGE_PATH") & "")%></span>
                        </div>
                    </div>
                </td>
                <td class="hide-sm"><span class="tag"><%=Icon("calendar")%> <%=H(bYear)%></span></td>
                <td><span class="plays"><%=Icon("music")%> <%=bCount%></span></td>
                <td class="cell-actions">
                    <div class="cell-actions__inner">
                        <a class="icon-btn icon-btn--sm" href="admin-albums.asp?edit=<%=bId%>" title="Editar" aria-label="Editar <%=H(bTitle)%>"><%=Icon("edit")%></a>
                        <form method="post" action="admin-albums.asp" style="display:inline"
                              data-confirm="¿Eliminar el álbum «<%=H(bTitle)%>»? Se eliminará su tracklist asociado.">
                            <input type="hidden" name="action" value="delete">
                            <input type="hidden" name="id" value="<%=bId%>">
                            <button type="submit" class="icon-btn icon-btn--sm icon-btn--danger" title="Eliminar" aria-label="Eliminar <%=H(bTitle)%>"><%=Icon("trash")%></button>
                        </form>
                    </div>
                </td>
            </tr>
        <%
            rsAlbums.MoveNext
        Loop
        rsAlbums.Close
        Set rsAlbums = Nothing
        %>
        </tbody>
    </table>
</div>

<div class="modal" id="albumForm" <% If openModal Then %>data-auto-open="true"<% End If %> role="dialog" aria-modal="true" aria-label="Formulario de álbum">
    <div class="modal__backdrop" data-close-modal></div>
    <div class="modal__card">
        <form method="post" action="admin-albums.asp">
            <div class="modal__head">
                <div>
                    <h3><% If isEdit Then %>Editar álbum<% Else %>Nuevo álbum<% End If %></h3>
                    <p>Define los datos del álbum y marca las canciones que forman su tracklist.</p>
                </div>
                <button type="button" class="icon-btn" data-close-modal aria-label="Cerrar"><%=Icon("close")%></button>
            </div>
            <div class="modal__body">
                <input type="hidden" name="action" value="save">
                <input type="hidden" name="id" value="<% If isEdit Then %><%=editId%><% End If %>">
                <div class="form-grid">
                    <div class="field field--full">
                        <label for="title_album">Título del álbum *</label>
                        <input type="text" id="title_album" name="title_album" maxlength="50" required
                               value="<%=H(formTitle)%>" placeholder="Ej. Rock Album">
                    </div>
                    <div class="field">
                        <label for="release_year">Año de lanzamiento</label>
                        <input type="number" id="release_year" name="release_year" min="1900" max="2100" step="1"
                               value="<%=H(formYear)%>" placeholder="Ej. 1976">
                    </div>
                    <div class="field">
                        <label for="album_image">Ruta de la portada</label>
                        <input type="text" id="album_image" name="album_image" maxlength="255"
                               value="<%=H(formPhoto)%>" placeholder="~/albumImages/album1.png">
                        <span class="field__hint">Se resuelve dentro de <code>/assets/img</code>.</span>
                    </div>
                    <div class="field field--full">
                        <label>Tracklist</label>
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
                <span class="modal__foot-note">Tabla: TBL_ALBUM · TBL_TRACKLIST</span>
                <button type="button" class="btn btn--ghost" data-close-modal>Cancelar</button>
                <button type="submit" class="btn btn--primary"><%=Icon("check")%> Guardar álbum</button>
            </div>
        </form>
    </div>
</div>

<!--#include file="includes/layout_bottom.asp"-->
