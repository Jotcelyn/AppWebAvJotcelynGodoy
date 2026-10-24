<%@ Language=VBScript CodePage=65001 %>
<% Dim PageDepth : PageDepth = "" %>
<!--#include file="includes/db.asp"-->
<!--#include file="includes/auth.asp"-->
<!--#include file="includes/queries.asp"-->
<!--#include file="includes/admin_helpers.asp"-->
<%
RequireAdmin

Dim PageTitle, ActiveNav
PageTitle = "Administrar canciones"
ActiveNav = "admin-songs"

Dim action, postId, errMsg, formName, formPath, formPlays, selArtists, selAlbums
Dim isEdit, openModal, editId, c, rsEdit, newId, rsNew

isEdit = False
openModal = False
editId = ""
formName = ""
formPath = "~/songFiles/bensound-adventure.mp3"
formPlays = "0"
selArtists = ""
selAlbums = ""

If UCase(Request.ServerVariables("REQUEST_METHOD")) = "POST" Then
    action = Request.Form("action") & ""
    postId = Trim(Request.Form("id") & "")

    If action = "delete" Then
        If postId <> "" And IsNumeric(postId) Then
            On Error Resume Next
            gConn.Execute "DELETE FROM TBL_PLAYER WHERE ID_SONG = " & CLng(postId)
            gConn.Execute "DELETE FROM TBL_TRACKLIST WHERE ID_SONG = " & CLng(postId)
            gConn.Execute "DELETE FROM TBL_PLAYLIST_SONG WHERE ID_SONG = " & CLng(postId)
            gConn.Execute "DELETE FROM TBL_SONG WHERE ID_SONG = " & CLng(postId)
            If Err.Number <> 0 Then
                SetFlash "No se pudo eliminar la canción: " & Err.Description, "error"
                Err.Clear
            Else
                SetFlash "Canción eliminada correctamente del catálogo.", "success"
            End If
            On Error GoTo 0
        End If
        Response.Redirect "admin-songs.asp"
    End If

    If action = "save" Then
        formName = TrimText(Request.Form("song_name"), 80)
        formPath = TrimText(Request.Form("song_path"), 255)
        formPlays = Request.Form("plays") & ""
        selArtists = CleanIntList(Request.Form("artists") & "")
        selAlbums = CleanIntList(Request.Form("albums") & "")

        If formName = "" Then
            errMsg = "El nombre de la canción es obligatorio."
        ElseIf Len(formName) > 50 Then
            errMsg = "El nombre no puede superar los 50 caracteres."
        End If

        If formPath = "" Then formPath = "~/songFiles/bensound-adventure.mp3"

        If errMsg <> "" Then
            SetFlash errMsg, "error"
            If postId <> "" Then
                Response.Redirect "admin-songs.asp?edit=" & postId
            Else
                Response.Redirect "admin-songs.asp?new=1"
            End If
        End If

        On Error Resume Next
        If postId = "" Then
            Set c = NewCmd("INSERT INTO TBL_SONG (SONG_NAME, SONG_PATH, PLAYS) OUTPUT INSERTED.ID_SONG AS NEW_ID VALUES (?, ?, ?)")
            AddStr c, "@name", formName
            AddStr c, "@path", formPath
            AddInt c, "@plays", NumOrZero(formPlays)
            Set rsNew = c.Execute
            If Err.Number = 0 Then
                newId = rsNew("NEW_ID")
                rsNew.Close
            End If
        Else
            newId = CLng(postId)
            Set c = NewCmd("UPDATE TBL_SONG SET SONG_NAME = ?, SONG_PATH = ?, PLAYS = ? WHERE ID_SONG = ?")
            AddStr c, "@name", formName
            AddStr c, "@path", formPath
            AddInt c, "@plays", NumOrZero(formPlays)
            AddInt c, "@id", CLng(postId)
            c.Execute
        End If

        If Err.Number <> 0 Then
            SetFlash "No se pudo guardar la canción: " & Err.Description, "error"
            Err.Clear
            On Error GoTo 0
            If postId <> "" Then
                Response.Redirect "admin-songs.asp?edit=" & postId
            Else
                Response.Redirect "admin-songs.asp?new=1"
            End If
        End If

        ' Sincronizar artistas
        gConn.Execute "DELETE FROM TBL_PLAYER WHERE ID_SONG = " & newId
        Dim aParts, ai
        If selArtists <> "" Then
            aParts = Split(selArtists, ",")
            For ai = 0 To UBound(aParts)
                gConn.Execute "INSERT INTO TBL_PLAYER (ID_SONG, ID_ARTIST) VALUES (" & newId & ", " & aParts(ai) & ")"
            Next
        End If

        ' Sincronizar álbumes (tracklist)
        gConn.Execute "DELETE FROM TBL_TRACKLIST WHERE ID_SONG = " & newId
        Dim bParts, bi
        If selAlbums <> "" Then
            bParts = Split(selAlbums, ",")
            For bi = 0 To UBound(bParts)
                gConn.Execute "INSERT INTO TBL_TRACKLIST (ID_ALBUM, ID_SONG) VALUES (" & bParts(bi) & ", " & newId & ")"
            Next
        End If
        On Error GoTo 0

        If postId = "" Then
            SetFlash "Canción «" & formName & "» creada correctamente.", "success"
        Else
            SetFlash "Canción «" & formName & "» actualizada correctamente.", "success"
        End If
        Response.Redirect "admin-songs.asp"
    End If
End If

' ---------- Cargar registro a editar ----------
editId = Request.QueryString("edit") & ""
If editId <> "" And IsNumeric(editId) Then
    Set c = NewCmd("SELECT ID_SONG, SONG_NAME, SONG_PATH, ISNULL(PLAYS, 0) AS PLAYS FROM TBL_SONG WHERE ID_SONG = ?")
    AddInt c, "@id", CLng(editId)
    Set rsEdit = c.Execute
    If Not rsEdit.EOF Then
        isEdit = True
        openModal = True
        editId = rsEdit("ID_SONG")
        formName = rsEdit("SONG_NAME") & ""
        formPath = rsEdit("SONG_PATH") & ""
        formPlays = rsEdit("PLAYS") & ""
        rsEdit.Close

        Dim relRs
        Set c = NewCmd("SELECT ID_ARTIST FROM TBL_PLAYER WHERE ID_SONG = ?")
        AddInt c, "@id", CLng(editId)
        Set relRs = c.Execute
        Do While Not relRs.EOF
            selArtists = selArtists & relRs(0) & ","
            relRs.MoveNext
        Loop
        relRs.Close

        Set c = NewCmd("SELECT ID_ALBUM FROM TBL_TRACKLIST WHERE ID_SONG = ?")
        AddInt c, "@id", CLng(editId)
        Set relRs = c.Execute
        Do While Not relRs.EOF
            selAlbums = selAlbums & relRs(0) & ","
            relRs.MoveNext
        Loop
        relRs.Close
        Set relRs = Nothing
    Else
        rsEdit.Close
    End If
End If

If Request.QueryString("new") = "1" Then openModal = True

Dim rsList, totalSongs, rsArtists, rsAlbums
Set rsList = SqlRs(SongSelectBase() & "FROM TBL_SONG s ORDER BY s.ID_SONG")
totalSongs = ScalarVal("SELECT COUNT(*) FROM TBL_SONG")
Set rsArtists = SqlRs("SELECT ID_ARTIST, ARTIST_NAME, ISNULL(ARTIST_LASTNAME,'') AS ARTIST_LASTNAME FROM TBL_ARTIST ORDER BY ARTIST_NAME")
Set rsAlbums = SqlRs("SELECT ID_ALBUM, TITLE_ALBUM, ISNULL(RELEASE_YEAR,0) AS RELEASE_YEAR FROM TBL_ALBUM ORDER BY TITLE_ALBUM")
%>
<!--#include file="includes/layout_top.asp"-->

<div class="page-head">
    <div>
        <span class="page-head__eyebrow"><%=Icon("shield")%> Panel administrativo</span>
        <h1>Canciones</h1>
        <p>Crea, edita y elimina canciones del catálogo. También puedes asignar artistas y álbumes a cada tema.</p>
    </div>
    <div class="page-head__actions">
        <a class="btn btn--ghost" href="home.asp"><%=Icon("arrow-left")%> Ir al inicio</a>
        <a class="btn btn--primary" href="admin-songs.asp?new=1"><%=Icon("plus")%> Nueva canción</a>
    </div>
</div>

<div class="admin-bar">
    <span class="admin-bar__count"><strong><%=totalSongs%></strong> canciones registradas</span>
    <span class="admin-bar__spacer"></span>
    <label class="mini-search">
        <%=Icon("search")%>
        <input type="search" placeholder="Filtrar canciones..." data-mini-search="#songsTable" aria-label="Filtrar canciones">
    </label>
</div>

<% If totalSongs = 0 Then %>
<div class="empty">
    <div class="empty__ico"><%=Icon("music")%></div>
    <strong>El catálogo está vacío</strong>
    <p>Crea la primera canción para comenzar a construir la biblioteca.</p>
    <a class="btn btn--primary" href="admin-songs.asp?new=1"><%=Icon("plus")%> Nueva canción</a>
</div>
<% Else %>
<div class="table-wrap">
    <table class="song-table" id="songsTable">
        <thead>
            <tr>
                <th style="width:54px">ID</th>
                <th>Canción</th>
                <th class="hide-sm">Artistas</th>
                <th class="hide-sm">Álbumes</th>
                <th>Reproducciones</th>
                <th class="cell-actions">Acciones</th>
            </tr>
        </thead>
        <tbody>
        <%
        Dim tId, tName, tArtist, tAlbum, tPlays, tCover, tPath, tSearch
        Do While Not rsList.EOF
            tId = rsList("ID_SONG")
            tName = rsList("SONG_NAME") & ""
            tArtist = rsList("ARTISTS") & ""
            tAlbum = rsList("ALBUMS") & ""
            tPlays = rsList("PLAYS")
            tPath = rsList("SONG_PATH") & ""
            tCover = ImgUrl(rsList("COVER") & "", PageDepth & "assets/img/default.jpg")
            If tArtist = "" Then tArtist = "Sin artista"
            If tAlbum = "" Then tAlbum = "Sin álbum"
            tSearch = tId & " " & tName & " " & tArtist & " " & tAlbum
        %>
            <tr data-search-item="<%=H(tSearch)%>">
                <td class="muted" style="font-variant-numeric:tabular-nums">#<%=tId%></td>
                <td>
                    <div class="song-cell">
                        <div class="song-cover"><img src="<%=tCover%>" alt=""></div>
                        <div class="song-name">
                            <strong><%=H(tName)%></strong>
                            <span><%=H(tPath)%></span>
                        </div>
                    </div>
                </td>
                <td class="hide-sm"><span class="tag"><%=Icon("mic")%> <%=H(tArtist)%></span></td>
                <td class="hide-sm"><span class="tag tag--plain"><%=H(tAlbum)%></span></td>
                <td><span class="plays"><%=Icon("headphones")%> <%=tPlays%></span></td>
                <td class="cell-actions">
                    <div class="cell-actions__inner">
                        <a class="icon-btn icon-btn--sm" href="admin-songs.asp?edit=<%=tId%>" title="Editar" aria-label="Editar <%=H(tName)%>"><%=Icon("edit")%></a>
                        <form method="post" action="admin-songs.asp" style="display:inline"
                              data-confirm="¿Eliminar la canción «<%=H(tName)%>»? Se eliminarán también sus vínculos con artistas, álbumes y playlists.">
                            <input type="hidden" name="action" value="delete">
                            <input type="hidden" name="id" value="<%=tId%>">
                            <button type="submit" class="icon-btn icon-btn--sm icon-btn--danger" title="Eliminar" aria-label="Eliminar <%=H(tName)%>"><%=Icon("trash")%></button>
                        </form>
                    </div>
                </td>
            </tr>
        <%
            rsList.MoveNext
        Loop
        rsList.Close
        Set rsList = Nothing
        %>
        </tbody>
    </table>
</div>
<% End If %>

<div class="modal" id="songForm" <% If openModal Then %>data-auto-open="true"<% End If %> role="dialog" aria-modal="true" aria-label="Formulario de canción">
    <div class="modal__backdrop" data-close-modal></div>
    <div class="modal__card">
        <form method="post" action="admin-songs.asp">
            <div class="modal__head">
                <div>
                    <h3><% If isEdit Then %>Editar canción<% Else %>Nueva canción<% End If %></h3>
                    <p>Los campos marcados con * son obligatorios.</p>
                </div>
                <button type="button" class="icon-btn" data-close-modal aria-label="Cerrar"><%=Icon("close")%></button>
            </div>
            <div class="modal__body">
                <input type="hidden" name="action" value="save">
                <input type="hidden" name="id" value="<% If isEdit Then %><%=editId%><% End If %>">
                <div class="form-grid">
                    <div class="field field--full">
                        <label for="song_name">Nombre de la canción *</label>
                        <input type="text" id="song_name" name="song_name" maxlength="50" required
                               value="<%=H(formName)%>" placeholder="Ej. Adventure">
                    </div>
                    <div class="field field--full">
                        <label for="song_path">Ruta del archivo de audio</label>
                        <input type="text" id="song_path" name="song_path" maxlength="255"
                               value="<%=H(formPath)%>" placeholder="~/songFiles/bensound-adventure.mp3">
                        <span class="field__hint">El reproductor resuelve el archivo dentro de <code>/assets/audio</code>. Si no existe, usa la pista disponible por defecto.</span>
                    </div>
                    <div class="field">
                        <label for="plays">Reproducciones</label>
                        <input type="number" id="plays" name="plays" min="0" step="1" value="<%=H(formPlays)%>">
                    </div>
                    <div class="field">
                        <label>Resumen</label>
                        <span class="field__hint" style="padding-top:10px">
                            <% If isEdit Then %>Editando el registro #<%=editId%><% Else %>Se creará un nuevo registro en TBL_SONG<% End If %>
                        </span>
                    </div>
                    <div class="field field--full">
                        <label>Artistas asociados</label>
                        <% If rsArtists.EOF Then %>
                        <span class="field__hint">No hay artistas registrados.</span>
                        <% Else %>
                        <div class="check-panel">
                            <%
                            Dim arId, arName
                            Do While Not rsArtists.EOF
                                arId = rsArtists("ID_ARTIST")
                                arName = Trim((rsArtists("ARTIST_NAME") & "") & " " & (rsArtists("ARTIST_LASTNAME") & ""))
                            %>
                            <label class="check-pill">
                                <input type="checkbox" name="artists" value="<%=arId%>"<%=SelectedAttr(selArtists, arId)%>>
                                <span><%=H(arName)%></span>
                            </label>
                            <%
                                rsArtists.MoveNext
                            Loop
                            %>
                        </div>
                        <% End If %>
                    </div>
                    <div class="field field--full">
                        <label>Álbumes asociados</label>
                        <% If rsAlbums.EOF Then %>
                        <span class="field__hint">No hay álbumes registrados.</span>
                        <% Else %>
                        <div class="check-panel">
                            <%
                            Dim alId, alTitle
                            Do While Not rsAlbums.EOF
                                alId = rsAlbums("ID_ALBUM")
                                alTitle = rsAlbums("TITLE_ALBUM") & ""
                            %>
                            <label class="check-pill">
                                <input type="checkbox" name="albums" value="<%=alId%>"<%=SelectedAttr(selAlbums, alId)%>>
                                <span><%=H(alTitle)%></span>
                            </label>
                            <%
                                rsAlbums.MoveNext
                            Loop
                            %>
                        </div>
                        <% End If %>
                    </div>
                </div>
            </div>
            <div class="modal__foot">
                <span class="modal__foot-note">Tabla: TBL_SONG</span>
                <button type="button" class="btn btn--ghost" data-close-modal>Cancelar</button>
                <button type="submit" class="btn btn--primary"><%=Icon("check")%> Guardar canción</button>
            </div>
        </form>
    </div>
</div>

<!--#include file="includes/layout_bottom.asp"-->
