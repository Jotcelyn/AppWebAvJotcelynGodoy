<%@ Language=VBScript CodePage=65001 %>
<% Dim PageDepth : PageDepth = "" %>
<!--#include file="includes/db.asp"-->
<!--#include file="includes/auth.asp"-->
<%
RequireLogin

Dim PageTitle, ActiveNav, displayName, todayText
PageTitle = "Inicio"
ActiveNav = "home"

displayName = Session("username") & ""
displayName = Replace(displayName, ".", " ")
displayName = Replace(displayName, "_", " ")
todayText = WeekdayName(Weekday(Now()), False, vbSunday) & ", " & Day(Now()) & " de " & MonthName(Month(Now()), False) & " de " & Year(Now())

' ---------- Métricas ----------
Dim totalSongs, totalArtists, totalAlbums, totalPlays, totalPlaylists
totalSongs = ScalarVal("SELECT COUNT(*) FROM TBL_SONG")
totalArtists = ScalarVal("SELECT COUNT(*) FROM TBL_ARTIST")
totalAlbums = ScalarVal("SELECT COUNT(*) FROM TBL_ALBUM")
totalPlays = ScalarVal("SELECT ISNULL(SUM(PLAYS), 0) FROM TBL_SONG")
totalPlaylists = ScalarVal("SELECT COUNT(*) FROM TBL_PLAYLIST")

' ---------- Canciones con artista, álbum y portada ----------
Dim songSql, songRs
songSql = "SELECT s.ID_SONG, s.SONG_NAME, s.SONG_PATH, ISNULL(s.PLAYS, 0) AS PLAYS, " & _
    "(SELECT TOP 1 al.ALBUM_IMAGE_PATH FROM TBL_TRACKLIST t JOIN TBL_ALBUM al ON al.ID_ALBUM = t.ID_ALBUM " & _
    " WHERE t.ID_SONG = s.ID_SONG ORDER BY al.ID_ALBUM) AS COVER, " & _
    "(SELECT STRING_AGG(a.ARTIST_NAME + ' ' + ISNULL(a.ARTIST_LASTNAME, ''), ', ') FROM TBL_PLAYER p " & _
    " JOIN TBL_ARTIST a ON a.ID_ARTIST = p.ID_ARTIST WHERE p.ID_SONG = s.ID_SONG) AS ARTISTS, " & _
    "(SELECT STRING_AGG(al.TITLE_ALBUM, ', ') FROM TBL_TRACKLIST t JOIN TBL_ALBUM al ON al.ID_ALBUM = t.ID_ALBUM " & _
    " WHERE t.ID_SONG = s.ID_SONG) AS ALBUMS " & _
    "FROM TBL_SONG s ORDER BY s.ID_SONG"
Set songRs = SqlRs(songSql)

' ---------- Playlists destacadas ----------
Dim plCmd, plRs
If IsAdmin() Then
    Set plCmd = NewCmd("SELECT TOP 4 p.ID_PLAYLIST, p.TITLE, p.PHOTO, p.TYPE, p.CREATION_DATE, " & _
        "(SELECT COUNT(*) FROM TBL_PLAYLIST_SONG ps WHERE ps.ID_PLAYLIST = p.ID_PLAYLIST) AS SONG_COUNT " & _
        "FROM TBL_PLAYLIST p ORDER BY p.ID_PLAYLIST")
Else
    Set plCmd = NewCmd("SELECT TOP 4 p.ID_PLAYLIST, p.TITLE, p.PHOTO, p.TYPE, p.CREATION_DATE, " & _
        "(SELECT COUNT(*) FROM TBL_PLAYLIST_SONG ps WHERE ps.ID_PLAYLIST = p.ID_PLAYLIST) AS SONG_COUNT " & _
        "FROM TBL_PLAYLIST p WHERE (p.TYPE = 0 OR p.ID_USER = ?) ORDER BY p.ID_PLAYLIST")
    AddInt plCmd, "@user", Session("user_id")
End If
Set plRs = ExecRs(plCmd)
%>
<!--#include file="includes/layout_top.asp"-->

<section class="hero">
    <div>
        <span class="hero__eyebrow"><%=Icon("sparkles")%> <%=RoleLabel()%> · <%=APP_NAME%></span>
        <h1>Bienvenido de nuevo, <em><%=H(displayName)%></em></h1>
        <p class="hero__lead">
            Tu sesión está activa como <strong><%=H(Session("username"))%></strong> con el rol de
            <strong><%=RoleLabel()%></strong>. Explora el catálogo, reproduce tus canciones favoritas
            y descubre las playlists que forman parte de tu biblioteca.
        </p>
        <div class="hero__actions">
            <button type="button" class="btn btn--primary" data-play-all="#songTable">
                <%=Icon("play")%> Reproducir todo
            </button>
            <a class="btn btn--ghost" href="playlists.asp">
                <%=Icon("playlist")%> Explorar playlists
            </a>
        </div>
        <div class="hero__meta">
            <span class="chip chip--muted"><%=Icon("calendar")%> <%=H(todayText)%></span>
            <span class="chip chip--muted"><%=Icon("music")%> <%=totalSongs%> canciones</span>
            <span class="chip chip--muted"><%=Icon("clock")%> <%=totalPlays%> reproducciones</span>
        </div>
    </div>
    <div class="hero__art" aria-hidden="true">
        <div class="hero__art-item"><img src="assets/img/album2.png" alt=""></div>
        <div class="hero__art-item"><img src="assets/img/album4.png" alt=""></div>
        <div class="hero__art-item"><img src="assets/img/album1.png" alt=""></div>
    </div>
</section>

<section class="stats" aria-label="Resumen de la biblioteca">
    <div class="stat">
        <span class="stat__label">Canciones</span>
        <span class="stat__value"><%=totalSongs%></span>
        <span class="stat__hint">En el catálogo</span>
        <span class="stat__ico"><%=Icon("music")%></span>
    </div>
    <div class="stat">
        <span class="stat__label">Artistas</span>
        <span class="stat__value"><%=totalArtists%></span>
        <span class="stat__hint">Voces registradas</span>
        <span class="stat__ico"><%=Icon("mic")%></span>
    </div>
    <div class="stat">
        <span class="stat__label">Álbumes</span>
        <span class="stat__value"><%=totalAlbums%></span>
        <span class="stat__hint">Colecciones activas</span>
        <span class="stat__ico"><%=Icon("disc")%></span>
    </div>
    <div class="stat">
        <span class="stat__label">Reproducciones</span>
        <span class="stat__value"><%=totalPlays%></span>
        <span class="stat__hint"><%=totalPlaylists%> playlists en total</span>
        <span class="stat__ico"><%=Icon("headphones")%></span>
    </div>
</section>

<div class="section-head" id="song-list">
    <div>
        <h2><%=Icon("music")%> Song List</h2>
        <p>
            A continuación se presenta el listado de canciones (Song List) disponible en tu biblioteca.
            Pulsa el botón de reproducción de cualquier fila para escucharla en el reproductor inferior.
            <% If IsAdmin() Then %>
            Como administrador también puedes editar o eliminar cada registro desde la columna de acciones.
            <% End If %>
        </p>
    </div>
    <div class="section-head__actions">
        <button type="button" class="btn btn--soft btn--sm" data-play-all="#songTable">
            <%=Icon("play")%> Reproducir lista
        </button>
        <span class="chip chip--muted"><%=Pluralize(totalSongs, "canción", "canciones")%></span>
    </div>
</div>

<div class="table-wrap">
    <table class="song-table" id="songTable">
        <thead>
            <tr>
                <th class="cell-rank">#</th>
                <th>Canción</th>
                <th class="hide-sm">Artista</th>
                <th class="hide-sm">Álbum</th>
                <th class="hide-sm">Reproducciones</th>
                <th class="cell-actions">Acciones</th>
            </tr>
        </thead>
        <tbody>
        <%
        Dim rowNum, songId, songName, artistName, albumName, coverUrl, audioSrc, playCount, searchBlob, likeId
        rowNum = 0
        Do While Not songRs.EOF
            rowNum = rowNum + 1
            songId = songRs("ID_SONG")
            songName = songRs("SONG_NAME") & ""
            artistName = songRs("ARTISTS") & ""
            albumName = songRs("ALBUMS") & ""
            playCount = songRs("PLAYS")
            If artistName = "" Then artistName = "Artista desconocido"
            If albumName = "" Then albumName = "Sencillo"
            coverUrl = ImgUrl(songRs("COVER") & "", PageDepth & "assets/img/default.jpg")
            audioSrc = AudioUrl(songRs("SONG_PATH") & "")
            searchBlob = songName & " " & artistName & " " & albumName
        %>
            <tr class="song-row"
                data-song-id="<%=songId%>"
                data-title="<%=H(songName)%>"
                data-artist="<%=H(artistName)%>"
                data-src="<%=audioSrc%>"
                data-cover="<%=coverUrl%>"
                data-plays="<%=playCount%>"
                data-search-item="<%=H(searchBlob)%>">
                <td class="cell-rank">
                    <span class="rank-num"><%=rowNum%></span>
                    <button type="button" class="row-play" data-play-id="<%=songId%>" aria-label="Reproducir <%=H(songName)%>">
                        <span class="ic-play"><%=Icon("play")%></span>
                        <span class="ic-pause"><%=Icon("pause")%></span>
                    </button>
                </td>
                <td>
                    <div class="song-cell">
                        <div class="song-cover">
                            <img src="<%=coverUrl%>" alt="Portada de <%=H(songName)%>" loading="lazy">
                            <span class="song-cover__eq" aria-hidden="true"><i></i><i></i><i></i><i></i></span>
                        </div>
                        <div class="song-name">
                            <strong><%=H(songName)%></strong>
                            <span><%=H(albumName)%></span>
                        </div>
                    </div>
                </td>
                <td class="hide-sm"><span class="tag"><%=Icon("mic")%> <%=H(artistName)%></span></td>
                <td class="hide-sm"><span class="tag tag--plain"><%=H(albumName)%></span></td>
                <td class="hide-sm">
                    <span class="plays"><%=Icon("headphones")%> <span data-plays-for="<%=songId%>"><%=playCount%></span></span>
                </td>
                <td class="cell-actions">
                    <div class="cell-actions__inner">
                        <button type="button" class="icon-btn icon-btn--sm js-like-row" data-like-id="<%=songId%>"
                                aria-label="Me gusta <%=H(songName)%>" title="Me gusta">
                            <%=Icon("heart")%>
                        </button>
                        <% If IsAdmin() Then %>
                        <a class="icon-btn icon-btn--sm" href="admin-songs.asp?edit=<%=songId%>" title="Editar" aria-label="Editar <%=H(songName)%>">
                            <%=Icon("edit")%>
                        </a>
                        <form method="post" action="admin-songs.asp" style="display:inline"
                              data-confirm="¿Eliminar la canción «<%=H(songName)%>»? Esta acción también quitará sus relaciones con artistas, álbumes y playlists.">
                            <input type="hidden" name="action" value="delete">
                            <input type="hidden" name="id" value="<%=songId%>">
                            <button type="submit" class="icon-btn icon-btn--sm icon-btn--danger" title="Eliminar" aria-label="Eliminar <%=H(songName)%>">
                                <%=Icon("trash")%>
                            </button>
                        </form>
                        <% End If %>
                    </div>
                </td>
            </tr>
        <%
            songRs.MoveNext
        Loop
        songRs.Close
        Set songRs = Nothing
        %>
        </tbody>
    </table>
</div>
<p class="no-results" data-no-results>No se encontraron canciones que coincidan con tu búsqueda.</p>

<div class="section-head">
    <div>
        <h2><%=Icon("playlist")%> Playlists para ti</h2>
        <p>Una selección curada de listas que puedes reproducir en cualquier momento.</p>
    </div>
    <div class="section-head__actions">
        <a class="btn btn--ghost btn--sm" href="playlists.asp">Ver todas <%=Icon("arrow-right")%></a>
    </div>
</div>

<%
If plRs.EOF Then
%>
<div class="empty">
    <div class="empty__ico"><%=Icon("playlist")%></div>
    <strong>Aún no hay playlists disponibles</strong>
    <p>Cuando existan playlists públicas o propias, aparecerán en esta sección.</p>
</div>
<%
Else
%>
<div class="cards">
    <%
    Dim plId, plTitle, plPhoto, plType, plCount
    Do While Not plRs.EOF
        plId = plRs("ID_PLAYLIST")
        plTitle = plRs("TITLE") & ""
        plPhoto = ImgUrl(plRs("PHOTO") & "", PageDepth & "assets/img/default.jpg")
        plType = plRs("TYPE")
        plCount = plRs("SONG_COUNT")
    %>
    <a class="card" href="playlists.asp#playlist-<%=plId%>">
        <div class="card__art">
            <img src="<%=plPhoto%>" alt="Portada de <%=H(plTitle)%>" loading="lazy">
            <span class="card__badge"><% If CStr(plType) = "1" Then %>Privada<% Else %>Pública<% End If %></span>
        </div>
        <div class="card__title"><%=H(plTitle)%></div>
        <div class="card__meta"><%=Pluralize(plCount, "canción", "canciones")%></div>
    </a>
    <%
        plRs.MoveNext
    Loop
    plRs.Close
    Set plRs = Nothing
    %>
</div>
<%
End If
%>

<!--#include file="includes/layout_bottom.asp"-->
