<%@ Language=VBScript CodePage=65001 %>
<% Dim PageDepth : PageDepth = "" %>
<!--#include file="includes/db.asp"-->
<!--#include file="includes/auth.asp"-->
<!--#include file="includes/queries.asp"-->
<!--#include file="includes/tracklist.asp"-->
<%
RequireLogin
Dim PageTitle, ActiveNav
PageTitle = "Playlists"
ActiveNav = "playlists"

Dim plSql, plCmd, plRs, totalVisible
plSql = "SELECT p.ID_PLAYLIST, p.TITLE, p.PHOTO, p.TYPE, p.CREATION_DATE, p.ID_USER, " & _
    "(SELECT COUNT(*) FROM TBL_PLAYLIST_SONG ps WHERE ps.ID_PLAYLIST = p.ID_PLAYLIST) AS SONG_COUNT, " & _
    "(SELECT u.USERNAME FROM TBL_USER u WHERE u.ID_USER = p.ID_USER) AS OWNER " & _
    "FROM TBL_PLAYLIST p "

Set plCmd = NewCmd("")
If IsAdmin() Then
    plCmd.CommandText = plSql & "ORDER BY p.ID_PLAYLIST"
Else
    plCmd.CommandText = plSql & "WHERE (p.TYPE = 0 OR p.ID_USER = ?) ORDER BY p.ID_PLAYLIST"
    AddInt plCmd, "@user", Session("user_id")
End If
Set plRs = ExecRs(plCmd)
If IsAdmin() Then
    totalVisible = ScalarVal("SELECT COUNT(*) FROM TBL_PLAYLIST")
Else
    totalVisible = ScalarVal("SELECT COUNT(*) FROM TBL_PLAYLIST WHERE (TYPE = 0 OR ID_USER = " & CLng(Session("user_id")) & ")")
End If
%>
<!--#include file="includes/layout_top.asp"-->

<div class="page-head">
    <div>
        <span class="page-head__eyebrow"><%=Icon("playlist")%> Biblioteca</span>
        <h1>Playlists</h1>
        <p>Colecciones creadas por los usuarios. Abre una playlist para ver su contenido completo y reproducirla de principio a fin.</p>
    </div>
    <div class="page-head__actions">
        <% If IsAdmin() Then %>
        <a class="btn btn--primary" href="admin-playlists.asp?new=1"><%=Icon("plus")%> Nueva playlist</a>
        <% End If %>
        <span class="chip chip--muted"><%=Pluralize(totalVisible, "playlist", "playlists")%></span>
    </div>
</div>

<% If totalVisible = 0 Then %>
<div class="empty">
    <div class="empty__ico"><%=Icon("playlist")%></div>
    <strong>No hay playlists para mostrar</strong>
    <p>Aún no existen playlists públicas ni propias asociadas a tu cuenta.</p>
</div>
<% Else %>
<div class="cards">
    <%
    Dim plId, plTitle, plPhoto, plType, plCount, plOwner, plTypeText
    Do While Not plRs.EOF
        plId = plRs("ID_PLAYLIST")
        plTitle = plRs("TITLE") & ""
        plPhoto = ImgUrl(plRs("PHOTO") & "", PageDepth & "assets/img/default.jpg")
        plType = CStr(plRs("TYPE"))
        plCount = plRs("SONG_COUNT")
        plOwner = plRs("OWNER") & ""
        If plOwner = "" Then plOwner = "Usuario"
        If plType = "1" Then plTypeText = "Privada" Else plTypeText = "Pública"
    %>
    <button type="button" class="card" data-open-modal="playlist-<%=plId%>"
            data-search-item="<%=H(plTitle & " " & plOwner & " " & plTypeText)%>">
        <div class="card__art">
            <img src="<%=plPhoto%>" alt="Portada de <%=H(plTitle)%>" loading="lazy">
            <span class="card__badge"><%=plTypeText%></span>
            <span class="card__play" aria-hidden="true"><%=Icon("play")%></span>
        </div>
        <div class="card__title"><%=H(plTitle)%></div>
        <div class="card__meta"><%=H(plOwner)%> · <%=Pluralize(plCount, "canción", "canciones")%></div>
    </button>
    <%
        plRs.MoveNext
    Loop
    %>
</div>
<p class="no-results" data-no-results>No se encontraron playlists con ese criterio.</p>
<%
End If

plRs.Close
Set plRs = Nothing
Set plCmd = Nothing

' ---------- Modales de detalle ----------
If totalVisible > 0 Then
    Dim mCmd, mRs, mPlId, mTitle, mOwner, mType, mCount, mSongs
    Set mCmd = NewCmd("")
    If IsAdmin() Then
        mCmd.CommandText = plSql & "ORDER BY p.ID_PLAYLIST"
    Else
        mCmd.CommandText = plSql & "WHERE (p.TYPE = 0 OR p.ID_USER = ?) ORDER BY p.ID_PLAYLIST"
        AddInt mCmd, "@user", Session("user_id")
    End If
    Set mRs = ExecRs(mCmd)

    Do While Not mRs.EOF
        mPlId = mRs("ID_PLAYLIST")
        mTitle = mRs("TITLE") & ""
        mOwner = mRs("OWNER") & ""
        mType = CStr(mRs("TYPE"))
        mCount = mRs("SONG_COUNT")
        Set mSongs = SongsByPlaylist(mPlId)
%>
<div class="modal" id="playlist-<%=mPlId%>" role="dialog" aria-modal="true" aria-label="Playlist <%=H(mTitle)%>">
    <div class="modal__backdrop" data-close-modal></div>
    <div class="modal__card modal__card--wide">
        <div class="modal__head">
            <div>
                <h3><%=H(mTitle)%></h3>
                <p><%=H(mOwner)%> · <%=Pluralize(mCount, "canción", "canciones")%> · <% If mType = "1" Then %>Playlist privada<% Else %>Playlist pública<% End If %></p>
            </div>
            <div style="display:flex;gap:8px;align-items:center">
                <button type="button" class="btn btn--primary btn--sm" data-play-all="#tracks-<%=mPlId%>">
                    <%=Icon("play")%> Reproducir
                </button>
                <button type="button" class="icon-btn" data-close-modal aria-label="Cerrar"><%=Icon("close")%></button>
            </div>
        </div>
        <div class="modal__body">
            <% If mSongs.EOF Then %>
            <div class="empty" style="padding:30px 18px">
                <div class="empty__ico"><%=Icon("music")%></div>
                <strong>Esta playlist está vacía</strong>
                <p>El administrador todavía no ha agregado canciones a esta colección.</p>
            </div>
            <% Else %>
            <% RenderTracklist mSongs, "tracks-" & mPlId %>
            <% End If %>
        </div>
        <div class="modal__foot">
            <span class="modal__foot-note">Creada el <%=PrettyDate(mRs("CREATION_DATE"))%></span>
            <button type="button" class="btn btn--ghost" data-close-modal>Cerrar</button>
        </div>
    </div>
</div>
<%
        mSongs.Close
        Set mSongs = Nothing
        mRs.MoveNext
    Loop
    mRs.Close
    Set mRs = Nothing
    Set mCmd = Nothing
End If
%>

<!--#include file="includes/layout_bottom.asp"-->
