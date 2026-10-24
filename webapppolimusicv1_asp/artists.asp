<%@ Language=VBScript CodePage=65001 %>
<% Dim PageDepth : PageDepth = "" %>
<!--#include file="includes/db.asp"-->
<!--#include file="includes/auth.asp"-->
<!--#include file="includes/queries.asp"-->
<!--#include file="includes/tracklist.asp"-->
<%
RequireLogin
Dim PageTitle, ActiveNav
PageTitle = "Artistas"
ActiveNav = "artists"

Dim arRs, arCount
Set arRs = SqlRs("SELECT a.ID_ARTIST, a.ARTIST_NAME, a.ARTIST_LASTNAME, a.ARTIST_IMAGE, " & _
    "(SELECT COUNT(*) FROM TBL_PLAYER p WHERE p.ID_ARTIST = a.ID_ARTIST) AS SONG_COUNT " & _
    "FROM TBL_ARTIST a ORDER BY a.ARTIST_NAME, a.ARTIST_LASTNAME")
arCount = ScalarVal("SELECT COUNT(*) FROM TBL_ARTIST")
%>
<!--#include file="includes/layout_top.asp"-->

<div class="page-head">
    <div>
        <span class="page-head__eyebrow"><%=Icon("mic")%> Biblioteca</span>
        <h1>Artistas</h1>
        <p>Las voces que dan forma al catálogo. Abre un artista para escuchar todas las canciones asociadas a su nombre.</p>
    </div>
    <div class="page-head__actions">
        <% If IsAdmin() Then %>
        <a class="btn btn--primary" href="admin-artists.asp?new=1"><%=Icon("plus")%> Nuevo artista</a>
        <% End If %>
        <span class="chip chip--muted"><%=Pluralize(arCount, "artista", "artistas")%></span>
    </div>
</div>

<% If arCount = 0 Then %>
<div class="empty">
    <div class="empty__ico"><%=Icon("mic")%></div>
    <strong>Todavía no hay artistas registrados</strong>
    <p>El administrador puede crear artistas desde el panel de administración.</p>
</div>
<% Else %>
<div class="cards">
    <%
    Dim arId, arName, arLast, arFull, arPhoto, arSongs
    Do While Not arRs.EOF
        arId = arRs("ID_ARTIST")
        arName = arRs("ARTIST_NAME") & ""
        arLast = arRs("ARTIST_LASTNAME") & ""
        arFull = Trim(arName & " " & arLast)
        arPhoto = ImgUrl(arRs("ARTIST_IMAGE") & "", "")
        arSongs = arRs("SONG_COUNT")
    %>
    <button type="button" class="card" data-open-modal="artist-<%=arId%>"
            data-search-item="<%=H(arFull)%>">
        <div class="card__art card__art--round">
            <% If arPhoto <> "" Then %>
            <img src="<%=arPhoto%>" alt="Foto de <%=H(arFull)%>" loading="lazy">
            <% Else %>
            <div class="card__initials"><%=H(Initials(arFull))%></div>
            <% End If %>
            <span class="card__play" aria-hidden="true"><%=Icon("play")%></span>
        </div>
        <div class="card__title"><%=H(arFull)%></div>
        <div class="card__meta"><%=Pluralize(arSongs, "canción", "canciones")%></div>
    </button>
    <%
        arRs.MoveNext
    Loop
    %>
</div>
<p class="no-results" data-no-results>No se encontraron artistas con ese criterio.</p>
<% End If %>
<%
arRs.Close
Set arRs = Nothing

If arCount > 0 Then
    Dim mArRs, mArId, mArName, mArLast, mArFull, mArSongs, mArPhoto
    Set mArRs = SqlRs("SELECT a.ID_ARTIST, a.ARTIST_NAME, a.ARTIST_LASTNAME, a.ARTIST_IMAGE, " & _
        "(SELECT COUNT(*) FROM TBL_PLAYER p WHERE p.ID_ARTIST = a.ID_ARTIST) AS SONG_COUNT " & _
        "FROM TBL_ARTIST a ORDER BY a.ARTIST_NAME, a.ARTIST_LASTNAME")

    Do While Not mArRs.EOF
        mArId = mArRs("ID_ARTIST")
        mArName = mArRs("ARTIST_NAME") & ""
        mArLast = mArRs("ARTIST_LASTNAME") & ""
        mArFull = Trim(mArName & " " & mArLast)
        mArSongs = mArRs("SONG_COUNT")
        mArPhoto = ImgUrl(mArRs("ARTIST_IMAGE") & "", "")
        Dim mSongs
        Set mSongs = SongsByArtist(mArId)
%>
<div class="modal" id="artist-<%=mArId%>" role="dialog" aria-modal="true" aria-label="Artista <%=H(mArFull)%>">
    <div class="modal__backdrop" data-close-modal></div>
    <div class="modal__card modal__card--wide">
        <div class="modal__head">
            <div style="display:flex;gap:14px;align-items:center">
                <div class="song-cover" style="width:54px;height:54px;border-radius:50%">
                    <% If mArPhoto <> "" Then %>
                    <img src="<%=mArPhoto%>" alt="">
                    <% Else %>
                    <div class="card__initials" style="font-size:18px"><%=H(Initials(mArFull))%></div>
                    <% End If %>
                </div>
                <div>
                    <h3><%=H(mArFull)%></h3>
                    <p><%=Pluralize(mArSongs, "canción asociada", "canciones asociadas")%></p>
                </div>
            </div>
            <div style="display:flex;gap:8px;align-items:center">
                <button type="button" class="btn btn--primary btn--sm" data-play-all="#artist-tracks-<%=mArId%>">
                    <%=Icon("play")%> Reproducir
                </button>
                <button type="button" class="icon-btn" data-close-modal aria-label="Cerrar"><%=Icon("close")%></button>
            </div>
        </div>
        <div class="modal__body">
            <% If mSongs.EOF Then %>
            <div class="empty" style="padding:30px 18px">
                <div class="empty__ico"><%=Icon("music")%></div>
                <strong>Sin canciones asociadas</strong>
                <p>Este artista aún no tiene canciones vinculadas en el catálogo.</p>
            </div>
            <% Else %>
            <% RenderTracklist mSongs, "artist-tracks-" & mArId %>
            <% End If %>
        </div>
        <div class="modal__foot">
            <span class="modal__foot-note">Catálogo PoliMusic</span>
            <button type="button" class="btn btn--ghost" data-close-modal>Cerrar</button>
        </div>
    </div>
</div>
<%
        mSongs.Close
        Set mSongs = Nothing
        mArRs.MoveNext
    Loop
    mArRs.Close
    Set mArRs = Nothing
End If
%>

<!--#include file="includes/layout_bottom.asp"-->
